"""渲染引擎：克隆模板页 → 原位替换 → 原子写出。

核心约束（见 docs/ARCHITECTURE.md INV-1..INV-6）：
- 输出页必须来自模板页 XML 深拷贝；
- 只写 ``<a:t>`` 文本（及显式开启的图片/表格结构）；
- 克隆时重映射所有关系 id；
- 一律原子写。
"""

from __future__ import annotations

import copy
import datetime as _dt
import itertools
import os
import re
import shutil
import tempfile
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, PartFactory
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

from . import text_fit
from .excel_images import ExcelImageError, extract_cell_images
from .excel_reader import ExcelDataError, SheetData, read_excel
from .model import Binding, Field, Profile
from .template_scanner import _safe_int, scan_template


class RenderError(Exception):
    """渲染失败（可读中文信息）。"""


@dataclass
class RenderResult:
    output: str
    rows: int
    slides: int
    replaced: int = 0
    warnings: list[str] = field(default_factory=list)


_RID_ATTRS = (qn("r:id"), qn("r:embed"), qn("r:link"))


def _is_textual(part: Part) -> bool:
    ct = (part.content_type or "").lower()
    return "xml" in ct or "text" in ct


def _unique_partname(package, src_part: Part, used: set[str]) -> PackURI:
    p = src_part.partname
    stem, ext = os.path.splitext(p.filename)
    m = re.match(r"([A-Za-z_]+)", stem)
    prefix = m.group(1) if m else "part"
    existing = {str(pp.partname) for pp in package.iter_parts()}
    n = 1
    while True:
        candidate = f"{p.baseURI}/{prefix}{n}{ext}"
        if candidate not in existing and candidate not in used:
            used.add(candidate)
            return PackURI(candidate)
        n += 1


def _clone_part_tree(package, src_part: Part, cache: dict, used: set[str]):
    """递归复制 part 及其关系树，返回 ``(新 part, {旧rId: 新rId})``。

    用 ``PartFactory`` 按 content type 构造正确的子类（例如 ChartPart），
    否则克隆后的图表会变成裸 ``Part``，``shape.chart`` 无法访问。
    """
    key = id(src_part)
    if key in cache:
        return cache[key]

    new_part = PartFactory(
        _unique_partname(package, src_part, used),
        src_part.content_type,
        package,
        src_part.blob,
    )
    cache[key] = (new_part, {})

    rid_map: dict[str, str] = {}
    for old_rid, rel in list(src_part.rels.items()):
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE, RT.SLIDE):
            # layout 复用；备注页不随页复制（其反向 slide 引用会拖入模板页副本）
            continue
        if rel.is_external:
            new_rid = new_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            child, _ = _clone_part_tree(package, rel.target_part, cache, used)
            new_rid = new_part.relate_to(child, rel.reltype)
        rid_map[old_rid] = new_rid

    element = getattr(new_part, "_element", None)
    if element is not None and rid_map:
        _remap_rids(element, rid_map)
    elif _is_textual(src_part) and rid_map:
        text = src_part.blob.decode("utf-8")
        for old, new in rid_map.items():
            text = re.sub(
                rf'(r:(?:id|embed|link)="){re.escape(old)}(")',
                rf"\g<1>{new}\g<2>",
                text,
            )
        new_part._blob = text.encode("utf-8")  # noqa: SLF001 - 官方无公开 API
    return new_part, rid_map


def _remap_rids(root, rid_map: dict[str, str]) -> None:
    for el in root.iter():
        for attr in _RID_ATTRS:
            value = el.get(attr)
            if value is not None and value in rid_map:
                el.set(attr, rid_map[value])


def clone_slide(prs: Presentation, src_slide):
    """克隆一页：复用同一 layout，深拷贝关系树并重映射 rId。"""
    package = prs.part.package
    dest = prs.slides.add_slide(src_slide.slide_layout)
    dest_part = dest.part
    # add_slide 只按页数取下一个编号，模板 partname 有空洞时会与已有页重名
    dest_part.partname = package.next_partname("/ppt/slides/slide%d.xml")
    root = parse_xml(src_slide.part.blob)

    rid_map: dict[str, str] = {}
    cache: dict = {}
    used: set[str] = set()
    for old_rid, rel in list(src_slide.part.rels.items()):
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE, RT.SLIDE):
            # layout 复用；备注页不复制（其反向 slide 引用会把模板页整页拖进包里）
            continue
        if rel.is_external:
            new_rid = dest_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            child, _ = _clone_part_tree(package, rel.target_part, cache, used)
            new_rid = dest_part.relate_to(child, rel.reltype)
        rid_map[old_rid] = new_rid

    _remap_rids(root, rid_map)
    dest_part._element = root  # noqa: SLF001
    dest_part.__dict__.pop("slide", None)  # lazyproperty 缓存必须失效
    return dest_part.slide


def _drop_slide(prs: Presentation, index: int) -> None:
    sld_id_lst = prs.part._element.get_or_add_sldIdLst()  # noqa: SLF001
    sld_id = sld_id_lst[index]
    r_id = sld_id.get(qn("r:id"))
    if r_id:
        prs.part.drop_rel(r_id)
    sld_id_lst.remove(sld_id)


def _iter_shapes(shapes) -> Iterable:
    for shape in shapes:
        yield shape
        if hasattr(shape, "shapes"):
            yield from _iter_shapes(shape.shapes)


def _find_shape(slide, field_id: str):
    m = re.match(r"^(?:sp|tbl|pic|chart|shape)-(\d+)", field_id)
    if not m:
        return None
    target = int(m.group(1))
    for shape in _iter_shapes(slide.shapes):
        if _safe_int(shape.shape_id) == target:
            return shape
    return None


def _page_var(name: str, page_index: int, total: int) -> str:
    """页码类变量：``page`` / ``total`` / ``date``。"""
    if name == "page":
        return str(page_index + 1)
    if name == "total":
        return str(total)
    if name == "date":
        return _dt.date.today().isoformat()
    return ""


def _to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, _dt.datetime):
        return value.strftime("%Y-%m-%d") if value.time() == _dt.time() else value.strftime(
            "%Y-%m-%d %H:%M"
        )
    if isinstance(value, _dt.date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _set_paragraph_text(paragraph, value: str) -> None:
    """整段替换：写进字号最大的 run，清空其余 run 的文本。

    模板里常见「14pt 缩进 run + 16pt 正文 run」，写第一个 run 会把正文带成 14pt；
    这里优先选字号最大的 run 作为载体（字号继承时视为 0，仍取第一个）。
    """
    runs = list(paragraph.runs)
    if not runs:
        paragraph.add_run().text = value
        return

    def _size(run) -> float:
        size = run.font.size
        return size.pt if size is not None else 0.0

    target = max(runs, key=_size)
    target.text = value
    for run in runs:
        if run is not target:
            run.text = ""


def _apply_text(
    shape, field: Field, value: str, options: dict[str, Any], slot_name: str | None = None
) -> int:
    """按槽位替换文本。返回替换点数量。

    - ``explicit`` / ``label``：run 内子串替换（保留标签格式）
    - ``paragraph``：整段替换，其它段落不动（无占位符模板的主力）
    - 无槽位：整段替换并清空其余段落
    """
    count = 0
    tf = shape.text_frame
    slots = field.slots
    if slot_name:
        slots = [s for s in slots if s.name == slot_name]
    if not slots:
        if not field.slots:
            _set_paragraph_text(tf.paragraphs[0], value)
            for paragraph in tf.paragraphs[1:]:
                for run in paragraph.runs:
                    run.text = ""
            return 1
        return 0

    for slot in slots:
        if slot.para >= len(tf.paragraphs):
            continue
        paragraph = tf.paragraphs[slot.para]
        if slot.kind in ("named", "alt", "paragraph"):
            _set_paragraph_text(paragraph, value)
            count += 1
            continue
        if slot.run >= len(paragraph.runs):
            continue
        run = paragraph.runs[slot.run]
        if slot.kind == "label":
            label = slot.name
            run.text = f"{label}：{value}"
        elif slot.raw in run.text:
            run.text = run.text.replace(slot.raw, value)
        else:
            continue
        count += 1
    return count


def _table_element(shape):
    return shape._element.graphic.graphicData.tbl  # noqa: SLF001


def _set_cell_text(cell, value: str) -> None:
    tf = cell.text_frame
    paragraph = tf.paragraphs[0]
    if paragraph.runs:
        paragraph.runs[0].text = value
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run().text = value


def _format_cell(column_name: str, value: Any) -> str:
    """单元格格式化：列名含「率」且值为 0~1 的小数时转百分比。"""
    if isinstance(value, float):
        name = (column_name or "").lower()
        if 0 <= value <= 1 and ("率" in name or "percent" in name):
            return f"{value * 100:.1f}%"
    return _to_text(value)


def _apply_table(
    shape, rows: list[dict[str, Any]], binding: Binding, warnings: list[str] | None = None
) -> int:
    table = shape.table
    columns = list(binding.source.get("columns") or [])
    if not columns:
        raise RenderError(f"表格 {binding.field_id} 未指定 source.columns。")
    if len(columns) > len(table.columns):
        raise RenderError(
            f"表格 {binding.field_id} 只有 {len(table.columns)} 列，映射给了 {len(columns)} 列。"
        )

    header = bool(binding.options.get("header", True))
    header_rows = 1 if header else 0
    tbl = _table_element(shape)
    trs = list(tbl.tr_lst)

    if not trs:
        raise RenderError(f"表格 {binding.field_id} 没有任何行。")

    template_rows = list(table.rows)
    header_height = (template_rows[0].height or 0) if header else 0
    row_height = template_rows[-1].height or 0
    template_row = copy.deepcopy(trs[-1])
    while len(tbl.tr_lst) > header_rows:
        tbl.remove(tbl.tr_lst[-1])

    if header and binding.options.get("header_names"):
        for c_index, name in enumerate(binding.options["header_names"]):
            if c_index < len(table.columns):
                _set_cell_text(table.cell(0, c_index), str(name))
    # 未绑定的表头列也要清空，否则会残留模板里的示例标题
    if header:
        for c_index in range(len(columns), len(table.columns)):
            _set_cell_text(table.cell(0, c_index), "")

    for row_data in rows:
        tr = copy.deepcopy(template_row)
        tbl.append(tr)
        r_index = len(tbl.tr_lst) - 1
        for c_index, col in enumerate(columns):
            _set_cell_text(
                table.cell(r_index, c_index), _format_cell(col, row_data.get(col))
            )
        # 模板列多于绑定列时，多出来的列必须清空，不能留模板样例文字
        for c_index in range(len(columns), len(table.columns)):
            _set_cell_text(table.cell(r_index, c_index), "")

    if warnings is not None and row_height:
        available = shape.height or 0
        needed = header_height + len(rows) * row_height
        if available and needed > available:
            max_rows = max(1, int((available - header_height) / row_height))
            msg = (
                f"表格超出高度：{len(rows)} 行需要 {needed / 360000:.1f}cm，"
                f"框高 {available / 360000:.1f}cm；建议每页 ≤ {max_rows} 行"
            )
            if msg not in warnings:
                warnings.append(msg)
    return len(rows) * len(columns)


def _apply_chart(shape, rows: list[dict[str, Any]], binding: Binding) -> int:
    """替换图表数据：保留图表类型、配色、坐标轴，只换 series 与类目。"""
    chart = shape.chart
    categories = binding.source.get("categories")
    series_cfg = binding.source.get("series") or []
    if not categories:
        raise RenderError(f"图表 {binding.field_id} 未指定类目列（source.categories）。")
    if not series_cfg:
        raise RenderError(f"图表 {binding.field_id} 未指定数据系列（source.series）。")

    data = CategoryChartData()
    data.categories = [_to_text(row.get(categories)) for row in rows]
    for cfg in series_cfg:
        value_col = cfg.get("value")
        if not value_col:
            continue
        values: list[float | None] = []
        for row in rows:
            raw = row.get(value_col)
            if raw in (None, ""):
                values.append(None)
                continue
            try:
                values.append(float(raw))
            except (TypeError, ValueError):
                values.append(None)
        data.add_series(cfg.get("name") or value_col, values)
    chart.replace_data(data)
    return len(rows) * max(1, len(series_cfg))


def _image_size(path: str) -> tuple[int, int]:
    with Image.open(path) as im:
        return im.size


def _set_src_rect(blip_fill, left: float, top: float, right: float, bottom: float) -> None:
    # 必须用顺序感知 API：CT_BlipFillProperties 规定 blip → srcRect → stretch/tile
    src_rect = blip_fill.get_or_add_srcRect()
    for name, value in (("l", left), ("t", top), ("r", right), ("b", bottom)):
        src_rect.set(name, str(int(value * 100000)))


def _apply_image_on_slide(slide, shape, path: str, options: dict[str, Any]) -> None:
    """替换图片字节并处理 contain / cover，保持裁剪与边框。"""
    image_part, r_id = slide.part.get_or_add_image_part(path)
    blip_fill = shape._element.blipFill  # noqa: SLF001
    blip = blip_fill.find(qn("a:blip"))
    if blip is None:
        raise RenderError("图片元素缺少 a:blip，无法替换。")
    old_rid = blip.get(qn("r:embed"))
    blip.set(qn("r:embed"), r_id)
    if old_rid and old_rid != r_id:
        # 模板原图的 part 已不再被本页引用，删除关系避免每页残留一份
        slide.part.drop_rel(old_rid)

    fit = options.get("fit", "contain")
    if fit not in ("contain", "cover", "stretch"):
        raise RenderError(f"不支持的图片填充方式：{fit}（可选 contain / cover / stretch）。")
    if fit == "stretch":
        _set_src_rect(blip_fill, 0, 0, 0, 0)
        return

    img_w, img_h = _image_size(path)
    if img_w <= 0 or img_h <= 0:
        return
    img_ratio = img_w / img_h
    box_w = _safe_int(shape.width)
    box_h = _safe_int(shape.height)
    if box_w <= 0 or box_h <= 0:
        return
    box_ratio = box_w / box_h

    if fit == "cover":
        if img_ratio > box_ratio:
            crop = (1 - box_ratio / img_ratio) / 2
            _set_src_rect(blip_fill, crop, 0, crop, 0)
        else:
            crop = (1 - img_ratio / box_ratio) / 2
            _set_src_rect(blip_fill, 0, crop, 0, crop)
        return

    # contain：等比缩小到框内并居中（框内留白，不变形）
    if img_ratio >= box_ratio:
        new_w = box_w
        new_h = int(round(box_w / img_ratio))
    else:
        new_h = box_h
        new_w = int(round(box_h * img_ratio))
    left = _safe_int(shape.left) + (box_w - new_w) // 2
    top = _safe_int(shape.top) + (box_h - new_h) // 2
    shape.left = left
    shape.top = top
    shape.width = new_w
    shape.height = new_h
    _set_src_rect(blip_fill, 0, 0, 0, 0)


def _apply_bindings(
    slide,
    fields: dict[str, Field],
    bindings: list[Binding],
    page_rows: list[dict[str, Any]],
    *,
    image_map: dict[str, str],
    excel_dir: str,
    warnings: list[str],
    page_index: int,
    total_pages: int,
    warned_overflow: set[str],
) -> int:
    """把一组 binding 应用到已克隆的页上，返回替换点数量。"""
    replaced = 0
    for binding in bindings:
        field = fields.get(binding.field_id)
        if field is None:
            continue
        shape = _find_shape(slide, binding.field_id)
        if shape is None:
            warnings.append(f"克隆后找不到元素：{binding.field_id}")
            continue
        if binding.kind == "text":
            if not field.slots and field.kind == "text":
                msg = f"元素 {binding.field_id} 没有可识别槽位，按整段替换（可能丢失多段格式）"
                if msg not in warnings:
                    warnings.append(msg)
            column = binding.source.get("column")
            row0 = page_rows[0]
            if binding.source.get("var"):
                value = _page_var(binding.source["var"], page_index, total_pages)
            elif binding.source.get("value_map"):
                key = _to_text(row0.get(binding.source.get("key_column", "序号")))
                mapped = (binding.source["value_map"] or {}).get(key)
                value = _to_text(mapped) if mapped else _to_text(row0.get(column))
            else:
                value = _to_text(row0.get(column)) if column else ""
            fit = text_fit.estimate(field, value)
            if fit is not None and fit.overflow and binding.field_id not in warned_overflow:
                warned_overflow.add(binding.field_id)
                warnings.append(
                    f"文本可能溢出：{binding.field_id} 约 {fit.lines} 行，"
                    f"需要 {fit.needed_emu / 12700:.0f}pt 高，框高 "
                    f"{fit.available_emu / 12700:.0f}pt（按「只告警不改」处理）"
                )
            replaced += _apply_text(
                shape, field, value, binding.options, binding.source.get("slot")
            )
        elif binding.kind == "table":
            replaced += _apply_table(shape, page_rows, binding, warnings)
        elif binding.kind == "chart":
            if len(page_rows) == 1 and not any("图表数据只有 1 个点" in w for w in warnings):
                warnings.append(
                    "图表数据只有 1 个点：图表通常配合「每 N 行一页」使用，"
                    "可在映射页调整分页模式。"
                )
            replaced += _apply_chart(shape, page_rows, binding)
        elif binding.kind == "image":
            column = binding.source.get("column")
            raw = page_rows[0].get(column) if column else None
            path = _resolve_one_image(raw, image_map, excel_dir)
            if path:
                _apply_image_on_slide(slide, shape, path, binding.options)
                replaced += 1
            elif raw not in (None, ""):
                warnings.append(
                    f"图片无法解析（文件不存在，或路径不在 Excel 所在目录内）："
                    f"{binding.field_id} ← {raw!r}"
                )
            elif binding.options.get("on_empty", "hide") == "hide":
                # 空图片位不能保留模板原图，否则会张冠李戴
                element = shape._element  # noqa: SLF001
                blip = element.find(".//" + qn("a:blip"))
                empty_rid = blip.get(qn("r:embed")) if blip is not None else None
                element.getparent().remove(element)
                if empty_rid:
                    slide.part.drop_rel(empty_rid)
                replaced += 1
        elif binding.kind == "constant":
            replaced += _apply_text(
                shape,
                field,
                _to_text(binding.source.get("value")),
                binding.options,
                binding.source.get("slot"),
            )
        else:
            warnings.append(f"暂不支持的绑定类型：{binding.kind}（{binding.field_id}）")
    return replaced


def _paginate(
    rows: list[dict[str, Any]], pagination: dict[str, Any]
) -> list[list[dict[str, Any]]]:
    mode = pagination.get("mode", "one_per_row")
    if mode == "one_per_row":
        return [[r] for r in rows]
    if mode == "n_per_page":
        size = max(1, int(pagination.get("size", 5)))
        return [rows[i : i + size] for i in range(0, len(rows), size)]
    raise RenderError(f"不支持的分页模式：{mode}（可选 one_per_row / n_per_page）。")


_IMAGE_ID_RE = re.compile(r"ID_[A-Fa-f0-9]{6,}")


def _extract_image_id(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value)
    match = _IMAGE_ID_RE.search(text)
    return match.group(0) if match else None


def _resolve_image_paths(
    rows: list[dict[str, Any]],
    bindings: list[Binding],
    excel_path: str,
    warnings: list[str],
) -> tuple[dict[str, str], str | None]:
    """按需要提取 DISPIMG 图片，返回 ``(图片ID→路径, 临时目录)``。

    只解压本页真正用到的图片：300MB 台账整包提取会拖垮预览。
    """
    image_bindings = [b for b in bindings if b.kind == "image"]
    if not image_bindings:
        return {}, None
    columns = {b.source.get("column") for b in image_bindings if b.source.get("column")}
    ids = {
        image_id
        for row in rows
        for col in columns
        if col
        for image_id in [_extract_image_id(row.get(col))]
        if image_id
    }
    if not ids:
        return {}, None
    tmp_dir = tempfile.mkdtemp(prefix="e2p_img_")
    try:
        mapping = extract_cell_images(excel_path, tmp_dir, only=ids)
    except ExcelImageError as exc:
        warnings.append(f"图片提取失败，图片绑定将被跳过：{exc}")
        return {}, tmp_dir
    skipped = mapping.pop("__skipped__", None)
    if skipped:
        warnings.append(f"{skipped} 张内嵌图片在 Excel 包内缺失，已跳过。")
    return mapping, tmp_dir


def _resolve_one_image(
    value: Any, image_map: dict[str, str], root_dir: str | None = None
) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    image_id = _extract_image_id(text)
    if image_id and image_id in image_map:
        return image_map[image_id]
    if os.path.isfile(text):
        # 台账里的绝对路径只允许落在 Excel 所在目录内，避免被诱导读取任意本机图片
        if root_dir:
            try:
                if not Path(text).resolve().is_relative_to(Path(root_dir).resolve()):
                    return None
            except (OSError, ValueError):
                return None
        return text
    return None


def render(
    template: str,
    profile: Profile,
    out: str | None = None,
    excel: str | None = None,
    on_progress: Callable[[int, int], None] | None = None,
    rows_override: list[dict[str, Any]] | None = None,
) -> RenderResult:
    """按 profile 把 Excel 逐行渲染为多页 PPT。

    ``rows_override`` 用于「只渲染某一页」的预览场景：传入该页已选好的行，
    跳过 Excel 全量行。正常生成时保持 ``None``。
    """
    excel_path = excel or profile.excel
    if not excel_path:
        raise RenderError("未指定 Excel 文件。")
    out_path = out or profile.output.get("path") or ""
    if not out_path:
        base, _ = os.path.splitext(excel_path)
        out_path = f"{base}.pptx"
    out_path = os.path.abspath(out_path)
    if os.path.abspath(excel_path) == out_path or os.path.abspath(template) == out_path:
        raise RenderError("输出路径不能与输入文件相同（会覆盖源文件）。")

    warnings: list[str] = []
    info = scan_template(template, profile.slide_index)
    warnings.extend(info.warnings)
    field_map = info.field_map

    sheet: SheetData
    try:
        sheet = read_excel(excel_path, profile.sheet or None, profile.header_row)
    except ExcelDataError as exc:
        raise RenderError(str(exc)) from exc
    warnings.extend(sheet.warnings)
    rows = list(rows_override) if rows_override is not None else sheet.rows
    if not rows:
        raise RenderError("没有可渲染的数据行。")

    binding_field_ids = {b.field_id for b in profile.bindings}
    for b in profile.bindings:
        if b.field_id not in field_map:
            warnings.append(f"绑定指向模板中不存在的元素：{b.field_id}（已跳过）")
    for field_id in field_map:
        if field_id not in binding_field_ids:
            field = field_map[field_id]
            if field.kind in ("text", "table", "image"):
                warnings.append(f"元素未映射：{field_id}（{field.name}）")

    excel_dir = os.path.dirname(os.path.abspath(excel_path))
    image_map, image_tmp = _resolve_image_paths(rows, profile.bindings, excel_path, warnings)
    try:
        try:
            prs = Presentation(template)
        except Exception as exc:  # noqa: BLE001
            raise RenderError(f"无法打开模板：{exc}") from exc
        # 记录模板原始页数：生成后要把这些母版页全部删掉（只留克隆页）
        template_pages = len(prs.slides)
        if profile.slide_index >= len(prs.slides):
            raise RenderError(f"模板页索引越界：{profile.slide_index}")
        src_slide = prs.slides[profile.slide_index]

        pages = _paginate(rows, profile.pagination)
        if not pages:
            raise RenderError("没有有效数据行。")

        replaced = 0
        layout = profile.layout if isinstance(profile.layout, dict) else {}
        cover_spec = layout.get("cover") if isinstance(layout.get("cover"), dict) else None
        section_spec = (
            layout.get("section") if isinstance(layout.get("section"), dict) else None
        )
        end_spec = layout.get("end") if isinstance(layout.get("end"), dict) else None
        group_by = profile.pagination.get("group_by")

        page_specs: dict[str, tuple[dict[str, Field], list[Binding], int]] = {}

        def _layout_page(name: str, spec: dict[str, Any], default_index: int):
            if name in page_specs:
                return page_specs[name]
            index = int(spec.get("slide_index", default_index))
            if index >= len(prs.slides):
                raise RenderError(f"编排页「{name}」索引越界：{index}")
            info = scan_template(template, index)
            fields = {f.id: f for f in info.fields}
            bindings = [
                b if isinstance(b, Binding) else Binding.from_dict(b)
                for b in (spec.get("bindings") or [])
            ]
            page_specs[name] = (fields, bindings, index)
            return page_specs[name]

        groups: list[tuple[str | None, list[dict[str, Any]]]] = []
        if group_by:
            for group_name, group_rows in itertools.groupby(
                rows, key=lambda r: _to_text(r.get(group_by))
            ):
                groups.append((group_name, list(group_rows)))
        else:
            groups = [(None, rows)]

        section_count = len(groups) if (section_spec and group_by) else 0
        total = len(pages) + (1 if cover_spec else 0) + section_count + (1 if end_spec else 0)
        warned_overflow: set[str] = set()
        slide_no = 0
        used_pages: set[int] = {profile.slide_index}

        if cover_spec:
            fields, bindings, index = _layout_page("cover", cover_spec, 1)
            slide = clone_slide(prs, prs.slides[index])
            replaced += _apply_bindings(
                slide,
                fields,
                bindings,
                [rows[0]],
                image_map=image_map,
                excel_dir=excel_dir,
                warnings=warnings,
                page_index=slide_no,
                total_pages=total,
                warned_overflow=warned_overflow,
            )
            used_pages.add(index)
            slide_no += 1

        done = 0
        for _group_name, group_rows in groups:
            if section_spec and group_by:
                fields, bindings, index = _layout_page("section", section_spec, 2)
                slide = clone_slide(prs, prs.slides[index])
                replaced += _apply_bindings(
                    slide,
                    fields,
                    bindings,
                    group_rows[:1],
                    image_map=image_map,
                    excel_dir=excel_dir,
                    warnings=warnings,
                    page_index=slide_no,
                    total_pages=total,
                    warned_overflow=warned_overflow,
                )
                used_pages.add(index)
                slide_no += 1
            for page_rows in _paginate(group_rows, profile.pagination):
                slide = clone_slide(prs, src_slide)
                replaced += _apply_bindings(
                    slide,
                    field_map,
                    profile.bindings,
                    page_rows,
                    image_map=image_map,
                    excel_dir=excel_dir,
                    warnings=warnings,
                    page_index=slide_no,
                    total_pages=total,
                    warned_overflow=warned_overflow,
                )
                slide_no += 1
                done += 1
                if on_progress:
                    on_progress(done, len(pages))

        if end_spec:
            fields, bindings, index = _layout_page("end", end_spec, 3)
            slide = clone_slide(prs, prs.slides[index])
            replaced += _apply_bindings(
                slide,
                fields,
                bindings,
                [rows[-1]],
                image_map=image_map,
                excel_dir=excel_dir,
                warnings=warnings,
                page_index=slide_no,
                total_pages=total,
                warned_overflow=warned_overflow,
            )
            used_pages.add(index)
            slide_no += 1

        # 删除全部母版页（封面/正文/章节/结尾的原始页都要删，只保留克隆出来的页）
        for index in range(template_pages - 1, -1, -1):
            _drop_slide(prs, index)
        if len(prs.slides) == 0:
            raise RenderError("生成后没有任何幻灯片（数据行可能为空）。")

        os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
        tmp_path = out_path + ".tmp"
        try:
            prs.save(tmp_path)
            os.replace(tmp_path, out_path)
        except Exception as exc:  # noqa: BLE001
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
            raise RenderError(f"保存失败：{exc}") from exc
    finally:
        if image_tmp:
            shutil.rmtree(image_tmp, ignore_errors=True)

    return RenderResult(
        output=out_path,
        rows=len(rows),
        slides=len(prs.slides),
        replaced=replaced,
        warnings=warnings,
    )
