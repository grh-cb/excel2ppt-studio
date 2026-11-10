"""模板解析：把一页 PPT 拆成可映射的 ``Field`` 列表。

坐标统一换算为 EMU 绝对坐标；组合形状按 ``chOff/chExt`` 逐层换算。
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

from .model import Field, ImageMeta, Paragraph, Rect, Run, TableMeta
from .slot_detect import scan_paragraphs, scan_shape_meta


class TemplateError(Exception):
    """模板无法解析（可读中文信息）。"""


@dataclass
class TemplateInfo:
    path: str
    slide_index: int
    slide_count: int
    width: int
    height: int
    fields: list[Field] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def field_map(self) -> dict[str, Field]:
        return {f.id: f for f in self.fields}

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "slide_index": self.slide_index,
            "slide_count": self.slide_count,
            "width": self.width,
            "height": self.height,
            "fields": [f.to_dict() for f in self.fields],
            "warnings": self.warnings,
        }


def _safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _extract_runs(paragraph) -> list[Run]:
    runs: list[Run] = []
    for r in paragraph.runs:
        color = None
        try:
            rgb = r.font.color
            if rgb is not None and rgb.type is not None and rgb.rgb is not None:
                color = str(rgb.rgb)
        except (AttributeError, TypeError, ValueError):
            color = None
        size = None
        try:
            size = r.font.size.pt if r.font.size is not None else None
        except (AttributeError, TypeError):
            size = None
        runs.append(
            Run(
                text=r.text or "",
                bold=r.font.bold,
                italic=r.font.italic,
                size_pt=size,
                color=color,
                font=r.font.name,
            )
        )
    return runs


def _extract_paragraphs(shape) -> list[Paragraph]:
    if not getattr(shape, "has_text_frame", False):
        return []
    paragraphs: list[Paragraph] = []
    for idx, p in enumerate(shape.text_frame.paragraphs):
        align = None
        try:
            align = str(p.alignment) if p.alignment is not None else None
        except (AttributeError, ValueError):
            align = None
        paragraphs.append(
            Paragraph(index=idx, align=align, level=_safe_int(p.level), runs=_extract_runs(p))
        )
    return paragraphs


def _group_transform(shape) -> tuple[tuple[int, int], tuple[int, int, int, int]] | None:
    """返回组形状的 ``(chOff, chExt)``；没有则 None。"""
    try:
        grp = shape._element.grpSpPr  # noqa: SLF001 - 官方无公开 API
        xfrm = grp.find("{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
        if xfrm is None:
            return None
        ch_off = xfrm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}chOff")
        ch_ext = xfrm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}chExt")
        if ch_off is None or ch_ext is None:
            return None
        off = (_safe_int(ch_off.get("x")), _safe_int(ch_off.get("y")))
        ext = (
            _safe_int(ch_ext.get("cx")),
            _safe_int(ch_ext.get("cy")),
            0,
            0,
        )
        return off, ext
    except (AttributeError, TypeError):
        return None


def _table_meta(shape) -> TableMeta:
    table = shape.table
    rows = len(table.rows)
    cols = len(table.columns)
    cells = [[table.cell(r, c).text for c in range(cols)] for r in range(rows)]
    col_widths = [_safe_int(col.width) for col in table.columns]
    row_heights = [_safe_int(row.height) for row in table.rows]
    first_row_header = bool(rows) and all(
        table.cell(0, c).text.strip() != "" for c in range(cols)
    )
    return TableMeta(
        rows=rows,
        cols=cols,
        cells=cells,
        col_widths=col_widths,
        row_heights=row_heights,
        first_row_header=first_row_header,
    )


def _image_meta(shape) -> ImageMeta:
    try:
        image = shape.image
        return ImageMeta(
            filename=image.filename or "",
            ext=image.ext or "",
            width_px=image.size[0] if image.size else None,
            height_px=image.size[1] if image.size else None,
        )
    except (AttributeError, KeyError, ValueError):
        return ImageMeta()


def _clean_name(name: str) -> str:
    """去掉形状名里的作者签名/括号备注，只留下可读的名字。"""
    text = re.sub(r"[（(][^)）]*[)）]", "", name or "")
    text = re.sub(r"\s+", " ", text).strip()
    return text or "形状"


def _is_picture_placeholder(shape) -> bool:
    """图片占位符的 shape_type 是 PLACEHOLDER，不能当文本处理。"""
    if not getattr(shape, "is_placeholder", False):
        return False
    try:
        return shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    except (AttributeError, ValueError):
        return False


def _chart_meta(shape) -> dict[str, Any]:
    chart = shape.chart
    info: dict[str, Any] = {
        "chart_type": str(chart.chart_type) if chart.chart_type else None,
        "has_legend": bool(chart.has_legend),
    }
    series = []
    for s in chart.series:
        values = None
        try:
            values = list(s.values)
        except (AttributeError, TypeError, ValueError):
            values = None
        series.append({"name": s.name, "values": values})
    info["series"] = series
    try:
        categories = list(chart.plots[0].categories)
    except (AttributeError, IndexError, TypeError):
        categories = []
    info["categories"] = [str(c) for c in categories]
    return info


def _role(shape, paragraphs: list[Paragraph]) -> str:
    if getattr(shape, "is_placeholder", False):
        try:
            ph_type = shape.placeholder_format.type
            if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                return "title"
            if ph_type == PP_PLACEHOLDER.SUBTITLE:
                return "subtitle"
            if ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT):
                return "body"
        except (AttributeError, ValueError):
            pass
    if getattr(shape, "has_table", False):
        return "table"
    if getattr(shape, "has_chart", False):
        return "chart"
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "image"
    sizes = [r.size_pt for p in paragraphs for r in p.runs if r.size_pt]
    if sizes and max(sizes) >= 24:
        return "title"
    if paragraphs:
        return "body"
    return "decoration"


class _Scanner:
    def __init__(self, warnings: list[str]) -> None:
        self.warnings = warnings
        self.fields: list[Field] = []
        self.z = 0
        self._used_ids: set[str] = set()

    def _field_id(self, shape, prefix: str) -> str:
        base = f"{prefix}-{shape.shape_id}"
        candidate = base
        n = 2
        while candidate in self._used_ids:
            candidate = f"{base}-{n}"
            n += 1
        self._used_ids.add(candidate)
        return candidate

    def walk(
        self,
        shape,
        origin: tuple[int, int],
        scale: tuple[float, float],
        depth: int = 0,
    ) -> None:
        left = origin[0] + _safe_int(getattr(shape, "left", 0)) * scale[0]
        top = origin[1] + _safe_int(getattr(shape, "top", 0)) * scale[1]
        width = _safe_int(getattr(shape, "width", 0)) * scale[0]
        height = _safe_int(getattr(shape, "height", 0)) * scale[1]
        rect = Rect(int(left), int(top), int(width), int(height))

        if shape.shape_type == MSO_SHAPE_TYPE.GROUP and hasattr(shape, "shapes"):
            transform = _group_transform(shape)
            if transform is None:
                self.warnings.append(f"组合形状缺少 chOff/chExt，按原坐标处理：{shape.name}")
                child_origin, child_scale = (rect.x, rect.y), (1.0, 1.0)
            else:
                ch_off, ch_ext = transform
                sx = (rect.w / ch_ext[0]) if ch_ext[0] else 1.0
                sy = (rect.h / ch_ext[1]) if ch_ext[1] else 1.0
                child_origin = (
                    rect.x - ch_off[0] * sx,
                    rect.y - ch_off[1] * sy,
                )
                child_scale = (sx, sy)
            for child in shape.shapes:
                self.walk(child, child_origin, child_scale, depth + 1)
            return

        if getattr(shape, "has_table", False):
            field_id = self._field_id(shape, "tbl")
            table = _table_meta(shape)
            text = "\n".join(" | ".join(row) for row in table.cells)
            self.fields.append(
                Field(
                    id=field_id,
                    kind="table",
                    name=_clean_name(shape.name),
                    role="table",
                    rect=rect,
                    z=self.z,
                    text=text,
                    table=table,
                )
            )
            self.z += 1
            return

        if getattr(shape, "has_chart", False):
            field_id = self._field_id(shape, "chart")
            chart = _chart_meta(shape)
            self.fields.append(
                Field(
                    id=field_id,
                    kind="chart",
                    name=_clean_name(shape.name),
                    role="chart",
                    rect=rect,
                    z=self.z,
                    text="",
                    chart=chart,
                )
            )
            self.z += 1
            return

        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or _is_picture_placeholder(shape):
            field_id = self._field_id(shape, "pic")
            image = _image_meta(shape)
            slots = scan_shape_meta(shape.name, _alt_text(shape), field_id, "")
            self.fields.append(
                Field(
                    id=field_id,
                    kind="image",
                    name=_clean_name(shape.name),
                    role="image",
                    rect=rect,
                    z=self.z,
                    text="",
                    image=image,
                    slots=slots,
                )
            )
            self.z += 1
            return

        paragraphs = _extract_paragraphs(shape)
        if paragraphs:
            field_id = self._field_id(shape, "sp")
            text = "\n".join(p.text for p in paragraphs)
            slots = scan_paragraphs(paragraphs, field_id)
            slots += scan_shape_meta(shape.name, _alt_text(shape), field_id, text)
            self.fields.append(
                Field(
                    id=field_id,
                    kind="text",
                    name=_clean_name(shape.name),
                    role=_role(shape, paragraphs),
                    rect=rect,
                    z=self.z,
                    text=text,
                    paragraphs=paragraphs,
                    slots=slots,
                )
            )
            self.z += 1
            return

        # 无文本、无表、无图的形状（装饰/线条）不产生 Field，但仍参与克隆
        self.z += 1


def _alt_text(shape) -> str:
    try:
        c_nv_pr = shape._element.find(
            ".//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr"
        )
        if c_nv_pr is None:
            c_nv_pr = shape._element.find(
                ".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr"
            )
        return c_nv_pr.get("descr", "") if c_nv_pr is not None else ""
    except AttributeError:
        return ""


def scan_template(path: str, slide_index: int = 0) -> TemplateInfo:
    """解析模板指定页，返回元素清单。"""
    if not os.path.isfile(path):
        raise TemplateError(f"模板文件不存在：{path}")
    if not path.lower().endswith(".pptx"):
        raise TemplateError("只支持 .pptx 模板（.ppt 请先另存为 .pptx）。")
    try:
        prs = Presentation(path)
    except Exception as exc:  # noqa: BLE001 - 统一转为可读错误
        raise TemplateError(f"无法打开模板（文件可能损坏或加密）：{exc}") from exc

    count = len(prs.slides)
    if count == 0:
        raise TemplateError("模板里没有任何幻灯片。")
    if not 0 <= slide_index < count:
        raise TemplateError(f"幻灯片索引越界：{slide_index}（共 {count} 页，从 0 开始）。")

    scanner = _Scanner([])
    slide = prs.slides[slide_index]
    for shape in slide.shapes:
        scanner.walk(shape, (0, 0), (1.0, 1.0))

    return TemplateInfo(
        path=os.path.abspath(path),
        slide_index=slide_index,
        slide_count=count,
        width=_safe_int(prs.slide_width),
        height=_safe_int(prs.slide_height),
        fields=scanner.fields,
        warnings=scanner.warnings,
    )


@dataclass
class SlideSummary:
    """一页模板的轻量摘要，用于自动识别正文页与缩略图。"""

    index: int
    kind: str  # cover | body | end | unknown
    score: int
    field_count: int
    slot_count: int
    table_count: int
    chart_count: int
    image_count: int
    text_chars: int


_END_WORDS = ("谢谢", "感谢", "thank", "THANK", "Thanks")


def _score_slide(info: TemplateInfo) -> tuple[int, int, int, int, int, int, str]:
    text_fields = [f for f in info.fields if f.kind == "text"]
    slots = sum(len(f.slots) for f in text_fields)
    real_slots = sum(
        1
        for f in text_fields
        for s in f.slots
        if s.kind in ("explicit", "label", "named", "alt")
    )
    tables = sum(1 for f in info.fields if f.kind == "table")
    charts = sum(1 for f in info.fields if f.kind == "chart")
    images = sum(1 for f in info.fields if f.kind == "image")
    text_chars = sum(len(f.text or "") for f in text_fields)
    joined = " ".join(f.text or "" for f in text_fields)

    score = (
        min(slots, 12) * 2
        + real_slots * 4
        + tables * 8
        + charts * 6
        + min(images, 4) * 2
        + min(text_chars // 40, 12)
    )
    if any(word in joined for word in _END_WORDS):
        score -= 25
    return score, slots, tables, charts, images, text_chars, joined


def summarize_slides(path: str) -> tuple[list[SlideSummary], int]:
    """扫描全部页，给每页打分并推荐正文页索引。

    规则：槽位/表格/图表/文字量越多越像正文页；含「谢谢/感谢」视为结尾；
    文字极少且无表格图表的视为封面。返回 ``(摘要列表, 推荐页索引)``。
    """
    first = scan_template(path, 0)
    summaries: list[SlideSummary] = []
    for idx in range(first.slide_count):
        info = first if idx == 0 else scan_template(path, idx)
        score, slots, tables, charts, images, text_chars, joined = _score_slide(info)
        if any(word in joined for word in _END_WORDS):
            kind = "end"
        elif text_chars < 40 and not tables and not charts:
            kind = "cover"
        else:
            kind = "body"
        # 最后一页如果不是正文，按结尾页展示（汇报模板的收尾页常与封面同版式）
        if idx == first.slide_count - 1 and kind != "body":
            kind = "end"
        summaries.append(
            SlideSummary(
                index=idx,
                kind=kind,
                score=score,
                field_count=len(info.fields),
                slot_count=slots,
                table_count=tables,
                chart_count=charts,
                image_count=images,
                text_chars=text_chars,
            )
        )

    bodies = [s for s in summaries if s.kind == "body"]
    best = max(bodies, key=lambda s: s.score) if bodies else max(summaries, key=lambda s: s.score)
    return summaries, best.index
