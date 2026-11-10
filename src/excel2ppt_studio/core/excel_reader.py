"""读取 Excel：表头校验、类型推断、行数据。

约定：首行（``header_row``）必须是表头；表头不得为空或重复。
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import Any

from openpyxl import load_workbook

from .model import Column

_IMAGE_ID = re.compile(r"^ID_[A-Fa-f0-9]{6,}$")
_DISPIMG_FORMULA = re.compile(r"DISPIMG\(\s*\"?(ID_[A-Fa-f0-9]+)\"?", re.IGNORECASE)
_IMAGE_PATH = re.compile(r"\.(png|jpe?g|bmp|gif|webp)$", re.IGNORECASE)
_MAX_SCAN = 200


class ExcelDataError(Exception):
    """Excel 结构或数据不符合约定。"""


@dataclass
class SheetData:
    path: str
    sheet: str
    columns: list[Column] = field(default_factory=list)
    rows: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def column_names(self) -> list[str]:
        return [c.name for c in self.columns]

    @property
    def header_columns(self) -> dict[str, str]:
        return {c.name: c.name for c in self.columns}


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.isoformat(sep=" ") if isinstance(value, _dt.datetime) else value.isoformat()
    return str(value).strip()


def _infer_dtype(name: str, values: list[Any]) -> str:
    sample = [v for v in values if v is not None and _cell_text(v) != ""]
    if not sample:
        return "text"
    texts = [_cell_text(v) for v in sample]
    if any(_IMAGE_ID.match(t) or _DISPIMG_FORMULA.search(t) for t in texts):
        return "image"
    path_like = sum(1 for t in texts if _IMAGE_PATH.search(t))
    if path_like >= max(1, int(len(texts) * 0.6)):
        return "image"
    nums = 0
    dates = 0
    bools = 0
    for v in sample:
        if isinstance(v, bool):
            bools += 1
        elif isinstance(v, (int, float)):
            nums += 1
        elif isinstance(v, (_dt.datetime, _dt.date)):
            dates += 1
    total = len(sample)
    if bools == total:
        return "bool"
    if dates >= total * 0.8:
        return "date"
    if nums >= total * 0.8:
        return "number"
    return "text"


def _read_header(ws, header_row: int) -> tuple[list[str], list[str]]:
    values = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
    # 右侧因残留格式撑出的空列不算表头：从尾部裁掉，中间的空列仍然报错
    while values and _cell_text(values[-1]) == "":
        values.pop()
    names: list[str] = []
    warnings: list[str] = []
    seen: set[str] = set()
    for idx, v in enumerate(values, start=1):
        text = _cell_text(v)
        if not text:
            raise ExcelDataError(
                f"第 {idx} 列表头为空（表头行 = 第 {header_row} 行）。"
                "请补全表头，或把 header_row 指到真正的表头行。"
            )
        if text in seen:
            # 真实台账常见重名列（如两个「归档率」），自动加序号并在结果里告警
            suffix = 2
            while f"{text}({suffix})" in seen:
                suffix += 1
            renamed = f"{text}({suffix})"
            warnings.append(f"表头「{text}」重复，已重命名为「{renamed}」")
            text = renamed
        seen.add(text)
        names.append(text)
    if not names:
        raise ExcelDataError(f"第 {header_row} 行没有检测到表头。")
    return names, warnings


def _header_merged_ranges(path: str, sheet: str, header_row: int) -> list[str]:
    """仅在表头出现空单元格时调用：用非只读模式确认是否由合并导致。"""
    try:
        wb = load_workbook(path, data_only=True, read_only=False)
    except Exception:  # noqa: BLE001 - 检测失败不影响主流程
        return []
    try:
        ws = wb[sheet]
        return [
            str(rng)
            for rng in ws.merged_cells.ranges
            if rng.min_row <= header_row <= rng.max_row
        ]
    except Exception:  # noqa: BLE001
        return []
    finally:
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass


def _uncached_formula_count(
    path: str, sheet: str, header_row: int, max_col: int
) -> int:
    """统计「有公式但没有缓存值」的单元格数（文件未在 Excel/WPS 中重算保存）。"""
    count = 0
    wb_f = None
    wb_v = None
    try:
        wb_f = load_workbook(path, data_only=False, read_only=True)
        wb_v = load_workbook(path, data_only=True, read_only=True)
        ws_f = wb_f[sheet]
        ws_v = wb_v[sheet]
        for row in ws_f.iter_rows(min_row=header_row + 1, max_col=max_col):
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    cached = ws_v.cell(row=cell.row, column=cell.column).value
                    if cached in (None, ""):
                        count += 1
    except Exception:  # noqa: BLE001 - 公式扫描失败不影响主流程
        return 0
    finally:
        for wb in (wb_f, wb_v):
            if wb is not None:
                try:
                    wb.close()
                except Exception:  # noqa: BLE001
                    pass
    return count


def read_excel(
    path: str,
    sheet: str | None = None,
    header_row: int = 1,
) -> SheetData:
    """读取指定工作表，返回列元数据与逐行字典。"""
    try:
        wb_values = load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001 - 统一转为可读错误
        raise ExcelDataError(f"无法打开 Excel：{path}（{exc}）") from exc

    try:
        if sheet is None:
            sheet = wb_values.sheetnames[0]
        if sheet not in wb_values.sheetnames:
            raise ExcelDataError(
                f"工作表不存在：{sheet}。可用工作表：{'、'.join(wb_values.sheetnames)}"
            )
        ws = wb_values[sheet]
        try:
            names, header_warnings = _read_header(ws, header_row)
        except ExcelDataError as exc:
            # 只有表头为空时才值得用非只读模式检查合并（大文件打开代价高）
            merged = _header_merged_ranges(path, sheet, header_row)
            if merged:
                raise ExcelDataError(
                    f"表头行（第 {header_row} 行）存在合并单元格：{'、'.join(merged[:5])}。"
                    "请先在 Excel/WPS 中取消合并，把表头拆成单行后再试。"
                ) from exc
            raise

        raw_rows: list[list[Any]] = []
        for row in ws.iter_rows(min_row=header_row + 1, max_col=len(names), values_only=True):
            if all(_cell_text(v) == "" for v in row):
                continue
            raw_rows.append(list(row))

        if not raw_rows:
            raise ExcelDataError(f"工作表「{sheet}」在表头之后没有有效数据行。")

        warnings: list[str] = list(header_warnings)
        if any(v is None for r in raw_rows for v in r):
            uncached = _uncached_formula_count(path, sheet, header_row, len(names))
            if uncached:
                warnings.append(
                    f"{uncached} 个公式单元格没有缓存值（文件可能未在 Excel/WPS 中重算保存），"
                    "这些单元格会被当作空值处理。"
                )

        columns: list[Column] = []
        for c_idx, name in enumerate(names):
            values = [r[c_idx] for r in raw_rows[:_MAX_SCAN]]
            dtype = _infer_dtype(name, values)
            samples = [_cell_text(v) for v in values if _cell_text(v)][:5]
            columns.append(Column(index=c_idx, name=name, dtype=dtype, samples=samples))

        rows: list[dict[str, Any]] = []
        for r in raw_rows:
            rows.append({name: r[c_idx] for c_idx, name in enumerate(names)})

        return SheetData(
            path=path, sheet=sheet, columns=columns, rows=rows, warnings=warnings
        )
    finally:
        wb_values.close()


def read_dispimg_formulas(path: str, sheet: str | None = None) -> dict[tuple[int, int], str]:
    """读取公式文本，返回 ``{(行, 列): 图片ID}``，用于识别 DISPIMG 单元格。

    ``sheet`` 为 None 时按工作表顺序查找，返回第一个含 DISPIMG 的工作表结果
    （台账常把「使用说明」放在第一页，不能只看第一页）。
    """
    wb = load_workbook(path, data_only=False, read_only=True)
    try:
        candidates = [sheet] if sheet else list(wb.sheetnames)
        for name in candidates:
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            out: dict[tuple[int, int], str] = {}
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        m = _DISPIMG_FORMULA.search(cell.value)
                        if m:
                            out[(cell.row, cell.column)] = m.group(1)
            if out:
                return out
        return {}
    finally:
        wb.close()
