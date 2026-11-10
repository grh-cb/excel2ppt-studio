from __future__ import annotations

import pytest
from openpyxl import Workbook

from excel2ppt_studio.core.excel_reader import ExcelDataError, read_excel


def test_read_basic(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["名称", "数量", "日期"])
    ws.append(["A", 1, __import__("datetime").date(2026, 1, 2)])
    ws.append(["B", 2, __import__("datetime").date(2026, 1, 3)])
    path = tmp_path / "ok.xlsx"
    wb.save(str(path))

    data = read_excel(str(path))
    assert data.column_names == ["名称", "数量", "日期"]
    assert len(data.rows) == 2
    assert [c.dtype for c in data.columns] == ["text", "number", "date"]


def test_empty_header_raises(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["名称", None, "数量"])
    ws.append(["A", "x", 1])
    path = tmp_path / "bad.xlsx"
    wb.save(str(path))
    with pytest.raises(ExcelDataError, match="表头为空"):
        read_excel(str(path))


def test_duplicate_header_renamed_with_warning(tmp_path):
    """真实台账常有重名列：自动加序号并告警，而不是直接拒绝文件。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["名称", "名称"])
    ws.append(["A", "B"])
    path = tmp_path / "dup.xlsx"
    wb.save(str(path))

    data = read_excel(str(path))
    assert data.column_names == ["名称", "名称(2)"]
    assert any("重复" in w for w in data.warnings)
    assert data.rows[0]["名称(2)"] == "B"


def test_no_data_rows_raises(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["名称"])
    path = tmp_path / "empty.xlsx"
    wb.save(str(path))
    with pytest.raises(ExcelDataError, match="没有有效数据行"):
        read_excel(str(path))


def test_missing_sheet_raises(tmp_path):
    wb = Workbook()
    wb.active.append(["名称"])
    path = tmp_path / "sheet.xlsx"
    wb.save(str(path))
    with pytest.raises(ExcelDataError, match="工作表不存在"):
        read_excel(str(path), "不存在")


def test_image_column_detected(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["名称", "照片"])
    ws.append(["A", "ID_1A2B3C4D5E"])
    ws.append(["B", "ID_1A2B3C4D5F"])
    path = tmp_path / "img.xlsx"
    wb.save(str(path))
    data = read_excel(str(path))
    assert data.columns[1].dtype == "image"


def test_merged_header_raises_clear_error(tmp_path):
    """表头行有合并单元格时，必须明确报「合并」，而不是含糊的「表头为空」。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["项目名称", None, "问题描述"])
    ws.merge_cells("A1:B1")
    ws.append(["A", "x", "y"])
    path = tmp_path / "merged.xlsx"
    wb.save(str(path))
    with pytest.raises(ExcelDataError, match="合并"):
        read_excel(str(path))


def test_formula_without_cached_value_warns(tmp_path):
    """公式单元格没有缓存值时不能静默当空，必须给出警告。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["名称", "数量"])
    ws.append(["A", "=1+1"])
    path = tmp_path / "formula.xlsx"
    wb.save(str(path))
    data = read_excel(str(path))
    assert any("公式" in w for w in data.warnings), data.warnings


def test_trailing_empty_header_columns_ignored(tmp_path):
    """表头右侧因残留格式撑出的空列应被裁剪，而不是报「表头为空」。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "检查表"
    ws.append(["名称", "数量", "备注"])
    ws.append(["A", 1, "x"])
    ws.cell(row=2, column=5, value="残留格式")
    path = tmp_path / "trailing.xlsx"
    wb.save(str(path))

    data = read_excel(str(path))
    assert data.column_names == ["名称", "数量", "备注"]
