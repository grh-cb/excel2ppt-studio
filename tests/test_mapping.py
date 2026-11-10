from __future__ import annotations

from excel2ppt_studio.core.mapping import (
    build_profile,
    load_profile,
    save_profile,
    similarity,
)
from excel2ppt_studio.core.template_scanner import scan_template


def test_similarity():
    assert similarity("项目名称", "项目名称") == 100
    assert similarity("项目名称", "项目名称 ") == 100
    assert similarity("单位", "施工单位") == 80
    assert similarity("完全无关", "数量") < 60


def test_suggest_bindings_cover_title_and_table(template_path, excel_path):
    from excel2ppt_studio.core.excel_reader import read_excel

    info = scan_template(template_path)
    sheet = read_excel(excel_path, "检查表", 1)
    profile, notes = build_profile(
        template_path, excel_path, info, sheet.sheet, sheet.columns
    )
    kinds = {b.kind for b in profile.bindings}
    assert "text" in kinds
    assert "table" in kinds
    assert "image" in kinds
    title = next(f for f in info.fields if f.role == "title")
    assert any(b.field_id == title.id and b.source.get("slot") == "项目名称" for b in profile.bindings)


def test_profile_roundtrip(template_path, excel_path, tmp_path):
    from excel2ppt_studio.core.excel_reader import read_excel

    info = scan_template(template_path)
    sheet = read_excel(excel_path, "检查表", 1)
    profile, _ = build_profile(template_path, excel_path, info, sheet.sheet, sheet.columns)
    path = tmp_path / "profile.json"
    save_profile(profile, path)
    loaded = load_profile(str(path))
    assert loaded.to_dict() == profile.to_dict()


def test_build_profile_keeps_header_row(template_path, excel_path):
    """suggest 时指定的表头行必须写进 Profile，否则渲染会回到第 1 行。"""
    from excel2ppt_studio.core.excel_reader import read_excel

    info = scan_template(template_path)
    sheet = read_excel(excel_path, "检查表", 1)
    profile, _ = build_profile(
        template_path, excel_path, info, sheet.sheet, sheet.columns, header_row=3
    )
    assert profile.header_row == 3


def test_load_profile_rejects_empty_bindings(tmp_path):
    import json

    import pytest

    from excel2ppt_studio.core.mapping import ProfileError

    path = tmp_path / "empty.json"
    path.write_text(json.dumps({"version": 1, "bindings": []}), encoding="utf-8")
    with pytest.raises(ProfileError, match="没有任何 bindings"):
        load_profile(str(path))
