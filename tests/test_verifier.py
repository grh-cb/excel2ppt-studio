from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.util import Pt

from excel2ppt_studio.core.model import Binding, Profile
from excel2ppt_studio.core.renderer import render
from excel2ppt_studio.core.template_scanner import scan_template
from excel2ppt_studio.core.verifier import fingerprint_slide, verify_file


def _build(template_path, excel_path, tmp_path):
    info = scan_template(template_path)
    title = next(f for f in info.fields if f.role == "title")
    body = next(f for f in info.fields if f.role == "body")
    profile = Profile(
        template=template_path,
        excel=excel_path,
        sheet="检查表",
        bindings=[
            Binding(title.id, "text", {"column": "项目名称"}),
            Binding(body.id, "text", {"column": "单位", "slot": "单位"}),
            Binding(body.id, "text", {"column": "问题描述", "slot": "问题描述"}),
        ],
    )
    out = tmp_path / "out.pptx"
    render(template_path, profile, out=str(out))
    return out


def test_verify_passes_for_cloned_output(template_path, excel_path, tmp_path):
    out = _build(template_path, excel_path, tmp_path)
    result = verify_file(template_path, str(out))
    assert result.ok is True
    assert result.pages_checked == 3
    assert result.shapes_checked >= 15
    assert result.errors == []


def test_verify_detects_font_change(template_path, excel_path, tmp_path):
    out = _build(template_path, excel_path, tmp_path)
    prs = Presentation(str(out))
    slide = prs.slides[0]
    for shape in slide.shapes:
        if shape.has_text_frame and "项目A" in shape.text_frame.text:
            for run in shape.text_frame.paragraphs[0].runs:
                run.font.size = Pt(44)
            break
    broken = tmp_path / "broken.pptx"
    prs.save(str(broken))

    result = verify_file(template_path, str(broken))
    assert result.ok is False
    codes = {i.code for i in result.errors}
    assert "font_changed" in codes


def test_verify_detects_geometry_change(template_path, excel_path, tmp_path):
    out = _build(template_path, excel_path, tmp_path)
    prs = Presentation(str(out))
    slide = prs.slides[0]
    for shape in slide.shapes:
        if shape.has_text_frame and "项目A" in shape.text_frame.text:
            shape.top = int(shape.top) + 500000
            break
    broken = tmp_path / "moved.pptx"
    prs.save(str(broken))

    result = verify_file(template_path, str(broken))
    assert any(i.code == "geometry_changed" for i in result.errors)


def test_verify_warns_residual_slot(template_path, excel_path, tmp_path):
    info = scan_template(template_path)
    title = next(f for f in info.fields if f.role == "title")
    profile = Profile(
        template=template_path,
        excel=excel_path,
        sheet="检查表",
        bindings=[Binding(title.id, "text", {"column": "项目名称"})],
    )
    out = tmp_path / "partial.pptx"
    render(template_path, profile, out=str(out))
    result = verify_file(template_path, str(out))
    assert any(i.code == "residual_slot" for i in result.warnings)


def test_fingerprint_ignores_text(template_path):
    prs = Presentation(template_path)
    slide = prs.slides[0]
    before = fingerprint_slide(slide)
    for shape in slide.shapes:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                for r in p.runs:
                    r.text = "被替换的文本"
    after = fingerprint_slide(slide)
    assert {k: v.digest for k, v in before.items()} == {k: v.digest for k, v in after.items()}


def test_verify_rejects_missing_files(tmp_path):
    from excel2ppt_studio.core.verifier import VerifyError

    with pytest.raises(VerifyError, match="不存在"):
        verify_file(str(tmp_path / "no.pptx"), str(tmp_path / "no2.pptx"))
