from __future__ import annotations

from excel2ppt_studio.core.template_health import check_template


def test_health_ok_template_scores_high(template_path):
    report = check_template(template_path)
    assert report.score >= 80
    assert report.stats["slots"] >= 2
    assert not report.errors


def test_health_flags_paragraph_only(tmp_path):
    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Emu(600000), Emu(300000), Emu(6000000), Emu(900000))
    box.text_frame.text = "固定标题，没有占位符"
    path = tmp_path / "plain.pptx"
    prs.save(str(path))

    report = check_template(str(path))
    assert any(i.code == "paragraph_only" for i in report.issues)
    assert 0 < report.score < 100


def test_health_flags_no_bindable_elements(tmp_path):
    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    prs.slides.add_slide(prs.slide_layouts[6])
    path = tmp_path / "empty.pptx"
    prs.save(str(path))

    report = check_template(str(path))
    assert any(i.code == "no_bindable" for i in report.issues)
    assert report.score < 100
