from __future__ import annotations

from excel2ppt_studio.core.model import Paragraph, Run
from excel2ppt_studio.core.slot_detect import normalize_name, scan_paragraphs, scan_shape_meta
from excel2ppt_studio.core.template_scanner import scan_template


def test_scan_finds_all_element_kinds(template_path):
    info = scan_template(template_path)
    kinds = {f.kind for f in info.fields}
    assert {"text", "table", "image", "chart"} <= kinds
    assert len(info.fields) == 5


def test_scan_extracts_slots_and_geometry(template_path):
    info = scan_template(template_path)
    title = next(f for f in info.fields if f.role == "title")
    assert title.rect.x == 600000
    assert title.rect.y == 300000
    assert "项目名称" in title.slot_names
    body = next(f for f in info.fields if f.role == "body")
    assert "单位" in body.slot_names
    assert "问题描述" in body.slot_names


def test_picture_placeholder_is_image(tmp_path):
    """图片占位符的 shape_type 是 PLACEHOLDER，必须识别为 image 而不是 text/body。"""
    from pptx import Presentation

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[8])  # Picture with Caption
    path = tmp_path / "pic_placeholder.pptx"
    prs.save(str(path))

    info = scan_template(str(path))
    pics = [f for f in info.fields if f.kind == "image"]
    assert pics, [(f.kind, f.role, f.name) for f in info.fields]
    assert pics[0].role == "image"


def test_scan_table_meta(template_path):
    info = scan_template(template_path)
    table = next(f for f in info.fields if f.kind == "table")
    assert table.table is not None
    assert table.table.rows == 2
    assert table.table.cols == 3
    assert table.table.cells[0] == ["单位", "问题描述", "备注"]


def test_scan_image_and_chart(template_path):
    info = scan_template(template_path)
    image = next(f for f in info.fields if f.kind == "image")
    assert image.image is not None
    assert image.image.width_px == 320
    chart = next(f for f in info.fields if f.kind == "chart")
    assert chart.chart is not None
    assert chart.chart["series"][0]["name"] == "系列1"


def test_scan_slide_index_out_of_range(template_path):
    import pytest

    from excel2ppt_studio.core.template_scanner import TemplateError

    with pytest.raises(TemplateError, match="越界"):
        scan_template(template_path, 99)


def test_slot_patterns():
    from excel2ppt_studio.core.model import Paragraph, Run

    paragraphs = [Paragraph(index=0, runs=[Run(text="项目：{{项目名称}}"), Run(text="${单位}")])]
    slots = scan_paragraphs(paragraphs, "sp-1")
    assert {s.name for s in slots} == {"项目名称", "单位"}


def test_label_slot():
    paragraphs = [Paragraph(index=0, runs=[Run(text="施工单位：")])]
    slots = scan_paragraphs(paragraphs, "sp-1")
    assert slots and slots[0].kind == "label"
    assert slots[0].name == "施工单位"


def test_named_shape_slot():
    slots = scan_shape_meta("fld_项目名称", "", "sp-1", "")
    assert slots[0].kind == "named"
    assert slots[0].name == "项目名称"


def test_normalize_name():
    assert normalize_name(" 项目 名称 ") == "项目名称"
    assert normalize_name("单位：") == "单位"


def test_paragraph_slots_for_plain_text():
    paragraphs = [
        Paragraph(index=0, runs=[Run(text="炼油五部新增氨液罐项目")]),
        Paragraph(index=1, runs=[Run(text="问题描述内容")]),
    ]
    slots = scan_paragraphs(paragraphs, "sp-1")
    assert [s.name for s in slots] == ["第1段", "第2段"]
    assert all(s.kind == "paragraph" for s in slots)


def test_paragraph_with_leading_spaces_is_one_slot():
    paragraphs = [Paragraph(index=0, runs=[Run(text="       "), Run(text="正文")])]
    slots = scan_paragraphs(paragraphs, "sp-1")
    assert len(slots) == 1
    assert slots[0].kind == "paragraph"


def test_summarize_slides_detects_body(tmp_path):
    """自动识别：封面文字少、正文元素多、结尾含「谢谢」。"""
    from pptx import Presentation
    from pptx.util import Emu

    from excel2ppt_studio.core.template_scanner import summarize_slides

    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)

    cover = prs.slides.add_slide(prs.slide_layouts[6])
    cover.shapes.add_textbox(Emu(600000), Emu(300000), Emu(5000000), Emu(900000)).text_frame.text = "封面"

    body = prs.slides.add_slide(prs.slide_layouts[6])
    body.shapes.add_textbox(Emu(600000), Emu(300000), Emu(5000000), Emu(600000)).text_frame.text = "{{标题}}"
    body.shapes.add_textbox(Emu(600000), Emu(1000000), Emu(8000000), Emu(1200000)).text_frame.text = "{{正文}}" * 5
    body.shapes.add_table(3, 3, Emu(600000), Emu(3000000), Emu(6000000), Emu(1500000))

    end = prs.slides.add_slide(prs.slide_layouts[6])
    end.shapes.add_textbox(Emu(600000), Emu(300000), Emu(5000000), Emu(900000)).text_frame.text = "谢谢大家的观看"

    path = tmp_path / "multi.pptx"
    prs.save(str(path))

    summaries, best = summarize_slides(str(path))
    kinds = [s.kind for s in summaries]
    assert kinds == ["cover", "body", "end"], kinds
    assert best == 1
