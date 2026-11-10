from __future__ import annotations

from pptx import Presentation
from pptx.util import Emu

from excel2ppt_studio.core.model import Binding, Profile
from excel2ppt_studio.core.renderer import render
from excel2ppt_studio.core.template_scanner import scan_template


def _find_text(slide, needle: str):
    for shape in slide.shapes:
        if shape.has_text_frame and needle in shape.text_frame.text:
            return shape
    raise AssertionError(f"未找到包含 {needle!r} 的形状")


def _make_profile(template_path, excel_path, **kwargs) -> Profile:
    info = scan_template(template_path)
    title = next(f for f in info.fields if f.role == "title")
    body = next(f for f in info.fields if f.role == "body")
    table = next(f for f in info.fields if f.kind == "table")
    image = next(f for f in info.fields if f.kind == "image")
    bindings = [
        Binding(title.id, "text", {"column": "项目名称"}),
        Binding(body.id, "text", {"column": "单位", "slot": "单位"}),
        Binding(body.id, "text", {"column": "问题描述", "slot": "问题描述"}),
        Binding(table.id, "table", {"columns": ["单位", "问题描述", "备注"]}),
        Binding(image.id, "image", {"column": "问题图片"}, {"fit": "contain"}),
    ]
    return Profile(
        template=template_path,
        excel=excel_path,
        sheet="检查表",
        bindings=bindings,
        **kwargs,
    )


def test_render_one_page_per_row(template_path, excel_path, tmp_path):
    profile = _make_profile(template_path, excel_path)
    out = tmp_path / "out.pptx"
    result = render(template_path, profile, out=str(out))
    assert result.rows == 3
    assert result.slides == 3
    assert out.exists()

    prs = Presentation(str(out))
    assert len(prs.slides) == 3
    first = prs.slides[0]
    assert _find_text(first, "项目A").text_frame.text == "项目A"
    assert "一公司" in _find_text(first, "一公司").text_frame.text


def test_clone_preserves_geometry_and_font(template_path, excel_path, tmp_path):
    profile = _make_profile(template_path, excel_path)
    out = tmp_path / "out.pptx"
    render(template_path, profile, out=str(out))

    tpl = Presentation(template_path)
    src = tpl.slides[0]
    src_title = next(
        s for s in src.shapes if s.has_text_frame and "{{项目名称}}" in s.text_frame.text
    )
    src_run = src_title.text_frame.paragraphs[0].runs[0]

    out_prs = Presentation(str(out))
    out_title = _find_text(out_prs.slides[0], "项目A")
    out_run = out_title.text_frame.paragraphs[0].runs[0]

    assert out_title.left == src_title.left
    assert out_title.top == src_title.top
    assert out_title.width == src_title.width
    assert out_run.font.size == src_run.font.size
    assert out_run.font.bold == src_run.font.bold
    assert str(out_run.font.color.rgb) == str(src_run.font.color.rgb) == "C00000"


def test_text_replacement_keeps_label_run_format(template_path, excel_path, tmp_path):
    profile = _make_profile(template_path, excel_path)
    out = tmp_path / "out.pptx"
    render(template_path, profile, out=str(out))

    prs = Presentation(str(out))
    body = _find_text(prs.slides[0], "单位：")
    p1 = body.text_frame.paragraphs[0]
    assert p1.runs[0].text == "单位："
    assert p1.runs[1].text == "一公司"
    assert p1.runs[1].font.size.pt == 16
    assert p1.runs[0].font.bold is True


def test_table_rows_follow_data(template_path, excel_path, tmp_path):
    profile = _make_profile(template_path, excel_path)
    out = tmp_path / "out.pptx"
    render(template_path, profile, out=str(out))

    prs = Presentation(str(out))
    table_shape = next(s for s in prs.slides[0].shapes if s.has_table)
    table = table_shape.table
    assert len(table.rows) == 2  # 表头 + 1 行数据
    assert table.cell(1, 0).text == "一公司"
    assert table.cell(0, 0).text == "单位"  # 表头保持模板文本


def test_image_replaced_with_contain(template_path, excel_path, tmp_path):
    profile = _make_profile(template_path, excel_path)
    out = tmp_path / "out.pptx"
    render(template_path, profile, out=str(out))

    prs = Presentation(str(out))
    pic = next(s for s in prs.slides[0].shapes if s.shape_type == 13)
    # 原图 400x300，框 3400000x2100000 -> contain 后高度撑满、宽度按比例
    assert pic.height == Emu(2100000)
    assert abs(pic.width - int(2100000 * 400 / 300)) <= 2


def test_missing_image_warns_but_renders(template_path, tmp_path):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "检查表"
    ws.append(["项目名称", "单位", "问题描述", "问题图片", "备注"])
    ws.append(["项目A", "一公司", "问题一", "不存在.png", ""])
    xlsx = tmp_path / "bad.xlsx"
    wb.save(str(xlsx))

    profile = _make_profile(template_path, str(xlsx))
    out = tmp_path / "out.pptx"
    result = render(template_path, profile, out=str(out))
    assert result.slides == 1
    assert any("图片无法解析" in w for w in result.warnings)


def test_overflow_warns_without_changing_font(template_path, tmp_path):
    from openpyxl import Workbook

    long_text = "这是一段很长的现场问题描述，" * 30
    wb = Workbook()
    ws = wb.active
    ws.title = "检查表"
    ws.append(["项目名称", "单位", "问题描述", "问题图片", "备注"])
    ws.append(["项目A", "一公司", long_text, "", ""])
    xlsx = tmp_path / "long.xlsx"
    wb.save(str(xlsx))

    profile = _make_profile(template_path, str(xlsx))
    out = tmp_path / "long_out.pptx"
    result = render(template_path, profile, out=str(out))
    assert any("可能溢出" in w for w in result.warnings)
    assert "只告警不改" in next(w for w in result.warnings if "可能溢出" in w)

    # 字号必须保持模板原值（18pt 正文 / 28pt 标题均未被改写）
    prs = Presentation(str(out))
    body = _find_text(prs.slides[0], "单位：")
    assert body.text_frame.paragraphs[0].runs[1].font.size.pt == 16


def test_render_rows_override_renders_only_given_rows(template_path, excel_path, tmp_path):
    profile = _make_profile(template_path, excel_path)
    rows = [
        {
            "项目名称": "覆盖行",
            "单位": "X公司",
            "问题描述": "覆盖描述",
            "问题图片": "",
            "备注": "",
        }
    ]
    out = tmp_path / "override.pptx"
    result = render(template_path, profile, out=str(out), rows_override=rows)
    assert result.slides == 1
    assert result.rows == 1

    prs = Presentation(str(out))
    assert _find_text(prs.slides[0], "覆盖行").text_frame.text == "覆盖行"
    assert "X公司" in _find_text(prs.slides[0], "X公司").text_frame.text


def test_extract_image_id_from_dispimg_formula():
    from excel2ppt_studio.core.renderer import _extract_image_id, _resolve_one_image

    assert _extract_image_id('=DISPIMG("ID_A6BAE75C9EF04E15",1)') == "ID_A6BAE75C9EF04E15"
    assert _extract_image_id("ID_ABC123") == "ID_ABC123"
    assert _extract_image_id("普通文本") is None
    assert _resolve_one_image('=DISPIMG("ID_ABC123",1)', {"ID_ABC123": "x.png"}) == "x.png"
    assert _resolve_one_image("普通文本", {}) is None


def test_page_var_binding(template_path, excel_path, tmp_path):
    info = scan_template(template_path)
    title = next(f for f in info.fields if f.role == "title")
    profile = Profile(
        template=template_path,
        excel=excel_path,
        sheet="检查表",
        bindings=[Binding(title.id, "text", {"var": "page", "slot": "项目名称"})],
    )
    out = tmp_path / "pagevar.pptx"
    render(template_path, profile, out=str(out))
    prs = Presentation(str(out))
    assert _find_text(prs.slides[0], "1").text_frame.text == "1"
    assert _find_text(prs.slides[1], "2").text_frame.text == "2"


def test_value_map_overrides_column(template_path, excel_path, tmp_path):
    info = scan_template(template_path)
    title = next(f for f in info.fields if f.role == "title")
    profile = Profile(
        template=template_path,
        excel=excel_path,
        sheet="检查表",
        bindings=[
            Binding(
                title.id,
                "text",
                {
                    "column": "项目名称",
                    "slot": "项目名称",
                    "key_column": "单位",
                    "value_map": {"一公司": "门窗框间隙过大"},
                },
            )
        ],
    )
    out = tmp_path / "valuemap.pptx"
    render(template_path, profile, out=str(out))
    prs = Presentation(str(out))
    assert _find_text(prs.slides[0], "门窗框间隙过大").text_frame.text == "门窗框间隙过大"
    # 未命中的行走原列值
    assert _find_text(prs.slides[1], "项目B").text_frame.text == "项目B"


def test_clone_skips_notes_slides_and_duplicate_parts(template_path, excel_path, tmp_path):
    import zipfile

    profile = _make_profile(template_path, excel_path)
    out = tmp_path / "notes_out.pptx"
    render(template_path, profile, out=str(out))

    with zipfile.ZipFile(str(out)) as z:
        names = z.namelist()
        slide_parts = [n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
        assert len(slide_parts) == len(set(slide_parts)), f"slide part 重名：{slide_parts}"
        assert not [n for n in names if "notesSlide" in n], "备注页不应随克隆复制"
        assert "ppt/slides/slide1.xml" not in names, "模板页应被删除"


def test_empty_image_hides_placeholder(template_path, tmp_path):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "检查表"
    ws.append(["项目名称", "单位", "问题描述", "问题图片", "备注", "数量"])
    ws.append(["项目A", "一公司", "问题一", "", "", 1])
    xlsx = tmp_path / "noimg.xlsx"
    wb.save(str(xlsx))

    profile = _make_profile(template_path, str(xlsx))
    out = tmp_path / "noimg_out.pptx"
    result = render(template_path, profile, out=str(out))
    assert result.slides == 1

    prs = Presentation(str(out))
    shapes = [s for s in prs.slides[0].shapes if s.has_text_frame or s.has_table]
    assert shapes, "文本与表格应保留"
    assert not any(s.shape_type == 13 for s in prs.slides[0].shapes), "空图片位应被隐藏"

    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.oxml.ns import qn

    used = {
        el.get(qn("r:embed"))
        for el in prs.slides[0].part._element.iter()
        if el.get(qn("r:embed"))
    }
    dead = [
        rid
        for rid, rel in prs.slides[0].part.rels.items()
        if rel.reltype == RT.IMAGE and rid not in used
    ]
    assert not dead, f"隐藏空图后残留死引用：{dead}"


def test_image_srcrect_child_order(template_path, excel_path, tmp_path):
    """a:srcRect 必须排在 a:stretch 之前（CT_BlipFillProperties 子元素顺序）。"""
    from lxml import etree

    profile = _make_profile(template_path, excel_path)
    out = tmp_path / "order.pptx"
    render(template_path, profile, out=str(out))

    prs = Presentation(str(out))
    pic = next(s for s in prs.slides[0].shapes if s.shape_type == 13)
    names = [etree.QName(c).localname for c in pic._element.blipFill]
    assert "srcRect" in names, f"contain 应写入 srcRect：{names}"
    assert names.index("srcRect") < names.index("stretch"), f"srcRect 顺序错误：{names}"


def test_gapped_slide_numbers_do_not_duplicate_partname(tmp_path):
    """模板 slideN 编号有空洞时（内存态构造），克隆页也必须分配唯一 partname。

    注意：从磁盘加载的 pptx 会被 python-pptx 规范化 partname，空洞消失；
    这里直接在内存对象上克隆，覆盖 clone_slide 被复用的场景。
    """
    import zipfile

    from pptx import Presentation
    from pptx.util import Emu

    from excel2ppt_studio.core.renderer import _drop_slide, clone_slide

    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    for i in range(3):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s.shapes.add_textbox(Emu(600000), Emu(300000), Emu(5000000), Emu(800000))
        tb.text_frame.text = f"第{i + 1}页"
    _drop_slide(prs, 1)  # 内存态留下 slide1 + slide3 空洞
    src = prs.slides[0]
    clone_slide(prs, src)
    clone_slide(prs, src)
    _drop_slide(prs, 0)

    out = tmp_path / "gap.pptx"
    prs.save(str(out))
    with zipfile.ZipFile(str(out)) as z:
        parts = [
            n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        ]
    assert len(parts) == len(set(parts)), f"slide part 重名：{parts}"



def test_image_tmp_cleaned_when_render_fails(template_path, excel_path, tmp_path, monkeypatch):
    """渲染中途失败时，已解压的 e2p_img_* 目录必须被清理。"""
    import os

    import pytest

    from excel2ppt_studio.core import renderer

    captured: dict[str, str] = {}

    def fake_resolve(rows, bindings, excel, warnings):
        d = renderer.tempfile.mkdtemp(prefix="e2p_img_test_")
        captured["dir"] = d
        return {}, d

    monkeypatch.setattr(renderer, "_resolve_image_paths", fake_resolve)
    monkeypatch.setattr(renderer, "_paginate", lambda rows, pagination: [])
    profile = _make_profile(template_path, excel_path)
    profile.slide_index = 0

    with pytest.raises(renderer.RenderError):
        render(template_path, profile, out=str(tmp_path / "fail.pptx"))

    assert "dir" in captured
    assert not os.path.exists(captured["dir"]), "失败的渲染泄漏了临时图片目录"


def test_image_rel_dropped_after_replacement(template_path, excel_path, tmp_path):
    """换图后旧的图片关系必须删除，否则每页都残留一份模板图。"""
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.oxml.ns import qn

    profile = _make_profile(template_path, excel_path)
    out = tmp_path / "rel.pptx"
    render(template_path, profile, out=str(out))

    prs = Presentation(str(out))
    slide = prs.slides[0]
    used = {
        el.get(qn("r:embed"))
        for el in slide.part._element.iter()
        if el.get(qn("r:embed"))
    }
    dead = [
        rid
        for rid, rel in slide.part.rels.items()
        if rel.reltype == RT.IMAGE and rid not in used
    ]
    assert not dead, f"换图后残留死引用：{dead}"


def test_image_path_outside_excel_dir_rejected(template_path, tmp_path):
    """Excel 里写的绝对图片路径必须在 Excel 所在目录内，防止读任意本机文件。"""
    from openpyxl import Workbook
    from PIL import Image

    outside = tmp_path / "outside"
    outside.mkdir()
    img = outside / "secret.png"
    Image.new("RGB", (10, 10)).save(img)

    data_dir = tmp_path / "data"
    data_dir.mkdir()
    wb = Workbook()
    ws = wb.active
    ws.title = "检查表"
    ws.append(["项目名称", "单位", "问题描述", "问题图片", "备注"])
    ws.append(["项目A", "一公司", "问题一", str(img), ""])
    xlsx = data_dir / "d.xlsx"
    wb.save(str(xlsx))

    profile = _make_profile(template_path, str(xlsx))
    out = tmp_path / "outside_img.pptx"
    result = render(template_path, profile, out=str(out))
    assert any("图片" in w for w in result.warnings), result.warnings


def test_paragraph_replace_prefers_largest_run_size(tmp_path):
    """整段替换时应写进字号最大的 run，避免被模板里的 14pt 缩进 run 带偏。"""
    from openpyxl import Workbook
    from pptx import Presentation
    from pptx.util import Emu, Pt

    from excel2ppt_studio.core.model import Binding, Profile
    from excel2ppt_studio.core.renderer import render
    from excel2ppt_studio.core.template_scanner import scan_template

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Emu(600000), Emu(600000), Emu(6000000), Emu(2000000))
    para = box.text_frame.paragraphs[0]
    r1 = para.add_run()
    r1.text = "缩进"
    r1.font.size = Pt(14)
    r2 = para.add_run()
    r2.text = "正文"
    r2.font.size = Pt(16)
    tpl = tmp_path / "mixed.pptx"
    prs.save(str(tpl))

    wb = Workbook()
    ws = wb.active
    ws.title = "检查表"
    ws.append(["内容"])
    ws.append(["新文本"])
    xlsx = tmp_path / "d.xlsx"
    wb.save(str(xlsx))

    info = scan_template(str(tpl))
    field = info.fields[0]
    profile = Profile(
        template=str(tpl),
        excel=str(xlsx),
        sheet="检查表",
        bindings=[Binding(field.id, "text", {"column": "内容"})],
    )
    out = tmp_path / "o.pptx"
    render(str(tpl), profile, out=str(out))

    prs2 = Presentation(str(out))
    shape = prs2.slides[0].shapes[0]
    sizes = [
        r.font.size.pt for r in shape.text_frame.paragraphs[0].runs if r.text.strip()
    ]
    assert sizes == [16.0], sizes


def test_layout_cover_section_end(tmp_path):
    """封面/章节/结尾编排：顺序与页数必须正确，且旧单页模式不受影响。"""
    from openpyxl import Workbook
    from pptx import Presentation
    from pptx.util import Emu

    from excel2ppt_studio.core.model import Binding, Profile
    from excel2ppt_studio.core.renderer import render
    from excel2ppt_studio.core.template_scanner import scan_template

    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    for label in ("BODY", "COVER", "SECTION", "END"):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Emu(600000), Emu(300000), Emu(6000000), Emu(900000))
        box.text_frame.text = label
    tpl = tmp_path / "multi.pptx"
    prs.save(str(tpl))

    wb = Workbook()
    ws = wb.active
    ws.title = "检查表"
    ws.append(["项目名称", "单位"])
    ws.append(["项目A", "一标段"])
    ws.append(["项目B", "一标段"])
    ws.append(["项目C", "二标段"])
    xlsx = tmp_path / "data.xlsx"
    wb.save(str(xlsx))

    def field_of(page: int) -> str:
        return scan_template(str(tpl), page).fields[0].id

    profile = Profile(
        template=str(tpl),
        excel=str(xlsx),
        sheet="检查表",
        bindings=[Binding(field_of(0), "text", {"column": "项目名称"})],
        pagination={"mode": "one_per_row", "group_by": "单位"},
        layout={
            "cover": {
                "slide_index": 1,
                "bindings": [Binding(field_of(1), "text", {"column": "项目名称"})],
            },
            "section": {
                "slide_index": 2,
                "bindings": [Binding(field_of(2), "text", {"column": "单位"})],
            },
            "end": {
                "slide_index": 3,
                "bindings": [
                    Binding(field_of(3), "constant", {"value": "谢谢"})
                ],
            },
        },
    )
    out = tmp_path / "multi_out.pptx"
    result = render(str(tpl), profile, out=str(out))

    prs2 = Presentation(str(out))
    texts = [s.text_frame.text for s in prs2.slides[0].shapes if s.has_text_frame]
    assert texts == ["项目A"], texts
    all_texts = []
    for slide in prs2.slides:
        vals = [s.text_frame.text for s in slide.shapes if s.has_text_frame]
        all_texts.append(vals[0] if vals else "")
    assert all_texts == ["项目A", "一标段", "项目A", "项目B", "二标段", "项目C", "谢谢"], all_texts
    assert result.slides == 7


def test_unused_template_pages_are_removed(tmp_path):
    """多页模板只用其中一页当母版时，其余模板页必须删除，不能残留在输出开头。"""
    from openpyxl import Workbook
    from pptx import Presentation
    from pptx.util import Emu

    from excel2ppt_studio.core.model import Binding, Profile
    from excel2ppt_studio.core.renderer import render
    from excel2ppt_studio.core.template_scanner import scan_template

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    for label in ("COVER", "BODY", "END"):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(Emu(600000), Emu(300000), Emu(6000000), Emu(900000)).text_frame.text = label
    tpl = tmp_path / "three.pptx"
    prs.save(str(tpl))

    wb = Workbook()
    ws = wb.active
    ws.title = "检查表"
    ws.append(["名称"])
    ws.append(["项目A"])
    xlsx = tmp_path / "data.xlsx"
    wb.save(str(xlsx))

    body_field = scan_template(str(tpl), 1).fields[0].id
    profile = Profile(
        template=str(tpl),
        slide_index=1,
        excel=str(xlsx),
        sheet="检查表",
        bindings=[Binding(body_field, "text", {"column": "名称"})],
    )
    out = tmp_path / "trimmed.pptx"
    render(str(tpl), profile, out=str(out))

    prs2 = Presentation(str(out))
    assert len(prs2.slides) == 1
    texts = [s.text_frame.text for s in prs2.slides[0].shapes if s.has_text_frame]
    assert texts == ["项目A"], texts
