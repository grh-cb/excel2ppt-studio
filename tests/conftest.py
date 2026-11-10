"""共享 fixtures：模板与 Excel 全部程序化生成，不依赖二进制样例。"""

from __future__ import annotations

import pytest
from openpyxl import Workbook
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu, Pt

from excel2ppt_studio.core.mapping import build_profile
from excel2ppt_studio.core.template_scanner import scan_template


@pytest.fixture()
def picture_file(tmp_path) -> str:
    path = tmp_path / "photo.png"
    im = Image.new("RGB", (400, 300), (235, 242, 250))
    draw = ImageDraw.Draw(im)
    draw.rectangle([10, 10, 390, 290], outline=(30, 80, 160), width=5)
    im.save(path)
    return str(path)


@pytest.fixture()
def template_path(tmp_path) -> str:
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = slide.shapes.add_textbox(Emu(600000), Emu(300000), Emu(10900000), Emu(900000))
    run = title.text_frame.paragraphs[0].add_run()
    run.text = "{{项目名称}}"
    run.font.name = "微软雅黑"
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    body = slide.shapes.add_textbox(Emu(600000), Emu(1400000), Emu(6500000), Emu(2000000))
    tf = body.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    r0 = p1.add_run()
    r0.text = "单位："
    r0.font.size = Pt(16)
    r0.font.bold = True
    r0.font.color.rgb = RGBColor(0x00, 0x70, 0xC0)
    r1 = p1.add_run()
    r1.text = "{{单位}}"
    r1.font.size = Pt(16)
    r1.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = "{{问题描述}}"
    r2.font.size = Pt(14)
    r2.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    frame = slide.shapes.add_table(
        2, 3, Emu(600000), Emu(3600000), Emu(6500000), Emu(1500000)
    )
    table = frame.table
    for c, name in enumerate(["单位", "问题描述", "备注"]):
        table.cell(0, c).text = name
        for para in table.cell(0, c).text_frame.paragraphs:
            for r in para.runs:
                r.font.size = Pt(12)
                r.font.bold = True
    for c in range(3):
        table.cell(1, c).text = f"占位{c + 1}"

    pic = tmp_path / "tpl.png"
    Image.new("RGB", (320, 240), (200, 210, 230)).save(pic)
    slide.shapes.add_picture(str(pic), Emu(8000000), Emu(1300000), Emu(3400000), Emu(2100000))

    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("系列1", (1, 2, 3))
    slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Emu(8000000),
        Emu(3600000),
        Emu(3400000),
        Emu(2500000),
        chart_data,
    )

    # 备注页：克隆时必须跳过，否则会经反向 slide 引用把模板页拖进输出包
    slide.notes_slide.notes_text_frame.text = "模板备注"

    path = tmp_path / "template.pptx"
    prs.save(str(path))
    return str(path)


@pytest.fixture()
def excel_path(tmp_path, picture_file) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "检查表"
    ws.append(["项目名称", "单位", "问题描述", "问题图片", "备注", "数量"])
    ws.append(["项目A", "一公司", "法兰偏差无法安装", picture_file, "限期整改", 1])
    ws.append(["项目B", "二公司", "防火涂料未报验", picture_file, "复查", 2])
    ws.append(["项目C", "三公司", "支架间距超标", picture_file, "", 3])
    path = tmp_path / "data.xlsx"
    wb.save(str(path))
    return str(path)


@pytest.fixture()
def profile(template_path, excel_path):
    info = scan_template(template_path)
    from excel2ppt_studio.core.excel_reader import read_excel

    sheet = read_excel(excel_path, "检查表", 1)
    prof, _ = build_profile(template_path, excel_path, info, sheet.sheet, sheet.columns)
    return prof, info, sheet
