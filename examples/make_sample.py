"""生成示例模板与 Excel，供快速试用与 README 演示。

用法：
    python examples/make_sample.py
产物：
    examples/sample/template.pptx
    examples/sample/sample_data.xlsx
    examples/sample/example_photo.png
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu, Pt

OUT = Path(__file__).resolve().parent / "sample"


def make_photo(path: Path) -> None:
    im = Image.new("RGB", (480, 320), (232, 240, 250))
    draw = ImageDraw.Draw(im)
    draw.rectangle([12, 12, 468, 308], outline=(30, 80, 160), width=6)
    draw.text((28, 28), "EXAMPLE PHOTO", fill=(190, 30, 30))
    im.save(path)


def make_template(path: Path, photo: Path) -> None:
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
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = "{{问题描述}}"
    r2.font.size = Pt(14)

    frame = slide.shapes.add_table(2, 3, Emu(600000), Emu(3600000), Emu(6500000), Emu(1500000))
    table = frame.table
    for c, name in enumerate(["单位", "问题描述", "备注"]):
        table.cell(0, c).text = name
        for para in table.cell(0, c).text_frame.paragraphs:
            for r in para.runs:
                r.font.size = Pt(12)
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        table.cell(0, c).fill.solid()
        table.cell(0, c).fill.fore_color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    for c in range(3):
        table.cell(1, c).text = f"示例{c + 1}"

    slide.shapes.add_picture(str(photo), Emu(8000000), Emu(1300000), Emu(3400000), Emu(2100000))

    data = CategoryChartData()
    data.categories = ["A", "B", "C"]
    data.add_series("示例系列", (1, 2, 3))
    slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Emu(8000000),
        Emu(3600000),
        Emu(3400000),
        Emu(2500000),
        data,
    )

    prs.save(str(path))


def make_excel(path: Path, photo: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "检查表"
    ws.append(["项目名称", "单位", "问题描述", "备注", "问题图片", "数量"])
    ws.append(["项目A", "一公司", "法兰偏差无法安装，需返工", "限期整改", str(photo), 1])
    ws.append(["项目B", "二公司", "防火涂料未按规定报验", "复查", str(photo), 2])
    ws.append(["项目C", "三公司", "支架间距超标", "", str(photo), 3])
    wb.save(str(path))


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    photo = OUT / "example_photo.png"
    make_photo(photo)
    make_template(OUT / "template.pptx", photo)
    make_excel(OUT / "sample_data.xlsx", photo)
    print(f"已生成示例到：{OUT}")
    print("下一步：python -m excel2ppt_studio web")


if __name__ == "__main__":
    main()
