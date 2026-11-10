"""模板体检：扫描结构、可映射性与格式一致性，给出分数与可操作建议。"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any

from .model import Field
from .slot_detect import find_cross_run_slots
from .template_scanner import scan_template


@dataclass
class HealthIssue:
    level: str  # error | warning | info
    code: str
    message: str
    field_id: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "level": self.level,
            "code": self.code,
            "message": self.message,
            "field_id": self.field_id,
        }


@dataclass
class HealthReport:
    template: str
    slide_index: int
    score: int
    issues: list[HealthIssue] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)

    @property
    def errors(self) -> list[HealthIssue]:
        return [i for i in self.issues if i.level == "error"]

    @property
    def warnings(self) -> list[HealthIssue]:
        return [i for i in self.issues if i.level == "warning"]

    def to_dict(self) -> dict[str, Any]:
        return {
            "template": self.template,
            "slide_index": self.slide_index,
            "score": self.score,
            "errors": len(self.errors),
            "warnings": len(self.warnings),
            "issues": [i.to_dict() for i in self.issues],
            "stats": self.stats,
        }


def _score(issues: list[HealthIssue]) -> int:
    penalty = sum(15 if i.level == "error" else 5 if i.level == "warning" else 1 for i in issues)
    return max(0, 100 - penalty)


def _fonts_of(fields: list[Field]) -> tuple[set[float], set[str]]:
    sizes: set[float] = set()
    names: set[str] = set()
    for fld in fields:
        for para in fld.paragraphs:
            for run in para.runs:
                if run.size_pt:
                    sizes.add(float(run.size_pt))
                if run.font:
                    names.add(run.font)
    return sizes, names


def _autofit_shapes(template: str, slide_index: int) -> list[str]:
    """返回开启了自动缩放/自动调整的顶层形状名。"""
    from pptx import Presentation

    hits: list[str] = []
    prs = Presentation(template)
    slide = prs.slides[slide_index]
    auto_tags = {
        "{http://schemas.openxmlformats.org/drawingml/2006/main}normAutofit",
        "{http://schemas.openxmlformats.org/drawingml/2006/main}spAutoFit",
    }
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        body = shape.text_frame._txBody.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr"
        )
        if body is None:
            continue
        if any(child.tag in auto_tags for child in body):
            hits.append(shape.name)
    return hits


def check_template(
    template: str, slide_index: int = 0, excel: str | None = None
) -> HealthReport:
    """扫描模板并返回体检报告；``excel`` 可选，用于交叉检查图片列/合并单元格。"""
    info = scan_template(template, slide_index)
    issues: list[HealthIssue] = []

    text_fields = [f for f in info.fields if f.kind == "text"]
    image_fields = [f for f in info.fields if f.kind == "image"]
    table_fields = [f for f in info.fields if f.kind == "table"]
    chart_fields = [f for f in info.fields if f.kind == "chart"]
    slot_total = sum(len(f.slots) for f in text_fields)

    def _real_slots(fld: Field) -> list[Any]:
        """显式/标签/命名/替代文字才算「真槽位」；段落与空白是兜底。"""
        return [s for s in fld.slots if s.kind in ("explicit", "label", "named", "alt")]

    zero_slot = [f for f in text_fields if not f.slots and f.text.strip()]
    for fld in zero_slot:
        issues.append(
            HealthIssue(
                "warning",
                "no_slots",
                f"「{fld.name}」没有任何可识别槽位，只能整体替换文本",
                fld.id,
            )
        )
    only_paragraph = [f for f in text_fields if f.slots and not _real_slots(f)]
    for fld in only_paragraph:
        issues.append(
            HealthIssue(
                "info",
                "paragraph_only",
                f"「{fld.name}」靠段落槽位映射（第N段），多段格式可能丢失",
                fld.id,
            )
        )

    for fld in text_fields:
        for para in fld.paragraphs:
            text = "".join(r.text or "" for r in para.runs)
            cross = [s for s in find_cross_run_slots(text) if "{{" in s or "${" in s]
            if cross and not fld.slots:
                issues.append(
                    HealthIssue(
                        "warning",
                        "cross_run_slot",
                        f"「{fld.name}」的 {'、'.join(cross[:3])} 被拆到多个 run，"
                        "无法识别为显式槽位",
                        fld.id,
                    )
                )
                break

    sizes, fonts = _fonts_of(text_fields)
    if len(sizes) > 3:
        issues.append(
            HealthIssue(
                "warning",
                "font_sizes",
                f"同一页出现 {len(sizes)} 种字号（{'/'.join(str(int(s)) if s == int(s) else str(s) for s in sorted(sizes))}），"
                "建议控制在 3 种以内",
            )
        )
    if len(fonts) > 2:
        issues.append(
            HealthIssue(
                "warning", "fonts", f"同一页出现 {len(fonts)} 种字体：{'、'.join(sorted(fonts))}"
            )
        )

    auto = _autofit_shapes(template, slide_index)
    if auto:
        issues.append(
            HealthIssue(
                "warning",
                "autofit",
                f"{len(auto)} 个文本框开启了自动调整（{'、'.join(auto[:3])}），"
                "长文本会被 WPS/PowerPoint 缩字",
            )
        )

    if image_fields and excel:
        try:
            from .excel_reader import read_dispimg_formulas, read_excel

            sheet = read_excel(excel)
            image_cols = [c for c in sheet.columns if c.dtype == "image"]
            if not image_cols:
                # data_only 读不到 DISPIMG 公式，改用公式模式确认（自动找含图片的工作表）
                formulas = read_dispimg_formulas(excel)
                if formulas:
                    image_cols = [None] * len({col for (_row, col) in formulas})
            if len(image_cols) < len(image_fields):
                issues.append(
                    HealthIssue(
                        "error",
                        "image_columns",
                        f"模板有 {len(image_fields)} 个图片位，Excel 只有 {len(image_cols)} 个图片列，"
                        "多出的图片位会留空",
                    )
                )
        except Exception:  # noqa: BLE001 - 交叉检查失败不影响主报告
            pass

    if not slot_total and not any(f.kind in ("table", "image", "chart") for f in info.fields):
        issues.append(
            HealthIssue("error", "no_bindable", "模板里没有可映射的元素（文本/表格/图片/图表都没有）")
        )

    stats = {
        "fields": len(info.fields),
        "text": len(text_fields),
        "table": len(table_fields),
        "image": len(image_fields),
        "chart": len(chart_fields),
        "slots": slot_total,
        "font_sizes": len(sizes),
        "fonts": len(fonts),
        "slide_count": info.slide_count,
    }
    if not issues:
        issues.append(HealthIssue("info", "ok", "结构完整，可直接映射"))

    return HealthReport(
        template=os.path.abspath(template),
        slide_index=slide_index,
        score=_score(issues),
        issues=issues,
        stats=stats,
    )
