from __future__ import annotations

from excel2ppt_studio.core.model import Field, Paragraph, Rect, Run
from excel2ppt_studio.core.text_fit import estimate


def _field(w=3_000_000, h=1_000_000, size=18.0) -> Field:
    return Field(
        id="sp-2",
        kind="text",
        name="TextBox",
        role="body",
        rect=Rect(0, 0, w, h),
        z=0,
        text="",
        paragraphs=[Paragraph(index=0, runs=[Run(text="x", size_pt=size)])],
    )


def test_short_text_fits():
    field = _field(w=4_000_000, h=1_200_000, size=14.0)
    fit = estimate(field, "短文本")
    assert fit is not None
    assert fit.overflow is False


def test_long_text_overflows():
    field = _field(w=1_200_000, h=400_000, size=18.0)
    fit = estimate(field, "这是一段很长很长的文字，用来验证溢出估算是否会触发告警。")
    assert fit is not None
    assert fit.overflow is True
    assert fit.lines >= 2


def test_newlines_count_as_lines():
    field = _field(w=6_000_000, h=300_000, size=12.0)
    fit = estimate(field, "第一行\n第二行\n第三行")
    assert fit is not None
    assert fit.lines >= 3


def test_non_text_or_empty_returns_none():
    field = _field()
    field.kind = "image"
    assert estimate(field, "x") is None
    field.kind = "text"
    assert estimate(field, "") is None
