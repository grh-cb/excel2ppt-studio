"""文本溢出估算：只告警，不修改字号。

按模板字号与文本框内边距估算行数；字符宽度用「中文/全角 1em、ASCII 0.55em」
的保守近似。结果用于预览标红与生成报告，不参与排版。
"""

from __future__ import annotations

import math
from dataclasses import dataclass

from .model import Field

EMU_PER_PT = 12_700
LINE_SPACING = 1.2
PAD_X_EMU = 91_440  # 0.1 inch
PAD_Y_EMU = 45_720  # 0.05 inch
DEFAULT_FONT_PT = 18.0


@dataclass
class FitResult:
    lines: int
    needed_emu: int
    available_emu: int

    @property
    def overflow(self) -> bool:
        return self.needed_emu > self.available_emu


def _visual_width_em(text: str) -> float:
    width = 0.0
    for ch in text:
        width += 1.0 if ord(ch) > 0x7F else 0.55
    return width


def _max_font_pt(field: Field) -> float:
    sizes = [r.size_pt for p in field.paragraphs for r in p.runs if r.size_pt]
    return max(sizes) if sizes else DEFAULT_FONT_PT


def estimate(field: Field, value: str) -> FitResult | None:
    """估算一个文本值填入该元素后的行数与所需高度。"""
    if field.kind != "text" or not value:
        return None
    font_pt = _max_font_pt(field)
    capacity_emu = max(1, field.rect.w - 2 * PAD_X_EMU)
    capacity_em = capacity_emu / (font_pt * EMU_PER_PT)
    if capacity_em <= 0:
        return None

    lines = 0
    for segment in str(value).split("\n"):
        width_em = _visual_width_em(segment)
        lines += max(1, math.ceil(width_em / capacity_em))
    available = max(1, field.rect.h - 2 * PAD_Y_EMU)
    needed = int(lines * font_pt * LINE_SPACING * EMU_PER_PT)
    return FitResult(lines=lines, needed_emu=needed, available_emu=available)
