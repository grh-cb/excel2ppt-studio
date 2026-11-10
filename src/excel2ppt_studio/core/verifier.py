"""格式指纹校验：把「零漂移」从目视变成可机检断言。

对模板页与输出页逐 shape 比对：
- shape 集合（按 cNvPr/@id）
- 位置与尺寸（EMU 精确相等）
- 字体签名集合（字号 / 加粗 / 斜体 / 颜色）
- 残留占位符（``{{...}}`` / ``${...}``）

文本内容本身不参与比较——那正是要替换的部分。
"""

from __future__ import annotations

import hashlib
import os
import re
from collections.abc import Iterable
from dataclasses import dataclass, field

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

_RESIDUAL = re.compile(r"\{\{|\$\{")
_SHAPE_ID = qn("p:cNvPr")


class VerifyError(Exception):
    """校验无法执行（文件不可读等）。"""


@dataclass
class ShapeFingerprint:
    shape_id: int
    tag: str
    rect: tuple[int, int, int, int] | None
    fonts: tuple[str, ...]
    residual: tuple[str, ...]

    @property
    def digest(self) -> str:
        payload = f"{self.tag}|{self.rect}|{','.join(self.fonts)}"
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


@dataclass
class VerifyIssue:
    level: str  # error | warning | info
    code: str
    message: str
    shape_id: int | None = None
    page: int | None = None

    def to_dict(self) -> dict:
        return {
            "level": self.level,
            "code": self.code,
            "message": self.message,
            "shape_id": self.shape_id,
            "page": self.page,
        }


@dataclass
class VerifyResult:
    ok: bool
    pages_checked: int
    shapes_checked: int
    issues: list[VerifyIssue] = field(default_factory=list)

    @property
    def errors(self) -> list[VerifyIssue]:
        return [i for i in self.issues if i.level == "error"]

    @property
    def warnings(self) -> list[VerifyIssue]:
        return [i for i in self.issues if i.level == "warning"]

    def to_dict(self) -> dict:
        return {
            "ok": self.ok,
            "pages_checked": self.pages_checked,
            "shapes_checked": self.shapes_checked,
            "issues": [i.to_dict() for i in self.issues],
        }


def _xfrm_of(el) -> tuple[int, int, int, int] | None:
    for path in (
        f"{qn('p:spPr')}/{qn('a:xfrm')}",
        f"{qn('p:grpSpPr')}/{qn('a:xfrm')}",
        qn("p:xfrm"),
    ):
        xfrm = el.find(path)
        if xfrm is None:
            continue
        off = xfrm.find(qn("a:off"))
        ext = xfrm.find(qn("a:ext"))
        if off is None or ext is None:
            continue
        try:
            return (
                int(off.get("x", 0)),
                int(off.get("y", 0)),
                int(ext.get("cx", 0)),
                int(ext.get("cy", 0)),
            )
        except (TypeError, ValueError):
            return None
    return None


def _color_of(rpr) -> str:
    fill = rpr.find(qn("a:solidFill"))
    if fill is None:
        return ""
    srgb = fill.find(qn("a:srgbClr"))
    if srgb is not None:
        return "srgb:" + (srgb.get("val") or "")
    scheme = fill.find(qn("a:schemeClr"))
    if scheme is not None:
        return "scheme:" + (scheme.get("val") or "")
    return ""


def _font_sig(el) -> tuple[str, ...]:
    sigs = set()
    for rpr in el.iter(qn("a:rPr")):
        sigs.add(
            "|".join(
                [
                    rpr.get("sz") or "",
                    rpr.get("b") or "",
                    rpr.get("i") or "",
                    _color_of(rpr),
                ]
            )
        )
    return tuple(sorted(sigs))


def _residual_text(el) -> tuple[str, ...]:
    hits = []
    for t in el.iter(qn("a:t")):
        text = t.text or ""
        if _RESIDUAL.search(text):
            hits.append(text.strip())
    return tuple(hits)


def _shape_id(el) -> int | None:
    c_nv_pr = el.find(".//" + _SHAPE_ID)
    if c_nv_pr is None:
        return None
    try:
        return int(c_nv_pr.get("id", ""))
    except (TypeError, ValueError):
        return None


def _top_level_shapes(slide) -> Iterable:
    sp_tree = slide.shapes._spTree  # noqa: SLF001 - 官方无公开 API
    for el in sp_tree:
        tag = etree.QName(el).localname
        if tag in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp", "contentPart"):
            yield el


def fingerprint_slide(slide) -> dict[int, ShapeFingerprint]:
    result: dict[int, ShapeFingerprint] = {}
    for el in _top_level_shapes(slide):
        sid = _shape_id(el)
        if sid is None:
            continue
        result[sid] = ShapeFingerprint(
            shape_id=sid,
            tag=etree.QName(el).localname,
            rect=_xfrm_of(el),
            fonts=_font_sig(el),
            residual=_residual_text(el),
        )
    return result


def _compare_page(
    base: dict[int, ShapeFingerprint],
    current: dict[int, ShapeFingerprint],
    page: int,
    allow_image_resize: bool,
) -> list[VerifyIssue]:
    issues: list[VerifyIssue] = []
    base_ids = set(base)
    cur_ids = set(current)

    for sid in sorted(base_ids - cur_ids):
        level = "warning" if (allow_image_resize and base[sid].tag == "pic") else "error"
        issues.append(
            VerifyIssue(
                level,
                "missing_shape",
                f"输出页缺少模板元素 id={sid}（图片位可能因空值被隐藏）",
                sid,
                page,
            )
        )
    for sid in sorted(cur_ids - base_ids):
        issues.append(
            VerifyIssue("error", "extra_shape", f"输出页多出元素 id={sid}", sid, page)
        )

    for sid in sorted(base_ids & cur_ids):
        b, c = base[sid], current[sid]
        if b.tag != c.tag:
            issues.append(
                VerifyIssue("error", "tag_changed", f"元素类型变化：{b.tag} → {c.tag}", sid, page)
            )
            continue
        if b.rect != c.rect:
            level = "warning" if (allow_image_resize and b.tag == "pic") else "error"
            issues.append(
                VerifyIssue(
                    level,
                    "geometry_changed",
                    f"位置/尺寸变化：{b.rect} → {c.rect}",
                    sid,
                    page,
                )
            )
        if b.fonts != c.fonts:
            issues.append(
                VerifyIssue(
                    "error",
                    "font_changed",
                    "字体签名变化：" + _font_diff(b.fonts, c.fonts),
                    sid,
                    page,
                )
            )
        if c.residual:
            issues.append(
                VerifyIssue(
                    "warning",
                    "residual_slot",
                    "残留未替换占位符：" + "、".join(c.residual[:3]),
                    sid,
                    page,
                )
            )
    return issues


def _font_diff(before: tuple[str, ...], after: tuple[str, ...]) -> str:
    removed = [x for x in before if x not in after]
    added = [x for x in after if x not in before]
    parts = []
    if removed:
        parts.append("移除 " + "；".join(removed[:2]))
    if added:
        parts.append("新增 " + "；".join(added[:2]))
    return "，".join(parts) or "顺序变化"


def verify_file(
    template: str,
    output: str,
    slide_index: int = 0,
    allow_image_resize: bool = True,
    max_pages: int = 0,
) -> VerifyResult:
    """逐页比对模板母版页与输出页。"""
    for label, path in (("模板", template), ("输出", output)):
        if not os.path.isfile(path):
            raise VerifyError(f"{label}文件不存在：{path}")
    try:
        tpl = Presentation(template)
    except Exception as exc:  # noqa: BLE001
        raise VerifyError(f"无法打开模板：{exc}") from exc
    if slide_index >= len(tpl.slides):
        raise VerifyError(f"模板页索引越界：{slide_index}（共 {len(tpl.slides)} 页）")
    try:
        out = Presentation(output)
    except Exception as exc:  # noqa: BLE001
        raise VerifyError(f"无法打开输出文件：{exc}") from exc
    if len(out.slides) == 0:
        raise VerifyError("输出文件没有任何幻灯片。")

    base = fingerprint_slide(tpl.slides[slide_index])
    limit = len(out.slides) if max_pages <= 0 else min(max_pages, len(out.slides))
    issues: list[VerifyIssue] = []
    shapes_checked = 0
    for page in range(limit):
        current = fingerprint_slide(out.slides[page])
        shapes_checked += len(current)
        issues.extend(_compare_page(base, current, page, allow_image_resize))

    return VerifyResult(
        ok=not any(i.level == "error" for i in issues),
        pages_checked=limit,
        shapes_checked=shapes_checked,
        issues=issues,
    )
