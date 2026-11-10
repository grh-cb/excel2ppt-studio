"""槽位探测：从文本与形状元数据识别「可填充位置」。

优先级：显式占位符 > 命名形状 > 替代文字 > 标签模式 > 空白/下划线。

只识别完整落在单个 run 内的槽位；跨 run 的槽位需要拆分 run，
会破坏原有格式，故不处理（报告为 ``cross_run`` 提示）。
"""

from __future__ import annotations

import re

from .model import Paragraph, Slot

# {{x}} / ${x} / %x%（避免误伤 100% 与 %Y-%m-%d 这类格式串）
_EXPLICIT = [
    re.compile(r"\{\{\s*([^{}\n]{1,64}?)\s*\}\}"),
    re.compile(r"\$\{\s*([^{}\n]{1,64}?)\s*\}"),
    re.compile(r"(?<![\w%])%([\u4e00-\u9fa5A-Za-z0-9_·\-/ ]{1,32})%(?![\w%])"),
]

_LABEL = re.compile(
    r"^\s*([\u4e00-\u9fa5A-Za-z0-9（）()·、\-/ ]{2,20}?)\s*[：:]\s*[_\s　]*$"
)
_BLANK = re.compile(r"^[_\s　]{3,}$")
_NAME_PREFIX = re.compile(r"^(?:fld|field|slot)[_-](.+)$", re.IGNORECASE)


def normalize_name(name: str) -> str:
    """统一槽位名，便于与 Excel 列名匹配。"""
    return re.sub(r"[\s\u3000]+", "", name).strip("：: ")


def scan_paragraphs(paragraphs: list[Paragraph], field_id: str) -> list[Slot]:
    """扫描段落列表，返回槽位。

    优先级：显式占位符 > 标签（仅单 run 段落）> 空白段落 > 段落槽位。

    没有占位符的模板（例如一页已填好的样例）靠「段落槽位」映射：
    每个有文字的段落生成一个 ``第N段`` 槽位，渲染时只替换该段，
    其余段落保持原样。
    """
    slots: list[Slot] = []
    for p in paragraphs:
        if not p.runs:
            continue
        explicit: list[Slot] = []
        for r_index, run in enumerate(p.runs):
            text = run.text
            if not text:
                continue
            for pattern in _EXPLICIT:
                for m in pattern.finditer(text):
                    explicit.append(
                        Slot(
                            name=normalize_name(m.group(1)),
                            raw=m.group(0),
                            field_id=field_id,
                            para=p.index,
                            run=r_index,
                            kind="explicit",
                        )
                    )
        if explicit:
            slots.extend(explicit)
            continue

        para_text = "".join(run.text or "" for run in p.runs)
        if len(p.runs) == 1:
            text = p.runs[0].text or ""
            label = _LABEL.match(text)
            if label and not _BLANK.match(text):
                slots.append(
                    Slot(
                        name=normalize_name(label.group(1)),
                        raw=text,
                        field_id=field_id,
                        para=p.index,
                        run=0,
                        kind="label",
                    )
                )
                continue

        if _BLANK.match(para_text):
            slots.append(
                Slot(
                    name="",
                    raw=para_text,
                    field_id=field_id,
                    para=p.index,
                    run=0,
                    kind="blank",
                )
            )
            continue

        if para_text.strip():
            slots.append(
                Slot(
                    name=f"第{p.index + 1}段",
                    raw=para_text.strip(),
                    field_id=field_id,
                    para=p.index,
                    run=0,
                    kind="paragraph",
                )
            )
    return slots


def scan_shape_meta(name: str, descr: str, field_id: str, text: str) -> list[Slot]:
    """从形状名 / 替代文字推断槽位（整形状一个槽位）。"""
    for candidate in (name, descr):
        if not candidate:
            continue
        m = _NAME_PREFIX.match(candidate.strip())
        if m:
            return [
                Slot(
                    name=normalize_name(m.group(1)),
                    raw=candidate,
                    field_id=field_id,
                    para=0,
                    run=0,
                    kind="named",
                )
            ]
        for pattern in _EXPLICIT[:2]:
            hit = pattern.search(candidate)
            if hit:
                return [
                    Slot(
                        name=normalize_name(hit.group(1)),
                        raw=hit.group(0),
                        field_id=field_id,
                        para=0,
                        run=0,
                        kind="alt",
                    )
                ]
    return []


def find_cross_run_slots(text: str) -> list[str]:
    """检测被拆到多个 run 的显式槽位，供报告提示（不参与替换）。"""
    found = []
    for pattern in _EXPLICIT[:2]:
        for m in pattern.finditer(text):
            found.append(m.group(0))
    return found
