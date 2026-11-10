"""映射建议与 Profile 存取。

自动建议只做「有把握」的匹配：
- 显式槽位名 ↔ 列名（归一化后相似度 ≥ 阈值）
- 图片元素 ↔ 图片类型列
- 表格 ↔ 表头匹配的列
其余留给人工在 GUI 中确认。
"""

from __future__ import annotations

import json
import os
from difflib import SequenceMatcher
from typing import Any

from .model import Binding, Column, Profile
from .slot_detect import normalize_name
from .template_scanner import TemplateInfo

DEFAULT_THRESHOLD = 60


class ProfileError(Exception):
    """Profile 读写或结构错误。"""


def similarity(a: str, b: str) -> int:
    """0–100 的相似度：完全相等 100，包含 80，否则字符序列比例。"""
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return 0
    if na == nb:
        return 100
    if na in nb or nb in na:
        return 80
    return int(SequenceMatcher(None, na, nb).ratio() * 100)


def _best_column(name: str, columns: list[Column]) -> tuple[Column | None, int]:
    best: Column | None = None
    best_score = 0
    for col in columns:
        score = similarity(name, col.name)
        if score > best_score:
            best, best_score = col, score
    return best, best_score


def suggest_bindings(
    info: TemplateInfo,
    columns: list[Column],
    threshold: int = DEFAULT_THRESHOLD,
) -> tuple[list[Binding], list[str]]:
    """按槽位名/列类型生成初始绑定与说明。"""
    bindings: list[Binding] = []
    notes: list[str] = []

    for field in info.fields:
        if field.kind == "text":
            if not field.slots:
                continue
            for slot in field.slots:
                if not slot.name:
                    continue
                col, score = _best_column(slot.name, columns)
                if col is None or score < threshold:
                    notes.append(
                        f"未自动匹配：元素 {field.id} 的槽位「{slot.name}」（最佳候选 "
                        f"{col.name if col else '无'}，相似度 {score}）"
                    )
                    continue
                bindings.append(
                    Binding(
                        field_id=field.id,
                        kind="text",
                        source={"column": col.name, "slot": slot.name},
                    )
                )
        elif field.kind == "image":
            image_cols = [c for c in columns if c.dtype == "image"]
            if image_cols:
                bindings.append(
                    Binding(
                        field_id=field.id,
                        kind="image",
                        source={"column": image_cols[0].name},
                        options={"fit": "contain"},
                    )
                )
            else:
                notes.append(f"未自动匹配：图片元素 {field.id}（Excel 没有图片类型列）")
        elif field.kind == "table":
            assert field.table is not None
            if not field.table.first_row_header:
                notes.append(f"表格 {field.id} 首行不像表头，需人工确认列顺序")
                continue
            header = field.table.cells[0] if field.table.cells else []
            chosen: list[str] = []
            for name in header:
                col, score = _best_column(name, columns)
                if col is None or score < threshold:
                    break
                chosen.append(col.name)
            if len(chosen) == len(header) and chosen:
                options: dict[str, Any] = {"header": True}
                if any("{{" in c or "${" in c for c in header):
                    # 模板表头本身是占位符时，用 Excel 列名覆盖，避免输出残留 {{...}}
                    options["header_names"] = chosen
                bindings.append(
                    Binding(
                        field_id=field.id,
                        kind="table",
                        source={"mode": "rows", "columns": chosen},
                        options=options,
                    )
                )
            else:
                notes.append(f"未自动匹配：表格 {field.id}（表头与 Excel 列名不一致）")
    return bindings, notes


def build_profile(
    template: str,
    excel: str,
    info: TemplateInfo,
    sheet_name: str,
    columns: list[Column],
    threshold: int = DEFAULT_THRESHOLD,
    header_row: int = 1,
) -> tuple[Profile, list[str]]:
    bindings, notes = suggest_bindings(info, columns, threshold)
    profile = Profile(
        template=os.path.abspath(template),
        slide_index=info.slide_index,
        excel=os.path.abspath(excel),
        sheet=sheet_name,
        header_row=header_row,
        bindings=bindings,
    )
    return profile, notes


def save_profile(profile: Profile, path: str) -> None:
    data = profile.to_dict()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")


def load_profile(path: str) -> Profile:
    if not os.path.isfile(path):
        raise ProfileError(f"Profile 不存在：{path}")
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as exc:
        raise ProfileError(f"Profile 不是合法 JSON：{path}（{exc}）") from exc
    if not isinstance(data, dict):
        raise ProfileError("Profile 顶层必须是 JSON 对象。")
    profile = Profile.from_dict(data)
    if not profile.bindings:
        raise ProfileError("Profile 里没有任何 bindings，无法生成。")
    return profile


def binding_summary(profile: Profile) -> list[dict[str, Any]]:
    return [
        {
            "field": b.field_id,
            "kind": b.kind,
            "column": b.source.get("column") or b.source.get("columns"),
            "options": b.options,
        }
        for b in profile.bindings
    ]
