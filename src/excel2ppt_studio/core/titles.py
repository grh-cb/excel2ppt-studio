"""语义短标题生成：把长问题描述压成 ≤16 字标题。

只依赖标准库（urllib），兼容 DeepSeek / 任意 OpenAI Chat Completions 端点。
无 API key 时可先用 ``--from-file`` 走人工/AI 预生成的 JSON。
"""

from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request

DEFAULT_BASE_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-chat"

_SYSTEM = (
    "你是工程质量问题标题编辑。把每条问题压缩成不超过 16 个汉字的短标题。\n"
    "规则：\n"
    "1. 结构=对象+问题，一眼看出「哪里、怎么了」；\n"
    "2. 不含时间、责任单位、规范条文号、过程描述；\n"
    "3. 不以句号/逗号/冒号结尾；同类问题句式保持一致；\n"
    "4. 禁止截断原文造成语义残缺。\n"
    '只输出 JSON 对象：{"<序号>": "<标题>"}，不要解释。'
)


class TitleError(Exception):
    """标题生成失败（可读中文信息）。"""


def _chat(messages: list[dict], api_key: str, base_url: str, model: str, timeout: int = 120) -> str:
    url = base_url.rstrip("/") + "/chat/completions"
    payload = json.dumps(
        {"model": model, "messages": messages, "temperature": 0.2, "stream": False},
        ensure_ascii=False,
    ).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=payload,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", "replace")[:300]
        raise TitleError(f"标题接口返回 {exc.code}：{body}") from exc
    except (urllib.error.URLError, TimeoutError) as exc:
        raise TitleError(f"标题接口不可达：{exc}") from exc
    try:
        return data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise TitleError(f"标题接口响应格式异常：{str(data)[:200]}") from exc


def _parse_json_object(text: str) -> dict[str, str]:
    text = text.strip()
    match = re.search(r"\{.*\}", text, re.S)
    if not match:
        raise TitleError("模型没有返回 JSON 对象。")
    try:
        data = json.loads(match.group(0))
    except json.JSONDecodeError as exc:
        raise TitleError(f"模型返回的 JSON 无法解析：{exc}") from exc
    if not isinstance(data, dict):
        raise TitleError("模型返回的不是 JSON 对象。")
    return {str(k): str(v).strip() for k, v in data.items() if str(v).strip()}


def generate_titles(
    items: list[tuple[str, str]],
    api_key: str,
    base_url: str = DEFAULT_BASE_URL,
    model: str = DEFAULT_MODEL,
    batch_size: int = 20,
) -> dict[str, str]:
    """``items`` 为 ``(序号, 问题描述)`` 列表，返回 ``{序号: 标题}``。"""
    if not items:
        return {}
    if not api_key:
        raise TitleError("缺少 API key：用 --api-key 或环境变量 DEEPSEEK_API_KEY 提供。")
    result: dict[str, str] = {}
    for start in range(0, len(items), batch_size):
        chunk = items[start : start + batch_size]
        lines = [f"{key}. {text}" for key, text in chunk]
        messages = [
            {"role": "system", "content": _SYSTEM},
            {"role": "user", "content": "\n".join(lines)},
        ]
        content = _chat(messages, api_key, base_url, model)
        result.update(_parse_json_object(content))
    return result


def load_titles(path: str) -> dict[str, str]:
    if not os.path.isfile(path):
        raise TitleError(f"标题文件不存在：{path}")
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise TitleError("标题文件必须是 {序号: 标题} 的 JSON 对象。")
    return {str(k): str(v) for k, v in data.items()}


def resolve_api_key(explicit: str | None) -> str:
    return (explicit or os.environ.get("DEEPSEEK_API_KEY") or "").strip()
