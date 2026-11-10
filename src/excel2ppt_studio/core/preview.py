"""渲染预览：调用 WPS / PowerPoint 导出单页 PNG。

COM 调用放在独立子进程（``preview_worker``）里执行：
- 隔离 COM 崩溃，主服务不受影响；
- 超时可控；
- 没有 Office/WPS 时明确报错，由 UI 降级为几何预览。
"""

from __future__ import annotations

import json
import os
import subprocess
import sys


class PreviewError(Exception):
    """预览导出失败（可读中文信息）。"""


def _progid_registered(progid: str) -> bool:
    try:
        import winreg
    except ImportError:  # 非 Windows
        return False
    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid)
        return True
    except OSError:
        return False


def available_backend() -> str | None:
    """返回可用的预览后端：``wps`` / ``powerpoint`` / ``None``。"""
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return None
    if _progid_registered("Kwpp.Application"):
        return "wps"
    if _progid_registered("PowerPoint.Application"):
        return "powerpoint"
    return None


def _kill_process_tree(pid: int) -> None:
    """结束 worker 及其通过 COM 拉起的 WPS/PowerPoint 子进程树。"""
    if not pid:
        return
    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            capture_output=True,
            timeout=15,
        )
    except (OSError, subprocess.SubprocessError):
        pass


def export_png(
    pptx_path: str,
    page: int,
    out_path: str,
    width: int = 1600,
    height: int = 900,
    timeout: int = 120,
) -> str:
    """把第 ``page`` 页（1 起）导出为 PNG，返回文件路径。"""
    if not os.path.isfile(pptx_path):
        raise PreviewError(f"文件不存在：{pptx_path}")
    backend = available_backend()
    if backend is None:
        raise PreviewError("未检测到 WPS 或 PowerPoint，无法渲染预览（已降级为几何预览）。")
    progid = "Kwpp.Application" if backend == "wps" else "PowerPoint.Application"

    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    cmd = [
        sys.executable,
        "-m",
        "excel2ppt_studio.core.preview_worker",
        "--pptx",
        os.path.abspath(pptx_path),
        "--page",
        str(page),
        "--out",
        out_path,
        "--progid",
        progid,
        "--width",
        str(width),
        "--height",
        str(height),
    ]
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    try:
        stdout, stderr = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        _kill_process_tree(proc.pid)
        try:
            proc.communicate(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
        raise PreviewError(
            f"预览超时（>{timeout}s），已结束相关进程，请重试。"
        ) from exc

    payload = {}
    for line in reversed((stdout or "").strip().splitlines()):
        try:
            payload = json.loads(line)
            break
        except json.JSONDecodeError:
            continue

    if proc.returncode != 0 or not payload.get("ok"):
        reason = payload.get("error") or (stderr or "").strip() or f"退出码 {proc.returncode}"
        raise PreviewError(f"预览导出失败：{reason}")
    if not os.path.isfile(out_path) or os.path.getsize(out_path) == 0:
        raise PreviewError("预览导出失败：输出文件为空。")
    return out_path
