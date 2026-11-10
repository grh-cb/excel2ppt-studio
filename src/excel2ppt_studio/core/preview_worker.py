"""独立进程执行 WPS / PowerPoint 导出，避免 COM 崩溃拖垮主进程。

由 ``preview.export_png`` 通过 ``python -m`` 调用；结果以 JSON 打到 stdout。
"""

from __future__ import annotations

import argparse
import json
import os
import sys


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--page", type=int, required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--progid", required=True)
    parser.add_argument("--width", type=int, default=1600)
    parser.add_argument("--height", type=int, default=900)
    args = parser.parse_args()

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        print(json.dumps({"ok": False, "error": f"缺少 pywin32：{exc}"}, ensure_ascii=False))
        return 2

    out = os.path.abspath(args.out)
    if os.path.exists(out):
        os.remove(out)

    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        app = win32.DispatchEx(args.progid)
        pres = app.Presentations.Open(os.path.abspath(args.pptx), ReadOnly=True, WithWindow=False)
        count = pres.Slides.Count
        if args.page < 1 or args.page > count:
            print(
                json.dumps(
                    {"ok": False, "error": f"页码越界：{args.page}（共 {count} 页）"},
                    ensure_ascii=False,
                )
            )
            return 3
        pres.Slides(args.page).Export(out, "PNG", args.width, args.height)
    except Exception as exc:  # noqa: BLE001 - 统一转成 JSON 错误
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 4
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:  # noqa: BLE001
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:  # noqa: BLE001
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass

    if not os.path.isfile(out) or os.path.getsize(out) == 0:
        print(json.dumps({"ok": False, "error": "导出文件为空"}, ensure_ascii=False))
        return 5
    print(json.dumps({"ok": True, "path": out}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
