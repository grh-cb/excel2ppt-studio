"""命令行入口：``e2p inspect|suggest|build``。"""

from __future__ import annotations

import argparse
import json
import os
import sys
from collections.abc import Sequence

from .core.excel_images import ExcelImageError
from .core.excel_reader import ExcelDataError, read_excel
from .core.mapping import ProfileError, build_profile, load_profile, save_profile
from .core.model import Profile
from .core.renderer import RenderError, render
from .core.template_health import check_template
from .core.template_scanner import TemplateError, scan_template
from .core.verifier import VerifyError, verify_file


def _print_fields(info) -> None:
    print(f"模板：{info.path}")
    print(
        f"第 {info.slide_index} 页 / 共 {info.slide_count} 页  "
        f"尺寸 {info.width / 360000:.2f}×{info.height / 360000:.2f} cm"
    )
    if info.warnings:
        print("警告：")
        for w in info.warnings:
            print(f"  - {w}")
    print()
    header = f"{'ID':<12}{'类型':<8}{'角色':<10}{'位置(cm)':<26}{'内容 / 槽位'}"
    print(header)
    print("-" * 96)
    for f in info.fields:
        x, y, w, h = f.rect.cm
        pos = f"({x},{y}) {w}×{h}"
        if f.kind == "table" and f.table:
            content = f"表格 {f.table.rows}×{f.table.cols}"
        elif f.kind == "image":
            content = f"图片 {f.image.width_px}×{f.image.height_px}" if f.image else "图片"
        elif f.kind == "chart":
            content = "图表"
        else:
            content = (f.text or "").replace("\n", " / ")[:40]
        slots = f"  [槽位: {', '.join(f.slot_names)}]" if f.slot_names else ""
        print(f"{f.id:<12}{f.kind:<8}{f.role:<10}{pos:<26}{content}{slots}")
    print()
    print(f"共 {len(info.fields)} 个元素，其中带槽位 {sum(1 for f in info.fields if f.slots)} 个。")


def cmd_inspect(args: argparse.Namespace) -> int:
    info = scan_template(args.template, args.slide)
    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(info.to_dict(), f, ensure_ascii=False, indent=2)
            f.write("\n")
        print(f"已写出：{args.json}")
    else:
        _print_fields(info)
    return 0


def cmd_suggest(args: argparse.Namespace) -> int:
    info = scan_template(args.template, args.slide)
    sheet = read_excel(args.excel, args.sheet or None, args.header_row)
    profile, notes = build_profile(
        args.template,
        args.excel,
        info,
        sheet.sheet,
        sheet.columns,
        args.threshold,
        header_row=args.header_row,
    )
    print(f"Excel：{os.path.abspath(args.excel)} / 工作表「{sheet.sheet}」")
    print(f"数据行：{len(sheet.rows)}  列：{', '.join(sheet.column_names)}")
    print()
    total_slots = sum(len(f.slots) for f in info.fields if f.kind == "text")
    print(f"自动匹配 {len(profile.bindings)} 项（文本槽位共 {total_slots} 个）")
    for note in notes:
        print(f"  ! {note}")
    out = args.out or "profile.json"
    save_profile(profile, out)
    print(f"\nProfile 已保存：{os.path.abspath(out)}")
    print("请人工检查后执行：e2p build --profile <该文件> --out <输出.pptx>")
    return 0


def cmd_titles(args: argparse.Namespace) -> int:
    from .core.titles import TitleError, generate_titles, load_titles, resolve_api_key

    sheet = read_excel(args.excel, args.sheet or None, args.header_row)
    if args.column not in sheet.column_names:
        print(f"错误：列不存在：{args.column}。可用列：{'、'.join(sheet.column_names)}", file=sys.stderr)
        return 1

    def to_text(value) -> str:
        if value is None:
            return ""
        if hasattr(value, "isoformat"):
            return value.isoformat()
        return str(value).strip()

    items = []
    for row in sheet.rows[: args.limit or None]:
        text = to_text(row.get(args.column))
        if text:
            items.append((to_text(row.get(args.key)) or str(len(items) + 1), text))

    try:
        if args.from_file:
            titles = load_titles(args.from_file)
        else:
            titles = generate_titles(
                items,
                resolve_api_key(args.api_key),
                args.base_url,
                args.model,
            )
    except TitleError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1

    out = args.out or "titles.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(titles, f, ensure_ascii=False, indent=1)
        f.write("\n")
    print(f"已生成 {len(titles)} 条短标题 → {os.path.abspath(out)}")
    long_titles = [(k, v) for k, v in titles.items() if len(v) > 16]
    if long_titles:
        print(f"警告：{len(long_titles)} 条超过 16 字，建议人工压缩。")
    return 0


def cmd_build(args: argparse.Namespace) -> int:
    profile: Profile = load_profile(args.profile)
    excel = args.excel or profile.excel
    template = args.template or profile.template
    if not excel or not template:
        print("错误：Profile 缺少 Excel 或模板路径，请用 --excel/--template 指定。", file=sys.stderr)
        return 1

    def progress(done: int, total: int) -> None:
        if not args.quiet and not getattr(args, "json", False):
            print(f"\r渲染 {done}/{total}", end="", flush=True)

    result = render(template, profile, out=args.out, excel=excel, on_progress=progress)
    if args.json:
        print(
            json.dumps(
                {
                    "ok": True,
                    "output": result.output,
                    "rows": result.rows,
                    "slides": result.slides,
                    "replaced": result.replaced,
                    "warnings": result.warnings,
                },
                ensure_ascii=False,
            )
        )
        return 0
    if not args.quiet:
        print()
    print(f"完成：{result.rows} 行数据 → {result.slides} 页幻灯片，替换 {result.replaced} 处")
    if result.warnings:
        print("警告：")
        for w in result.warnings:
            print(f"  - {w}")
    print(f"输出：{result.output}")
    return 0


def cmd_web(args: argparse.Namespace) -> int:
    from .web.app import run

    run(host=args.host, port=args.port, open_browser=not args.no_browser)
    return 0


def cmd_verify(args: argparse.Namespace) -> int:
    result = verify_file(
        args.template,
        args.output,
        slide_index=args.slide,
        allow_image_resize=not args.strict_images,
        max_pages=args.max_pages,
    )
    if getattr(args, "json", False):
        data = result.to_dict()
        data["ok"] = result.ok
        print(json.dumps(data, ensure_ascii=False))
        return 0 if result.ok else 1
    print(f"校验：{result.pages_checked} 页 / {result.shapes_checked} 个元素")
    for issue in result.issues:
        page = f"（第 {issue.page + 1} 页）" if issue.page is not None else ""
        print(f"  [{issue.level}] {issue.message}{page}")
    if result.ok:
        print("结论：通过（无结构性差异）")
        return 0
    print("结论：失败（存在版式/格式差异）")
    return 1


def cmd_doctor(args: argparse.Namespace) -> int:
    report = check_template(args.template, args.slide, args.excel)
    if getattr(args, "json", False):
        print(json.dumps(report.to_dict(), ensure_ascii=False))
        return 0 if not report.errors else 1
    print(f"模板体检：{report.score} 分（{report.template}）")
    for issue in report.issues:
        tag = {"error": "错误", "warning": "警告", "info": "信息"}.get(issue.level, issue.level)
        ref = f" [{issue.field_id}]" if issue.field_id else ""
        print(f"  [{tag}] {issue.message}{ref}")
    return 0 if not report.errors else 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="e2p",
        description="Excel2PPT Studio：按单页 PPT 模板把 Excel 逐行生成多页幻灯片。",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_inspect = sub.add_parser("inspect", help="查看模板元素清单")
    p_inspect.add_argument("template", help="模板 .pptx 路径")
    p_inspect.add_argument("--slide", type=int, default=0, help="母版页索引，从 0 开始（默认 0）")
    p_inspect.add_argument("--json", help="把清单写成 JSON")
    p_inspect.set_defaults(func=cmd_inspect)

    p_suggest = sub.add_parser("suggest", help="按列名自动生成初始映射")
    p_suggest.add_argument("--excel", required=True, help="Excel 路径")
    p_suggest.add_argument("--template", required=True, help="模板 .pptx 路径")
    p_suggest.add_argument("--slide", type=int, default=0, help="母版页索引（默认 0）")
    p_suggest.add_argument("--sheet", help="工作表名（默认第一个）")
    p_suggest.add_argument("--header-row", type=int, default=1, help="表头行号（默认 1）")
    p_suggest.add_argument("--threshold", type=int, default=60, help="相似度阈值 0-100（默认 60）")
    p_suggest.add_argument("--out", help="Profile 输出路径（默认 profile.json）")
    p_suggest.set_defaults(func=cmd_suggest)

    p_titles = sub.add_parser("titles", help="用 AI 生成 ≤16 字语义短标题")
    p_titles.add_argument("--excel", required=True, help="Excel 路径")
    p_titles.add_argument("--sheet", help="工作表名（默认第一个）")
    p_titles.add_argument("--column", required=True, help="问题描述列名")
    p_titles.add_argument("--key", default="序号", help="序号列名（默认「序号」）")
    p_titles.add_argument("--header-row", type=int, default=1, help="表头行号（默认 1）")
    p_titles.add_argument("--limit", type=int, default=0, help="只处理前 N 行（0 = 全部）")
    p_titles.add_argument("--out", help="输出 JSON（默认 titles.json）")
    p_titles.add_argument("--api-key", help="DeepSeek/OpenAI 兼容 API key（或 DEEPSEEK_API_KEY）")
    p_titles.add_argument("--base-url", default="https://api.deepseek.com", help="API 基址")
    p_titles.add_argument("--model", default="deepseek-chat", help="模型名")
    p_titles.add_argument("--from-file", help="不调 API，直接读取已生成的 {序号:标题} JSON")
    p_titles.set_defaults(func=cmd_titles)

    p_build = sub.add_parser("build", help="按 Profile 生成 PPT")
    p_build.add_argument("--profile", required=True, help="Profile JSON 路径")
    p_build.add_argument("--excel", help="覆盖 Profile 里的 Excel 路径")
    p_build.add_argument("--template", help="覆盖 Profile 里的模板路径")
    p_build.add_argument("--out", help="输出 .pptx 路径")
    p_build.add_argument("--quiet", action="store_true", help="不打印进度")
    p_build.add_argument("--json", action="store_true", help="输出 JSON（供脚本/插件解析）")
    p_build.set_defaults(func=cmd_build)

    p_web = sub.add_parser("web", help="启动本地 Web 向导")
    p_web.add_argument("--host", default="127.0.0.1", help="监听地址（默认 127.0.0.1）")
    p_web.add_argument("--port", type=int, default=8765, help="端口（默认 8765）")
    p_web.add_argument("--no-browser", action="store_true", help="不自动打开浏览器")
    p_web.set_defaults(func=cmd_web)

    p_verify = sub.add_parser("verify", help="校验输出与模板是否零漂移")
    p_verify.add_argument("--template", required=True, help="模板 .pptx")
    p_verify.add_argument("--output", required=True, help="生成结果 .pptx")
    p_verify.add_argument("--slide", type=int, default=0, help="模板母版页索引（默认 0）")
    p_verify.add_argument("--max-pages", type=int, default=0, help="最多校验几页（0 = 全部）")
    p_verify.add_argument(
        "--strict-images",
        action="store_true",
        help="图片尺寸变化也视为错误（默认只告警，因为 contain 会等比缩放）",
    )
    p_verify.add_argument("--json", action="store_true", help="输出 JSON（供脚本/插件解析）")
    p_verify.set_defaults(func=cmd_verify)

    p_doctor = sub.add_parser("doctor", help="模板体检：结构、可映射性、格式一致性")
    p_doctor.add_argument("--template", required=True, help="模板 .pptx")
    p_doctor.add_argument("--slide", type=int, default=0, help="模板页索引（默认 0）")
    p_doctor.add_argument("--excel", help="可选：Excel 路径，用于交叉检查图片列")
    p_doctor.add_argument("--json", action="store_true", help="输出 JSON")
    p_doctor.set_defaults(func=cmd_doctor)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if getattr(args, "json", False):
        # JSON 供脚本/插件解析：固定 UTF-8，避免 Windows 管道按 GBK 编码
        for stream in (sys.stdout, sys.stderr):
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8")
    try:
        return int(args.func(args))
    except (TemplateError, ExcelDataError, ExcelImageError, ProfileError, RenderError, VerifyError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("已取消。", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
