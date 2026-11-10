"""本地 Web 服务：浏览器四步向导的后端。

只监听 127.0.0.1，不做鉴权（本机工具）；所有错误转成可读中文 JSON。
"""

from __future__ import annotations

import base64
import datetime as _dt
import hashlib
import json
import os
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from collections import OrderedDict
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from pydantic import Field as PydanticField

from .. import __version__
from ..core import preview
from ..core.excel_images import ExcelImageError
from ..core.excel_reader import ExcelDataError, read_excel
from ..core.mapping import ProfileError, build_profile
from ..core.model import Profile
from ..core.renderer import RenderError, _paginate, render
from ..core.template_health import check_template
from ..core.template_scanner import TemplateError, scan_template, summarize_slides
from ..core.titles import TitleError, generate_titles, resolve_api_key
from ..core.verifier import VerifyError, verify_file

STATIC_DIR = Path(__file__).parent / "static"
UPLOAD_DIR = Path(tempfile.gettempdir()) / "excel2ppt_studio_uploads"
MAX_UPLOAD_BYTES = 1024 * 1024 * 1024
UPLOAD_TTL_SECONDS = 24 * 3600
_MAX_UNPACKED_BYTES = 2 * 1024 * 1024 * 1024

app = FastAPI(title="Excel2PPT Studio", version=__version__)
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


class ScanIn(BaseModel):
    template: str
    slide: int = PydanticField(default=0, ge=0)


class DoctorIn(BaseModel):
    template: str
    slide: int = PydanticField(default=0, ge=0)
    excel: str | None = None


class SlidesIn(BaseModel):
    template: str


class ExcelIn(BaseModel):
    path: str
    sheet: str | None = None
    header_row: int = PydanticField(default=1, ge=1)


class SuggestIn(BaseModel):
    template: str
    slide: int = 0
    excel: str
    sheet: str | None = None
    header_row: int = 1
    threshold: int = PydanticField(default=60, ge=0, le=100)


class BuildIn(BaseModel):
    profile: dict
    out: str | None = None


class VerifyIn(BaseModel):
    template: str
    output: str
    slide: int = 0
    max_pages: int = 0
    strict_images: bool = False


class PreviewIn(BaseModel):
    path: str
    page: int = PydanticField(default=1, ge=1)


class PreviewPageIn(BaseModel):
    profile: dict
    page_index: int = PydanticField(default=0, ge=0)
    full: bool = False


class TitlesIn(BaseModel):
    excel: str
    sheet: str | None = None
    column: str
    key: str = "序号"
    header_row: int = 1
    limit: int = 0
    api_key: str | None = None
    base_url: str = "https://api.deepseek.com"
    model: str = "deepseek-chat"


_PREVIEW_CACHE: OrderedDict[str, dict] = OrderedDict()
_PREVIEW_CACHE_MAX = 8
_PREVIEW_LOCK = threading.Lock()
_FULL_PREVIEW: OrderedDict[str, dict] = OrderedDict()
_FULL_PREVIEW_MAX = 3


def _preview_key(profile: Profile, page_index: int) -> str:
    payload = json.dumps(profile.to_dict(), ensure_ascii=False, sort_keys=True)
    try:
        mtime = os.path.getmtime(profile.excel)
    except OSError:
        mtime = 0
    raw = f"{payload}|{page_index}|{mtime}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def _jsonable(value):
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return value.isoformat()
    if isinstance(value, float) and value != value:  # NaN
        return None
    return value


def _fail(message: str, status: int = 400) -> JSONResponse:
    return JSONResponse({"ok": False, "error": message}, status_code=status)


def _guard_file(path: str, label: str) -> str:
    if not path or not path.strip():
        raise HTTPException(status_code=400, detail=f"请填写{label}路径。")
    resolved = os.path.abspath(os.path.expanduser(path.strip().strip('"')))
    if not os.path.isfile(resolved):
        raise HTTPException(status_code=400, detail=f"{label}不存在：{resolved}")
    return resolved


def _cleanup_uploads(max_age_seconds: int = UPLOAD_TTL_SECONDS) -> int:
    """删除上传目录里超过 TTL 的文件，返回删除数量。"""
    if not UPLOAD_DIR.exists():
        return 0
    now = time.time()
    removed = 0
    for item in UPLOAD_DIR.iterdir():
        try:
            if item.is_file() and now - item.stat().st_mtime > max_age_seconds:
                item.unlink()
                removed += 1
        except OSError:
            continue
    return removed


def _validate_ooxml(path: Path, kind: str) -> None:
    """确认上传文件确实是 OOXML 包，避免 ZIP bomb 与非 Office 文件进入解析链。"""
    need = "ppt/presentation.xml" if kind == "template" else "xl/workbook.xml"
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            if "[Content_Types].xml" not in names or need not in names:
                raise ValueError(f"不是有效的 {'PPTX' if kind == 'template' else 'XLSX'} 文件（缺少 {need}）。")
            unpacked = sum(info.file_size for info in archive.infolist())
            if unpacked > _MAX_UNPACKED_BYTES:
                raise ValueError("解压后体积过大，已拒绝（疑似 ZIP bomb）。")
    except zipfile.BadZipFile as exc:
        raise ValueError("文件不是有效的 ZIP/OOXML 包。") from exc


@app.get("/api/health")
def health() -> dict:
    return {"ok": True, "version": app.version}


@app.get("/")
def index() -> FileResponse:
    return FileResponse(str(STATIC_DIR / "index.html"))


@app.post("/api/upload")
async def upload(kind: str, file: UploadFile = File(...), request: Request = None) -> dict:
    if kind not in ("template", "excel"):
        return _fail("kind 只能是 template 或 excel。")
    suffix = Path(file.filename or "").suffix.lower()
    allowed = (".pptx",) if kind == "template" else (".xlsx", ".xlsm")
    if suffix not in allowed:
        return _fail(f"文件类型不支持：{suffix or '无扩展名'}（需要 {'/'.join(allowed)}）")

    declared = request.headers.get("content-length") if request is not None else None
    if declared and declared.isdigit() and int(declared) > MAX_UPLOAD_BYTES:
        return _fail(
            f"文件超过 {MAX_UPLOAD_BYTES // (1024 * 1024)}MB，请改用路径直接填写本机文件。",
            status=413,
        )

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    _cleanup_uploads()
    target = UPLOAD_DIR / f"{kind}_{uuid.uuid4().hex[:10]}{suffix}"
    part = target.with_suffix(target.suffix + ".part")

    total = 0
    try:
        with part.open("wb") as handle:
            while chunk := await file.read(1024 * 1024):
                total += len(chunk)
                if total > MAX_UPLOAD_BYTES:
                    raise ValueError("文件超过上限，请改用路径直接填写本机文件。")
                handle.write(chunk)
        _validate_ooxml(part, kind)
        os.replace(part, target)
    except ValueError as exc:
        part.unlink(missing_ok=True)
        return _fail(str(exc), status=413 if "超过" in str(exc) else 400)
    finally:
        if part.exists():
            part.unlink(missing_ok=True)

    return {"ok": True, "path": str(target), "size": total, "name": file.filename}


@app.post("/api/scan")
def scan(payload: ScanIn) -> dict:
    template = _guard_file(payload.template, "模板")
    try:
        info = scan_template(template, payload.slide)
        summaries, auto_slide = summarize_slides(template)
    except TemplateError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    data = info.to_dict()
    return {
        "ok": True,
        "template": os.path.abspath(template),
        "width": data["width"],
        "height": data["height"],
        "slide_index": data["slide_index"],
        "slide_count": data["slide_count"],
        "auto_slide": auto_slide,
        "slides": [
            {
                "index": s.index,
                "kind": s.kind,
                "score": s.score,
                "fields": s.field_count,
                "slots": s.slot_count,
                "tables": s.table_count,
                "charts": s.chart_count,
                "images": s.image_count,
            }
            for s in summaries
        ],
        "fields": data["fields"],
        "warnings": data["warnings"],
    }


@app.post("/api/slides")
def slides(payload: SlidesIn) -> dict:
    """返回全部页的轻量几何，用于 Step2 缩略图条。"""
    template = _guard_file(payload.template, "模板")
    try:
        summaries, auto_slide = summarize_slides(template)
        out = []
        for summary in summaries:
            info = scan_template(template, summary.index)
            out.append(
                {
                    "index": summary.index,
                    "kind": summary.kind,
                    "score": summary.score,
                    "width": info.width,
                    "height": info.height,
                    "fields": [
                        {
                            "id": f.id,
                            "kind": f.kind,
                            "rect": f.rect.to_dict(),
                            "text": (f.text or "")[:60],
                        }
                        for f in info.fields
                    ],
                }
            )
    except TemplateError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"ok": True, "auto_slide": auto_slide, "slides": out}


@app.post("/api/doctor")
def doctor(payload: DoctorIn) -> dict:
    template = _guard_file(payload.template, "模板")
    excel = _guard_file(payload.excel, "Excel") if payload.excel else None
    try:
        report = check_template(template, payload.slide, excel)
    except TemplateError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"ok": True, **report.to_dict()}


@app.post("/api/excel")
def excel(payload: ExcelIn) -> dict:
    path = _guard_file(payload.path, "Excel")
    try:
        data = read_excel(path, payload.sheet or None, payload.header_row)
    except ExcelDataError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    preview = [
        {k: _jsonable(v) for k, v in row.items()} for row in data.rows[:20]
    ]
    return {
        "ok": True,
        "path": os.path.abspath(path),
        "sheet": data.sheet,
        "columns": [c.to_dict() for c in data.columns],
        "rows": preview,
        "total": len(data.rows),
    }


@app.post("/api/suggest")
def suggest(payload: SuggestIn) -> dict:
    template = _guard_file(payload.template, "模板")
    excel_path = _guard_file(payload.excel, "Excel")
    try:
        info = scan_template(template, payload.slide)
        sheet = read_excel(excel_path, payload.sheet or None, payload.header_row)
        profile, notes = build_profile(
            template,
            excel_path,
            info,
            sheet.sheet,
            sheet.columns,
            payload.threshold,
            header_row=payload.header_row,
        )
    except (TemplateError, ExcelDataError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {
        "ok": True,
        "profile": profile.to_dict(),
        "notes": notes,
        "slots_total": sum(len(f.slots) for f in info.fields if f.kind == "text"),
        "matched": len(profile.bindings),
        "sheet": sheet.sheet,
    }


@app.post("/api/build")
def build(payload: BuildIn) -> dict:
    try:
        profile = Profile.from_dict(payload.profile)
    except (KeyError, TypeError, ValueError) as exc:
        raise HTTPException(status_code=400, detail=f"Profile 结构非法：{exc}") from exc
    if not profile.bindings:
        raise HTTPException(status_code=400, detail="还没有任何映射，无法生成。")
    out = payload.out
    if not out:
        stem = os.path.splitext(profile.excel)[0] or "output"
        out = f"{stem}_{_dt.datetime.now():%Y%m%d_%H%M%S}.pptx"
    elif not out.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="输出文件必须是 .pptx。")
    try:
        result = render(profile.template, profile, out=out, excel=profile.excel)
    except (TemplateError, ExcelDataError, ExcelImageError, RenderError, ProfileError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {
        "ok": True,
        "output": result.output,
        "rows": result.rows,
        "slides": result.slides,
        "replaced": result.replaced,
        "warnings": result.warnings,
    }


@app.post("/api/verify")
def verify(payload: VerifyIn) -> dict:
    template = _guard_file(payload.template, "模板")
    output = _guard_file(payload.output, "输出文件")
    try:
        result = verify_file(
            template,
            output,
            slide_index=payload.slide,
            allow_image_resize=not payload.strict_images,
            max_pages=payload.max_pages,
        )
    except VerifyError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    data = result.to_dict()
    data["ok"] = result.ok
    data["errors"] = len(result.errors)
    data["warnings_count"] = len(result.warnings)
    return data


@app.post("/api/preview")
def preview_page(payload: PreviewIn) -> dict:
    path = _guard_file(payload.path, "PPT")
    tmp_dir = tempfile.mkdtemp(prefix="e2p_preview_")
    out = os.path.join(tmp_dir, "page.png")
    try:
        preview.export_png(path, payload.page, out)
        with open(out, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("ascii")
        return {
            "ok": True,
            "image": "data:image/png;base64," + encoded,
            "backend": preview.available_backend() or "unknown",
        }
    except preview.PreviewError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _preview_full(profile: Profile, page_index: int) -> dict:
    """完整渲染整册并缓存，返回指定页的 PNG。"""
    from pptx import Presentation

    key = _preview_key(profile, -1)
    with _PREVIEW_LOCK:
        entry = _FULL_PREVIEW.get(key)
        if entry is None:
            tmp_dir = tempfile.mkdtemp(prefix="e2p_full_")
            pptx = os.path.join(tmp_dir, "full.pptx")
            try:
                render(profile.template, profile, out=pptx, excel=profile.excel)
                total = len(Presentation(pptx).slides)
            except (TemplateError, ExcelDataError, ExcelImageError, RenderError, ProfileError) as exc:
                shutil.rmtree(tmp_dir, ignore_errors=True)
                raise HTTPException(status_code=400, detail=str(exc)) from exc
            entry = {"dir": tmp_dir, "pptx": pptx, "total": total}
            _FULL_PREVIEW[key] = entry
            while len(_FULL_PREVIEW) > _FULL_PREVIEW_MAX:
                _, old = _FULL_PREVIEW.popitem(last=False)
                shutil.rmtree(old["dir"], ignore_errors=True)

        page = max(0, min(int(entry["total"]) - 1, page_index))
        png = os.path.join(entry["dir"], f"p{page}.png")
        if not os.path.isfile(png):
            try:
                preview.export_png(entry["pptx"], page + 1, png)
            except preview.PreviewError as exc:
                return {
                    "ok": True,
                    "page_index": page,
                    "page_total": entry["total"],
                    "page_rows": 0,
                    "engine": "none",
                    "image": None,
                    "error": str(exc),
                }
        with open(png, "rb") as handle:
            encoded = base64.b64encode(handle.read()).decode("ascii")
        return {
            "ok": True,
            "page_index": page,
            "page_total": entry["total"],
            "page_rows": 0,
            "engine": preview.available_backend() or "unknown",
            "image": "data:image/png;base64," + encoded,
        }


@app.post("/api/preview-page")
def preview_page_of_profile(payload: PreviewPageIn) -> dict:
    """按当前映射预览第 N 页；``full=True`` 时按完整编排（封面/章节/正文/结尾）渲染整册。"""
    try:
        profile = Profile.from_dict(payload.profile)
    except (KeyError, TypeError, ValueError) as exc:
        raise HTTPException(status_code=400, detail=f"Profile 结构非法：{exc}") from exc
    if not profile.template or not profile.excel:
        raise HTTPException(status_code=400, detail="Profile 缺少模板或 Excel 路径。")
    _guard_file(profile.template, "模板")
    _guard_file(profile.excel, "Excel")

    if payload.full:
        return _preview_full(profile, payload.page_index)

    key = _preview_key(profile, payload.page_index)
    with _PREVIEW_LOCK:
        cached = _PREVIEW_CACHE.get(key)
        if cached is not None:
            _PREVIEW_CACHE.move_to_end(key)
            return {**cached, "cached": True}

        try:
            sheet = read_excel(profile.excel, profile.sheet or None, profile.header_row)
        except ExcelDataError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        pages = _paginate(sheet.rows, profile.pagination)
        if not pages:
            raise HTTPException(status_code=400, detail="没有可预览的数据页。")
        total = len(pages)
        if payload.page_index >= total:
            raise HTTPException(status_code=400, detail=f"页码越界：{payload.page_index}（共 {total} 页）")
        page_rows = pages[payload.page_index]

        sub = Profile(
            version=profile.version,
            template=profile.template,
            slide_index=profile.slide_index,
            excel=profile.excel,
            sheet=profile.sheet,
            header_row=profile.header_row,
            bindings=profile.bindings,
            pagination={"mode": "n_per_page", "size": max(1, len(page_rows))},
            output={},
        )
        tmp_dir = tempfile.mkdtemp(prefix="e2p_preview_page_")
        try:
            pptx = os.path.join(tmp_dir, "page.pptx")
            result = render(
                profile.template,
                sub,
                out=pptx,
                excel=profile.excel,
                rows_override=page_rows,
            )
            data: dict = {
                "ok": True,
                "page_index": payload.page_index,
                "page_total": total,
                "page_rows": len(page_rows),
                "warnings": result.warnings,
            }
            try:
                png = preview.export_png(pptx, 1, os.path.join(tmp_dir, "page.png"))
                with open(png, "rb") as f:
                    encoded = base64.b64encode(f.read()).decode("ascii")
                data.update(
                    {
                        "engine": preview.available_backend() or "unknown",
                        "image": "data:image/png;base64," + encoded,
                    }
                )
            except preview.PreviewError as exc:
                data.update(
                    {
                        "engine": "none",
                        "image": None,
                        "error": str(exc),
                        "rows": [
                            {k: _jsonable(v) for k, v in row.items()} for row in page_rows
                        ],
                    }
                )
        except (TemplateError, RenderError, ExcelDataError, ExcelImageError) as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

        _PREVIEW_CACHE[key] = data
        while len(_PREVIEW_CACHE) > _PREVIEW_CACHE_MAX:
            _PREVIEW_CACHE.popitem(last=False)
        return {**data, "cached": False}


def _plain_text(value) -> str:
    if value is None:
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value).strip()


@app.post("/api/titles")
def make_titles(payload: TitlesIn) -> dict:
    """调 LLM 批量生成 ≤16 字短标题，返回 {序号: 标题}。"""
    excel = _guard_file(payload.excel, "Excel")
    try:
        sheet = read_excel(excel, payload.sheet or None, payload.header_row)
    except ExcelDataError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    if payload.column not in sheet.column_names:
        raise HTTPException(status_code=400, detail=f"列不存在：{payload.column}")
    rows = sheet.rows[: payload.limit or None]
    items = []
    for row in rows:
        text = _plain_text(row.get(payload.column))
        if text:
            items.append((_plain_text(row.get(payload.key)) or str(len(items) + 1), text))
    if not items:
        raise HTTPException(status_code=400, detail="没有可用的问题描述。")
    try:
        titles = generate_titles(
            items,
            resolve_api_key(payload.api_key),
            payload.base_url,
            payload.model,
        )
    except TitleError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    over = [k for k, v in titles.items() if len(v) > 16]
    return {"ok": True, "titles": titles, "count": len(titles), "over16": len(over)}


def run(host: str = "127.0.0.1", port: int = 8765, open_browser: bool = True) -> None:
    import webbrowser

    import uvicorn

    url = f"http://{host}:{port}"
    if open_browser:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    uvicorn.run(app, host=host, port=port, log_level="info")
