from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from excel2ppt_studio.web.app import app


@pytest.fixture()
def client():
    return TestClient(app)


def test_health(client):
    res = client.get("/api/health")
    assert res.status_code == 200
    assert res.json()["ok"] is True


def test_index_served(client):
    res = client.get("/")
    assert res.status_code == 200
    assert "Excel2PPT Studio" in res.text


def test_static_assets(client):
    for name in ("app.js", "style.css"):
        res = client.get(f"/static/{name}")
        assert res.status_code == 200
        assert len(res.content) > 500


def test_scan_endpoint(client, template_path):
    res = client.post("/api/scan", json={"template": template_path, "slide": 0})
    assert res.status_code == 200
    data = res.json()
    assert data["ok"] is True
    assert len(data["fields"]) == 5
    assert data["width"] == 12192000


def test_scan_rejects_missing_file(client, tmp_path):
    res = client.post("/api/scan", json={"template": str(tmp_path / "none.pptx"), "slide": 0})
    assert res.status_code == 400
    assert "不存在" in res.json()["detail"]


def test_excel_endpoint(client, excel_path):
    res = client.post("/api/excel", json={"path": excel_path, "header_row": 1})
    assert res.status_code == 200
    data = res.json()
    assert data["total"] == 3
    assert [c["name"] for c in data["columns"]][:2] == ["项目名称", "单位"]


def test_suggest_endpoint(client, template_path, excel_path):
    res = client.post(
        "/api/suggest",
        json={"template": template_path, "slide": 0, "excel": excel_path, "threshold": 60},
    )
    assert res.status_code == 200
    data = res.json()
    assert data["ok"] is True
    assert data["matched"] >= 3
    kinds = {b["kind"] for b in data["profile"]["bindings"]}
    assert {"text", "table", "image"} <= kinds


def test_build_endpoint(client, template_path, excel_path, tmp_path):
    suggest = client.post(
        "/api/suggest",
        json={"template": template_path, "slide": 0, "excel": excel_path},
    ).json()
    out = tmp_path / "web_out.pptx"
    res = client.post("/api/build", json={"profile": suggest["profile"], "out": str(out)})
    assert res.status_code == 200
    data = res.json()
    assert data["slides"] == 3
    assert out.exists()


def test_build_rejects_empty_bindings(client, tmp_path):
    res = client.post("/api/build", json={"profile": {"bindings": []}})
    assert res.status_code == 400
    assert "映射" in res.json()["detail"]


def test_upload_template(client, template_path, tmp_path):
    with open(template_path, "rb") as f:
        res = client.post(
            "/api/upload?kind=template",
            files={"file": ("模板.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
        )
    assert res.status_code == 200
    data = res.json()
    assert data["ok"] is True
    assert data["path"].endswith(".pptx")


def test_upload_rejects_wrong_type(client):
    res = client.post(
        "/api/upload?kind=excel",
        files={"file": ("x.txt", b"hello", "text/plain")},
    )
    assert res.status_code == 400


def test_verify_endpoint(client, template_path, excel_path, tmp_path):
    suggest = client.post(
        "/api/suggest",
        json={"template": template_path, "slide": 0, "excel": excel_path},
    ).json()
    out = tmp_path / "verify_out.pptx"
    build = client.post("/api/build", json={"profile": suggest["profile"], "out": str(out)})
    assert build.status_code == 200

    res = client.post(
        "/api/verify",
        json={"template": template_path, "output": str(out)},
    )
    assert res.status_code == 200
    data = res.json()
    assert data["ok"] is True
    assert data["pages_checked"] == 3


def test_verify_rejects_missing_output(client, template_path, tmp_path):
    res = client.post(
        "/api/verify",
        json={"template": template_path, "output": str(tmp_path / "none.pptx")},
    )
    assert res.status_code == 400


def test_preview_endpoint_or_graceful_failure(client, template_path):
    res = client.post("/api/preview", json={"path": template_path, "page": 1})
    if res.status_code == 200:
        data = res.json()
        assert data["image"].startswith("data:image/png;base64,")
    else:
        assert res.status_code == 400
        assert "WPS" in res.json()["detail"] or "PowerPoint" in res.json()["detail"]


def test_preview_rejects_bad_page(client, template_path):
    res = client.post("/api/preview", json={"path": template_path, "page": 99})
    assert res.status_code == 400


def test_preview_page_endpoint(client, template_path, excel_path):
    suggest = client.post(
        "/api/suggest",
        json={"template": template_path, "slide": 0, "excel": excel_path},
    ).json()
    res = client.post(
        "/api/preview-page",
        json={"profile": suggest["profile"], "page_index": 0},
    )
    assert res.status_code == 200
    data = res.json()
    assert data["ok"] is True
    assert data["page_total"] == 3
    if data.get("image"):
        assert data["image"].startswith("data:image/png;base64,")
        assert data["engine"] in ("wps", "powerpoint")
    else:
        assert data["engine"] == "none"
        assert data["rows"]


def test_preview_page_out_of_range(client, template_path, excel_path):
    suggest = client.post(
        "/api/suggest",
        json={"template": template_path, "slide": 0, "excel": excel_path},
    ).json()
    res = client.post(
        "/api/preview-page",
        json={"profile": suggest["profile"], "page_index": 99},
    )
    assert res.status_code == 400
    assert "越界" in res.json()["detail"]


def test_upload_rejects_fake_pptx(client):
    """扩展名是 .pptx 但内容不是 OOXML 包，必须拒绝。"""
    res = client.post(
        "/api/upload?kind=template",
        files={"file": ("fake.pptx", b"not a zip at all", "application/octet-stream")},
    )
    assert res.status_code == 400
    body = res.json()
    assert body["ok"] is False
    assert "PPTX" in body["error"] or "xlsx" in body["error"].lower() or "有效" in body["error"]


def test_upload_rejects_oversized(client, monkeypatch):
    from excel2ppt_studio.web import app as webapp

    monkeypatch.setattr(webapp, "MAX_UPLOAD_BYTES", 1024)
    res = client.post(
        "/api/upload?kind=excel",
        files={"file": ("big.xlsx", b"x" * 4096, "application/octet-stream")},
    )
    assert res.status_code == 413


def test_upload_dir_cleanup(tmp_path, monkeypatch):
    import os
    import time

    from excel2ppt_studio.web import app as webapp

    up = tmp_path / "up"
    up.mkdir()
    monkeypatch.setattr(webapp, "UPLOAD_DIR", up)
    old = up / "old.xlsx"
    old.write_bytes(b"x")
    old_time = time.time() - 7200
    os.utime(old, (old_time, old_time))

    removed = webapp._cleanup_uploads(max_age_seconds=3600)
    assert removed == 1
    assert not old.exists()


def test_build_rejects_non_pptx_output(client, template_path, excel_path, tmp_path):
    suggest = client.post(
        "/api/suggest",
        json={"template": template_path, "slide": 0, "excel": excel_path},
    ).json()
    res = client.post(
        "/api/build",
        json={"profile": suggest["profile"], "out": str(tmp_path / "evil.exe")},
    )
    assert res.status_code == 400
    assert "pptx" in res.json()["detail"].lower()


def test_preview_timeout_kills_process_tree(tmp_path, monkeypatch):
    """预览超时必须杀整棵进程树（含 COM 拉起的 WPS），不能只杀 worker。"""
    import subprocess

    import pytest

    from excel2ppt_studio.core import preview

    pptx = tmp_path / "deck.pptx"
    pptx.write_bytes(b"x")
    killed: list[int] = []

    class FakeProc:
        pid = 4321
        returncode = 1

        def communicate(self, timeout=None):
            if timeout is not None:
                raise subprocess.TimeoutExpired("cmd", timeout)
            return "", ""

        def kill(self):
            return None

    monkeypatch.setattr(preview, "available_backend", lambda: "wps")
    monkeypatch.setattr(preview, "_kill_process_tree", lambda pid: killed.append(pid))
    monkeypatch.setattr(subprocess, "Popen", lambda *a, **k: FakeProc())

    with pytest.raises(preview.PreviewError, match="超时"):
        preview.export_png(str(pptx), 1, str(tmp_path / "out.png"), timeout=1)
    assert killed == [4321]
