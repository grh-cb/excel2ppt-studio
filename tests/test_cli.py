from __future__ import annotations

from pptx import Presentation

from excel2ppt_studio.cli import main


def test_inspect_cli(template_path, capsys):
    code = main(["inspect", template_path])
    assert code == 0
    out = capsys.readouterr().out
    assert "共 5 个元素" in out
    assert "槽位" in out


def test_suggest_and_build_cli(template_path, excel_path, tmp_path, capsys):
    profile_path = tmp_path / "profile.json"
    assert main(["suggest", "--excel", excel_path, "--template", template_path,
                 "--out", str(profile_path)]) == 0
    assert profile_path.exists()

    out_path = tmp_path / "out.pptx"
    code = main(["build", "--profile", str(profile_path), "--out", str(out_path), "--quiet"])
    assert code == 0
    assert out_path.exists()
    assert len(Presentation(str(out_path)).slides) == 3


def test_build_cli_reports_readable_error(template_path, tmp_path, capsys):
    import json

    bad = tmp_path / "bad.json"
    bad.write_text(json.dumps({"version": 1, "bindings": []}), encoding="utf-8")
    code = main(["build", "--profile", str(bad), "--out", str(tmp_path / "x.pptx")])
    assert code == 1
    assert "错误" in capsys.readouterr().err
