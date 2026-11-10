from __future__ import annotations

from pptx import Presentation

from excel2ppt_studio.core.model import Binding, Profile
from excel2ppt_studio.core.renderer import render
from excel2ppt_studio.core.template_scanner import scan_template


def _chart_field(template_path):
    info = scan_template(template_path)
    return next(f for f in info.fields if f.kind == "chart")


def test_chart_binding_replaces_categories_and_values(template_path, excel_path, tmp_path):
    chart = _chart_field(template_path)
    profile = Profile(
        template=template_path,
        excel=excel_path,
        sheet="检查表",
        bindings=[
            Binding(
                chart.id,
                "chart",
                {
                    "categories": "项目名称",
                    "series": [{"name": "数量", "value": "数量"}],
                },
            )
        ],
        pagination={"mode": "n_per_page", "size": 3},
    )
    out = tmp_path / "chart_out.pptx"
    result = render(template_path, profile, out=str(out))
    assert result.slides == 1
    assert result.replaced == 3

    prs = Presentation(str(out))
    shape = next(s for s in prs.slides[0].shapes if s.has_chart)
    assert [str(c) for c in shape.chart.plots[0].categories] == ["项目A", "项目B", "项目C"]
    assert list(shape.chart.series[0].values) == [1.0, 2.0, 3.0]


def test_chart_binding_keeps_shape_geometry(template_path, excel_path, tmp_path):
    chart = _chart_field(template_path)
    profile = Profile(
        template=template_path,
        excel=excel_path,
        sheet="检查表",
        bindings=[
            Binding(
                chart.id,
                "chart",
                {"categories": "项目名称", "series": [{"name": "数量", "value": "数量"}]},
            )
        ],
        pagination={"mode": "n_per_page", "size": 3},
    )
    out = tmp_path / "chart_geo.pptx"
    render(template_path, profile, out=str(out))

    tpl = Presentation(template_path)
    src = next(s for s in tpl.slides[0].shapes if s.has_chart)
    prs = Presentation(str(out))
    got = next(s for s in prs.slides[0].shapes if s.has_chart)
    assert (got.left, got.top, got.width, got.height) == (
        src.left,
        src.top,
        src.width,
        src.height,
    )


def test_chart_binding_requires_columns(template_path, excel_path, tmp_path):
    import pytest

    from excel2ppt_studio.core.renderer import RenderError

    chart = _chart_field(template_path)
    profile = Profile(
        template=template_path,
        excel=excel_path,
        sheet="检查表",
        bindings=[Binding(chart.id, "chart", {})],
    )
    with pytest.raises(RenderError, match="类目列"):
        render(template_path, profile, out=str(tmp_path / "x.pptx"))
