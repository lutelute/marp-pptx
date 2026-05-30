"""Tests for the MCP server tools."""
import asyncio

import pytest

pytest.importorskip("mcp")

from marp_pptx import mcp as srv
from marp_pptx.render import tools_available


def test_slide_types():
    allt = srv.slide_types()
    assert len(allt) == 52
    assert {"name", "category", "meaning", "use_when", "geometry"} <= set(allt[0])
    evals = srv.slide_types("evaluation")
    assert evals and all(t["category"] == "evaluation" for t in evals)


def test_slide_template():
    assert "kpi-container" in srv.slide_template("kpi")
    assert "unknown type" in srv.slide_template("not-a-type")


def test_presets():
    presets = srv.list_presets()
    ids = {p["id"] for p in presets}
    assert {"minimal", "academic-talk", "product", "lecture"} <= ids
    assert "marp: true" in srv.get_preset("minimal")
    assert "unknown preset" in srv.get_preset("nope")


def test_build_pptx(tmp_path):
    out = tmp_path / "deck.pptx"
    res = srv.build_pptx(srv.get_preset("minimal"), output_path=str(out))
    assert res["slide_count"] == 3
    assert out.is_file()
    assert isinstance(res["lint_warnings"], list)


def test_tools_registered():
    tools = asyncio.run(srv.mcp.list_tools())
    names = {t.name for t in tools}
    assert {"slide_types", "slide_template", "list_presets",
            "get_preset", "build_pptx", "preview_png"} <= names


@pytest.mark.skipif(not tools_available(), reason="LibreOffice/pdftoppm not installed")
def test_preview_png_returns_images():
    imgs = srv.preview_png(srv.get_preset("minimal"), max_slides=2)
    assert 1 <= len(imgs) <= 2
    assert imgs[0].data  # PNG bytes present
