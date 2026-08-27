"""Tests for the MCP server tools."""
import asyncio

import pytest

pytest.importorskip("mcp")

from marp_pptx import mcp as srv
from marp_pptx.render import tools_available


def test_slide_types():
    allt = srv.slide_types()
    assert len(allt) == 57
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
    assert isinstance(res["authoring_warnings"], list)


def test_build_pptx_surfaces_authoring_warnings(tmp_path):
    # An agent building via MCP must learn when the tool couldn't honour the
    # deck: an unknown _class and a missing image both surface (they used to
    # be swallowed on stderr, leaving the agent with a silently-wrong deck).
    md = ("---\nmarp: true\n---\n\n"
          "<!-- _class: bogus-type -->\n# Hi\n\n---\n\n"
          "<!-- _class: diagram -->\n# Fig\n\n![w:800](nope.png)\n")
    res = srv.build_pptx(md, math="png", output_path=str(tmp_path / "d.pptx"))
    aw = res["authoring_warnings"]
    assert any("unknown type" in w and "bogus-type" in w for w in aw)
    assert any("image not found" in w for w in aw)


def test_tools_registered():
    tools = asyncio.run(srv.mcp.list_tools())
    names = {t.name for t in tools}
    assert {"slide_types", "slide_template", "list_presets",
            "get_preset", "build_pptx", "preview_png"} <= names


def test_build_pptx_rejects_oversized_markdown():
    with pytest.raises(RuntimeError):
        srv.build_pptx("x" * (srv._MAX_MARKDOWN + 1))


def test_palette_name_traversal_rejected():
    from marp_pptx.theme import get_palette_path
    assert get_palette_path("../../etc/passwd") is None
    assert get_palette_path("../academic") is None
    assert get_palette_path("foo/bar") is None
    assert get_palette_path("navy") is not None  # legit name still resolves


@pytest.mark.skipif(not tools_available(), reason="LibreOffice/pdftoppm not installed")
def test_preview_png_returns_images():
    imgs = srv.preview_png(srv.get_preset("minimal"), max_slides=2)
    assert 1 <= len(imgs) <= 2
    assert imgs[0].data  # PNG bytes present
