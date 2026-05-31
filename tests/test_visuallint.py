"""Tests for the deterministic visual lint (synthetic PNGs, no rendering)."""
import pytest

from marp_pptx.visuallint import visual_lint

Image = pytest.importorskip("PIL.Image")
ImageDraw = pytest.importorskip("PIL.ImageDraw")

W, H = 800, 450
CREAM = (250, 249, 245)
INK = (20, 20, 19)


def _slide(tmp_path, name, rects):
    im = Image.new("RGB", (W, H), CREAM)
    d = ImageDraw.Draw(im)
    for r in rects:
        d.rectangle(r, fill=INK)
    p = tmp_path / name
    im.save(str(p))
    return str(p)


def test_clean_balanced_slide(tmp_path):
    # content fills a centred band -> no warnings
    p = _slide(tmp_path, "clean.png", [(80, 110, 720, 340)])
    assert visual_lint(p) == []


def test_sparse_slide_flagged(tmp_path):
    p = _slide(tmp_path, "sparse.png", [(90, 200, 280, 250)])
    warns = visual_lint(p)
    assert any("sparse" in w for w in warns)


def test_overflow_to_edge_flagged(tmp_path):
    # content runs to the bottom edge
    p = _slide(tmp_path, "overflow.png", [(80, 120, 720, 446)])
    assert any("edge" in w for w in visual_lint(p))


def test_top_skew_flagged(tmp_path):
    p = _slide(tmp_path, "topskew.png", [(80, 15, 720, 110)])
    assert any("top" in w for w in visual_lint(p))


def test_minimal_type_not_flagged_sparse(tmp_path):
    # a title slide is intentionally minimal -> sparse/skew suppressed
    p = _slide(tmp_path, "title.png", [(250, 190, 550, 240)])
    assert visual_lint(p, slide_class="title") == []


def test_blank_slide(tmp_path):
    p = _slide(tmp_path, "blank.png", [])
    assert any("blank" in w for w in visual_lint(p))
