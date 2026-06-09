"""Regression tests for the 2026-06-08 quality-verification findings.

Each test pins one of the four real-content bugs found while authoring
demo/hosting-capacity.md:

1. title: inline $math$ / HTML entities leaked verbatim into author lines
2. timeline-h: <br> inside tl-h-detail was collapsed to a single line
3. takeaway: a wrapped ta-main overlapped ta-points (height under-estimate)
4. table-slide: the rendered table grew into the box below it (no headroom)
"""
import tempfile
from pathlib import Path

from pptx.util import Inches, Pt

from marp_pptx.parser import parse_slide, strip_html, text_with_breaks
from marp_pptx.theme import ThemeConfig
from marp_pptx.builder import PptxBuilder


def _make_builder(tmp_path=None):
    if tmp_path is None:
        tmp_path = Path(tempfile.mkdtemp())
    return PptxBuilder(base_path=tmp_path, theme=ThemeConfig())


def _all_text(slide):
    return "\n".join(s.text_frame.text for s in slide.shapes
                     if getattr(s, "has_text_frame", False))


# --- 1. entities + inline math on the title slide ---------------------------

def test_strip_html_unescapes_entities():
    assert strip_html("A &ensp;|&ensp; B") == "A \u2002|\u2002 B"
    assert strip_html("R&amp;D") == "R&D"


def test_title_author_line_renders_inline_markup():
    md = ("<!-- _class: title -->\n"
          "# Test Title\n"
          "## Subtitle\n\n"
          "Alice $^{1}$, Bob $^{2}$\n\n"
          "$^{1}$ Univ A &ensp;|&ensp; 2026")
    b = _make_builder()
    b.build_title(parse_slide(0, md))
    joined = _all_text(b.prs.slides[0])
    assert "Alice" in joined
    assert "$" not in joined          # raw $...$ must not leak as plain text
    assert "&ensp;" not in joined     # entities must be unescaped


# --- 2. <br> inside timeline details ----------------------------------------

def test_text_with_breaks_preserves_br():
    assert text_with_breaks("foo<br>bar") == "foo\nbar"
    assert text_with_breaks("foo <br/> bar") == "foo\nbar"
    assert text_with_breaks("a  <BR >  b") == "a\nb"


TL_H = """<!-- _class: timeline-h -->
# History
<div class="tl-h-container">
<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2012</span>
    <span class="tl-h-text">CNN</span>
    <div class="tl-h-detail">line1<br>line2</div>
  </div>
</div>
</div>"""


def test_timeline_h_detail_keeps_br_as_newline():
    sd = parse_slide(0, TL_H)
    assert sd.timeline_items[0]["detail"] == "line1\nline2"


# --- 3. takeaway main wrapping must push the points down --------------------

def test_estimate_height_accounts_for_wrapping():
    b = _make_builder()
    long_jp = "感" * 60     # 60 em of full-width text
    h_nowrap = b._estimate_text_height([long_jp], Pt(28))
    h_wrap = b._estimate_text_height([long_jp], Pt(28), width=Inches(10))
    # 60 em at 28pt in a 720pt-wide box needs >= 3 lines
    assert h_wrap.pt >= h_nowrap.pt * 2


def test_takeaway_points_follow_wrapped_main():
    long_main = "とても長いキーメッセージで折り返しが必ず発生する" * 4   # ~96 em
    md = ("<!-- _class: takeaway -->\n# Key\n"
          f'<div class="ta-main">{long_main}</div>\n'
          '<div class="ta-points">\n<li>p1</li>\n<li>p2</li>\n</div>')
    b = _make_builder()
    b.build_takeaway(parse_slide(0, md))
    slide = b.prs.slides[0]
    boxes = [s for s in slide.shapes
             if getattr(s, "has_text_frame", False) and s.text_frame.text]
    main = next(s for s in boxes if "とても長い" in s.text_frame.text)
    pts = next(s for s in boxes if "p1" in s.text_frame.text)
    # the main box must be tall enough for >= 3 wrapped lines ...
    assert main.height >= int(Pt(28 * 1.25 * 3))
    # ... and the points box must start below it
    assert pts.top >= main.top + main.height


# --- 4. table block keeps headroom above whatever sits below it -------------

def test_table_block_reserves_headroom():
    rows = "| A | B |\n|---|---|\n" + "| a | b |\n" * 5   # header + 5 body rows
    md = ("<!-- _class: table-slide -->\n# Compare\n\n" + rows +
          '\n<div class="box-accent">\n\nNote below the table\n\n</div>')
    b = _make_builder()
    b.build_table(parse_slide(0, md))
    slide = b.prs.slides[0]
    tbl = next(s for s in slide.shapes if getattr(s, "has_table", False))
    # 6 rows: bare 0.5in/row estimate must carry extra headroom so renderers
    # that grow rows (CJK line height) don't eat the gap below the table
    assert tbl.height >= int(Inches(0.5) * 6 + Inches(0.2))


# --- 5. appendix with a table must stay ONE slide ----------------------------

def test_appendix_table_stays_on_one_slide():
    md = ("<!-- _class: appendix -->\n# Params\n\n"
          '<span class="appendix-label">Appendix A</span>\n\n'
          "| k | v |\n|---|---|\n| lr | 3e-4 |\n| bs | 64 |\n")
    b = _make_builder()
    b.build_appendix(parse_slide(0, md))
    assert len(b.prs.slides) == 1          # was split into 2 slides before
    slide = b.prs.slides[0]
    assert any(getattr(s, "has_table", False) for s in slide.shapes)


# --- 6. unclosed <div> must terminate (used to spin forever) ----------------

def test_unclosed_div_terminates_and_salvages():
    from marp_pptx.parser import extract_child_divs
    out = extract_child_divs('<div class="a">x<div class="b">content')
    assert out and "content" in out[0]      # salvaged, and returned at all
    md = ('<!-- _class: kpi -->\n# K\n<div class="kpi-container">\n'
          '<div class="kpi-item">\n<span class="kpi-value">42</span>\n'
          '<span class="kpi-label">answer</span>\n')
    sd = parse_slide(0, md)                  # must not hang
    assert sd.slide_class == "kpi"


# --- 7. non-fatal authoring warnings are surfaced, not silent ---------------

def test_unknown_slide_type_warns():
    b = _make_builder()
    b.build_all([parse_slide(0, "<!-- _class: no-such-type -->\n# H\n\nbody")])
    assert any("unknown type" in w and "no-such-type" in w for w in b.warnings)
    assert len(b.prs.slides) == 1            # still renders (as plain slide)


def test_missing_image_warns():
    b = _make_builder()
    assert b._resolve_image("does-not-exist.png") is None
    assert any("image not found" in w for w in b.warnings)


def test_warnings_are_deduplicated():
    b = _make_builder()
    b._resolve_image("nope.png")
    b._resolve_image("nope.png")
    assert sum("nope.png" in w for w in b.warnings) == 1


def test_known_type_does_not_warn():
    b = _make_builder()
    b.build_all([parse_slide(0, "<!-- _class: title -->\n# T\n## sub")])
    assert b.warnings == []


# --- 8. visual_lint must ignore the page-number footer chrome ---------------

def _synth_png(rects, size=(1280, 720), bg=(250, 249, 245)):
    """White slide with black rects [(x0,y0,x1,y1), …]. Returns a temp path."""
    from PIL import Image
    im = Image.new("RGB", size, bg)
    px = im.load()
    for x0, y0, x1, y1 in rects:
        for y in range(y0, y1):
            for x in range(x0, x1):
                px[x, y] = (20, 20, 20)
    p = tempfile.mktemp(suffix=".png")
    im.save(p)
    return p


def test_visual_lint_ignores_page_number_corner():
    from marp_pptx.visuallint import visual_lint
    # A page-number-like blob in the bottom-right corner + normal centered
    # content must NOT be flagged as edge overflow (regression: it was).
    png = _synth_png([(400, 250, 880, 470),      # centered content, well inside
                      (1180, 690, 1260, 712)])    # "N / M" footer chrome
    warns = visual_lint(png, "kpi")
    assert not any("reaches the slide edge" in w for w in warns)


def test_visual_lint_still_flags_real_bottom_overflow():
    from marp_pptx.visuallint import visual_lint
    # Content spanning most of the width and running to the bottom edge is a
    # genuine overflow and must still be flagged.
    png = _synth_png([(100, 400, 900, 718)])
    warns = visual_lint(png, "kpi")
    assert any("reaches the slide edge" in w for w in warns)
