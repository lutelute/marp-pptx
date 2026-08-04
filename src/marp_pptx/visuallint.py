"""Deterministic visual lint — flag design problems from a rendered slide PNG.

No AI / no key: pure pixel analysis of the rendered image. Detects the failure
modes that text-level lint can't see — content overflowing the safe area,
slides that are nearly empty, and badly top/bottom-skewed balance. Intended both
as a standalone check (anyone with Pillow) and as the deterministic floor under
the optional LLM visual-critique loop.

    visual_lint(png_path, slide_class=None) -> list[str]   # human-readable warnings
    lint_deck(markdown, palette=...) -> list[{slide, class, warnings}]
"""
from __future__ import annotations

from pathlib import Path

# Types that are intentionally minimal / centered — don't flag them as sparse.
# references is here because its fill is citation-count-dependent: 3 refs on a
# vertically centered list is a normal deck, not a layout bug.
_MINIMAL = {"title", "divider", "statement", "big-statement", "dark", "end",
            "rq", "quote", "big-number", "big-number-dark", "takeaway", "section",
            "references"}


def visual_lint(png_path: str, slide_class: str | None = None, *, edge: float = 0.035,
                bottom_exclude: float = 0.0, top_exclude: float = 0.0) -> list[str]:
    """Return design warnings for one rendered slide image (empty list = clean).

    bottom_exclude / top_exclude: fraction of the slide height to ignore at
    the bottom / top — chrome bands, not author content. Use ~0.05 at the
    bottom for themes with a footer bar (layout.footer_bar) and ~0.16 at the
    top for a full-width frametitle band (h1_deco: band).
    """
    from PIL import Image

    im = Image.open(png_path).convert("RGB")
    w, h = im.size
    px = im.load()
    # background = a corner colour INSIDE the chrome-exclusion bands — with a
    # full-width frametitle band the literal top-left corner is the band color,
    # which would invert the content mask (everything white reads as content).
    bg_y = min(h - 3, int(h * top_exclude) + 4 if top_exclude else 2)
    bg = px[2, bg_y]
    tol = 42  # sum-of-abs-channel-diff threshold for "this pixel is content"

    # The builder stamps a right-aligned page number ("N / M") in the
    # bottom-right corner of every interior slide. That chrome sits ~3% from
    # the bottom edge and used to trip the edge-overflow check on essentially
    # every slide (worse at keynote density, where the bigger type pushes it
    # past the threshold) — a false positive that drowned out real overflow.
    # Skip that corner so the lint measures the AUTHOR's content, not chrome.
    foot_y = h * 0.93
    foot_x = w * 0.78
    band_y = h * (1.0 - bottom_exclude) if bottom_exclude else None
    head_y = h * top_exclude if top_exclude else None
    step = max(1, w // 380)
    minx = miny = 10 ** 9
    maxx = maxy = -1
    content = total = ysum = 0
    for y in range(0, h, step):
        for x in range(0, w, step):
            if y > foot_y and x > foot_x:
                continue  # page-number footer chrome, not content
            if band_y is not None and y >= band_y:
                continue  # full-width footer bar chrome (beamer themes)
            if head_y is not None and y <= head_y:
                continue  # frametitle band chrome (h1_deco: band)
            r, g, b = px[x, y]
            if abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2]) > tol:
                content += 1
                ysum += y
                minx, maxx = min(minx, x), max(maxx, x)
                miny, maxy = min(miny, y), max(maxy, y)
            total += 1

    if content == 0:
        return ["blank slide — no content rendered"]

    warns: list[str] = []
    # Use the content BOUNDING BOX area, not ink ratio — text is mostly
    # whitespace, so ink ratio is ~always low even on full slides.
    bbox_area = ((maxx - minx) * (maxy - miny)) / (w * h)
    cy = (ysum / content) / h  # vertical centre of mass, 0 (top) .. 1 (bottom)

    if (maxy > h * (1 - edge) or miny < h * edge
            or maxx > w * (1 - edge) or minx < w * edge):
        warns.append("content reaches the slide edge (overflow / margins too tight)")

    if slide_class not in _MINIMAL:
        if bbox_area < 0.28:
            warns.append(f"very sparse slide — content fills only ~{bbox_area * 100:.0f}% of the frame (large empty area)")
        if cy < 0.35:
            warns.append(f"content skewed to the top (vertical centre ~{cy * 100:.0f}%); empty space below")
        elif cy > 0.65:
            warns.append(f"content skewed to the bottom (vertical centre ~{cy * 100:.0f}%); empty space above")

    return warns


def lint_deck(markdown: str, palette: str = "claude", dpi: int = 96,
              density: str = "academic") -> list[dict]:
    """Render a deck and visual-lint every slide. Requires LibreOffice + pdftoppm.

    `density` must match what the deck is actually built with ("academic" or
    "keynote") — linting a keynote deck under academic scaling reports false
    overflow/skew because the type is larger in the real output.

    Returns [{slide, class, warnings}] for slides that have warnings.
    """
    import tempfile

    from marp_pptx.parser import parse_marp
    from marp_pptx.builder import PptxBuilder
    from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
    from marp_pptx.render import pptx_to_pngs, tools_available

    if not tools_available():
        raise RuntimeError("lint_deck needs LibreOffice (soffice) + pdftoppm")

    tc = ThemeConfig.from_css(get_default_theme_path())
    tc.math_mode = "png"
    tc.density = density
    # Mirror the CLI's keynote scaling so the linted render matches the deck.
    if density == "keynote":
        tc.font_scale = getattr(tc, "font_scale", 1.0) * 1.22
        tc.margin_scale = 1.12
    pp = get_palette_path(palette)
    if pp:
        tc.apply_palette(pp)

    with tempfile.TemporaryDirectory() as td:
        tdp = Path(td)
        md = tdp / "deck.md"
        md.write_text(markdown, encoding="utf-8")
        slides = parse_marp(str(md))
        builder = PptxBuilder(base_path=tdp, theme=tc)
        builder.build_all(slides)
        pptx = tdp / "deck.pptx"
        builder.save(str(pptx))
        pngs = pptx_to_pngs(pptx, tdp, dpi=dpi)

        # Full-width chrome bands (beamer themes) are excluded from linting.
        bottom = 0.05 if getattr(tc.layout, "footer_bar", False) else 0.0
        top = 0.13 if getattr(tc.layout, "h1_deco", "") == "band" else 0.0
        out: list[dict] = []
        for i, (png, sd) in enumerate(zip(pngs, slides), 1):
            cls = getattr(sd, "slide_class", None)
            warns = visual_lint(str(png), cls, bottom_exclude=bottom,
                                top_exclude=top)
            if warns:
                out.append({"slide": i, "class": cls or "default", "warnings": warns})
        return out


# --- geometric text-collision detector --------------------------------------
# Deterministic, no rendering: catches the "wrap-overlap" bug class (a text box
# sized too short for its wrapped content, so the text spills onto the element
# below it). Operates on a built python-pptx Presentation by re-estimating each
# text box's required height at its own width and checking whether that height
# intrudes into a lower text box that shares its horizontal span.

import re as _re

from .audit import needed_height_emu as _needed_height_emu   # noqa: E402,F401
from .metrics import measure_em as _em_width                 # noqa: E402,F401

_PAGE_NO = _re.compile(r"\d+\s*/\s*\d+")  # footer chrome — never a collision target

# Height estimation lives in `audit` — one measurement model for the builder
# (which reserves the box), this detector (which checks the reservation held)
# and `marp-pptx doctor` (which reports it). They used to have two, and drifted:
# the local one charged 0.55 em for every Latin glyph, so a line that really
# fits could be reported as wrapping.


def detect_text_collisions(prs, *, intrude_emu: int = 152400) -> list[dict]:
    """Flag text whose wrapped content overflows where it shouldn't — either
    onto a lower text box (collision) or off the bottom of the slide.

    Both come from the same root (a box estimated too short for its wrapped
    content). The collision case catches stacked layouts (rq/takeaway/…); the
    off-slide case catches single blocks with nothing below them but the slide
    edge (e.g. a long profile bio that runs off the bottom) — which the
    collision check alone cannot see.

    `intrude_emu` (default 0.12in) is how far a collision overflow must reach
    into the box below before it counts — tuned to ignore intentionally-tight
    label placements (eyebrows above their heading) while catching real spill.

    Returns [{slide, over, onto}]. `onto` is the text it collides with, or
    "(off the slide bottom)" for off-slide overflow. Empty list = clean.
    """
    out: list[dict] = []
    sh = prs.slide_height
    # Text sitting in the bottom ~0.35in is footer-bar chrome (beamer themes):
    # not author content, and its bar-height boxes false-trip the off-slide
    # check when a long deck title wraps in the estimator.
    bar_chrome = sh - 320040
    for i, slide in enumerate(prs.slides, 1):
        boxes = [s for s in slide.shapes
                 if getattr(s, "has_text_frame", False)
                 and s.text_frame.text.strip()
                 and not _PAGE_NO.fullmatch(s.text_frame.text.strip())
                 and s.top < bar_chrome]
        for a in boxes:
            a_bottom = a.top + _needed_height_emu(a)
            # (1) off-slide bottom overflow — text runs past the slide edge with
            #     nothing below it to collide with. +0.2in margin past the edge:
            #     bottom-anchored footnotes over-estimate by ~0.07in (they fit on
            #     one line), while a real run-off (e.g. a too-long profile bio)
            #     clears the edge by ~0.4in — the margin separates them cleanly.
            if a.top < sh and a_bottom > sh + 182880:
                out.append({"slide": i,
                            "over": a.text_frame.text.strip()[:40],
                            "onto": "(off the slide bottom)"})
                continue
            # (2) collision onto a lower text box
            for b in boxes:
                if b is a or b.top < a.top + 6350:          # b must sit below a
                    continue
                if a.left >= b.left + b.width or b.left >= a.left + a.width:
                    continue                                 # no horizontal overlap
                if a_bottom > b.top + intrude_emu:
                    out.append({"slide": i,
                                "over": a.text_frame.text.strip()[:40],
                                "onto": b.text_frame.text.strip()[:40]})
                    break
    return out
