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
_MINIMAL = {"title", "divider", "statement", "big-statement", "dark", "end",
            "rq", "quote", "big-number", "big-number-dark", "takeaway", "section"}


def visual_lint(png_path: str, slide_class: str | None = None, *, edge: float = 0.035) -> list[str]:
    """Return design warnings for one rendered slide image (empty list = clean)."""
    from PIL import Image

    im = Image.open(png_path).convert("RGB")
    w, h = im.size
    px = im.load()
    # background = the slide's corner colour (cream / white / dark all work)
    bg = px[2, 2]
    tol = 42  # sum-of-abs-channel-diff threshold for "this pixel is content"

    step = max(1, w // 380)
    minx = miny = 10 ** 9
    maxx = maxy = -1
    content = total = ysum = 0
    for y in range(0, h, step):
        for x in range(0, w, step):
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


def lint_deck(markdown: str, palette: str = "claude", dpi: int = 96) -> list[dict]:
    """Render a deck and visual-lint every slide. Requires LibreOffice + pdftoppm.

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

        out: list[dict] = []
        for i, (png, sd) in enumerate(zip(pngs, slides), 1):
            cls = getattr(sd, "slide_class", None)
            warns = visual_lint(str(png), cls)
            if warns:
                out.append({"slide": i, "class": cls or "default", "warnings": warns})
        return out
