"""Pure-Python LaTeX-to-PNG renderer using matplotlib.

Replaces the Node.js KaTeX+Playwright dependency entirely.
"""
from __future__ import annotations

import hashlib
import tempfile
from pathlib import Path

_CACHE_DIR = Path(tempfile.gettempdir()) / "marp_math_png"
_CACHE_DIR.mkdir(exist_ok=True)


def _matching_brace(s: str, open_idx: int) -> int:
    """Index of the '}' matching the '{' at ``open_idx`` (handles nesting), or -1.

    Braces escaped as literal LaTeX (``\\{`` / ``\\}``) are skipped so they don't
    throw off the depth count.
    """
    depth = 0
    for i in range(open_idx, len(s)):
        if i > 0 and s[i - 1] == "\\":  # escaped literal brace, not a delimiter
            continue
        if s[i] == "{":
            depth += 1
        elif s[i] == "}":
            depth -= 1
            if depth == 0:
                return i
    return -1


def _strip_colorbox(tex: str) -> str:
    r"""``\colorbox{c}{X}`` -> ``X`` (drop the highlight box; mathtext has none)."""
    while True:
        i = tex.find("\\colorbox")
        if i < 0:
            return tex
        b1 = tex.find("{", i)
        e1 = _matching_brace(tex, b1) if b1 >= 0 else -1
        b2 = tex.find("{", e1 + 1) if e1 >= 0 else -1
        e2 = _matching_brace(tex, b2) if b2 >= 0 else -1
        if e2 < 0:
            return tex  # malformed — leave as-is rather than loop forever
        content = tex[b2 + 1:e2].strip().strip("$")
        tex = tex[:i] + content + tex[e2 + 1:]


def _strip_brace_annotation(tex: str, cmd: str) -> str:
    r"""``\underbrace{X}_{Y}`` / ``\overbrace{X}^{Y}`` -> ``X``.

    mathtext supports neither macro, and the annotation Y is often CJK ``\text``
    which mathtext can't render either — so collapse to the base term X.
    """
    token = "\\" + cmd
    while True:
        i = tex.find(token)
        if i < 0:
            return tex
        b1 = tex.find("{", i)
        e1 = _matching_brace(tex, b1) if b1 >= 0 else -1
        if e1 < 0:
            return tex
        base = tex[b1 + 1:e1]
        j = e1 + 1
        # optionally consume a trailing _{...} or ^{...} annotation
        if j < len(tex) and tex[j] in "_^" and j + 1 < len(tex) and tex[j + 1] == "{":
            ann_end = _matching_brace(tex, j + 1)
            if ann_end >= 0:
                j = ann_end + 1
        tex = tex[:i] + base + tex[j:]


def _sanitize_for_mathtext(tex: str) -> str:
    """Rewrite LaTeX macros matplotlib's mathtext can't parse into ones it can.

    mathtext is a subset of LaTeX: no \\tfrac, \\bigl, \\operatorname, etc.
    This keeps the common academic equations renderable without a TeX install."""
    import re as _re
    tex = _strip_colorbox(tex)
    tex = _strip_brace_annotation(tex, "underbrace")
    tex = _strip_brace_annotation(tex, "overbrace")
    repl = [
        (r"\\tfrac", r"\\frac"),
        (r"\\dfrac", r"\\frac"),
        (r"\\bigl", ""), (r"\\bigr", ""), (r"\\Bigl", ""), (r"\\Bigr", ""),
        (r"\\big", ""), (r"\\Big", ""), (r"\\biggl", ""), (r"\\biggr", ""),
        (r"\\Biggl", ""), (r"\\Biggr", ""), (r"\\bigg", ""), (r"\\Bigg", ""),
        (r"\\!", ""), (r"\\,", r"\\ "), (r"\\;", r"\\ "), (r"\\:", r"\\ "),
        (r"\\qquad", r"\\quad"),
        (r"\\lVert", r"\\|"), (r"\\rVert", r"\\|"),
        (r"\\lvert", r"|"), (r"\\rvert", r"|"),
        (r"\\colon", ":"),
    ]
    for pat, sub in repl:
        tex = _re.sub(pat, sub, tex)
    # relation aliases mathtext may not know
    tex = _re.sub(r"\\le(?![a-zA-Z])", r"\\leq", tex)
    tex = _re.sub(r"\\ge(?![a-zA-Z])", r"\\geq", tex)
    tex = _re.sub(r"\\to(?![a-zA-Z])", r"\\rightarrow", tex)
    # \operatorname{xyz} -> \mathrm{xyz}
    tex = _re.sub(r"\\operatorname\s*\{([^}]*)\}", r"\\mathrm{\1}", tex)
    return tex


def render_latex_png(
    latex: str,
    fontsize: int = 28,
    display: bool = False,
    color: str = "#1a1a2e",
    dpi: int = 150,
) -> str | None:
    """Render LaTeX to a PNG file. Returns path to the PNG, or None on failure.

    Uses matplotlib's mathtext engine (no TeX installation required).
    """
    key = hashlib.md5(f"{latex}:{fontsize}:{display}:{dpi}".encode()).hexdigest()
    png_path = _CACHE_DIR / f"{key}.png"
    if png_path.exists():
        return str(png_path)

    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        # Sanitize unsupported macros, then wrap in math delimiters.
        # (mathtext has no \displaystyle; \sum already renders with limits.)
        tex = _sanitize_for_mathtext(latex.strip())
        if not tex.startswith("$"):
            tex = f"${tex}$"

        fig = plt.figure(figsize=(0.01, 0.01))
        fig.patch.set_alpha(0.0)

        text = fig.text(
            0, 0, tex,
            fontsize=fontsize,
            color=color,
            math_fontfamily="cm",
        )

        # Render to get bounding box
        fig.canvas.draw()
        bbox = text.get_window_extent(fig.canvas.get_renderer())

        # Resize figure to fit text
        fig.set_size_inches(
            (bbox.width / dpi) + 0.1,
            (bbox.height / dpi) + 0.1,
        )

        # Re-position text
        text.set_position((0.05 * dpi / bbox.width if bbox.width > 0 else 0,
                          0.05 * dpi / bbox.height if bbox.height > 0 else 0))

        fig.savefig(
            str(png_path),
            dpi=dpi,
            transparent=True,
            bbox_inches="tight",
            pad_inches=0.05,
        )
        plt.close(fig)

        if png_path.exists() and png_path.stat().st_size > 0:
            return str(png_path)
        return None
    except Exception as e:
        import sys
        print(f"  Math PNG render failed: {latex[:40]}... ({e})", file=sys.stderr)
        return None
