"""Real font metrics — deterministic text measurement.

The single most common defect in generated decks is text that does not fit its
box.  Everything else (overlap, collision, a caption riding over a figure) is
usually that defect one step downstream.  Catching it needs an honest answer to
three questions, *before* the .pptx is written:

    how wide does this string actually render?
    where will PowerPoint break the line?
    how tall is the resulting block?

This module answers them from the font's own advance widths (``hmtx``/``cmap``
via fontTools, or Pillow's rasteriser), not from a character-class guess.  No
LibreOffice, no rendering, no AI — the same numbers on every machine that has
the font, and a documented fallback when it doesn't.

    w  = measure_pt("Hello 世界", "Helvetica Neue", 18, ea_font="Hiragino Sans")
    ls = wrap_text("長い日本語の本文…", "Helvetica Neue", 18, 400.0, ea_font=...)
    h  = block_height_pt(ls, 18)
    sz = fit_size("長いタイトル", "Helvetica Neue", 400.0, 60.0, max_size=32)

Widths are returned in *em* units internally (1.0 em == the font size), so a
measurement is scale-free until a point size is applied.
"""

from __future__ import annotations

import re
import unicodedata
from functools import lru_cache
from pathlib import Path

__all__ = [
    "measure_em", "measure_pt", "wrap_text", "line_count", "block_height_pt",
    "fit_size", "overflow_pt", "line_height_em", "font_file", "font_status",
    "SAFE_FONTS", "QA_UNRELIABLE", "DEFAULT_LINE_FACTOR",
]

# ── Fallback width table ────────────────────────────────────────────────────
# Adobe's Helvetica AFM advance widths (units/1000 em).  Used when the named
# font has no file on this machine, so a deck authored on a bare CI box still
# measures far closer than a two-bucket 0.55/1.0 guess (which is wrong by 3x
# between "llll" and "WWWW").
_HELV = {
    " ": 278, "!": 278, '"': 355, "#": 556, "$": 556, "%": 889, "&": 667,
    "'": 191, "(": 333, ")": 333, "*": 389, "+": 584, ",": 278, "-": 333,
    ".": 278, "/": 278, ":": 278, ";": 278, "<": 584, "=": 584, ">": 584,
    "?": 556, "@": 1015, "[": 278, "\\": 278, "]": 278, "^": 469, "_": 556,
    "`": 333, "{": 334, "|": 260, "}": 334, "~": 584,
    "A": 667, "B": 667, "C": 722, "D": 722, "E": 667, "F": 611, "G": 778,
    "H": 722, "I": 278, "J": 500, "K": 667, "L": 556, "M": 833, "N": 722,
    "O": 778, "P": 667, "Q": 778, "R": 722, "S": 667, "T": 611, "U": 722,
    "V": 667, "W": 944, "X": 667, "Y": 667, "Z": 611,
    "a": 556, "b": 556, "c": 500, "d": 556, "e": 556, "f": 278, "g": 556,
    "h": 556, "i": 222, "j": 222, "k": 500, "l": 222, "m": 833, "n": 556,
    "o": 556, "p": 556, "q": 556, "r": 333, "s": 500, "t": 278, "u": 556,
    "v": 500, "w": 722, "x": 500, "y": 500, "z": 500,
}
for _d in "0123456789":
    _HELV[_d] = 556

# East-Asian width class -> em advance, for anything outside the table.
_EAW_EM = {"W": 1.0, "F": 1.0, "H": 0.5, "Na": 0.5, "A": 0.55, "N": 0.55}

# Line pitch the decks are laid out against (line height ÷ font size).  Real
# fonts report anywhere from 1.0 (Hiragino's hhea) to 1.5 (its win metrics);
# PowerPoint lands near 1.2–1.25 for the faces used here, and every reserved
# box in `builder` is sized with this figure — so measuring against it keeps
# prediction and layout in the same units.
DEFAULT_LINE_FACTOR = 1.25

# Fonts that ship with Office *and* render true-to-width under the LibreOffice
# substitution used for visual QA.  Anything else may look different in the
# preview than in the user's PowerPoint — worth a warning, not an error.
SAFE_FONTS = {
    "Arial", "Calibri", "Cambria", "Times New Roman", "Courier New",
    "Bookman Old Style", "Century Schoolbook",
}
# Substitutes have different metrics — a fit check on these is approximate.
QA_UNRELIABLE = {
    "Georgia", "Trebuchet MS", "Impact", "Arial Black", "Garamond",
    "Consolas", "Palatino Linotype", "Calibri Light", "Aptos",
}

# Names worth trying when the exact family isn't installed.  Keeps metrics
# meaningful on Linux CI, where "Helvetica Neue" and "Hiragino Sans" are absent.
_ALIASES = {
    "Helvetica Neue": ("Helvetica", "Arial", "Liberation Sans", "DejaVu Sans"),
    "Helvetica": ("Arial", "Liberation Sans", "DejaVu Sans"),
    "Arial": ("Liberation Sans", "Helvetica", "DejaVu Sans"),
    "Hiragino Sans": ("Hiragino Kaku Gothic ProN", "Noto Sans CJK JP",
                      "Noto Sans JP", "YuGothic", "Yu Gothic", "MS Gothic"),
    "Hiragino Mincho ProN": ("YuMincho", "Noto Serif CJK JP", "MS Mincho"),
    "Yu Gothic": ("YuGothic", "Noto Sans CJK JP", "Hiragino Sans"),
    "Times New Roman": ("Liberation Serif", "DejaVu Serif", "Times"),
    "SF Mono": ("Menlo", "Consolas", "DejaVu Sans Mono", "Courier New"),
}


# ── Font file resolution ────────────────────────────────────────────────────
@lru_cache(maxsize=256)
def font_file(name: str, bold: bool = False, italic: bool = False) -> str | None:
    """Absolute path to the file backing `name`, or None if not installed.

    Resolution is exact-family-only: matplotlib's ``findfont`` happily returns
    DejaVu for a name it has never heard of, and silently measuring the wrong
    font is worse than knowing we're on the fallback table.
    """
    try:
        from matplotlib import font_manager as fm
    except Exception:                                   # pragma: no cover
        return None
    from matplotlib.font_manager import FontProperties
    for cand in (name, *_ALIASES.get(name, ())):
        if not cand:
            continue
        try:
            prop = FontProperties(family=cand,
                                  weight="bold" if bold else "normal",
                                  style="italic" if italic else "normal")
            path = fm.findfont(prop, fallback_to_default=False)
        except Exception:
            continue
        if path and Path(path).exists():
            return path
    return None


def font_status(name: str) -> str:
    """``"safe"`` / ``"unreliable"`` / ``"missing"`` / ``"ok"``.

    Reported by the audit so a deck says *why* its fit numbers are approximate
    rather than leaving the author to discover it in PowerPoint.
    """
    if name in QA_UNRELIABLE:
        return "unreliable"
    if font_file(name) is None:
        return "missing"
    return "safe" if name in SAFE_FONTS else "ok"


_EM_UNITS = 1000   # rasterise at this size so an advance in px == em/1000


class _Metrics:
    """Advance widths for one font file, in em units.

    Widths come from FreeType (via Pillow), not from fontTools' ``cmap``:
    decompiling the cmap of a system TTC costs ~70 s in pure Python, while
    FreeType answers per glyph in microseconds and agrees with it.  fontTools
    is still used for the *vertical* metrics, which decompile instantly.
    """

    def __init__(self, path: str | None):
        self.path = path
        self._w: dict[str, float] = {}
        self._pil = None
        self.line_em = 1.2
        if path:
            self._load(path)

    def _load(self, path: str) -> None:
        try:
            from PIL import ImageFont
            self._pil = ImageFont.truetype(path, _EM_UNITS)
        except Exception:
            self._pil = None
        self.line_em = self._line_em(path)

    def _line_em(self, path: str) -> float:
        """Single-spaced line height, from hhea (ascent - descent + lineGap)."""
        try:
            from fontTools.ttLib import TTFont
            font = TTFont(path, fontNumber=0, lazy=True)
            upem = float(font["head"].unitsPerEm) or 1000.0
            hhea = font["hhea"]
            line = (hhea.ascent - hhea.descent + getattr(hhea, "lineGap", 0)) / upem
            font.close()
            return min(1.7, max(1.0, line))
        except Exception:
            pass
        if self._pil is not None:
            try:
                asc, desc = self._pil.getmetrics()
                return min(1.7, max(1.0, (asc + desc) / _EM_UNITS))
            except Exception:
                pass
        return 1.2

    def advance(self, ch: str) -> float | None:
        w = self._w.get(ch)
        if w is not None:
            return w
        if self._pil is None:
            return None
        try:
            val = self._pil.getlength(ch) / _EM_UNITS
        except Exception:
            return None
        self._w[ch] = val
        return val


@lru_cache(maxsize=64)
def _metrics(name: str, bold: bool = False, italic: bool = False) -> _Metrics:
    return _Metrics(font_file(name, bold, italic))


def _fallback_em(ch: str) -> float:
    w = _HELV.get(ch)
    if w is not None:
        return w / 1000.0
    if ch in ("　",):
        return 1.0
    return _EAW_EM.get(unicodedata.east_asian_width(ch), 0.55)


def _is_ea(ch: str) -> bool:
    """True for characters a CJK font should measure (and PowerPoint will
    render with the ``<a:ea>`` typeface)."""
    if unicodedata.east_asian_width(ch) in ("W", "F"):
        return True
    return "　" <= ch <= "〿" or "＀" <= ch <= "￯"


# ── Measurement ─────────────────────────────────────────────────────────────
def measure_em(text: str, font: str = "Helvetica Neue", *, ea_font: str | None = None,
               bold: bool = False, italic: bool = False) -> float:
    """Width of `text` in em units, measuring CJK with `ea_font` when given."""
    if not text:
        return 0.0
    latin = _metrics(font, bold, italic)
    ea = _metrics(ea_font or font, bold, italic) if ea_font else latin
    total = 0.0
    for ch in text:
        if ch == "\t":
            total += 2.0
            continue
        m = ea if _is_ea(ch) else latin
        val = m.advance(ch)
        if val is None and m is not latin:
            val = latin.advance(ch)
        total += val if val is not None else _fallback_em(ch)
    return total


def measure_pt(text: str, font: str, size_pt: float, *, ea_font: str | None = None,
               bold: bool = False, italic: bool = False) -> float:
    """Rendered width of `text` in points at `size_pt`."""
    return measure_em(text, font, ea_font=ea_font, bold=bold, italic=italic) * size_pt


def line_height_em(font: str = "Helvetica Neue", ea_font: str | None = None) -> float:
    """Single-spaced line height in em, from the font's own vertical metrics.

    CJK faces are typically ~1.5 em against ~1.17 for a Latin sans, and a run
    that mixes them takes the taller — which is exactly what PowerPoint does.
    """
    h = _metrics(font).line_em
    if ea_font:
        h = max(h, _metrics(ea_font).line_em)
    return h


# ── Line breaking ───────────────────────────────────────────────────────────
# 行頭禁則: characters that may not open a line.
_NO_LINE_START = set(
    "、。，．・：；？！ヽヾゝゞ々ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ"
    "）］｝」』】〉》〕〗〙〟”’»‐－ー～〜…‥"
    ")]}>,.;:!?%°'\"’”"
)
# 行末禁則: characters that may not close a line.
_NO_LINE_END = set("（［｛「『【〈《〔〖〘〝“‘«([{<$¥£¥“‘")

_WORD_SPLIT = re.compile(r"(\s+)")


def wrap_text(text: str, font: str, size_pt: float, width_pt: float, *,
              ea_font: str | None = None, bold: bool = False,
              italic: bool = False) -> list[str]:
    """Break `text` the way PowerPoint would inside a `width_pt` box.

    Latin breaks at spaces (with per-character breaking for a word that cannot
    fit at all); CJK breaks between characters, subject to 禁則処理 — a line
    never opens with 。」 or closes with 「（.  Explicit newlines are kept.
    """
    if width_pt <= 0 or not text:
        return [text] if text else [""]

    def w(s: str) -> float:
        return measure_pt(s, font, size_pt, ea_font=ea_font, bold=bold, italic=italic)

    out: list[str] = []
    for raw in text.split("\n"):
        if not raw:
            out.append("")
            continue
        line = ""
        for tok in _tokens(raw):
            cand = line + tok
            if line and w(cand) > width_pt:
                if tok.isspace():
                    out.append(line.rstrip())
                    line = ""
                    continue
                broken, line = _break(line, tok, w, width_pt)
                out.extend(broken)
            else:
                line = cand
        if line.strip() or not out:
            out.append(line.rstrip())
    return out


def _tokens(s: str) -> list[str]:
    """Split into break-eligible units: Latin words, whitespace, single CJK."""
    toks: list[str] = []
    buf = ""
    for ch in s:
        if _is_ea(ch) or ch.isspace():
            if buf:
                toks.append(buf)
                buf = ""
            toks.append(ch)
        else:
            buf += ch
    if buf:
        toks.append(buf)
    return toks


def _break(line: str, tok: str, w, width_pt: float) -> tuple[list[str], str]:
    """Emit `line` (kinsoku-adjusted) and return the carry for the next line."""
    emitted: list[str] = []
    # 追い出し: pull the trailing char down when it may not close a line, and
    # keep a leading 。」 up when it may not open one.
    if tok and tok[0] in _NO_LINE_START and len(line) > 1:
        # The offending char stays on this line if that only overhangs slightly
        # (ぶら下げ), which is what PowerPoint does for a single punctuation mark.
        if w(line + tok[0]) <= width_pt * 1.06:
            emitted.append((line + tok[0]).rstrip())
            return emitted, tok[1:]
    carry = ""
    while line and line[-1] in _NO_LINE_END:
        carry = line[-1] + carry
        line = line[:-1]
    emitted.append(line.rstrip())
    rest = carry + tok
    # A single token wider than the box still has to go somewhere: split it.
    while w(rest) > width_pt and len(rest) > 1:
        cut = len(rest) - 1
        while cut > 1 and w(rest[:cut]) > width_pt:
            cut -= 1
        emitted.append(rest[:cut])
        rest = rest[cut:]
    return emitted, rest


def line_count(text: str, font: str, size_pt: float, width_pt: float, **kw) -> int:
    return len(wrap_text(text, font, size_pt, width_pt, **kw))


def block_height_pt(lines, size_pt: float, *, line_spacing: float = 1.0,
                    space_before: float = 0.0, font: str | None = None,
                    ea_font: str | None = None,
                    line_em: float | None = None) -> float:
    """Height of an already-wrapped block, in points.

    The per-line factor defaults to `DEFAULT_LINE_FACTOR`, the figure the deck
    is laid out against.  Pass `line_em=line_height_em(font, ea_font)` to model
    the font's own vertical metrics instead — useful when checking a deck this
    library did not build.
    """
    n = len(lines) if not isinstance(lines, str) else 1
    if n == 0:
        return 0.0
    lh = (line_em or DEFAULT_LINE_FACTOR) * size_pt * line_spacing
    return n * lh + max(0, n - 1) * space_before


def fit_size(text: str, font: str, width_pt: float, height_pt: float, *,
             max_size: float, min_size: float | None = None,
             ea_font: str | None = None, bold: bool = False,
             line_spacing: float = 1.0, step: float = 0.5) -> float:
    """Largest size ≤ `max_size` at which `text` fits `width_pt` × `height_pt`.

    Returns `min_size` when even that overflows — the caller then knows the
    content is genuinely too long for the box and should be split, not shrunk.
    """
    if min_size is None:
        min_size = max(8.0, max_size * 0.6)
    size = max_size
    while size > min_size:
        lines = wrap_text(text, font, size, width_pt, ea_font=ea_font, bold=bold)
        h = block_height_pt(lines, size, line_spacing=line_spacing,
                            font=font, ea_font=ea_font)
        if h <= height_pt:
            return size
        size -= step
    return min_size


def overflow_pt(text: str, font: str, size_pt: float, width_pt: float,
                height_pt: float, *, ea_font: str | None = None,
                bold: bool = False, line_spacing: float = 1.0) -> float:
    """How far the wrapped text exceeds `height_pt` (0.0 when it fits)."""
    lines = wrap_text(text, font, size_pt, width_pt, ea_font=ea_font, bold=bold)
    h = block_height_pt(lines, size_pt, line_spacing=line_spacing,
                        font=font, ea_font=ea_font)
    return max(0.0, h - height_pt)
