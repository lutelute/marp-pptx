"""board — a slide composed of zones, the way a hand-built report slide is.

One slide carries a definition band, then columns of cards / code / figures,
a conclusion callout, numbered steps, an emphasis band and its source line —
each in its own tinted zone. The Markdown stays plain; a few div classes say
where a zone starts:

    <div class="defn">          **Term**<br>one-line definition
    <div class="row">           columns side by side
      <div class="col w2">      a column (w1..w4 = width ratio)
      <div class="panel maroon">a column drawn as a tinted panel with an
                                accent bar (blue / maroon / navy)
    <div class="callout">       conclusion card (a leading "→ " draws an arrow)
    <div class="steps">         numbered step panels (### title | formula)
    <div class="band maroon">   emphasis band across the slide

Inside a column: `#### label`, `### card title` (+ text under it; a leading
`:point:` / `:line:` / `:polygon:` / `:network:` draws an icon), ``` code
fences (dark panel, JSON highlighted), images, bullets.
"""

from __future__ import annotations

import re

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from marp_pptx.parser import (
    parse_markdown_lines, split_code_line, strip_html, text_with_breaks,
)

_DIV_TAG = re.compile(r'<div\s+class="([^"]*)"\s*>|</div>', re.IGNORECASE)
_ICON = re.compile(r"^:(point|line|polygon|network|arrow):\s*")
_ACCENTS = ("blue", "maroon", "navy", "accent", "primary", "secondary")

# code panel (dark), from the reference deck
CODE_BG = RGBColor(0x22, 0x27, 0x3A)
CODE_COLORS = {"base": RGBColor(0xE8, 0xE8, 0xE8), "key": RGBColor(0x7E, 0xC8, 0xFF),
               "string": RGBColor(0xC3, 0xE8, 0x8D), "literal": RGBColor(0xF0, 0xB3, 0x7E),
               "comment": RGBColor(0xF0, 0x71, 0x78)}


# ── parsing ────────────────────────────────────────────────────────────────

def _split_divs(text: str):
    """[(classes | None, inner)] for the top-level divs of `text`, in order;
    None marks Markdown between divs."""
    out, pos, depth, start, cls = [], 0, 0, 0, ""
    for m in _DIV_TAG.finditer(text):
        if m.group(0).lower().startswith("<div"):
            if depth == 0:
                if text[pos:m.start()].strip():
                    out.append((None, text[pos:m.start()]))
                start, cls = m.end(), m.group(1)
            depth += 1
        elif depth:
            depth -= 1
            if depth == 0:
                out.append((cls.split(), text[start:m.start()]))
                pos = m.end()
    if text[pos:].strip():
        out.append((None, text[pos:]))
    return out


def _lines(chunk: str) -> list[str]:
    return [l for l in text_with_breaks(chunk).split("\n") if l.strip()]


def _steps(chunk: str) -> list[dict]:
    items = []
    for line in parse_markdown_lines(chunk):
        s = line.strip()
        if s.startswith("### "):
            title = s[4:]
            icon = None
            m = _ICON.match(title)
            if m:
                icon, title = m.group(1), title[m.end():]
            parts = re.split(r"\s*[|｜]\s*", title, maxsplit=1)
            items.append({"title": parts[0].strip(),
                          "sub": parts[1].strip() if len(parts) > 1 else "",
                          "body": [], "icon": icon})
        elif s and items:
            items[-1]["body"].append(re.sub(r"^[-*]\s+", "", s))
    return items


def _md_items(chunk: str) -> list[dict]:
    items, card, text = [], None, None
    for line in parse_markdown_lines(chunk):
        s = line.strip()
        if not s:
            continue
        code = split_code_line(s)
        img = re.match(r"!\[[^\]]*\]\(([^)]+)\)", s)
        if s.startswith("#### "):
            items.append({"t": "label", "text": s[5:]})
            card = text = None
        elif s.startswith("### "):
            title = s[4:]
            m = _ICON.match(title)
            card = {"t": "card", "title": title[m.end():] if m else title,
                    "icon": m.group(1) if m else None, "body": []}
            items.append(card)
            text = None
        elif code:
            items.append({"t": "code", "lang": code[0], "code": code[1]})
            card = text = None
        elif img:
            items.append({"t": "image", "path": img.group(1)})
            card = text = None
        elif card is not None:
            card["body"].append(s)
        else:
            if text is None:
                text = {"t": "text", "lines": []}
                items.append(text)
            text["lines"].append(s)
    return items


def _items(chunk: str) -> list[dict]:
    items = []
    for classes, inner in _split_divs(chunk):
        if classes is None:
            items += _md_items(inner)
        elif "callout" in classes:
            items.append({"t": "callout", "lines": _lines(inner)})
        elif "steps" in classes:
            items.append({"t": "steps", "items": _steps(inner)})
        else:
            items += _items(inner)
    return items


def _col(classes, inner) -> dict:
    w = next((int(c[1:]) for c in classes if re.fullmatch(r"w[1-6]", c)), 1)
    accent = next((c for c in classes if c in _ACCENTS), "")
    return {"w": w, "panel": "panel" in classes, "accent": accent,
            "items": _items(inner)}


def parse_board(content: str) -> list[dict]:
    blocks = []
    for classes, inner in _split_divs(content):
        if classes is None:
            its = _md_items(inner)
            if its:
                blocks.append({"kind": "row", "cols": [{"w": 1, "panel": False,
                                                        "accent": "", "items": its}]})
        elif "defn" in classes:
            ls = _lines(inner)
            if ls:
                blocks.append({"kind": "defn", "term": ls[0].strip("*").strip(),
                               "body": ls[1:]})
        elif "row" in classes:
            cols = [_col(c, i) for c, i in _split_divs(inner) if c]
            if cols:
                blocks.append({"kind": "row", "cols": cols})
        elif "panel" in classes or "col" in classes:
            blocks.append({"kind": "row", "cols": [_col(classes, inner)]})
        elif "band" in classes:
            blocks.append({"kind": "band", "lines": _lines(inner),
                           "accent": next((c for c in classes if c in _ACCENTS), "")})
        elif "steps" in classes:
            blocks.append({"kind": "steps", "items": _steps(inner)})
        elif "callout" in classes:
            blocks.append({"kind": "row", "cols": [{"w": 1, "panel": False, "accent": "",
                                                    "items": [{"t": "callout",
                                                               "lines": _lines(inner)}]}]})
    return blocks


# ── rendering ──────────────────────────────────────────────────────────────

class _Board:
    GAP = Inches(0.15)        # between blocks
    COL_GAP = Inches(0.3)
    ITEM_GAP = Inches(0.10)
    PAD = Inches(0.18)
    # base sizes (pt) — the reference deck's own scale
    S_LABEL, S_CARD_T, S_BODY, S_CALLOUT, S_CODE = 15, 19, 15, 20, 14
    S_TERM, S_BAND, S_STEP_T, S_STEP_SUB, S_STEP_B = 16, 15, 14, 13, 12.5

    def __init__(self, b, slide, k=1.0):
        self.b, self.slide, self.k = b, slide, k

    # sizes at this scale
    def pt(self, v):
        return Pt(v * self.k)

    def h(self, lines, size, width, bold=False):
        """Text height with a 2pt pad (the builder's generic 6pt tail stacks
        up badly in a card of several blocks)."""
        return int(self.b._estimate_text_height(lines, size, width=int(width),
                                                gap=Pt(3), bold=bold)) - int(Pt(4))

    def color(self, name, default):
        b = self.b
        return {"blue": b.SECONDARY, "secondary": b.SECONDARY, "maroon": b.ACCENT,
                "accent": b.ACCENT, "navy": b.PRIMARY, "primary": b.PRIMARY}.get(
                    name, default)

    @property
    def card_fill(self):
        return self.b._tint(self.b.MUTED, 0.9)

    # measurement --------------------------------------------------------
    def item_h(self, it, w):
        b, k = self.b, self.k
        t = it["t"]
        if t == "label":
            return self.h([it["text"]], self.pt(self.S_LABEL), w, bold=True)
        if t == "text":
            if self._chips(it):
                return len(it["lines"]) * int(Inches(0.33 * self.k + 0.06))
            return self.h(it["lines"], self.pt(self.S_BODY), w)
        if t == "caption":
            return self.h(it["lines"], self.pt(10.5), w)
        if t == "card":
            iw = w - 2 * int(Inches(0.22)) - (int(Inches(1.05)) if it["icon"] else 0)
            return (self.h([it["title"]], self.pt(self.S_CARD_T), iw, bold=True)
                    + (self.h(it["body"], self.pt(self.S_BODY), iw) if it["body"] else 0)
                    + 2 * int(Inches(0.13)))
        if t == "callout":
            iw = w - 2 * int(Inches(0.22)) - (int(Inches(0.9)) if self._arrow(it) else 0)
            return self.h(self._callout_lines(it), self.pt(self.S_CALLOUT), iw) \
                + 2 * int(Inches(0.16))
        if t == "code":
            n = len(it["code"].split("\n"))
            return int(Pt(n * self.S_CODE * k * 1.35)) + 2 * int(self.PAD)
        if t == "image":
            return int(Inches(1.6))
        if t == "steps":
            return self.steps_h(it["items"], w)
        return 0

    def steps_h(self, items, w):
        tw = w - int(Inches(0.75)) - int(Inches(1.1))
        hs = [int(Inches(0.2)) * 2 + self.h([s["title"] + "　" + s["sub"]], self.pt(self.S_STEP_T), tw, bold=True)
              + (self.h(s["body"], self.pt(self.S_STEP_B), tw) if s["body"] else 0) for s in items]
        return sum(hs) + int(Inches(0.12)) * max(0, len(hs) - 1)

    @staticmethod
    def resolved(items):
        """Text right under an image is its caption."""
        return [dict(it, t="caption") if it["t"] == "text" and i and
                items[i - 1]["t"] == "image" else it for i, it in enumerate(items)]

    def col_h(self, col, w):
        inner = w - (2 * int(Inches(0.22)) + int(Pt(3)) if col["panel"] else 0)
        hs = [self.item_h(it, inner) for it in self.resolved(col["items"])]
        pad = 2 * int(self.PAD) if col["panel"] else 0
        return sum(hs) + int(self.ITEM_GAP) * max(0, len(hs) - 1) + pad

    def col_widths(self, cols, width):
        total = sum(c["w"] for c in cols)
        avail = width - int(self.COL_GAP) * (len(cols) - 1)
        return [int(avail * c["w"] / total) for c in cols]

    def block_h(self, blk, width):
        if blk["kind"] == "defn":
            tw = width - 2 * int(Inches(0.3))
            return (self.h([blk["term"]], self.pt(self.S_TERM), tw, bold=True)
                    + (self.h(blk["body"], self.pt(self.S_BODY), tw) if blk["body"] else 0)
                    + 2 * int(Inches(0.1)))
        if blk["kind"] == "band":
            return self.h(blk["lines"], self.pt(self.S_BAND), width - int(Inches(0.5)), bold=True) \
                + 2 * int(Inches(0.14))
        if blk["kind"] == "steps":
            return self.steps_h(blk["items"], width)
        ws = self.col_widths(blk["cols"], width)
        return max(self.col_h(c, w) for c, w in zip(blk["cols"], ws))

    # helpers ------------------------------------------------------------
    _CHIP = re.compile(r"^[-*]\s+\*\*([^*]+)\*\*\s*(.*)$")

    def _chips(self, it):
        """A list whose every item is `- **label** values` reads as label
        chips + values (who uses it / which files / …), not as prose."""
        ls = it.get("lines", [])
        return len(ls) >= 2 and all(self._CHIP.match(l) for l in ls)

    def draw_chips(self, it, x, y, w, h):
        b = self.b
        rows = [self._CHIP.match(l).groups() for l in it["lines"]]
        rh = int(Inches(0.33 * self.k))
        step = rh + int(Inches(0.06))
        lab = self.pt(10.5)
        cw = max(int(Pt(b._visual_em_width(r[0]) * lab.pt + 22)) for r in rows)
        cur = y + max(0, (h - step * len(rows)) // 2)
        for label, val in rows:
            pill = self.rect(x, cur, cw, rh, b.WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                             radius=0.2, line=b.BORDER)
            pp = pill.text_frame.paragraphs[0]
            b._add_plain_run(pp, label, lab, b.PRIMARY, bold=True)
            pp.alignment = PP_ALIGN.CENTER
            pill.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            pill.text_frame.margin_left = pill.text_frame.margin_right = 0
            self.text(x + cw + int(Inches(0.18)), cur, w - cw - int(Inches(0.18)), rh,
                      [val], self.pt(12), b.FG, anchor=MSO_ANCHOR.MIDDLE)
            cur += step

    @staticmethod
    def _arrow(it):
        return bool(it["lines"]) and it["lines"][0].lstrip().startswith(("→", "->"))

    def _callout_lines(self, it):
        ls = list(it["lines"])
        if self._arrow(it):
            ls[0] = re.sub(r"^\s*(→|->)\s*", "", ls[0])
        return ls

    def rect(self, x, y, w, h, fill, *, shape=MSO_SHAPE.RECTANGLE, radius=None,
             shadow=False, line=None):
        s = self.slide.shapes.add_shape(shape, int(x), int(y), int(w), int(h))
        if radius is not None:
            s.adjustments[0] = radius
        s.fill.solid()
        s.fill.fore_color.rgb = fill
        if line is None:
            s.line.fill.background()
        else:
            s.line.color.rgb = line
            s.line.width = Pt(0.8)
        if shadow:
            self.b._soft_shadow(s)
        else:
            self.b._no_shadow(s)
        return s

    def text(self, x, y, w, h, lines, size, color, *,
             anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, bullets=False, spacing=None):
        tb = self.b._add_textbox(self.slide, int(x), int(y), int(w), int(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        for i, ln in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            bullet = bullets and re.match(r"^[-*]\s+", ln)
            self.b._set_rich_text(p, re.sub(r"^[-*]\s+", "", ln) if bullet else ln,
                                  size, color)
            if bullet:
                self.b._hang_bullet(p, "・", int(size.pt * 12700), color=color)
            p.alignment = align
            if i:
                p.space_before = Pt(3)
            if spacing:
                p.line_spacing = spacing
        return tb

    # icons (native shapes, so they recolor with the deck) ------------------
    def icon(self, kind, x, y, w, h):
        b, sh = self.b, self.slide.shapes
        c = b.SECONDARY
        dot = int(Inches(0.13))

        def at(fx, fy):
            return int(x + w * fx), int(y + h * fy)

        def node(px, py, d=dot, fill=c):
            e = sh.add_shape(MSO_SHAPE.OVAL, px - d // 2, py - d // 2, d, d)
            e.fill.solid(); e.fill.fore_color.rgb = fill
            e.line.fill.background(); b._no_shadow(e)

        def seg(p, q, width=2.0):
            ln = sh.add_connector(MSO_CONNECTOR.STRAIGHT, p[0], p[1], q[0], q[1])
            ln.line.color.rgb = c
            ln.line.width = Pt(width)

        if kind == "point":
            node(*at(0.5, 0.5), d=int(Inches(0.24)))
        elif kind == "line":
            pts = [at(0.15, 0.72), at(0.5, 0.3), at(0.85, 0.62)]
            seg(pts[0], pts[1]); seg(pts[1], pts[2])
            for p in pts:
                node(*p)
        elif kind == "polygon":
            pts = [at(0.15, 0.32), at(0.7, 0.18), at(0.88, 0.62), at(0.3, 0.86)]
            ff = sh.build_freeform(pts[0][0], pts[0][1])
            ff.add_line_segments(pts[1:], close=True)
            poly = ff.convert_to_shape()
            poly.fill.solid(); poly.fill.fore_color.rgb = b._tint(c, 0.7)
            poly.line.color.rgb = c; poly.line.width = Pt(2)
            b._no_shadow(poly)
            for p in pts:
                node(*p)
        elif kind == "network":
            pts = [at(0.12, 0.2), at(0.55, 0.45), at(0.9, 0.55), at(0.4, 0.88)]
            for i, j in ((0, 1), (1, 2), (1, 3), (0, 3)):
                seg(pts[i], pts[j], 1.4)
            for p in pts:
                node(*p, fill=b.PRIMARY)
        elif kind == "arrow":
            a = sh.add_shape(MSO_SHAPE.RIGHT_ARROW, int(x), int(y + h * 0.3),
                             int(w), int(h * 0.4))
            a.fill.solid(); a.fill.fore_color.rgb = c
            a.line.fill.background(); b._no_shadow(a)

    # drawing --------------------------------------------------------------
    def draw_item(self, it, x, y, w, h, next_is_code=False):
        b = self.b
        t = it["t"]
        if t == "label":
            tb = self.text(x, y, w, h, [it["text"]], self.pt(self.S_LABEL),
                           b.SECONDARY if next_is_code else b.PRIMARY)
            for r in tb.text_frame.paragraphs[0].runs:
                r.font.bold = True
        elif t == "text":
            if self._chips(it):
                self.draw_chips(it, x, y, w, h)
            else:
                self.text(x, y, w, h, it["lines"], self.pt(self.S_BODY), b.FG, bullets=True)
        elif t == "caption":
            self.text(x, y, w, h, it["lines"], self.pt(10.5), b.MUTED)
        elif t == "card":
            self.rect(x, y, w, h, self.card_fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                      radius=0.08, shadow=True)
            pad = int(Inches(0.22))
            tx = x + pad
            if it["icon"]:
                self.icon(it["icon"], x + int(Inches(0.12)), y + int(Inches(0.1)),
                          int(Inches(0.9)), h - int(Inches(0.2)))
                tx = x + int(Inches(1.05)) + pad
            tw = x + w - pad - tx
            lines = ["**" + it["title"] + "**"] + it["body"]
            tb = self.text(tx, y, tw, h, lines, self.pt(self.S_BODY), b.FG,
                           anchor=MSO_ANCHOR.MIDDLE)
            first = tb.text_frame.paragraphs[0]
            for r in first.runs:
                r.font.size = self.pt(self.S_CARD_T)
                r.font.color.rgb = b.SECONDARY
        elif t == "callout":
            self.rect(x, y, w, h, b.LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                      radius=0.08, shadow=True)
            pad = int(Inches(0.22))
            tx = x + pad
            if self._arrow(it):
                self.icon("arrow", x + pad, y + h // 2 - int(Inches(0.22)),
                          int(Inches(0.55)), int(Inches(0.44)))
                tx = x + pad + int(Inches(0.75))
            self.text(tx, y, x + w - pad - tx, h, self._callout_lines(it), self.pt(self.S_CALLOUT),
                      b.FG, anchor=MSO_ANCHOR.MIDDLE)
        elif t == "code":
            self.rect(x, y, w, h, CODE_BG)
            tb = b._add_textbox(self.slide, int(x + self.PAD), int(y + self.PAD),
                                int(w - 2 * self.PAD), int(h - 2 * self.PAD))
            tf = tb.text_frame
            tf.word_wrap = False
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            size = self.pt(self.S_CODE)
            for i, ln in enumerate(it["code"].split("\n")):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                if ln.strip().startswith(("//", "#")) and it["lang"].lower() != "python":
                    b._code_runs(p, ln, "", size, {"base": CODE_COLORS["comment"]})
                else:
                    b._code_runs(p, ln, it["lang"] or "json", size, CODE_COLORS)
                p.line_spacing = Pt(size.pt * 1.35)
        elif t == "image":
            img = b._image_or_placeholder(it["path"])
            if img:
                from PIL import Image
                with Image.open(img) as im:
                    iw, ih = im.size
                s = min(w / iw, h / ih)
                pw, ph = int(iw * s), int(ih * s)
                self.slide.shapes.add_picture(img, int(x + (w - pw) // 2),
                                              int(y + (h - ph) // 2), pw, ph)
        elif t == "steps":
            self.draw_steps(it["items"], x, y, w, h)

    def draw_steps(self, items, x, y, w, h):
        b = self.b
        gap = int(Inches(0.12))
        tw = w - int(Inches(0.75)) - int(Inches(1.1))
        nat = [int(Inches(0.2)) * 2 + self.h([s["title"] + "　" + s["sub"]], self.pt(self.S_STEP_T), tw, bold=True)
               + (self.h(s["body"], self.pt(self.S_STEP_B), tw) if s["body"] else 0) for s in items]
        extra = max(0, h - sum(nat) - gap * (len(items) - 1))
        cur = y
        for i, (s, nh) in enumerate(zip(items, nat), 1):
            sh = nh + extra // max(1, len(items))
            self.rect(x, cur, w, sh, b.LIGHT)
            self.rect(x, cur, w, int(Pt(3)), b.ACCENT)
            d = int(Inches(0.36))
            circ = self.rect(x + int(Inches(0.2)), cur + int(Inches(0.2)), d, d, b.ACCENT,
                             shape=MSO_SHAPE.OVAL)
            cp = circ.text_frame.paragraphs[0]
            b._add_plain_run(cp, str(i), self.pt(12), b.WHITE, bold=True)
            cp.alignment = PP_ALIGN.CENTER
            circ.text_frame.margin_left = circ.text_frame.margin_right = 0
            circ.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            tx = x + int(Inches(0.75))
            tb = self.text(tx, cur + int(Inches(0.2)), tw, sh - int(Inches(0.3)),
                           [""] + s["body"], self.pt(self.S_STEP_B), b.FG, anchor=MSO_ANCHOR.TOP)
            # header: title (navy) + formula / key figure (accent)
            hp = tb.text_frame.paragraphs[0]
            hp.clear()
            b._append_rich_text(hp, s["title"], self.pt(self.S_STEP_T), b.PRIMARY)
            if s["sub"]:
                b._append_rich_text(hp, "　" + s["sub"], self.pt(self.S_STEP_SUB), b.ACCENT)
            for r in hp.runs:
                r.font.bold = True
            if s["icon"]:
                self.icon(s["icon"], x + w - int(Inches(1.05)), cur + int(Inches(0.15)),
                          int(Inches(0.9)), sh - int(Inches(0.3)))
            cur += sh + gap

    def draw_col(self, col, x, y, w, h):
        b = self.b
        if col["panel"]:
            self.rect(x, y, w, h, b.LIGHT)
            self.rect(x, y, int(Pt(3)), h, self.color(col["accent"], b.SECONDARY))
            x, y = x + int(Inches(0.22)) + int(Pt(3)), y + int(self.PAD)
            w, h = w - 2 * int(Inches(0.22)) - int(Pt(3)), h - 2 * int(self.PAD)
        items = self.resolved(col["items"])
        hs = [self.item_h(it, w) for it in items]
        natural = sum(hs) + int(self.ITEM_GAP) * max(0, len(hs) - 1)
        extra = max(0, h - natural)
        grow = [i for i, it in enumerate(items) if it["t"] in ("code", "image", "steps")]
        if not grow:
            grow = [i for i, it in enumerate(items) if it["t"] == "card"]
        for i in grow:
            hs[i] += extra // len(grow)
        cur = y
        for i, (it, ih) in enumerate(zip(items, hs)):
            nxt = items[i + 1]["t"] if i + 1 < len(items) else ""
            self.draw_item(it, x, cur, w, ih, next_is_code=(nxt == "code"))
            cur += ih + int(self.ITEM_GAP)

    def draw_block(self, blk, x, y, w, h):
        b = self.b
        if blk["kind"] == "defn":
            self.rect(x, y, w, h, b._tint(b.SECONDARY, 0.9))
            lines = ["**" + blk["term"] + "**"] + blk["body"]
            tb = self.text(x + int(Inches(0.3)), y, w - int(Inches(0.6)), h, lines,
                           self.pt(self.S_BODY), b.FG, anchor=MSO_ANCHOR.MIDDLE)
            for r in tb.text_frame.paragraphs[0].runs:
                r.font.size = self.pt(self.S_TERM)
                r.font.color.rgb = b.SECONDARY
        elif blk["kind"] == "band":
            self.rect(x, y, w, h, b.LIGHT)
            self.rect(x, y, int(Pt(3)), h, self.color(blk["accent"], b.ACCENT))
            tb = self.text(x + int(Inches(0.28)), y, w - int(Inches(0.5)), h,
                           blk["lines"], self.pt(self.S_BAND), b.FG, anchor=MSO_ANCHOR.MIDDLE)
            for i, p in enumerate(tb.text_frame.paragraphs):
                for r in p.runs:
                    if r.font.bold:
                        r.font.color.rgb = b.PRIMARY if i == 0 else b.ACCENT
        elif blk["kind"] == "steps":
            self.draw_steps(blk["items"], x, y, w, h)
        else:
            ws = self.col_widths(blk["cols"], w)
            cx = x
            for col, cw in zip(blk["cols"], ws):
                self.draw_col(col, cx, y, cw, h)
                cx += cw + int(self.COL_GAP)


def render_board(b, sd) -> None:
    slide = b._blank_slide()
    if sd.h1:
        b._add_title(slide, sd.h1)
    rleft, rtop, rwidth, rheight = b._content_region_with_lead(slide, sd)
    blocks = sd.board or []
    if not blocks:
        return
    gap = int(_Board.GAP)
    for k in (1.0, 0.94, 0.88, 0.82, 0.76, 0.7, 0.64):
        bd = _Board(b, slide, k)
        hs = [bd.block_h(blk, rwidth) for blk in blocks]
        total = sum(hs) + gap * (len(hs) - 1)
        if total <= rheight:
            break
    else:
        b._warn(f'board "{sd.h1}": more than one slide can hold — split it')
    # Zones fill the slide: rows (and step stacks) take the free height in
    # proportion to their size, up to +45% — bands and definitions keep theirs.
    extra = max(0, rheight - total)
    growable = [i for i, blk in enumerate(blocks) if blk["kind"] in ("row", "steps")]
    base = sum(hs[i] for i in growable) or 1
    for i in growable:
        hs[i] += min(int(extra * hs[i] / base), int(hs[i] * 0.45))
    y = rtop
    for blk, h in zip(blocks, hs):
        bd.draw_block(blk, rleft, y, rwidth, h)
        y += h + gap
    if sd.footnote:
        b._add_footnote(slide, sd.footnote)
