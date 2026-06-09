"""PptxBuilder: converts parsed SlideData into PowerPoint slides.

All theme-dependent values come from self.theme (ThemeConfig instance),
eliminating the global state of the original convert_v2.py.
"""
from __future__ import annotations

import re
import sys
import tempfile
import hashlib
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

from marp_pptx.parser import SlideData, strip_html
from marp_pptx.theme import ThemeConfig
from marp_pptx.layout import (
    SW, SH, MARGIN_L, MARGIN_R, MARGIN_T, MARGIN_B, CONTENT_W,
    TITLE_H, TITLE_TOP, TITLE_GAP, KICKER_H, BODY_TOP, BODY_H, FOOTER_RESERVE,
    SZ_DISPLAY, SZ_TITLE, SZ_KICKER, SZ_H2, SZ_H3, SZ_BODY, SZ_COL, SZ_SMALL,
    SZ_FOOT, SZ_METRIC, SZ_EQ, SZ_EQ_VAR, SZ_ZONE_L, SZ_ZONE_B,
    LINE_BODY, LINE_TITLE, PARA_GAP, BLOCK_GAP, SECTION_GAP,
    CARD_GAP, CARD_PAD, CARD_RADIUS, CARD_ACCENT_H, HAIRLINE_W, ACCENT_RULE_W,
)
from marp_pptx.math.omml import latex_to_omml_element, set_omml_size, omml_glyph_count, OmmlError
from marp_pptx.math.renderer import render_latex_png

try:
    import cairosvg
    HAS_CAIROSVG = True
except ImportError:
    HAS_CAIROSVG = False

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


class PptxBuilder:
    def __init__(self, base_path: Path, theme: ThemeConfig):
        self.prs = Presentation()
        self.prs.slide_width = SW
        self.prs.slide_height = SH
        self.base_path = base_path
        self.theme = theme
        self._img_cache: dict = {}
        self._divider_no = 0
        # Non-fatal authoring problems surfaced during build (unknown slide
        # type, missing image, …). Printed once each to stderr and collected
        # here so callers (CLI / web / tests) can report them.
        self.warnings: list[str] = []
        self._warned: set = set()

    def _warn(self, msg: str):
        """Record a non-fatal authoring warning (deduplicated) and echo it to
        stderr. The build still succeeds — this just stops the tool from
        silently producing something other than what the author wrote."""
        if msg in self._warned:
            return
        self._warned.add(msg)
        self.warnings.append(msg)
        print(f"[warn] {msg}", file=sys.stderr)

    def _fs(self, pt_val):
        """Scale a Pt value by theme.font_scale. Min 8pt.

        Accepts either a Pt object or a raw int (points).
        """
        from pptx.util import Pt as _Pt
        scale = getattr(self.theme, "font_scale", 1.0)
        try:
            base = pt_val.pt
        except AttributeError:
            base = float(pt_val)
        return _Pt(max(8, base * scale))

    def _ms(self, inch_val):
        """Scale an Inches value by theme.margin_scale. Min 0.1 inch.

        Accepts either an Inches object (EMU) or a raw float (inches).
        Used for margins, padding, gaps, and box positions/sizes.
        """
        from pptx.util import Inches as _Inches
        scale = getattr(self.theme, "margin_scale", 1.0)
        try:
            base = inch_val.inches
        except AttributeError:
            base = float(inch_val)
        return _Inches(max(0.1, base * scale))

    # ══════════════════════════════════════════════
    # Design-system helpers (refined-minimal, center-balanced)
    # ══════════════════════════════════════════════
    def _tint(self, rgb, t):
        """Blend rgb toward white by t (0=rgb, 1=white). For hairlines/ghosts."""
        return RGBColor(
            int(rgb[0] + (255 - rgb[0]) * t),
            int(rgb[1] + (255 - rgb[1]) * t),
            int(rgb[2] + (255 - rgb[2]) * t),
        )

    @property
    def HAIRLINE(self):
        return getattr(self.theme, "hairline", None) or self._tint(self.MUTED, 0.62)

    @property
    def SURFACE(self):
        # Slightly off-white panel fill that reads against pure-white bg.
        return self.LIGHT

    def _no_shadow(self, shape):
        """Kill python-pptx's heavy default preset shadow (mandatory for minimal)."""
        try:
            shape.shadow.inherit = False
        except Exception:
            pass
        return shape

    def _content_region(self, has_title: bool = True, *, full: bool = False):
        """(left, top, width, height) EMU for the body canvas.

        has_title: start below the title band; else from the top margin.
        full: don't reserve the footer band (figures / full-bleed images).
        Single source of truth — every build_* should start from this.
        """
        left = int(MARGIN_L)
        width = int(CONTENT_W)
        footer = 0 if full else int(FOOTER_RESERVE)
        if has_title:
            top = int(MARGIN_T + TITLE_H + TITLE_GAP)
        else:
            top = int(MARGIN_T)
        height = int(SH - top - MARGIN_B - footer)
        return (left, top, width, height)

    def _stack_tops(self, heights, region, mode="center", gap=None):
        """Return the top (EMU) for each block stacked vertically in region.

        mode: "center" (block group centered), "justify" (free space spread
        between blocks), "top" (legacy top-pile), "fill" (each block expanded
        to share the region equally — caller uses the returned step as height).
        This is the core fix for the "empty bottom 40-60%" defect.
        """
        left, rtop, width, rheight = region
        if gap is None:
            gap = int(BLOCK_GAP)
        n = len(heights)
        if n == 0:
            return []
        content_h = sum(int(h) for h in heights)
        if mode == "justify" and n > 1:
            free = max(0, rheight - content_h)
            g = free // (n - 1)
            tops, y = [], rtop
            for h in heights:
                tops.append(int(y)); y += int(h) + g
            return tops
        if mode == "fill":
            free = max(0, rheight - gap * (n - 1))
            each = free // n
            tops, y = [], rtop
            for _ in heights:
                tops.append(int(y)); y += each + gap
            return tops
        total = content_h + gap * (n - 1)
        if mode == "center":
            y = rtop + max(0, (rheight - total) // 2)
        else:  # "top"
            y = rtop
        tops = []
        for h in heights:
            tops.append(int(y)); y += int(h) + gap
        return tops

    def _vcenter_shapes(self, shapes, region):
        """Shift already-placed shapes so their bounding box is vertically
        centered in region. Use ONLY for shapes with known height (tables,
        pictures, explicit-height textboxes) — never autofit textboxes."""
        shapes = [s for s in shapes if s is not None]
        if not shapes:
            return
        _, rtop, _, rheight = region
        tops = [s.top for s in shapes]
        bottoms = [s.top + s.height for s in shapes]
        bbox_top, bbox_bot = min(tops), max(bottoms)
        offset = rtop + (rheight - (bbox_bot - bbox_top)) // 2 - bbox_top
        for s in shapes:
            s.top = int(s.top + offset)

    def _hairline(self, slide, left, top, width, thickness=None, color=None,
                  vertical=False):
        """Thin accent/divider rule. Horizontal by default."""
        if thickness is None:
            thickness = HAIRLINE_W
        if color is None:
            color = self.HAIRLINE
        w, h = (int(thickness), int(width)) if vertical else (int(width), int(thickness))
        ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(left), int(top), w, h)
        ln.fill.solid()
        ln.fill.fore_color.rgb = color
        ln.line.fill.background()
        self._no_shadow(ln)
        return ln

    _GREEK = {
        "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
        "zeta": "ζ", "eta": "η", "theta": "θ", "iota": "ι", "kappa": "κ",
        "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π", "rho": "ρ",
        "sigma": "σ", "tau": "τ", "phi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
        "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Pi": "Π",
        "Sigma": "Σ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
        "times": "×", "cdot": "·", "leq": "≤", "geq": "≥", "le": "≤", "ge": "≥",
        "in": "∈", "forall": "∀", "infty": "∞", "nabla": "∇", "partial": "∂",
        "sum": "Σ", "approx": "≈", "neq": "≠", "to": "→", "top": "⊤",
        "sqrt": "√", "odot": "⊙", "otimes": "⊗", "oplus": "⊕", "pm": "±",
        "cup": "∪", "cap": "∩", "subset": "⊂", "supset": "⊃", "propto": "∝",
        "sim": "∼", "equiv": "≡", "langle": "⟨", "rangle": "⟩", "prod": "∏",
        "int": "∫", "exists": "∃", "leftarrow": "←", "rightarrow": "→",
    }

    def _math_text_fallback(self, latex: str) -> str:
        """Approximate a LaTeX snippet as plain text/unicode (png-mode fallback
        when OMML is disabled). Maps common greek/operators, strips the rest."""
        s = latex
        for name, ch in self._GREEK.items():
            s = re.sub(r"\\" + name + r"(?![a-zA-Z])", ch, s)
        s = s.replace("\\,", " ").replace("\\;", " ").replace("\\ ", " ")
        s = re.sub(r"[\\{}$]", "", s)
        s = s.replace("_", "").replace("^", "")
        return s.strip()

    def _set_tracking(self, para, spc=160):
        """Wide letter-spacing (1/100 pt units) on a paragraph's runs."""
        for run in para.runs:
            run._r.get_or_add_rPr().set("spc", str(int(spc)))

    def _kicker_para(self, para, text, size=None, color=None, align=PP_ALIGN.LEFT,
                     tracking=180):
        """Small-caps-style label: UPPERCASE (latin only) + wide tracking."""
        if size is None:
            size = SZ_KICKER
        if color is None:
            color = self.ACCENT_TEXT
        t = (text or "")
        para.text = t.upper() if t.isascii() else t
        para.alignment = align
        para.font.name = self.FONT_HEAD
        para.font.size = self._fs(size)
        para.font.bold = True
        para.font.color.rgb = color
        self._set_tracking(para, tracking)
        return para

    def _eyebrow(self, slide, text, left, top, width, align=PP_ALIGN.LEFT,
                 color=None, size=None):
        """Standalone small-caps eyebrow label textbox."""
        tb = self._add_textbox(slide, int(left), int(top), int(width),
                               int(KICKER_H))
        self._kicker_para(tb.text_frame.paragraphs[0], text, size=size,
                          color=color or self.ACCENT_TEXT, align=align)
        return tb

    def _add_card(self, slide, region, *, label="", body="", body_lines=None,
                  value=None, accent=None, accent_bar=False, fill=None,
                  anchor="top", label_size=None, body_size=None,
                  value_size=None, label_color=None):
        """Refined-minimal card: surface fill + hairline border, no drop shadow.

        Elevation cue is the optional top accent bar (accent_bar=True), used
        sparingly (KPI, highlighted panels). Returns (bg_shape, text_shape).
        """
        left, top, width, height = (int(v) for v in region)
        if accent is None:
            accent = self.ACCENT
        if fill is None:
            fill = self.SURFACE
        pad = int(CARD_PAD)
        acc_h = int(CARD_ACCENT_H) if accent_bar else 0

        bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        bg.adjustments[0] = CARD_RADIUS
        bg.fill.solid()
        bg.fill.fore_color.rgb = fill
        bg.line.color.rgb = self.HAIRLINE
        bg.line.width = HAIRLINE_W
        self._no_shadow(bg)

        if accent_bar:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, acc_h)
            bar.fill.solid(); bar.fill.fore_color.rgb = accent
            bar.line.fill.background()
            self._no_shadow(bar)

        tx = left + pad
        ty = top + acc_h + pad
        tw = width - pad * 2
        th = height - acc_h - pad * 2
        tb = self._add_textbox(slide, tx, ty, max(int(Inches(0.3)), tw),
                               max(int(Inches(0.2)), th))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE if anchor == "middle" else MSO_ANCHOR.TOP

        first = True
        if value is not None:
            p = tf.paragraphs[0]; first = False
            run = p.add_run(); run.text = value
            run.font.name = self.FONT_HEAD
            run.font.size = value_size or self._fs(SZ_METRIC)
            run.font.bold = True
            run.font.color.rgb = accent
            p.alignment = PP_ALIGN.CENTER
        if label:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            if not first:
                p.space_before = Pt(6)
            first = False
            align = PP_ALIGN.CENTER if value is not None else PP_ALIGN.LEFT
            self._kicker_para(p, label, size=label_size or SZ_ZONE_L,
                              color=label_color or self.SECONDARY, align=align,
                              tracking=120)
        if body or body_lines:
            if first:
                p2 = tf.paragraphs[0]
            else:
                p2 = tf.add_paragraph(); p2.space_before = Pt(6)
            if body_lines:
                self._fill_multiline_box(tf, "\n".join(body_lines),
                                         body_size or SZ_ZONE_B, self.FG)
            else:
                self._set_text_with_inline_math(p2, body, body_size or SZ_ZONE_B,
                                                self.FG)
        return bg, tb

    def save(self, path: str):
        self._ensure_ea_font()
        self.prs.save(path)

    # ── Theme shortcuts ──
    @property
    def PRIMARY(self): return self.theme.primary
    @property
    def SECONDARY(self): return self.theme.secondary
    @property
    def ACCENT(self): return self.theme.accent
    @property
    def ACCENT_TEXT(self):
        # Darker accent for small text (AA contrast); falls back to ACCENT.
        return getattr(self.theme, "accent_text", None) or self.theme.accent
    @property
    def FG(self): return self.theme.fg
    @property
    def MUTED(self): return self.theme.muted
    @property
    def LIGHT(self): return self.theme.light
    @property
    def WHITE(self): return self.theme.white
    @property
    def FONT(self): return self.theme.font
    @property
    def FONT_HEAD(self): return self.theme.font_head
    @property
    def FONT_EA(self): return self.theme.font_ea
    @property
    def FONT_MONO(self): return self.theme.font_mono
    @property
    def LAYOUT(self): return self.theme.layout

    # ── EA font injection ──
    def _ensure_ea_font(self):
        """Patch every run-property element with an East-Asian typeface so
        CJK text uses the EA font in PowerPoint. Covers slides AND the slide
        master/layout defaults (so inherited runs are handled too)."""
        rpr_tags = (f"{{{NS_A}}}rPr", f"{{{NS_A}}}defRPr", f"{{{NS_A}}}endParaRPr")
        roots = [s._element for s in self.prs.slides]
        for layout in self.prs.slide_layouts:
            roots.append(layout._element)
            roots.append(layout.slide_master._element)
        seen = set()
        for root in roots:
            if id(root) in seen:
                continue
            seen.add(id(root))
            for tag in rpr_tags:
                for rpr in root.iter(tag):
                    self._patch_rpr(rpr)

    def _patch_rpr(self, rpr):
        """Ensure <a:ea> and <a:cs> typefaces exist, in correct schema order
        (latin, ea, cs). Works even when <a:latin> is absent."""
        A = f"{{{NS_A}}}"

        def _insert_ordered(child, after_tags):
            for t in after_tags:
                prev = rpr.find(f"{A}{t}")
                if prev is not None:
                    prev.addnext(child)
                    return
            rpr.insert(0, child)

        ea = rpr.find(f"{A}ea")
        if ea is None:
            ea = etree.Element(f"{A}ea")
            _insert_ordered(ea, ["latin"])
        ea.set("typeface", self.FONT_EA)
        cs = rpr.find(f"{A}cs")
        if cs is None:
            cs = etree.Element(f"{A}cs")
            _insert_ordered(cs, ["ea", "latin"])
        cs.set("typeface", self.FONT_EA)

    # ── Math helpers ──
    def _omml_element(self, latex: str, display: bool, pt_size: int | None = None):
        # Allow the theme to force PNG fallback (preview mode for viewers
        # with poor OMML support like LibreOffice).
        if getattr(self.theme, "math_mode", "omml") == "png":
            return None
        try:
            el = latex_to_omml_element(latex, display=display)
        except OmmlError as e:
            print(f"  OMML failed: {latex[:40]}... ({e}) — falling back to PNG", file=sys.stderr)
            return None
        if pt_size is not None:
            # Pin run size in half-points so every renderer draws it identically.
            eff = max(8.0, float(pt_size) * getattr(self.theme, "font_scale", 1.0))
            set_omml_size(el, int(round(eff * 2)))
        return el

    def _add_math_omml_display(self, slide, latex, left, top, width, pt_size=30):
        el = self._omml_element(latex, display=True, pt_size=pt_size)
        if el is None:
            return None
        from marp_pptx.math.omml import NS_M
        has_frac = el.find(f".//{{{NS_M}}}f") is not None
        line_factor = 2.4 if has_frac else 1.6
        eff_pt = max(8.0, pt_size * getattr(self.theme, "font_scale", 1.0))
        height = Emu(int(eff_pt * line_factor * 12700))   # 1pt = 12700 EMU
        tb = self._add_textbox(slide, left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = ""
        run.font.name = self.FONT
        run.font.size = self._fs(Pt(pt_size))
        run.font.color.rgb = self.FG
        p._p.append(el)
        return tb

    def _append_math_omml_inline(self, para, latex, size, color):
        pt = size.pt if hasattr(size, "pt") else float(size)
        el = self._omml_element(latex, display=False, pt_size=int(round(pt)))
        if el is None:
            return False
        run = para.add_run()
        run.text = ""
        run.font.name = self.FONT
        run.font.size = size
        if color is not None:
            run.font.color.rgb = color
        para._p.append(el)
        return True

    _MATH_DPI = 220   # crisp on retina / when projected

    def _render_math(self, latex: str, display: bool = False, fontsize: int = 28) -> str | None:
        """Render LaTeX to PNG via matplotlib. Returns path to PNG."""
        color_hex = f"#{self.FG}"
        return render_latex_png(latex, fontsize=fontsize, display=display,
                                color=color_hex, dpi=self._MATH_DPI)

    def _add_math_image(self, slide, latex, left, top, max_width, display=True, fontsize=28):
        png = self._render_math(latex, display=display, fontsize=fontsize)
        if not png:
            return None
        from PIL import Image
        with Image.open(png) as im:
            iw, ih = im.size
        dpi = self._MATH_DPI
        pw = int(iw * 914400 / dpi)
        ph = int(ih * 914400 / dpi)
        if pw > max_width:
            scale = max_width / pw
            pw = int(pw * scale)
            ph = int(ph * scale)
        img_left = left + (max_width - pw) // 2
        slide.shapes.add_picture(png, img_left, top, pw, ph)
        return (pw, ph)

    # ── Basic shape helpers ──
    def _blank_slide(self):
        layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(layout)
        # Apply the theme background (e.g. warm cream) to every slide. Hero
        # builders may override this with their own bg.
        bg = self.theme.bg
        if (bg[0], bg[1], bg[2]) != (0xff, 0xff, 0xff):
            self._set_bg(slide, bg)
        return slide

    def _add_textbox(self, slide, left, top, width, height):
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        return tb

    def _set_bg(self, slide, color):
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = color

    def _set_gradient_bg(self, slide, c1, c2):
        bg = slide.background
        bg.fill.gradient()
        bg.fill.gradient_stops[0].color.rgb = c1
        bg.fill.gradient_stops[0].position = 0.0
        bg.fill.gradient_stops[1].color.rgb = c2
        bg.fill.gradient_stops[1].position = 1.0

    def _add_title(self, slide, text, top=None, color=None, kicker=None):
        """Slide title: a small-caps kicker (provides the color cue) + the H1,
        separated from the body by whitespace.

        We deliberately AVOID a decorative accent line under the title — a
        repeated underline rule is a hallmark of AI-generated decks; hierarchy
        comes from whitespace + the kicker + size contrast instead. Optional
        under-title rules are only drawn when a palette explicitly opts in via
        ThemeLayout (h1_deco='bottom-line' for a neutral hairline, or
        accent_rule='short-left' for a colored tick)."""
        if color is None:
            color = self.PRIMARY
        if top is None:
            top = TITLE_TOP
        text_left = int(MARGIN_L)
        text_w = int(CONTENT_W)

        # Small-caps eyebrow — the title's color accent without a line.
        if kicker:
            kb = self._add_textbox(slide, text_left, int(top - KICKER_H + Inches(0.02)),
                                   text_w, int(KICKER_H))
            self._kicker_para(kb.text_frame.paragraphs[0], kicker, color=self.ACCENT_TEXT)

        left_bar = self.LAYOUT.h1_deco == "left-bar"
        if left_bar:
            deco_w = Pt(6)
            bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, int(MARGIN_L), int(top), int(deco_w), int(TITLE_H))
            bar.fill.solid(); bar.fill.fore_color.rgb = self.ACCENT
            bar.line.fill.background(); self._no_shadow(bar)
            text_left = int(MARGIN_L + deco_w + Pt(12))
            text_w = int(CONTENT_W - deco_w - Pt(12))

        tb = self._add_textbox(slide, text_left, int(top), text_w, int(TITLE_H))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = self.FONT_HEAD
        p.font.size = self._fs(SZ_TITLE)
        p.font.bold = True
        p.font.color.rgb = color
        p.line_spacing = LINE_TITLE

        # Opt-in under-title rules only (off by default to avoid the AI-deck look).
        if not left_bar:
            rule_y = int(top + TITLE_H + Inches(0.04))
            if self.LAYOUT.h1_deco == "bottom-line":
                self._hairline(slide, MARGIN_L, rule_y, CONTENT_W, color=self.HAIRLINE)
            if self.LAYOUT.accent_rule == "short-left":
                self._hairline(slide, MARGIN_L, rule_y, Inches(0.7),
                               thickness=ACCENT_RULE_W, color=self.ACCENT)
        return tb

    def _add_para(self, tf, text, size=None, color=None, bold=False, italic=False, space_before=Pt(4)):
        if size is None:
            size = SZ_BODY
        if color is None:
            color = self.FG
        p = tf.add_paragraph()
        p.text = text
        p.font.name = self.FONT
        p.font.size = size
        p.font.color.rgb = color
        p.font.bold = bold
        p.space_before = space_before
        return p

    @staticmethod
    def _visual_em_width(s: str) -> float:
        """Approximate display width of a string in em units (CJK-aware:
        full-width chars count 1.0 em, everything else ~0.55 em)."""
        import unicodedata
        return sum(1.0 if unicodedata.east_asian_width(c) in ("W", "F") else 0.55
                   for c in s)

    def _estimate_text_height(self, lines, size, width=None):
        """Tight height for a block of markdown lines at given font size.

        Slightly over-estimates so the shape hugs content but never clips.
        When `width` (EMU) is given, soft wrapping of long lines is taken
        into account, so callers stacking blocks below this one don't
        overlap when a line wraps (e.g. a long Japanese takeaway).
        """
        from pptx.util import Pt as _Pt
        base = size.pt if hasattr(size, "pt") else float(size)
        scale = getattr(self.theme, "font_scale", 1.0)
        line_h = base * 1.25 * scale
        cap_em = None
        if width:
            width_pt = width / 12700.0          # EMU -> pt
            cap_em = max(4.0, width_pt / (base * scale))
        total = 0.0
        first = True
        for line in lines:
            s = line.strip()
            if not s:
                continue
            mult = 1.35 if s.startswith(("## ", "### ")) else 1.0
            wraps = 1
            if cap_em:
                plain = s.replace("**", "").replace("`", "")
                wraps = max(1, -(-int(self._visual_em_width(plain) * 100) // int(cap_em * 100)))
            total += line_h * mult * wraps
            if not first:
                total += 4 * scale
            first = False
        # +6pt tail breathing room to prevent clipping when autofit is off
        return _Pt(max(18, total + 6))

    def _add_body_text(self, slide, lines, left=None, top=None, width=None, height=None, size=None):
        if size is None:
            size = SZ_BODY
        if left is None: left = MARGIN_L
        if top is None: top = BODY_TOP
        if width is None: width = CONTENT_W
        explicit_height = height is not None
        if height is None:
            estimated = self._estimate_text_height(lines, size)
            height = min(BODY_H, int(estimated))

        tb = self._add_textbox(slide, left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = True
        # When the caller specified a height (zone reservation), lock the size
        # so viewers that ignore spAutoFit don't grow the textbox and break
        # downstream placement (conclusion boxes, footnotes, etc.).
        if explicit_height:
            tf.auto_size = MSO_AUTO_SIZE.NONE

        first = True
        for line in lines:
            s = line.strip()
            if not s:
                continue
            if first:
                p = tf.paragraphs[0]
                first = False
            else:
                p = tf.add_paragraph()

            is_h2 = s.startswith("## ")
            is_h3 = s.startswith("### ")
            is_bullet = s.startswith("- ") or s.startswith("* ")
            is_numbered = re.match(r"^\d+\.\s", s)

            if is_h2:
                p.text = strip_html(s[3:])
                p.font.name = self.FONT_HEAD
                p.font.size = self._fs(SZ_H2)
                p.font.bold = True
                p.font.color.rgb = self.SECONDARY
                p.space_before = Pt(10)
            elif is_h3:
                p.text = strip_html(s[4:])
                p.font.name = self.FONT_HEAD
                p.font.size = self._fs(SZ_H3)
                p.font.bold = True
                p.font.color.rgb = self.MUTED
                p.space_before = Pt(6)
            elif is_bullet:
                # Apply rich text (bold markdown **...**) to bullet content
                self._set_rich_text(p, s[2:], size, self.FG)
                p.level = 0
                p.space_before = Pt(4)
                pPr = p._p.get_or_add_pPr()
                buChar = pPr.makeelement(qn("a:buChar"), {"char": "\u2022"})
                for existing in pPr.findall(qn("a:buChar")):
                    pPr.remove(existing)
                for existing in pPr.findall(qn("a:buNone")):
                    pPr.remove(existing)
                pPr.append(buChar)
            elif is_numbered:
                p.text = s
                p.font.name = self.FONT
                p.font.size = size
                p.font.color.rgb = self.FG
                p.space_before = Pt(4)
            else:
                self._set_rich_text(p, s, size, self.FG)
                p.space_before = Pt(4)

        # Only enable autofit when caller left sizing to us
        if not explicit_height:
            tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        return tb

    def _add_plain_run(self, para, text, size, color, bold=False, mono=False):
        """Append a single styled run to para. No-op if text is empty."""
        if not text:
            return
        run = para.add_run()
        run.text = text
        run.font.name = self.FONT_MONO if mono else self.FONT
        run.font.size = self._fs(size) if hasattr(size, "pt") else size
        if mono:
            # Make inline code visually distinct from prose.
            run.font.color.rgb = self.SECONDARY
        elif color is not None:
            run.font.color.rgb = color
        if bold:
            run.font.bold = True

    # Combined inline markup: **bold**, `code`, $math$
    _RICH_PATTERN = re.compile(
        r"(\*\*[^\*\n]+?\*\*)"
        r"|(`[^`\n]+?`)"
        r"|(\$[^\$\n]+?\$)"
    )

    def _set_rich_text(self, para, text, size=None, color=None):
        """Render inline markup in a SINGLE paragraph / SINGLE textbox.

        Handles **bold**, `code` (monospace), and $math$ (OMML) without
        breaking the containing textbox. Runs co-exist with Japanese+Latin
        mixed text via the ea-font patch applied at save time.
        """
        if size is None:
            size = SZ_BODY
        if color is None:
            color = self.FG
        para.clear()
        if not text:
            return

        pos = 0
        for m in self._RICH_PATTERN.finditer(text):
            if m.start() > pos:
                self._add_plain_run(para, text[pos:m.start()], size, color)
            if m.group(1):  # **bold**
                self._add_plain_run(para, m.group(1)[2:-2], size, color, bold=True)
            elif m.group(2):  # `code`
                self._add_plain_run(para, m.group(2)[1:-1], size, color, mono=True)
            elif m.group(3):  # $math$
                latex = m.group(3)[1:-1]
                if not self._append_math_omml_inline(para, latex, size, color):
                    # OMML disabled (png preview mode): show the symbol without
                    # the $ delimiters, italicized, as a readable approximation.
                    run = para.add_run()
                    run.text = self._math_text_fallback(latex)
                    run.font.name = self.FONT
                    run.font.size = self._fs(size) if hasattr(size, "pt") else size
                    run.font.italic = True
                    if color is not None:
                        run.font.color.rgb = color
            pos = m.end()
        if pos < len(text):
            self._add_plain_run(para, text[pos:], size, color)

    # Backward-compatible alias for callers that used the math-only name
    def _set_text_with_inline_math(self, para, text, size, color):
        return self._set_rich_text(para, text, size=size, color=color)

    def _resolve_image(self, img_path: str) -> str | None:
        p = self.base_path / img_path
        if not p.exists():
            self._warn(f"image not found, skipping: {img_path}")
            return None
        if p.suffix.lower() == ".svg":
            png_path = p.with_suffix(".png")
            if not png_path.exists() or png_path.stat().st_mtime < p.stat().st_mtime:
                if HAS_CAIROSVG:
                    cairosvg.svg2png(url=str(p), write_to=str(png_path), output_width=1400, dpi=300)
                else:
                    self._warn(f"SVG needs cairosvg to embed, skipping: {img_path} "
                               "(install marp-pptx[ingest] or pre-render to PNG)")
                    return None
            return str(png_path)
        return str(p)

    def _fill_multiline_box(self, tf, text, size, color):
        """Fill a textbox with text that may contain bullets and continuation lines.

        Handles:
        - `- item` and `* item` as bullets
        - Continuation lines indented MORE than their parent bullet
        - **bold** markdown via _set_rich_text
        """
        # Strip common leading whitespace (dedent) so HTML-source indentation
        # doesn't get mistaken for content continuation.
        raw = [line for line in text.split("\n") if line.strip()]
        if not raw:
            return
        indents = [len(l) - len(l.lstrip()) for l in raw]
        base_indent = min(indents) if indents else 0
        dedented = [l[base_indent:] if len(l) >= base_indent else l for l in raw]

        # Merge continuation: a line is a continuation only if the previous
        # line was a bullet AND this line is further indented.
        merged: list[str] = []
        last_was_bullet = False
        for line in dedented:
            stripped = line.lstrip()
            line_indent = len(line) - len(stripped)
            is_bullet = stripped.startswith("- ") or stripped.startswith("* ")
            if last_was_bullet and line_indent > 0 and not is_bullet and merged:
                merged[-1] = merged[-1] + " " + stripped
            else:
                merged.append(stripped)
                last_was_bullet = is_bullet

        first = True
        for line in merged:
            s = line.strip()
            if not s:
                continue
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            is_bullet = s.startswith("- ") or s.startswith("* ")
            if is_bullet:
                self._set_rich_text(p, s[2:], size, color)
                p.space_before = Pt(4)
                pPr = p._p.get_or_add_pPr()
                buChar = pPr.makeelement(qn("a:buChar"), {"char": "\u2022"})
                for existing in pPr.findall(qn("a:buChar")):
                    pPr.remove(existing)
                pPr.append(buChar)
            else:
                self._set_rich_text(p, s, size, color)
                p.space_before = Pt(4)

    def _add_accent_box(self, slide, text, left, top, width, height, border_color=None):
        if border_color is None:
            border_color = self.ACCENT
        left, top, width, height = int(left), int(top), int(width), int(height)
        bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        bg.adjustments[0] = CARD_RADIUS
        bg.fill.solid(); bg.fill.fore_color.rgb = self.SURFACE
        bg.line.color.rgb = self.HAIRLINE; bg.line.width = HAIRLINE_W
        self._no_shadow(bg)
        bdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, int(Pt(4)), height)
        bdr.fill.solid(); bdr.fill.fore_color.rgb = border_color
        bdr.line.fill.background(); self._no_shadow(bdr)
        tb = self._add_textbox(slide, left + int(Pt(18)), top + int(Pt(8)),
                               width - int(Pt(34)), height - int(Pt(16)))
        tf = tb.text_frame; tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        self._fill_multiline_box(tf, text, SZ_ZONE_B, self.FG)
        return tb

    def _add_conclusion_box(self, slide, text, left, top, width, height):
        left, top, width, height = int(left), int(top), int(width), int(height)
        bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        bg.adjustments[0] = CARD_RADIUS
        bg.fill.solid(); bg.fill.fore_color.rgb = self.SURFACE
        bg.line.color.rgb = self.HAIRLINE; bg.line.width = HAIRLINE_W
        self._no_shadow(bg)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, int(CARD_ACCENT_H))
        bar.fill.solid(); bar.fill.fore_color.rgb = self.ACCENT
        bar.line.fill.background(); self._no_shadow(bar)
        tb = self._add_textbox(slide, left + int(Pt(18)), top + int(Pt(10)),
                               width - int(Pt(36)), height - int(Pt(20)))
        tf = tb.text_frame; tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        self._fill_multiline_box(tf, text, SZ_ZONE_B, self.FG)
        return tb

    def _add_footnote(self, slide, text):
        left = int(MARGIN_L)
        top = int(SH - Inches(0.62))
        width = int(CONTENT_W)
        self._hairline(slide, left, top, Inches(2.8), color=self.HAIRLINE)
        tb = self._add_textbox(slide, left, top + int(Pt(5)), width, int(Inches(0.4)))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.name = self.FONT
        p.font.size = self._fs(SZ_FOOT)
        p.font.color.rgb = self.MUTED
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

    def _add_zone_box(self, slide, left, top, width, height,
                      label="", body="", fill_color=None, fill=None,
                      label_size=None, body_size=None,
                      accent_bar=False, accent=None, anchor="top"):
        """Backward-compatible wrapper over the refined-minimal card.

        Every zone-based builder (zone-flow/compare/matrix/process, steps,
        stack, card-grid, split-text, before-after, multi-result) routes
        through here, so they all gain the visible surface + hairline at once.
        """
        return self._add_card(
            slide, (left, top, width, height),
            label=label, body=body, fill=fill if fill is not None else fill_color,
            label_size=label_size, body_size=body_size,
            accent_bar=accent_bar, accent=accent, anchor=anchor,
        )

    def _set_cell_border(self, cell, edges=("bottom",), color=None, pt=0.75):
        """Inject cell borders via lxml (python-pptx has no public API)."""
        if color is None:
            color = self.HAIRLINE
        hexcol = str(color)
        tcPr = cell._tc.get_or_add_tcPr()
        tag = {"left": "a:lnL", "right": "a:lnR", "top": "a:lnT", "bottom": "a:lnB"}
        width_emu = int(Pt(pt))
        for edge in edges:
            qname = qn(tag[edge])
            for old in tcPr.findall(qname):
                tcPr.remove(old)
            ln = tcPr.makeelement(qname, {"w": str(width_emu), "cap": "flat",
                                          "cmpd": "sng", "algn": "ctr"})
            fill = ln.makeelement(qn("a:solidFill"), {})
            clr = fill.makeelement(qn("a:srgbClr"), {"val": hexcol})
            fill.append(clr); ln.append(fill)
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": "solid"}))
            tcPr.insert(0, ln)

    def _styled_table(self, slide, rows_data, left, top, width, height):
        """Themed table: optional filled header, zebra body, hairline rules.
        Header style follows theme.layout.table_header_style (fill | rule)."""
        rows = len(rows_data)
        cols = max(len(r) for r in rows_data) if rows_data else 1
        gf = slide.shapes.add_table(rows, cols, int(left), int(top), int(width), int(height))
        table = gf.table
        try:
            table.first_row = False
            table.horz_banding = False
            tblPr = table._tbl.find(qn("a:tblPr"))
            if tblPr is not None:
                for sid in tblPr.findall(qn("a:tableStyleId")):
                    tblPr.remove(sid)
        except Exception:
            pass
        rule_header = getattr(self.LAYOUT, "table_header_style", "fill") == "rule"
        # Right-align columns whose body cells are predominantly numeric
        # (numbers read better right/decimal-aligned — CSE table guidance).
        num_re = re.compile(r"^[\s$¥€£]*[-+]?[\d,]+(\.\d+)?\s*[%x×kKMB]?\s*$")
        clean = lambda s: strip_html(s).replace("**", "").strip()
        numeric_cols = set()
        for ci in range(cols):
            body = [clean(r[ci]) for r in rows_data[1:] if ci < len(r) and r[ci].strip()]
            if body and sum(bool(num_re.match(b)) for b in body) >= max(1, len(body) * 0.6):
                numeric_cols.add(ci)
        for ri, row in enumerate(rows_data):
            for ci in range(cols):
                cell = table.cell(ri, ci)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = Pt(10); cell.margin_right = Pt(10)
                cell.margin_top = Pt(5); cell.margin_bottom = Pt(5)
                raw = row[ci] if ci < len(row) else ""
                cell.text = clean(raw)
                is_head = (ri == 0)
                emph = ("**" in raw)   # markdown-bolded cell (e.g. **Ours**)
                align = PP_ALIGN.RIGHT if ci in numeric_cols else PP_ALIGN.LEFT
                for p in cell.text_frame.paragraphs:
                    p.alignment = align
                    p.font.name = self.FONT_HEAD if (is_head or emph) else self.FONT
                    p.font.size = self._fs(SZ_SMALL)
                    p.font.bold = is_head or emph or (ci == 0 and not is_head)
                    if emph and not is_head:
                        p.font.color.rgb = self.ACCENT_TEXT
                    if is_head:
                        p.font.color.rgb = self.FG if rule_header else self.WHITE
                    elif ci == 0:
                        p.font.color.rgb = self.SECONDARY
                    else:
                        p.font.color.rgb = self.FG
                cell.fill.solid()
                if is_head and not rule_header:
                    cell.fill.fore_color.rgb = self.PRIMARY
                elif (not is_head) and ri % 2 == 0:
                    cell.fill.fore_color.rgb = self.LIGHT
                else:
                    cell.fill.fore_color.rgb = self.WHITE
                if is_head:
                    self._set_cell_border(cell, ("bottom",), color=self.PRIMARY, pt=1.5)
                else:
                    self._set_cell_border(cell, ("bottom",), color=self.HAIRLINE, pt=0.5)
        return gf

    def _add_column_content(self, slide, lines, left, top, width, height, size=None):
        if size is None:
            size = SZ_COL
        text_lines = []
        images = []
        for line in lines:
            img_m = re.match(r"!\[(?:w:\d+)?\]\(([^)]+)\)", line.strip())
            if img_m:
                images.append(img_m.group(1))
            else:
                text_lines.append(line)
        cur_top = top
        for img_path in images:
            img_file = self._resolve_image(img_path)
            if img_file:
                from PIL import Image
                with Image.open(img_file) as im:
                    iw, ih = im.size
                max_w = int(width * 0.95)
                max_h = int(Inches(2.5))
                scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
                pw = int(iw * scale * 914400 / 96)
                ph = int(ih * scale * 914400 / 96)
                img_left = left + (width - pw) // 2
                slide.shapes.add_picture(img_file, img_left, cur_top, pw, ph)
                cur_top += ph + Inches(0.1)
        remaining_h = top + height - cur_top
        if text_lines and remaining_h > 0:
            # Hug content height (capped at remaining), don't force full column
            estimated = self._estimate_text_height(text_lines, size)
            use_h = min(int(remaining_h), int(estimated))
            tb = self._add_body_text(slide, text_lines, left=left, top=int(cur_top),
                                     width=int(width), height=use_h, size=size)
            # Lock the column textbox size — some viewers (Keynote, LibreOffice)
            # ignore spAutoFit and render the stored height. Disable autofit so
            # the zone stays predictable for downstream placement.
            if tb is not None:
                tb.text_frame.auto_size = MSO_AUTO_SIZE.NONE

    # ══════════════════════════════════════════════
    # Slide type builders
    # ══════════════════════════════════════════════

    def _hero_meta(self, sd):
        """Split a title slide's body into (subtitle_lines, meta_tokens).
        A line with a 4-digit year or an author/date marker becomes meta."""
        subs, meta = [], []
        past_h1 = False
        for line in sd.raw.split("\n"):
            s = line.strip()
            if s.startswith("# ") and not past_h1:
                past_h1 = True
                continue
            if s.startswith("## "):
                continue   # H2 is used as the kicker
            if past_h1 and s and not s.startswith(("<!--", "<div", "</div", ">")):
                t = strip_html(s)
                if not t:
                    continue
                if re.search(r"\b(19|20)\d{2}\b", t) or t.startswith(("by ", "@", "—", "·")):
                    meta.append(t.lstrip("—@·· ").strip())
                else:
                    subs.append(t)
        return subs, meta

    def build_title(self, sd: SlideData):
        slide = self._blank_slide()
        if self.LAYOUT.title_bg == "gradient":
            self._set_gradient_bg(slide, self.PRIMARY, self.SECONDARY)
        elif self.LAYOUT.title_bg == "dark":
            self._set_bg(slide, self.PRIMARY)
        elif self.LAYOUT.title_bg == "light":
            self._set_bg(slide, self.LIGHT)
        is_dark = self.LAYOUT.title_bg in ("gradient", "dark")
        h_color = self.WHITE if is_dark else self.PRIMARY
        sub_color = RGBColor(0xD8, 0xD8, 0xDE) if is_dark else self.MUTED
        accent = self.WHITE if is_dark else self.ACCENT           # for the rule (graphic)
        kicker_color = self.WHITE if is_dark else self.ACCENT_TEXT  # for small text
        cx = int(SW // 2)
        align = PP_ALIGN.CENTER

        subs, meta = self._hero_meta(sd)
        kicker = sd.h2 if (sd.h2 and len(sd.h2.split()) <= 6) else None

        # Vertically center the kicker→title→rule→subtitle→meta stack.
        # (1) kicker
        if kicker:
            kb = self._add_textbox(slide, int(Inches(1.0)), int(Inches(2.55)),
                                   int(SW - Inches(2.0)), int(KICKER_H))
            self._kicker_para(kb.text_frame.paragraphs[0], kicker, color=kicker_color,
                              align=align, tracking=200)
        # (2) accent hairline above the title
        self._hairline(slide, int(cx - Inches(0.55)), int(Inches(3.16)),
                       Inches(1.1), thickness=Pt(1.6), color=accent)
        # (3) title
        title_h = int(self._fs(Pt(SZ_DISPLAY.pt * 1.25 * 2)))
        tb = self._add_textbox(slide, int(Inches(0.8)), int(Inches(3.34)),
                               int(SW - Inches(1.6)), title_h)
        tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
        p = tf.paragraphs[0]; p.text = sd.h1
        p.font.name = self.FONT_HEAD; p.font.size = self._fs(SZ_DISPLAY)
        p.font.bold = True; p.font.color.rgb = h_color; p.alignment = align
        p.line_spacing = LINE_TITLE
        # (4) subtitle — rich text so $math$ (e.g. author superscripts) renders
        if subs:
            sb = self._add_textbox(slide, int(Inches(1.2)), int(Inches(5.05)),
                                   int(SW - Inches(2.4)), int(Inches(0.95)))
            sb.text_frame.word_wrap = True
            for i, line in enumerate(subs[:3]):
                sp = sb.text_frame.paragraphs[0] if i == 0 else sb.text_frame.add_paragraph()
                self._set_rich_text(sp, line, SZ_BODY, sub_color)
                sp.alignment = align
                if i:
                    sp.space_before = Pt(4)
        # (5) author / affiliation / date
        if meta:
            mb = self._add_textbox(slide, int(Inches(1.0)), int(Inches(6.35)),
                                   int(SW - Inches(2.0)), int(Inches(0.45)))
            mp = mb.text_frame.paragraphs[0]
            self._set_rich_text(mp, "   ·   ".join(meta), SZ_SMALL, sub_color)
            mp.alignment = align

    def build_divider(self, sd: SlideData):
        slide = self._blank_slide()
        is_center = self.LAYOUT.divider_align != "left"
        align = PP_ALIGN.CENTER if is_center else PP_ALIGN.LEFT
        cx = int(SW // 2)
        # ghost section number
        m = re.match(r"^\s*0?(\d{1,2})\b", sd.h2 or "")
        ghost = (m.group(1) if m else str(getattr(self, "_divider_no", 1))).zfill(2)
        if self.LAYOUT.divider_number:
            gb = self._add_textbox(slide, int(Inches(0.85)), int(Inches(0.85)),
                                   int(Inches(5.0)), int(Inches(3.2)))
            gp = gb.text_frame.paragraphs[0]; gp.text = ghost
            gp.font.name = self.FONT_HEAD; gp.font.size = self._fs(Pt(150))
            gp.font.bold = True; gp.font.color.rgb = self._tint(self.PRIMARY, 0.88)
            gp.alignment = PP_ALIGN.LEFT
        x = int(Inches(1.0)) if is_center else int(Inches(1.4))
        w = int(SW - Inches(2.0)) if is_center else int(SW - Inches(2.8))
        # kicker
        kb = self._add_textbox(slide, x, int(Inches(3.5)), w, int(KICKER_H))
        self._kicker_para(kb.text_frame.paragraphs[0],
                          ("Section " + ghost) if self.LAYOUT.divider_number else "Section",
                          color=self.ACCENT_TEXT, align=align, tracking=200)
        # title
        tb = self._add_textbox(slide, x, int(Inches(3.92)), w,
                               int(self._fs(Pt(34 * 1.3 * 2))))
        tf = tb.text_frame; tf.word_wrap = True
        tp = tf.paragraphs[0]; tp.text = sd.h1
        tp.font.name = self.FONT_HEAD; tp.font.size = self._fs(Pt(34))
        tp.font.bold = True; tp.font.color.rgb = self.PRIMARY; tp.alignment = align
        tp.line_spacing = LINE_TITLE
        # accent rule
        rx = cx if is_center else int(Inches(1.4) + Inches(0.35))
        self._hairline(slide, int(rx - Inches(0.35)), int(Inches(4.95)), Inches(0.7),
                       thickness=ACCENT_RULE_W, color=self.ACCENT)
        # subtitle (only when h2 wasn't just the number)
        if sd.h2 and not m:
            sb = self._add_textbox(slide, x, int(Inches(5.18)), w, int(Inches(0.6)))
            sb.text_frame.word_wrap = True
            sp = sb.text_frame.paragraphs[0]; sp.text = sd.h2
            sp.font.name = self.FONT; sp.font.size = self._fs(SZ_BODY)
            sp.font.color.rgb = self.MUTED; sp.alignment = align

    def build_default(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        region = self._content_region(has_title=bool(sd.h1))
        left, _, width, _ = region

        # Measure each block, then center/justify the stack in the body region.
        blocks = []  # (kind, height)
        if sd.body_lines:
            bh = int(self._estimate_text_height(sd.body_lines, SZ_BODY))
            blocks.append(("body", min(bh, region[3])))
        if sd.table_rows:
            th = int(min(Inches(4.6), Inches(0.5) * len(sd.table_rows)))
            blocks.append(("table", th))
        if sd.bottom_text:
            blocks.append(("accent", int(Inches(1.0))))

        heights = [h for _, h in blocks]
        mode = "center" if len(heights) <= 2 else "justify"
        tops = self._stack_tops(heights, region, mode=mode)
        for (kind, h), t in zip(blocks, tops):
            if kind == "body":
                self._add_body_text(slide, sd.body_lines, left=left, top=t,
                                    width=width, height=h)
            elif kind == "table":
                self._styled_table(slide, sd.table_rows, left, t, width, h)
            elif kind == "accent":
                self._add_accent_box(slide, sd.bottom_text, left, t, width, h)
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_equation(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        region = self._content_region(has_title=bool(sd.h1))
        _, rtop, _, rheight = region
        EQ_PT = 34
        # Estimate the equation block + variable legend, center the stack.
        eq_h_est = int(Emu(int(EQ_PT * 2.4 * 12700)))
        var_n = len(sd.eq_vars)
        row_h = int(Inches(0.58))
        vars_h = row_h * var_n
        gap = int(Inches(0.35))
        total = eq_h_est + (gap + vars_h if var_n else 0)
        eq_top = rtop + max(0, (rheight - total) // 2)
        omml_box = self._add_math_omml_display(slide, sd.eq_main, MARGIN_L, eq_top, CONTENT_W, pt_size=EQ_PT)
        if omml_box is not None:
            var_top = eq_top + omml_box.height + Inches(0.25)
        else:
            result = self._add_math_image(slide, sd.eq_main, MARGIN_L, eq_top, CONTENT_W, display=True, fontsize=EQ_PT)
            if result:
                _, eq_h = result
                var_top = eq_top + eq_h + Inches(0.3)
            else:
                tb = self._add_textbox(slide, MARGIN_L, eq_top, CONTENT_W, Inches(1.2))
                tf = tb.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = sd.eq_main
                p.font.name = self.FONT
                p.font.size = self._fs(Pt(28))
                p.font.color.rgb = self.FG
                p.alignment = PP_ALIGN.CENTER
                var_top = eq_top + Inches(1.6)
        if sd.eq_vars:
            self._add_var_legend(slide, sd.eq_vars, var_top, sym_pt=20, desc_pt=17)
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def _add_var_legend(self, slide, eq_vars, top, sym_pt=18, desc_pt=15,
                        row_h=None):
        """Refined "where: <sym> | <description>" legend, centered as a group:
        right-aligned accent symbol, thin vertical hairline, left description."""
        if row_h is None:
            row_h = int(Inches(0.5))
        sym_w = int(Inches(1.7))
        desc_w = int(Inches(6.2))
        group_w = sym_w + int(Inches(0.5)) + desc_w
        gx = int(MARGIN_L + (CONTENT_W - group_w) // 2)
        # "where" eyebrow
        self._eyebrow(slide, "where", gx, int(top - KICKER_H + Inches(0.04)),
                      sym_w, align=PP_ALIGN.RIGHT, color=self.MUTED)
        n = len(eq_vars)
        # vertical hairline divider spanning the rows
        div_x = gx + sym_w + int(Inches(0.25))
        self._hairline(slide, div_x, int(top + Inches(0.04)),
                       int(row_h * n - Inches(0.08)), thickness=HAIRLINE_W,
                       color=self.HAIRLINE, vertical=True)
        for vi, (sym, desc) in enumerate(eq_vars):
            row_top = int(top + row_h * vi)
            sym_latex = sym.strip().strip("$")
            stb = self._add_textbox(slide, gx, row_top, sym_w, row_h)
            stb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            sp = stb.text_frame.paragraphs[0]
            sp.alignment = PP_ALIGN.RIGHT
            if not self._append_math_omml_inline(sp, sym_latex, Pt(sym_pt), self.ACCENT_TEXT):
                sp.text = self._math_text_fallback(sym_latex)
                sp.font.name = self.FONT; sp.font.size = self._fs(Pt(sym_pt - 2))
                sp.font.bold = True; sp.font.italic = True
                sp.font.color.rgb = self.ACCENT_TEXT
            dtb = self._add_textbox(slide, div_x + int(Inches(0.25)), row_top, desc_w, row_h)
            dtb.text_frame.word_wrap = True
            dtb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            self._set_text_with_inline_math(dtb.text_frame.paragraphs[0], desc,
                                            Pt(desc_pt), self.FG)

    def build_equations(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.eq_system)
        if n == 0:
            return
        region = self._content_region(has_title=bool(sd.h1))
        _, rtop, _, rheight = region
        label_left = MARGIN_L + Inches(0.3)
        label_w = Inches(1.9)
        eq_left = label_left + label_w + Inches(0.2)
        eq_w = CONTENT_W - (eq_left - MARGIN_L) - Inches(0.3)
        vars_h = int(Inches(0.56)) * max(len(sd.eq_vars), 0)
        row_h = int(min(Inches(1.2), max(Inches(0.8), (rheight - vars_h - int(Inches(0.25))) // n)))
        pt_size = 30 if n <= 3 else (26 if n <= 4 else 22)
        # Center the whole rows+legend block in the body region.
        block_h = row_h * n + (int(Inches(0.25)) + vars_h if sd.eq_vars else 0)
        top = rtop + max(0, (rheight - block_h) // 2)
        for i, (label, latex) in enumerate(sd.eq_system):
            row_top = top + int(row_h * i)
            if label:
                ltb = self._add_textbox(slide, label_left, row_top, label_w, row_h)
                ltf = ltb.text_frame
                ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
                ltf.word_wrap = True
                lp = ltf.paragraphs[0]
                lp.text = label
                lp.font.name = self.FONT
                lp.font.size = self._fs(Pt(max(14, pt_size - 10)))
                lp.font.color.rgb = self.SECONDARY
                lp.alignment = PP_ALIGN.RIGHT
            el = self._omml_element(latex, display=True, pt_size=pt_size)
            etb = self._add_textbox(slide, eq_left, row_top, eq_w, row_h)
            etf = etb.text_frame
            etf.word_wrap = True
            etf.vertical_anchor = MSO_ANCHOR.MIDDLE
            ep = etf.paragraphs[0]
            ep.alignment = PP_ALIGN.LEFT
            erun = ep.add_run()
            erun.text = ""
            erun.font.name = self.FONT
            erun.font.size = self._fs(Pt(pt_size))
            erun.font.color.rgb = self.FG
            if el is not None:
                ep._p.append(el)
            else:
                etb.element.getparent().remove(etb.element)
                result = self._add_math_image(slide, latex, eq_left, row_top, eq_w, display=True, fontsize=pt_size)
                if not result:
                    tb2 = self._add_textbox(slide, eq_left, row_top, eq_w, row_h)
                    tf2 = tb2.text_frame
                    tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
                    p2 = tf2.paragraphs[0]
                    p2.text = latex
                    p2.font.name = self.FONT
                    p2.font.size = self._fs(Pt(pt_size - 4))
                    p2.font.color.rgb = self.FG
        if sd.eq_vars:
            var_top = top + int(row_h * n) + int(Inches(0.32))
            self._add_var_legend(slide, sd.eq_vars, var_top, sym_pt=18, desc_pt=15,
                                 row_h=int(Inches(0.46)))
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_columns(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        region = self._content_region(has_title=bool(sd.h1))
        rleft, rtop, rwidth, rheight = region
        n = len(sd.columns)
        if n == 0:
            return
        cls = sd.slide_class or ""
        gap = int(CARD_GAP)
        if cls == "cols-2-wide-l":
            widths = [int(rwidth * 0.62) - gap // 2, int(rwidth * 0.38) - gap // 2]
        elif cls == "cols-2-wide-r":
            widths = [int(rwidth * 0.38) - gap // 2, int(rwidth * 0.62) - gap // 2]
        else:
            cw = (rwidth - gap * (n - 1)) // n
            widths = [cw] * n
        size = SZ_COL
        x = rleft
        for i, col_lines in enumerate(sd.columns):
            col_w = widths[i] if i < len(widths) else widths[-1]
            ch = int(self._estimate_text_height(
                [l for l in col_lines if not l.strip().startswith("![")], size))
            ch = min(ch, rheight)
            col_top = rtop + max(0, (rheight - ch) // 2)
            self._add_column_content(slide, col_lines, left=int(x), top=int(col_top),
                                     width=int(col_w), height=int(ch), size=size)
            x += col_w + gap
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_sandwich(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        n = len(sd.columns)
        lead_h = int(Inches(0.6)) if sd.top_text else 0
        concl_h = int(Inches(1.0)) if sd.bottom_text else 0
        gaps = (int(BLOCK_GAP) if sd.top_text and n else 0) + (int(BLOCK_GAP) if sd.bottom_text and n else 0)
        col_h = int(min(Inches(3.2), rheight - lead_h - concl_h - gaps))
        block_h = lead_h + (int(BLOCK_GAP) if sd.top_text and n else 0) + \
                  (col_h if n else 0) + (int(BLOCK_GAP) if sd.bottom_text and n else 0) + concl_h
        cur = rtop + max(0, (rheight - block_h) // 2)
        if sd.top_text:
            tb = self._add_textbox(slide, rleft, cur, rwidth, lead_h)
            tf = tb.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            self._set_rich_text(p, sd.top_text, SZ_BODY, self.SECONDARY)
            p.alignment = PP_ALIGN.CENTER
            cur += lead_h + int(BLOCK_GAP)
        if n > 0:
            gap = int(CARD_GAP)
            col_w = (rwidth - gap * (n - 1)) // n
            for i, col_lines in enumerate(sd.columns):
                left = rleft + i * (col_w + gap)
                self._add_body_text(slide, col_lines, left=int(left), top=int(cur),
                                    width=int(col_w), height=int(col_h), size=SZ_COL)
            cur += col_h + int(BLOCK_GAP)
        if sd.bottom_text:
            self._add_conclusion_box(slide, sd.bottom_text, rleft, int(cur), rwidth, concl_h)

    def build_figure(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        cap_h = int(Inches(0.45)) if sd.caption else 0
        desc_h = int(self._estimate_text_height(sd.body_lines, SZ_COL)) if sd.body_lines else 0
        img_file = self._resolve_image(sd.image_path) if sd.image_path else None
        img_top = rtop; ph = 0; pw = 0
        if img_file:
            from PIL import Image
            with Image.open(img_file) as im:
                iw, ih = im.size
            max_w = int(rwidth * 0.9)
            max_h = int(rheight - cap_h - desc_h - Inches(0.3))
            scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
            pw = int(iw * scale * 914400 / 96); ph = int(ih * scale * 914400 / 96)
        block_h = ph + cap_h + desc_h + (int(Inches(0.15)) if cap_h else 0) + (int(Inches(0.2)) if desc_h else 0)
        img_top = rtop + max(0, (rheight - block_h) // 2)
        if img_file:
            slide.shapes.add_picture(img_file, (SW - pw) // 2, int(img_top), pw, ph)
        cap_top = img_top + ph + int(Inches(0.15))
        if sd.caption:
            tb = self._add_textbox(slide, rleft, int(cap_top), rwidth, cap_h)
            p = tb.text_frame.paragraphs[0]; p.text = sd.caption
            p.font.name = self.FONT; p.font.size = self._fs(SZ_SMALL)
            p.font.color.rgb = self.MUTED; p.alignment = PP_ALIGN.CENTER
            tb.text_frame.word_wrap = True
        if sd.body_lines:
            self._add_body_text(slide, sd.body_lines, left=rleft,
                                top=int(cap_top + cap_h + Inches(0.2)),
                                width=rwidth, size=SZ_COL)

    def build_table(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        if not sd.table_rows:
            return
        region = self._content_region(has_title=bool(sd.h1))
        left, _, width, _ = region
        blocks = []
        if sd.h2:
            blocks.append(("sub", int(Inches(0.4))))
        # +0.25in headroom: renderers grow rows to fit CJK line height, so a
        # bare 0.5in/row estimate lets the real table eat the gap below it.
        blocks.append(("table", int(min(Inches(4.8), Inches(0.5) * len(sd.table_rows) + Inches(0.25)))))
        if sd.bottom_text:
            blocks.append(("accent", int(Inches(0.9))))
        tops = self._stack_tops([h for _, h in blocks], region, mode="center",
                                gap=int(Inches(0.2)))
        for (kind, h), t in zip(blocks, tops):
            if kind == "sub":
                tb = self._add_textbox(slide, left, t, width, h)
                p = tb.text_frame.paragraphs[0]
                p.text = sd.h2
                p.font.name = self.FONT; p.font.size = self._fs(SZ_H3)
                p.font.bold = True; p.font.color.rgb = self.SECONDARY
            elif kind == "table":
                self._styled_table(slide, sd.table_rows, left, t, width, h)
            elif kind == "accent":
                self._add_accent_box(slide, sd.bottom_text, left, t, width, h)
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_references(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        if not sd.ref_items:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        items = sd.ref_items
        # Two columns when the list is long.
        two_col = len(items) > 6
        gap = int(Inches(0.5))
        col_w = (rwidth - gap) // 2 if two_col else rwidth
        ref_size = Pt(13)

        def fill(col_items, start_idx, left):
            tb = self._add_textbox(slide, int(left), rtop, int(col_w), rheight)
            tf = tb.text_frame; tf.word_wrap = True
            for j, (author, title, venue) in enumerate(col_items):
                p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                p.space_before = PARA_GAP
                r = p.add_run(); r.text = f"[{start_idx + j + 1}] "
                r.font.name = self.FONT; r.font.size = self._fs(ref_size)
                r.font.bold = True; r.font.color.rgb = self.ACCENT
                if author:
                    r = p.add_run(); r.text = author + " "
                    r.font.name = self.FONT; r.font.size = self._fs(ref_size)
                    r.font.bold = True; r.font.color.rgb = self.FG
                if title:
                    r = p.add_run(); r.text = title + " "
                    r.font.name = self.FONT; r.font.size = self._fs(ref_size)
                    r.font.color.rgb = self.FG
                if venue:
                    r = p.add_run(); r.text = venue
                    r.font.name = self.FONT; r.font.size = self._fs(ref_size)
                    r.font.italic = True; r.font.color.rgb = self.MUTED

        if two_col:
            half = (len(items) + 1) // 2
            fill(items[:half], 0, rleft)
            fill(items[half:], half, rleft + col_w + gap)
        else:
            fill(items, 0, rleft)

    def build_timeline_h(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.timeline_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        line_y = rtop + rheight // 2 - int(Inches(0.4))
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, rleft, line_y, rwidth, int(Pt(3)))
        line.fill.solid(); line.fill.fore_color.rgb = self.HAIRLINE
        line.line.fill.background(); self._no_shadow(line)
        item_w = rwidth // n
        for i, item in enumerate(sd.timeline_items):
            cx = rleft + int(item_w * i) + int(item_w / 2)
            hl = item.get("highlight")
            dot_color = self.ACCENT if hl else self.SECONDARY
            d = int(Pt(22) if hl else Pt(16))
            dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx - d // 2, line_y + int(Pt(1.5)) - d // 2, d, d)
            dot.fill.solid(); dot.fill.fore_color.rgb = dot_color
            dot.line.color.rgb = self.WHITE; dot.line.width = Pt(2)
            self._no_shadow(dot)
            tb_left = rleft + int(item_w * i)
            yr = self._add_textbox(slide, tb_left, line_y - int(Inches(0.55)), int(item_w), int(Inches(0.4)))
            p = yr.text_frame.paragraphs[0]
            p.text = item.get("year", "")
            p.font.name = self.FONT_HEAD; p.font.size = self._fs(SZ_H3)
            p.font.bold = True; p.font.color.rgb = self.ACCENT if hl else self.PRIMARY
            p.alignment = PP_ALIGN.CENTER
            txt = self._add_textbox(slide, tb_left + int(Pt(6)), line_y + int(Inches(0.3)),
                                    int(item_w) - int(Pt(12)), int(Inches(0.55)))
            p2 = txt.text_frame.paragraphs[0]
            p2.text = item.get("text", "")
            p2.font.name = self.FONT; p2.font.size = self._fs(SZ_SMALL)
            p2.font.bold = True; p2.font.color.rgb = self.FG
            p2.alignment = PP_ALIGN.CENTER
            txt.text_frame.word_wrap = True
            if item.get("detail"):
                dtl = self._add_textbox(slide, tb_left + int(Pt(6)), line_y + int(Inches(0.85)),
                                        int(item_w) - int(Pt(12)), int(Inches(0.9)))
                p3 = dtl.text_frame.paragraphs[0]
                p3.text = item["detail"]
                p3.font.name = self.FONT; p3.font.size = self._fs(SZ_FOOT)
                p3.font.color.rgb = self.MUTED; p3.alignment = PP_ALIGN.CENTER
                dtl.text_frame.word_wrap = True

    def build_timeline_v(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.timeline_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        row_h = int(min(Inches(1.15), rheight / max(n, 1)))
        block_h = row_h * n
        top = rtop + max(0, (rheight - block_h) // 2)
        line_x = rleft + int(Inches(0.16))
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, line_x, top + int(Pt(8)),
                                      int(Pt(2.5)), int(block_h - Pt(16)))
        line.fill.solid(); line.fill.fore_color.rgb = self.HAIRLINE
        line.line.fill.background(); self._no_shadow(line)
        content_left = rleft + int(Inches(0.7))
        year_w = int(Inches(1.3))
        for i, item in enumerate(sd.timeline_items):
            ry = top + int(row_h * i)
            hl = item.get("highlight")
            d = int(Pt(16) if hl else Pt(13))
            dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, line_x + int(Pt(1.25)) - d // 2,
                                         ry + int(Pt(8)), d, d)
            dot.fill.solid(); dot.fill.fore_color.rgb = self.ACCENT if hl else self.SECONDARY
            dot.line.color.rgb = self.WHITE; dot.line.width = Pt(2); self._no_shadow(dot)
            yr = self._add_textbox(slide, content_left, ry, year_w, int(Inches(0.4)))
            p = yr.text_frame.paragraphs[0]
            p.text = item.get("year", "")
            p.font.name = self.FONT_HEAD; p.font.size = self._fs(SZ_H3)
            p.font.bold = True; p.font.color.rgb = self.ACCENT if hl else self.PRIMARY
            tx = content_left + year_w + int(Inches(0.25))
            tw = rleft + rwidth - tx
            txt = self._add_textbox(slide, tx, ry, tw, int(Inches(0.4)))
            p2 = txt.text_frame.paragraphs[0]
            self._set_rich_text(p2, item.get("text", ""), SZ_BODY, self.FG)
            txt.text_frame.word_wrap = True
            if item.get("detail"):
                dtl = self._add_textbox(slide, tx, ry + int(Inches(0.38)), tw, int(Inches(0.5)))
                p3 = dtl.text_frame.paragraphs[0]; p3.text = item["detail"]
                p3.font.name = self.FONT; p3.font.size = self._fs(SZ_SMALL)
                p3.font.color.rgb = self.MUTED
                dtl.text_frame.word_wrap = True

    def build_end(self, sd: SlideData):
        slide = self._blank_slide()
        if self.LAYOUT.end_bg == "dark":
            self._set_bg(slide, self.PRIMARY)
        elif self.LAYOUT.end_bg == "light":
            self._set_bg(slide, self.LIGHT)
        is_dark = self.LAYOUT.end_bg == "dark"
        h_color = self.WHITE if is_dark else self.PRIMARY
        sub_color = RGBColor(0xD8, 0xD8, 0xDE) if is_dark else self.MUTED
        accent = self.WHITE if is_dark else self.ACCENT
        cx = int(SW // 2)
        # accent hairline above the thank-you line
        self._hairline(slide, int(cx - Inches(0.55)), int(Inches(2.95)),
                       Inches(1.1), thickness=Pt(1.6), color=accent)
        # thank-you
        tb = self._add_textbox(slide, int(Inches(1)), int(Inches(3.15)),
                               int(SW - Inches(2)), int(self._fs(Pt(44 * 1.3))))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = sd.h1 or "Thank You"
        p.font.name = self.FONT_HEAD; p.font.size = self._fs(SZ_DISPLAY)
        p.font.bold = True; p.font.color.rgb = h_color; p.alignment = PP_ALIGN.CENTER
        # sub lines
        remaining = [strip_html(s.strip()) for s in sd.raw.split("\n")
                     if s.strip() and not s.strip().startswith(("#", "<!--", "<div", "</div", ">"))]
        if remaining:
            sb = self._add_textbox(slide, int(Inches(1)), int(Inches(4.55)),
                                   int(SW - Inches(2)), int(Inches(0.9)))
            sb.text_frame.word_wrap = True
            for i, line in enumerate(remaining[:2]):
                sp = sb.text_frame.paragraphs[0] if i == 0 else sb.text_frame.add_paragraph()
                sp.text = line; sp.font.name = self.FONT
                sp.font.size = self._fs(SZ_BODY); sp.font.color.rgb = sub_color
                sp.alignment = PP_ALIGN.CENTER
                if i:
                    sp.space_before = Pt(4)
        if len(remaining) > 2:
            cb = self._add_textbox(slide, int(Inches(1)), int(Inches(5.65)),
                                   int(SW - Inches(2)), int(Inches(0.4)))
            self._kicker_para(cb.text_frame.paragraphs[0], remaining[2],
                              color=sub_color, align=PP_ALIGN.CENTER, tracking=160)

    def build_zone_flow(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.zone_flow_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        gap = int(Inches(0.3))
        arrow_w = int(Inches(0.34))
        box_w = (rwidth - arrow_w * (n - 1) - gap * (n - 1)) // n
        box_h = int(min(Inches(3.6), rheight))
        box_top = rtop + max(0, (rheight - box_h) // 2)
        for i, item in enumerate(sd.zone_flow_items):
            x = rleft + i * (box_w + gap + arrow_w)
            self._add_zone_box(slide, int(x), box_top, int(box_w), int(box_h),
                              label=item.get("label", ""), body=item.get("body", ""))
            if i < n - 1:
                ax = int(x + box_w + gap // 2)
                ay = int(box_top + box_h // 2 - int(Pt(11)))
                arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, ax, ay, int(arrow_w), int(Pt(22)))
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = self.ACCENT
                arrow.line.fill.background()
                self._no_shadow(arrow)
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_zone_compare(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        mid = int(Inches(1.0))
        box_w = (rwidth - mid) // 2
        box_h = int(min(Inches(4.2), rheight))
        box_top = rtop + max(0, (rheight - box_h) // 2)
        self._add_zone_box(slide, rleft, box_top, int(box_w), int(box_h),
                          label=sd.zone_compare.get("left_label", ""),
                          body=sd.zone_compare.get("left_body", ""))
        # VS badge (circular)
        d = int(Inches(0.72))
        vs_x = int(rleft + box_w + (mid - d) // 2)
        vs_y = int(box_top + box_h // 2 - d // 2)
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, vs_x, vs_y, d, d)
        badge.fill.solid(); badge.fill.fore_color.rgb = self.ACCENT
        badge.line.fill.background(); self._no_shadow(badge)
        bp = badge.text_frame.paragraphs[0]
        bp.text = sd.zone_compare.get("vs_text", "VS")
        bp.font.name = self.FONT_HEAD; bp.font.size = self._fs(SZ_H3)
        bp.font.bold = True; bp.font.color.rgb = self.WHITE
        bp.alignment = PP_ALIGN.CENTER
        self._add_zone_box(slide, int(rleft + box_w + mid), box_top, int(box_w), int(box_h),
                          label=sd.zone_compare.get("right_label", ""),
                          body=sd.zone_compare.get("right_body", ""))
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_zone_matrix(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        cells = sd.zone_matrix.get("cells", [{}, {}, {}, {}])
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        gap = int(CARD_GAP)
        cell_w = (rwidth - gap) // 2
        cell_h = (rheight - gap) // 2
        positions = [
            (rleft, rtop),
            (int(rleft + cell_w + gap), rtop),
            (rleft, int(rtop + cell_h + gap)),
            (int(rleft + cell_w + gap), int(rtop + cell_h + gap)),
        ]
        for i, (x, y) in enumerate(positions):
            if i < len(cells):
                self._add_zone_box(slide, int(x), int(y), int(cell_w), int(cell_h),
                                  label=cells[i].get("label", ""),
                                  body=cells[i].get("body", ""))
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_zone_process(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.zone_process_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        gap = int(Inches(0.28))
        box_w = (rwidth - gap * (n - 1)) // n
        badge_d = int(Inches(0.56))
        block_h = int(min(Inches(4.0) + badge_d, rheight))
        box_h = block_h - badge_d - int(Inches(0.18))
        top = rtop + max(0, (rheight - block_h) // 2)
        for i, item in enumerate(sd.zone_process_items):
            x = rleft + i * (box_w + gap)
            bx = int(x + box_w // 2 - badge_d // 2)
            badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, bx, int(top), badge_d, badge_d)
            badge.fill.solid(); badge.fill.fore_color.rgb = self.SECONDARY
            badge.line.fill.background(); self._no_shadow(badge)
            bp = badge.text_frame.paragraphs[0]
            bp.text = item.get("step", str(i + 1))
            bp.font.name = self.FONT_HEAD; bp.font.size = self._fs(SZ_H3)
            bp.font.bold = True; bp.font.color.rgb = self.WHITE
            bp.alignment = PP_ALIGN.CENTER
            self._add_zone_box(slide, int(x), int(top + badge_d + Inches(0.18)),
                              int(box_w), int(box_h),
                              label=item.get("title", ""), body=item.get("body", ""))
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_agenda(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        if not sd.agenda_items:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        n = len(sd.agenda_items)
        row_h = int(min(Inches(0.92), rheight / max(n, 1)))
        block_h = row_h * n
        top = rtop + max(0, (rheight - block_h) // 2)
        num_w = int(Inches(0.9))
        for i, item in enumerate(sd.agenda_items):
            y = top + row_h * i
            nb = self._add_textbox(slide, rleft, y, num_w, row_h)
            nb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            np_ = nb.text_frame.paragraphs[0]
            np_.text = f"{i + 1:02d}"
            np_.font.name = self.FONT_HEAD; np_.font.size = self._fs(Pt(26))
            np_.font.bold = True; np_.font.color.rgb = self.ACCENT
            tb = self._add_textbox(slide, rleft + num_w + int(Inches(0.2)), y,
                                   rwidth - num_w - int(Inches(0.2)), row_h)
            tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            tb.text_frame.word_wrap = True
            tp = tb.text_frame.paragraphs[0]
            tp.text = item
            tp.font.name = self.FONT; tp.font.size = self._fs(SZ_H2)
            tp.font.color.rgb = self.FG
            if i < n - 1:
                self._hairline(slide, rleft, y + row_h - int(Pt(1)), rwidth,
                               color=self._tint(self.MUTED, 0.78))

    def build_rq(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        main_h = int(self._estimate_text_height([sd.rq_main], SZ_H2)) + int(Inches(0.4)) if sd.rq_main else 0
        sub_h = int(Inches(0.8)) if sd.rq_sub else 0
        gap = int(Inches(0.3)) if sd.rq_sub else 0
        card_w = int(rwidth * 0.82)
        card_x = rleft + (rwidth - card_w) // 2
        block_h = main_h + gap + sub_h
        top = rtop + max(0, (rheight - block_h) // 2)
        if sd.rq_main:
            self._add_card(slide, (card_x, top, card_w, main_h),
                           body=sd.rq_main, accent_bar=True, anchor="middle",
                           body_size=SZ_H2)
            # center & emphasize the question text
            # (re-style: the card body paragraph)
        if sd.rq_sub:
            tb2 = self._add_textbox(slide, card_x, top + main_h + gap, card_w, sub_h)
            tf2 = tb2.text_frame; tf2.word_wrap = True
            tf2.vertical_anchor = MSO_ANCHOR.TOP
            p2 = tf2.paragraphs[0]; p2.text = sd.rq_sub
            p2.font.name = self.FONT; p2.font.size = self._fs(SZ_SMALL)
            p2.font.color.rgb = self.MUTED; p2.alignment = PP_ALIGN.CENTER

    def build_result_dual(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.result_dual_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        gap = int(CARD_GAP)
        col_w = (rwidth - gap * (n - 1)) // n
        cap_h = int(Inches(0.45))
        for i, item in enumerate(sd.result_dual_items):
            x = rleft + i * (col_w + gap)
            img_file = self._resolve_image(item.get("image", ""))
            ph = 0
            img_top = rtop
            if img_file:
                from PIL import Image
                with Image.open(img_file) as im:
                    iw, ih = im.size
                max_w = int(col_w * 0.98)
                max_h = int(rheight - cap_h - Inches(0.15))
                scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
                pw = int(iw * scale * 914400 / 96)
                ph = int(ih * scale * 914400 / 96)
                block_h = ph + (cap_h if item.get("caption") else 0)
                img_top = rtop + max(0, (rheight - block_h) // 2)
                img_left = int(x) + (int(col_w) - pw) // 2
                slide.shapes.add_picture(img_file, img_left, int(img_top), pw, ph)
            cap_text = item.get("caption", "")
            if cap_text:
                ctb = self._add_textbox(slide, int(x), int(img_top + ph + Inches(0.12)),
                                        int(col_w), cap_h)
                cp = ctb.text_frame.paragraphs[0]
                cp.text = cap_text
                cp.font.name = self.FONT
                cp.font.size = self._fs(SZ_SMALL)
                cp.font.color.rgb = self.MUTED
                cp.alignment = PP_ALIGN.CENTER

    def build_summary(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        if not sd.summary_points:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        n = len(sd.summary_points)
        row_h = int(min(Inches(1.0), rheight / max(n, 1)))
        block_h = row_h * n
        top = rtop + max(0, (rheight - block_h) // 2)
        bar_w = int(Pt(5))
        text_x = rleft + bar_w + int(Inches(0.28))
        for i, pt in enumerate(sd.summary_points):
            y = top + row_h * i
            # accent bar synced to the row top (fixes the old desync)
            bh = int(min(row_h - Inches(0.12), Inches(0.7)))
            bdr = self._hairline(slide, rleft, y + (row_h - bh) // 2, bh,
                                 thickness=bar_w, color=self.ACCENT, vertical=True)
            tb = self._add_textbox(slide, text_x, y, rwidth - (text_x - rleft), row_h)
            tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            tb.text_frame.word_wrap = True
            self._set_rich_text(tb.text_frame.paragraphs[0], pt, SZ_H3, self.FG)

    def build_appendix(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1, color=self.MUTED)
        if sd.appendix_label:
            lbl = self._add_textbox(slide, int(SW - Inches(2.6)), int(MARGIN_T),
                                    int(Inches(2.0)), int(KICKER_H))
            self._kicker_para(lbl.text_frame.paragraphs[0], sd.appendix_label,
                              color=self.MUTED, align=PP_ALIGN.RIGHT)
        if sd.table_rows:
            # Render the table on THIS slide — delegating to build_table()
            # would open a fresh slide and split the appendix in two.
            region = self._content_region(has_title=bool(sd.h1))
            left, _, width, _ = region
            th = int(min(Inches(4.8), Inches(0.5) * len(sd.table_rows) + Inches(0.25)))
            blocks = [("table", th)]
            if sd.bottom_text:
                blocks.append(("accent", int(Inches(0.9))))
            tops = self._stack_tops([h for _, h in blocks], region, mode="center",
                                    gap=int(Inches(0.25)))
            for (kind, h), t in zip(blocks, tops):
                if kind == "table":
                    self._styled_table(slide, sd.table_rows, left, t, width, h)
                else:
                    self._add_accent_box(slide, sd.bottom_text, left, t, width, h)
            if sd.footnote:
                self._add_footnote(slide, sd.footnote)
        elif sd.body_lines:
            region = self._content_region(has_title=bool(sd.h1))
            left, _, width, _ = region
            bh = int(self._estimate_text_height(sd.body_lines, SZ_SMALL))
            top = self._stack_tops([min(bh, region[3])], region, mode="center")[0]
            self._add_body_text(slide, sd.body_lines, left=left, top=top,
                                width=width, height=min(bh, region[3]), size=SZ_SMALL)

    def _build_image_points(self, sd, lead_text, image_path, points):
        """Shared layout for overview/result: optional lead line, then a
        figure on the left and bullet points on the right, vertically centered."""
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        cur_top = rtop
        avail_h = rheight
        if lead_text:
            tb = self._add_textbox(slide, rleft, cur_top, rwidth, int(Inches(0.55)))
            p = tb.text_frame.paragraphs[0]
            p.text = lead_text
            p.font.name = self.FONT; p.font.size = self._fs(SZ_H3)
            p.font.color.rgb = self.SECONDARY
            tb.text_frame.word_wrap = True
            cur_top += int(Inches(0.7)); avail_h -= int(Inches(0.7))
        left_w = int(rwidth * 0.56)
        right_w = int(rwidth * 0.4)
        right_x = rleft + rwidth - right_w
        img_file = self._resolve_image(image_path) if image_path else None
        if img_file:
            from PIL import Image
            with Image.open(img_file) as im:
                iw, ih = im.size
            max_w = int(left_w * 0.98); max_h = int(avail_h * 0.96)
            scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
            pw = int(iw * scale * 914400 / 96); ph = int(ih * scale * 914400 / 96)
            img_top = cur_top + max(0, (avail_h - ph) // 2)
            slide.shapes.add_picture(img_file, rleft, int(img_top), pw, ph)
        if points:
            ph_pts = int(self._estimate_text_height([f"\u2022 {p}" for p in points], SZ_COL))
            pts_top = cur_top + max(0, (avail_h - min(ph_pts, avail_h)) // 2)
            tb = self._add_textbox(slide, right_x, int(pts_top), right_w,
                                   int(min(ph_pts, avail_h)))
            tf = tb.text_frame; tf.word_wrap = True
            for i, pt in enumerate(points):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.space_before = PARA_GAP
                self._set_rich_text(p, f"\u2022 {pt}", SZ_COL, self.FG)
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_overview(self, sd: SlideData):
        self._build_image_points(sd, sd.overview_text, sd.image_path, sd.overview_points)

    def build_result(self, sd: SlideData):
        self._build_image_points(sd, sd.result_text, sd.result_figure, sd.result_analysis)

    def build_steps(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.steps_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        gap = int(Inches(0.28))
        box_w = (rwidth - gap * (n - 1)) // n
        badge_d = int(Inches(0.56))
        block_h = int(min(Inches(3.9) + badge_d, rheight))
        box_h = block_h - badge_d - int(Inches(0.18))
        top = rtop + max(0, (rheight - block_h) // 2)
        for i, item in enumerate(sd.steps_items):
            x = rleft + i * (box_w + gap)
            bx = int(x + box_w // 2 - badge_d // 2)
            badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, bx, int(top), badge_d, badge_d)
            badge.fill.solid(); badge.fill.fore_color.rgb = self.ACCENT
            badge.line.fill.background(); self._no_shadow(badge)
            bp = badge.text_frame.paragraphs[0]
            bp.text = item.get("num", str(i + 1))
            bp.font.name = self.FONT_HEAD; bp.font.size = self._fs(SZ_H3)
            bp.font.bold = True; bp.font.color.rgb = self.WHITE
            bp.alignment = PP_ALIGN.CENTER
            self._add_zone_box(slide, int(x), int(top + badge_d + Inches(0.18)),
                              int(box_w), int(box_h),
                              label=item.get("title", ""), body=item.get("body", ""))
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_quote(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        qx = rleft + int(Inches(1.0))
        qw = rwidth - int(Inches(1.6))
        quote_h = int(self._estimate_text_height([sd.quote_text or ""], SZ_H2))
        block_h = int(Inches(0.9)) + quote_h + (int(Inches(0.5)) if sd.quote_source else 0)
        top = rtop + max(0, (rheight - block_h) // 2)
        # opening quotation mark (ghost)
        mark = self._add_textbox(slide, rleft, top, int(Inches(1.2)), int(Inches(1.0)))
        mp = mark.text_frame.paragraphs[0]
        mp.text = "\u201C"
        mp.font.name = self.FONT_HEAD; mp.font.size = self._fs(Pt(80))
        mp.font.bold = True; mp.font.color.rgb = self._tint(self.ACCENT, 0.55)
        # left accent bar beside the quote
        qtop = top + int(Inches(0.55))
        self._hairline(slide, qx - int(Inches(0.25)), qtop, max(quote_h, int(Inches(0.5))),
                       thickness=Pt(3), color=self.ACCENT, vertical=True)
        if sd.quote_text:
            tb = self._add_textbox(slide, qx, qtop, qw, quote_h + int(Inches(0.2)))
            tf = tb.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.text = sd.quote_text
            p.font.name = self.FONT; p.font.size = self._fs(SZ_H2)
            p.font.color.rgb = self.FG; p.line_spacing = LINE_BODY
        if sd.quote_source:
            stb = self._add_textbox(slide, qx, qtop + quote_h + int(Inches(0.25)),
                                    qw, int(Inches(0.5)))
            sp = stb.text_frame.paragraphs[0]
            sp.text = f"\u2014 {sd.quote_source}"
            sp.font.name = self.FONT; sp.font.size = self._fs(SZ_SMALL)
            sp.font.color.rgb = self.MUTED

    def build_history(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.history_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        row_h = int(min(Inches(0.92), rheight / max(n, 1)))
        block_h = row_h * n
        top = rtop + max(0, (rheight - block_h) // 2)
        year_w = int(Inches(1.5))
        ev_x = rleft + year_w + int(Inches(0.35))
        for i, item in enumerate(sd.history_items):
            y = top + row_h * i
            yr = self._add_textbox(slide, rleft, y, year_w, row_h)
            yr.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = yr.text_frame.paragraphs[0]
            p.text = item.get("year", "")
            p.font.name = self.FONT_HEAD; p.font.size = self._fs(SZ_H3)
            p.font.bold = True; p.font.color.rgb = self.ACCENT
            p.alignment = PP_ALIGN.RIGHT
            # vertical connector tick
            self._hairline(slide, ev_x - int(Inches(0.18)), y + int(Inches(0.04)),
                           row_h - int(Inches(0.08)), thickness=Pt(2),
                           color=self.HAIRLINE, vertical=True)
            ev = self._add_textbox(slide, ev_x, y, rwidth - (ev_x - rleft), row_h)
            ev.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            ev.text_frame.word_wrap = True
            p2 = ev.text_frame.paragraphs[0]
            self._set_rich_text(p2, item.get("event", ""), SZ_BODY, self.FG)
            if i < n - 1:
                self._hairline(slide, rleft, y + row_h - int(Pt(1)), rwidth,
                               color=self._tint(self.MUTED, 0.82))

    def build_panorama(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1), full=True)
        cap_h = int(Inches(0.5)) if sd.panorama_text else 0
        img_file = self._resolve_image(sd.image_path) if sd.image_path else None
        img_top = rtop; ph = 0
        if img_file:
            from PIL import Image
            with Image.open(img_file) as im:
                iw, ih = im.size
            max_w = int(rwidth); max_h = int(rheight - cap_h - Inches(0.15))
            scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
            pw = int(iw * scale * 914400 / 96); ph = int(ih * scale * 914400 / 96)
            img_top = rtop + max(0, (rheight - ph - cap_h) // 2)
            slide.shapes.add_picture(img_file, (SW - pw) // 2, int(img_top), pw, ph)
        if sd.panorama_text:
            tb = self._add_textbox(slide, rleft, int(img_top + ph + Inches(0.12)), rwidth, cap_h)
            p = tb.text_frame.paragraphs[0]; p.text = sd.panorama_text
            p.font.name = self.FONT; p.font.size = self._fs(SZ_SMALL)
            p.font.color.rgb = self.MUTED; p.alignment = PP_ALIGN.CENTER
            tb.text_frame.word_wrap = True

    def build_kpi(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.kpi_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        gap = int(CARD_GAP)
        box_w = (rwidth - gap * (n - 1)) // n
        box_h = int(min(Inches(2.7), rheight))
        kpi_top = rtop + max(0, (rheight - box_h) // 2)
        for i, item in enumerate(sd.kpi_items):
            x = rleft + i * (box_w + gap)
            self._add_card(slide, (int(x), kpi_top, int(box_w), box_h),
                           value=item.get("value", ""), label=item.get("label", ""),
                           accent_bar=True, anchor="middle",
                           value_size=self._fs(SZ_METRIC),
                           label_size=SZ_SMALL, label_color=self.MUTED)

    def build_pros_cons(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        gap = int(CARD_GAP)
        half_w = (rwidth - gap) // 2
        box_h = int(min(Inches(4.4), rheight))
        top = rtop + max(0, (rheight - box_h) // 2)
        head_h = int(Inches(0.62))
        pad = int(Pt(16))
        green = RGBColor(0x3F, 0x7A, 0x52)
        red = RGBColor(0xA8, 0x45, 0x3A)
        for idx, (items, color, label_text, mark) in enumerate([
            (sd.pros_items, green, "Pros", "\u2713"),
            (sd.cons_items, red, "Cons", "\u2717"),
        ]):
            x = rleft + idx * (half_w + gap)
            # card body
            bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, int(x), top, int(half_w), box_h)
            bg.adjustments[0] = CARD_RADIUS
            bg.fill.solid(); bg.fill.fore_color.rgb = self._tint(color, 0.92)
            bg.line.color.rgb = self._tint(color, 0.5); bg.line.width = HAIRLINE_W
            self._no_shadow(bg)
            # header band
            hb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(x), top, int(half_w), head_h)
            hb.fill.solid(); hb.fill.fore_color.rgb = color
            hb.line.fill.background(); self._no_shadow(hb)
            hp = hb.text_frame.paragraphs[0]
            hb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            hb.text_frame.margin_left = pad
            hp.text = f"{mark}  {label_text}"
            hp.font.name = self.FONT_HEAD; hp.font.size = self._fs(SZ_ZONE_L)
            hp.font.bold = True; hp.font.color.rgb = self.WHITE
            # items
            itb = self._add_textbox(slide, int(x) + pad, top + head_h + int(Pt(12)),
                                    int(half_w) - pad * 2, box_h - head_h - int(Pt(24)))
            tf = itb.text_frame; tf.word_wrap = True
            for i, item in enumerate(items):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.space_before = PARA_GAP
                self._set_rich_text(p, f"{mark}  {item}", SZ_ZONE_B, self.FG)

    def build_definition(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        bar_x = rleft
        tx = rleft + int(Inches(0.4))
        tw = rwidth - int(Inches(0.4))
        term_h = int(Inches(0.7)) if sd.def_term else 0
        body_h = int(self._estimate_text_height([sd.def_body], SZ_BODY)) if sd.def_body else 0
        note_h = int(Inches(0.6)) if sd.def_note else 0
        gaps = (int(Inches(0.2)) if sd.def_term and sd.def_body else 0) + \
               (int(Inches(0.3)) if sd.def_note else 0)
        block_h = term_h + body_h + note_h + gaps
        top = rtop + max(0, (rheight - block_h) // 2)
        # left accent bar spanning term+body
        self._hairline(slide, bar_x, top, max(term_h + body_h, int(Inches(0.6))),
                       thickness=Pt(3), color=self.ACCENT, vertical=True)
        cur = top
        if sd.def_term:
            self._eyebrow(slide, "Definition", tx, cur - int(KICKER_H) + int(Inches(0.02)), tw)
            tb = self._add_textbox(slide, tx, cur, tw, term_h)
            p = tb.text_frame.paragraphs[0]; p.text = sd.def_term
            p.font.name = self.FONT_HEAD; p.font.size = self._fs(Pt(28))
            p.font.bold = True; p.font.color.rgb = self.PRIMARY
            cur += term_h + int(Inches(0.2))
        if sd.def_body:
            tb2 = self._add_textbox(slide, tx, cur, tw, body_h)
            tf2 = tb2.text_frame; tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            self._set_rich_text(p2, sd.def_body, SZ_BODY, self.FG)
            p2.line_spacing = LINE_BODY
            cur += body_h + int(Inches(0.3))
        if sd.def_note:
            ntb = self._add_textbox(slide, tx, cur, tw, note_h)
            ntb.text_frame.word_wrap = True
            p = ntb.text_frame.paragraphs[0]; p.text = sd.def_note
            p.font.name = self.FONT; p.font.size = self._fs(SZ_SMALL)
            p.font.color.rgb = self.MUTED

    def build_diagram(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1), full=not sd.caption)
        cap_h = int(Inches(0.5)) if sd.caption else 0
        img_file = self._resolve_image(sd.image_path) if sd.image_path else None
        img_top = rtop; ph = 0
        if img_file:
            from PIL import Image
            with Image.open(img_file) as im:
                iw, ih = im.size
            max_w = int(rwidth * 0.94)
            max_h = int(rheight - cap_h - Inches(0.15))
            scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
            pw = int(iw * scale * 914400 / 96); ph = int(ih * scale * 914400 / 96)
            block_h = ph + cap_h
            img_top = rtop + max(0, (rheight - block_h) // 2)
            left = (SW - pw) // 2
            slide.shapes.add_picture(img_file, left, int(img_top), pw, ph)
        if sd.caption:
            ctb = self._add_textbox(slide, rleft, int(img_top + ph + Inches(0.12)),
                                    rwidth, cap_h)
            p = ctb.text_frame.paragraphs[0]
            p.text = sd.caption
            p.font.name = self.FONT; p.font.size = self._fs(SZ_SMALL)
            p.font.color.rgb = self.MUTED; p.alignment = PP_ALIGN.CENTER

    def build_gallery_img(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.gallery_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        cols = min(n, 3)
        rows_n = (n + cols - 1) // cols
        gap = int(CARD_GAP)
        cell_w = (rwidth - gap * (cols - 1)) // cols
        cell_h = (rheight - gap * (rows_n - 1)) // rows_n
        for idx, item in enumerate(sd.gallery_items):
            r, c = divmod(idx, cols)
            x = rleft + c * (cell_w + gap)
            y = rtop + r * (cell_h + gap)
            # subtle card behind each image
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, int(x), int(y),
                                          int(cell_w), int(cell_h))
            card.adjustments[0] = CARD_RADIUS
            card.fill.solid(); card.fill.fore_color.rgb = self.SURFACE
            card.line.color.rgb = self.HAIRLINE; card.line.width = HAIRLINE_W
            self._no_shadow(card)
            img_file = self._resolve_image(item.get("image", ""))
            if img_file:
                from PIL import Image
                with Image.open(img_file) as im:
                    iw, ih = im.size
                max_w = int(cell_w * 0.86)
                max_h = int(cell_h * 0.82)
                scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
                pw = int(iw * scale * 914400 / 96); ph = int(ih * scale * 914400 / 96)
                ix = int(x) + (int(cell_w) - pw) // 2
                iy = int(y) + (int(cell_h) - ph) // 2
                slide.shapes.add_picture(img_file, ix, iy, pw, ph)

    def build_highlight(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        if not sd.highlight_text:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        text_h = int(self._estimate_text_height([sd.highlight_text], Pt(30)))
        band_h = min(rheight, text_h + int(Inches(1.0)))
        band_top = rtop + max(0, (rheight - band_h) // 2)
        # soft surface band
        band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, rleft, band_top, rwidth, band_h)
        band.fill.solid(); band.fill.fore_color.rgb = self.SURFACE
        band.line.fill.background(); self._no_shadow(band)
        cx = int(SW // 2)
        self._hairline(slide, cx - int(Inches(0.6)), band_top + int(Inches(0.18)),
                       Inches(1.2), thickness=ACCENT_RULE_W, color=self.ACCENT)
        tb = self._add_textbox(slide, rleft + int(Inches(0.5)), band_top,
                               rwidth - int(Inches(1.0)), band_h)
        tf = tb.text_frame; tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        self._set_rich_text(p, sd.highlight_text, Pt(30), self.PRIMARY)
        for r in p.runs:
            r.font.bold = True; r.font.name = self.FONT_HEAD
        p.alignment = PP_ALIGN.CENTER; p.line_spacing = LINE_TITLE

    def build_checklist(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        if not sd.checklist_items:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        n = len(sd.checklist_items)
        row_h = int(min(Inches(0.7), rheight / max(n, 1)))
        block_h = row_h * n
        top = rtop + max(0, (rheight - block_h) // 2)
        mark_w = int(Inches(0.5))
        for i, item in enumerate(sd.checklist_items):
            y = top + row_h * i
            done = item.get("done")
            d = int(Inches(0.26))
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, rleft,
                                         y + (row_h - d) // 2, d, d)
            box.adjustments[0] = 0.2
            if done:
                box.fill.solid(); box.fill.fore_color.rgb = self.ACCENT
                box.line.fill.background()
                cp = box.text_frame.paragraphs[0]; cp.text = "\u2713"
                cp.font.size = self._fs(Pt(13)); cp.font.bold = True
                cp.font.color.rgb = self.WHITE; cp.alignment = PP_ALIGN.CENTER
            else:
                box.fill.background()
                box.line.color.rgb = self.MUTED; box.line.width = Pt(1)
            self._no_shadow(box)
            tb = self._add_textbox(slide, rleft + mark_w + int(Inches(0.18)), y,
                                   rwidth - mark_w - int(Inches(0.18)), row_h)
            tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            tb.text_frame.word_wrap = True
            self._set_rich_text(tb.text_frame.paragraphs[0], item.get("text", ""),
                                SZ_H3, self.MUTED if done else self.FG)

    def build_annotation(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        fig_w = int(rwidth * 0.56)
        notes_w = rwidth - fig_w - int(Inches(0.35))
        notes_x = rleft + fig_w + int(Inches(0.35))
        img_file = self._resolve_image(sd.annotation_figure) if sd.annotation_figure else None
        if img_file:
            from PIL import Image
            with Image.open(img_file) as im:
                iw, ih = im.size
            max_w = int(fig_w * 0.98); max_h = int(rheight * 0.96)
            scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
            pw = int(iw * scale * 914400 / 96); ph = int(ih * scale * 914400 / 96)
            slide.shapes.add_picture(img_file, rleft, int(rtop + max(0, (rheight - ph) // 2)), pw, ph)
        if sd.annotation_notes:
            nh = int(self._estimate_text_height([f"\u2022 {x}" for x in sd.annotation_notes], SZ_COL))
            ntop = rtop + max(0, (rheight - min(nh, rheight)) // 2)
            tb = self._add_textbox(slide, notes_x, int(ntop), notes_w, int(min(nh, rheight)))
            tf = tb.text_frame; tf.word_wrap = True
            for i, note in enumerate(sd.annotation_notes):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.space_before = PARA_GAP
                self._set_rich_text(p, f"\u2022 {note}", SZ_COL, self.FG)

    def build_before_after(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        mid = int(Inches(0.8))
        half_w = (rwidth - mid) // 2
        box_h = int(min(Inches(4.2), rheight))
        top = rtop + max(0, (rheight - box_h) // 2)
        for idx, (data, label_default, fill, bar) in enumerate([
            (sd.ba_before, "Before", self.SURFACE, False),
            (sd.ba_after, "After", self._tint(self.ACCENT, 0.9), True),
        ]):
            x = rleft + idx * (half_w + mid)
            label = data.get("label", label_default) if data else label_default
            body = data.get("body", "") if data else ""
            self._add_zone_box(slide, int(x), top, int(half_w), box_h,
                               label=label, body=body, fill=fill, accent_bar=bar)
        arrow_x = int(rleft + half_w + (mid - Inches(0.55)) // 2)
        arrow_y = int(top + box_h // 2 - Pt(16))
        arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, arrow_x, arrow_y, int(Inches(0.55)), int(Pt(32)))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = self.ACCENT
        arrow.line.fill.background(); self._no_shadow(arrow)

    def build_funnel(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.funnel_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        max_w = rwidth
        row_h = int(min(Inches(0.95), rheight / max(n, 1)))
        block_h = row_h * n
        top = rtop + max(0, (rheight - block_h) // 2)
        for i, item in enumerate(sd.funnel_items):
            frac = 1 - (i / max(n - 1, 1)) * 0.55
            w = int(max_w * frac)
            x = int(rleft + (max_w - w) / 2)
            y = int(top + row_h * i)
            bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, int(row_h * 0.85))
            bg.adjustments[0] = 0.05
            bg.fill.solid()
            t = i / max(n - 1, 1)
            r = int(self.PRIMARY.red * (1 - t) + self.ACCENT.red * t) if hasattr(self.PRIMARY, 'red') else 0x16
            g = int(self.PRIMARY.green * (1 - t) + self.ACCENT.green * t) if hasattr(self.PRIMARY, 'green') else 0x21
            b = int(self.PRIMARY.blue * (1 - t) + self.ACCENT.blue * t) if hasattr(self.PRIMARY, 'blue') else 0x3e
            bg.fill.fore_color.rgb = RGBColor(min(r, 255), min(g, 255), min(b, 255))
            bg.line.fill.background()
            tb = self._add_textbox(slide, x + Pt(16), y + Pt(4), w - Pt(32), int(row_h * 0.85) - Pt(8))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            run_l = tf.paragraphs[0].add_run()
            run_l.text = item.get("label", "")
            run_l.font.name = self.FONT_HEAD
            run_l.font.size = self._fs(SZ_ZONE_B)
            run_l.font.bold = True
            run_l.font.color.rgb = self.WHITE
            if item.get("value"):
                run_v = tf.paragraphs[0].add_run()
                run_v.text = f"  {item['value']}"
                run_v.font.name = self.FONT
                run_v.font.size = self._fs(SZ_SMALL)
                run_v.font.color.rgb = RGBColor(0xEE, 0xEE, 0xEE)

    def build_stack(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.stack_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        gap = int(Inches(0.12))
        row_h = int(min(Inches(1.0), (rheight - gap * (n - 1)) / max(n, 1)))
        block_h = row_h * n + gap * (n - 1)
        top = rtop + max(0, (rheight - block_h) // 2)
        for i, item in enumerate(sd.stack_items):
            y = int(top + (row_h + gap) * (n - 1 - i))
            # deeper layers get a slightly stronger surface for depth
            shade = self._tint(self.SECONDARY, 0.93 - 0.04 * i)
            self._add_zone_box(slide, rleft, y, rwidth, int(row_h), fill=shade,
                              label=item.get("name", ""), body=item.get("desc", ""),
                              anchor="middle")

    def build_card_grid(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.card_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        cols = 2 if n <= 4 else 3
        rows_n = (n + cols - 1) // cols
        gap = int(CARD_GAP)
        card_w = (rwidth - gap * (cols - 1)) // cols
        card_h = int(min(Inches(2.4), (rheight - gap * (rows_n - 1)) / max(rows_n, 1)))
        block_h = card_h * rows_n + gap * (rows_n - 1)
        top = rtop + max(0, (rheight - block_h) // 2)
        for idx, item in enumerate(sd.card_items):
            r, c = divmod(idx, cols)
            x = rleft + c * (card_w + gap)
            y = top + r * (card_h + gap)
            self._add_zone_box(slide, int(x), int(y), int(card_w), int(card_h),
                              label=item.get("title", ""), body=item.get("body", ""))

    def build_split_text(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        gap = int(CARD_GAP)
        half_w = (rwidth - gap) // 2
        for idx, data in enumerate([sd.split_left, sd.split_right]):
            x = rleft + idx * (half_w + gap)
            label = data.get("label", "") if data else ""
            body = data.get("body", "") if data else ""
            self._add_zone_box(slide, int(x), rtop, int(half_w), rheight,
                               label=label, body=body, anchor="middle")

    def build_code(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        desc_h = int(Inches(0.7)) if sd.code_desc else 0
        lines = (sd.code_text or "").count("\n") + 1
        code_h = int(min(rheight - desc_h - (int(Inches(0.2)) if desc_h else 0),
                         max(Inches(1.5), Inches(0.3) + Pt(14) * 1.5 * lines / 72 * 914400)))
        block_h = code_h + (desc_h + int(Inches(0.2)) if desc_h else 0)
        top = rtop + max(0, (rheight - block_h) // 2)
        bar_h = int(Inches(0.34))
        bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, rleft, top, rwidth, code_h)
        bg.adjustments[0] = 0.02
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(0x1E, 0x1E, 0x2E)
        bg.line.fill.background(); self._no_shadow(bg)
        # window dots
        for j, col in enumerate([RGBColor(0xFF, 0x5F, 0x57), RGBColor(0xFE, 0xBC, 0x2E),
                                 RGBColor(0x28, 0xC8, 0x40)]):
            dot = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                         rleft + int(Pt(16)) + j * int(Pt(18)),
                                         top + int(Pt(11)), int(Pt(10)), int(Pt(10)))
            dot.fill.solid(); dot.fill.fore_color.rgb = col
            dot.line.fill.background(); self._no_shadow(dot)
        tb = self._add_textbox(slide, rleft + int(Pt(20)), top + bar_h + int(Pt(6)),
                               rwidth - int(Pt(40)), code_h - bar_h - int(Pt(14)))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = sd.code_text
        p.font.name = self.FONT_MONO; p.font.size = self._fs(Pt(13))
        p.font.color.rgb = RGBColor(0xCD, 0xD6, 0xF4)
        if sd.code_desc:
            dtb = self._add_textbox(slide, rleft, int(top + code_h + Inches(0.2)),
                                    rwidth, desc_h)
            dtb.text_frame.word_wrap = True
            p2 = dtb.text_frame.paragraphs[0]; p2.text = sd.code_desc
            p2.font.name = self.FONT; p2.font.size = self._fs(SZ_SMALL)
            p2.font.color.rgb = self.MUTED

    def build_multi_result(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        n = len(sd.multi_result_items)
        if n == 0:
            return
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        gap = int(CARD_GAP)
        box_w = (rwidth - gap * (n - 1)) // n
        box_h = int(min(Inches(3.4), rheight))
        top = rtop + max(0, (rheight - box_h) // 2)
        for i, item in enumerate(sd.multi_result_items):
            x = rleft + i * (box_w + gap)
            bg, tb = self._add_card(slide, (int(x), top, int(box_w), box_h),
                                    accent_bar=True, anchor="middle")
            tf = tb.text_frame
            tf.clear()
            pm = tf.paragraphs[0]
            self._kicker_para(pm, item.get("metric", ""), size=SZ_SMALL,
                              color=self.MUTED, align=PP_ALIGN.CENTER, tracking=120)
            pv = tf.add_paragraph(); pv.space_before = Pt(6)
            rv = pv.add_run(); rv.text = item.get("value", "")
            rv.font.name = self.FONT_HEAD; rv.font.size = self._fs(Pt(38))
            rv.font.bold = True; rv.font.color.rgb = self.ACCENT
            pv.alignment = PP_ALIGN.CENTER
            if item.get("desc"):
                pd = tf.add_paragraph(); pd.space_before = Pt(8)
                self._set_rich_text(pd, item.get("desc", ""), SZ_ZONE_B, self.FG)
                pd.alignment = PP_ALIGN.CENTER

    def build_takeaway(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        main_w = rwidth - int(Inches(1.0))
        pts_w = rwidth - int(Inches(2.8))
        main_h = int(self._estimate_text_height([sd.takeaway_main], Pt(28), width=main_w)) if sd.takeaway_main else 0
        pts_h = int(self._estimate_text_height([f"\u2022 {p}" for p in sd.takeaway_points], SZ_H3, width=pts_w)) if sd.takeaway_points else 0
        gap = int(Inches(0.4)) if (main_h and pts_h) else 0
        block_h = main_h + gap + pts_h
        top = rtop + max(0, (rheight - block_h) // 2)
        cx = int(SW // 2)
        if sd.takeaway_main:
            self._hairline(slide, cx - int(Inches(0.6)), top - int(Inches(0.22)),
                           Inches(1.2), thickness=ACCENT_RULE_W, color=self.ACCENT)
            tb = self._add_textbox(slide, rleft + int(Inches(0.5)), top,
                                   main_w, main_h)
            tf = tb.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            self._set_rich_text(p, sd.takeaway_main, Pt(28), self.PRIMARY)
            for r in p.runs:
                r.font.bold = True; r.font.name = self.FONT_HEAD
            p.alignment = PP_ALIGN.CENTER; p.line_spacing = LINE_TITLE
        if sd.takeaway_points:
            ptb = self._add_textbox(slide, rleft + int(Inches(1.4)), top + main_h + gap,
                                    pts_w, pts_h)
            tf = ptb.text_frame; tf.word_wrap = True
            for i, pt in enumerate(sd.takeaway_points):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.space_before = PARA_GAP
                self._set_rich_text(p, f"\u2022  {pt}", SZ_H3, self.FG)

    def build_profile(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        rleft, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1))
        left_w = int(Inches(3.2))
        right_x = rleft + left_w + int(Inches(0.5))
        right_w = rwidth - left_w - int(Inches(0.5))
        img_file = self._resolve_image(sd.image_path) if sd.image_path else None
        if img_file:
            from PIL import Image
            with Image.open(img_file) as im:
                iw, ih = im.size
            max_w = int(left_w * 0.95); max_h = int(rheight * 0.92)
            scale = min(max_w / (iw * 914400 / 96), max_h / (ih * 914400 / 96))
            pw = int(iw * scale * 914400 / 96); ph = int(ih * scale * 914400 / 96)
            iy = rtop + max(0, (rheight - ph) // 2)
            slide.shapes.add_picture(img_file, rleft, int(iy), pw, ph)
        # Right column: name / affiliation / bio, vertically centered as a block
        name_h = int(Inches(0.6)) if sd.profile_name else 0
        affil_h = int(Inches(0.4)) if sd.profile_affiliation else 0
        bio_h = int(self._estimate_text_height(["\u2022 " + b for b in sd.profile_bio], SZ_ZONE_B)) if sd.profile_bio else 0
        block_h = name_h + affil_h + (int(Inches(0.2)) if bio_h else 0) + bio_h
        cur_y = rtop + max(0, (rheight - block_h) // 2)
        if sd.profile_name:
            ntb = self._add_textbox(slide, int(right_x), int(cur_y), int(right_w), name_h)
            p = ntb.text_frame.paragraphs[0]; p.text = sd.profile_name
            p.font.name = self.FONT_HEAD; p.font.size = self._fs(Pt(24))
            p.font.bold = True; p.font.color.rgb = self.PRIMARY
            cur_y += name_h
        if sd.profile_affiliation:
            atb = self._add_textbox(slide, int(right_x), int(cur_y), int(right_w), affil_h)
            p = atb.text_frame.paragraphs[0]
            self._kicker_para(p, sd.profile_affiliation, size=SZ_SMALL,
                              color=self.ACCENT_TEXT, tracking=120)
            cur_y += affil_h + int(Inches(0.2))
        if sd.profile_bio:
            btb = self._add_textbox(slide, int(right_x), int(cur_y), int(right_w), bio_h)
            btf = btb.text_frame; btf.word_wrap = True
            for i, item in enumerate(sd.profile_bio):
                p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
                p.space_before = PARA_GAP
                self._set_rich_text(p, "\u2022  " + item, SZ_ZONE_B, self.FG)

    # ── Okabe-Ito color-vision-safe categorical palette ──
    OKABE_ITO = [
        RGBColor(0xE6, 0x9F, 0x00), RGBColor(0x56, 0xB4, 0xE9),
        RGBColor(0x00, 0x9E, 0x73), RGBColor(0xF0, 0xE4, 0x42),
        RGBColor(0x00, 0x72, 0xB2), RGBColor(0xD5, 0x5E, 0x00),
        RGBColor(0xCC, 0x79, 0xA7), RGBColor(0x00, 0x00, 0x00),
    ]

    def build_statement(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.dark:
            self._set_bg(slide, self.PRIMARY)
        fg = self.WHITE if sd.dark else self.PRIMARY
        accent = self.WHITE if sd.dark else self.ACCENT
        text = sd.statement_text or sd.h1 or ""
        cx = int(SW // 2)
        n = len(text)
        pt = 44 if n <= 40 else (36 if n <= 90 else 28)
        th = int(self._estimate_text_height([text], Pt(pt)))
        top = int(MARGIN_T) + max(0, (int(SH - 2 * MARGIN_T) - th - int(Inches(0.3))) // 2)
        self._hairline(slide, cx - int(Inches(0.5)), top - int(Inches(0.28)),
                       Inches(1.0), thickness=ACCENT_RULE_W, color=accent)
        tb = self._add_textbox(slide, int(Inches(1.2)), top, int(SW - Inches(2.4)),
                               th + int(Inches(0.4)))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]
        self._set_rich_text(p, text, Pt(pt), fg)
        for r in p.runs:
            r.font.bold = True; r.font.name = self.FONT_HEAD
        p.alignment = PP_ALIGN.CENTER; p.line_spacing = LINE_TITLE

    def build_big_number(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.dark:
            self._set_bg(slide, self.PRIMARY)
        if sd.h1:
            self._add_title(slide, sd.h1)
        _, rtop, rwidth, rheight = self._content_region(has_title=bool(sd.h1),
                                                        full=not sd.footnote)
        cx = int(SW // 2)
        big = sd.bignum_value or "—"
        num_pt = 110 if len(big) <= 4 else (88 if len(big) <= 7 else 64)
        num_h = int(self._fs(Pt(num_pt)) * 1.15)
        lbl_h = int(Inches(0.5)) if sd.bignum_label else 0
        cap_h = int(Inches(0.45)) if sd.bignum_caption else 0
        block = num_h + lbl_h + cap_h
        top = rtop + max(0, (rheight - block) // 2)
        tb = self._add_textbox(slide, int(MARGIN_L), top, int(CONTENT_W), num_h)
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tb.text_frame.paragraphs[0]; p.text = big
        p.font.name = self.FONT_HEAD; p.font.size = self._fs(Pt(num_pt))
        p.font.bold = True; p.font.color.rgb = self.ACCENT
        p.alignment = PP_ALIGN.CENTER
        self._hairline(slide, cx - int(Inches(0.5)), top + num_h + int(Inches(0.02)),
                       Inches(1.0), thickness=ACCENT_RULE_W, color=self.ACCENT)
        cur = top + num_h + int(Inches(0.16))
        if sd.bignum_label:
            lb = self._add_textbox(slide, int(MARGIN_L), cur, int(CONTENT_W), lbl_h)
            self._kicker_para(lb.text_frame.paragraphs[0], sd.bignum_label,
                              size=SZ_H3, color=(self.WHITE if sd.dark else self.SECONDARY),
                              align=PP_ALIGN.CENTER, tracking=140)
            cur += lbl_h
        if sd.bignum_caption:
            cb = self._add_textbox(slide, int(MARGIN_L), cur, int(CONTENT_W), cap_h)
            cp = cb.text_frame.paragraphs[0]; cp.text = sd.bignum_caption
            cp.font.name = self.FONT; cp.font.size = self._fs(SZ_SMALL)
            cp.font.color.rgb = (RGBColor(0xD8, 0xD8, 0xDE) if sd.dark else self.MUTED)
            cp.alignment = PP_ALIGN.CENTER
            cb.text_frame.word_wrap = True
        if sd.footnote:
            self._add_footnote(slide, sd.footnote)

    def build_chart(self, sd: SlideData):
        slide = self._blank_slide()
        if sd.h1:
            self._add_title(slide, sd.h1)
        left, rtop, width, rheight = self._content_region(has_title=bool(sd.h1))
        if not sd.chart_series or not sd.chart_categories:
            self._add_body_text(slide, ["(chart: データ未指定 — Markdown表で系列を渡してください)"],
                                left=left, top=rtop, width=width)
            return
        from pptx.chart.data import CategoryChartData
        from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
        kind = {"line": XL_CHART_TYPE.LINE_MARKERS,
                "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                "column": XL_CHART_TYPE.COLUMN_CLUSTERED}.get(
                    sd.chart_kind, XL_CHART_TYPE.COLUMN_CLUSTERED)
        data = CategoryChartData()
        data.categories = sd.chart_categories
        for name, vals in sd.chart_series:
            data.add_series(name, vals)
        cap_h = int(Inches(0.4)) if (sd.chart_caption or sd.footnote) else 0
        ch_h = rheight - cap_h
        gf = slide.shapes.add_chart(kind, left, rtop, width, ch_h, data)
        chart = gf.chart
        chart.has_title = False
        n_series = len(sd.chart_series)
        chart.has_legend = n_series > 1
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = self._fs(SZ_SMALL)
        for i, plot_series in enumerate(chart.series):
            col = self.ACCENT if n_series == 1 else self.OKABE_ITO[i % len(self.OKABE_ITO)]
            try:
                fmt = plot_series.format
                fmt.fill.solid(); fmt.fill.fore_color.rgb = col
                fmt.line.color.rgb = col
            except Exception:
                pass
        try:
            for ax in (chart.category_axis, chart.value_axis):
                ax.tick_labels.font.size = self._fs(SZ_SMALL)
                ax.format.line.color.rgb = self.HAIRLINE
        except Exception:
            pass
        if sd.chart_caption or sd.footnote:
            cb = self._add_textbox(slide, left, int(rtop + ch_h + Inches(0.05)),
                                   width, cap_h)
            cp = cb.text_frame.paragraphs[0]
            cp.text = sd.chart_caption or sd.footnote
            cp.font.name = self.FONT; cp.font.size = self._fs(SZ_SMALL)
            cp.font.color.rgb = self.MUTED; cp.alignment = PP_ALIGN.CENTER

    # ══════════════════════════════════════════════
    # Build all slides
    # ══════════════════════════════════════════════
    BUILDERS = {
        "title": "build_title",
        "divider": "build_divider",
        "cols-2": "build_columns",
        "cols-2-wide-l": "build_columns",
        "cols-2-wide-r": "build_columns",
        "cols-3": "build_columns",
        "sandwich": "build_sandwich",
        "equation": "build_equation",
        "equations": "build_equations",
        "figure": "build_figure",
        "table-slide": "build_table",
        "references": "build_references",
        "timeline-h": "build_timeline_h",
        "timeline": "build_timeline_v",
        "end": "build_end",
        "zone-flow": "build_zone_flow",
        "zone-compare": "build_zone_compare",
        "zone-matrix": "build_zone_matrix",
        "zone-process": "build_zone_process",
        "agenda": "build_agenda",
        "rq": "build_rq",
        "result-dual": "build_result_dual",
        "summary": "build_summary",
        "appendix": "build_appendix",
        "overview": "build_overview",
        "result": "build_result",
        "steps": "build_steps",
        "quote": "build_quote",
        "history": "build_history",
        "panorama": "build_panorama",
        "kpi": "build_kpi",
        "pros-cons": "build_pros_cons",
        "definition": "build_definition",
        "diagram": "build_diagram",
        "gallery-img": "build_gallery_img",
        "highlight": "build_highlight",
        "checklist": "build_checklist",
        "annotation": "build_annotation",
        "before-after": "build_before_after",
        "funnel": "build_funnel",
        "stack": "build_stack",
        "card-grid": "build_card_grid",
        "split-text": "build_split_text",
        "code": "build_code",
        "multi-result": "build_multi_result",
        "takeaway": "build_takeaway",
        "profile": "build_profile",
        "statement": "build_statement",
        "big-statement": "build_statement",
        "dark": "build_statement",
        "big-number": "build_big_number",
        "big-number-dark": "build_big_number",
        "chart": "build_chart",
    }

    # Generic noun-label titles that should be assertions instead.
    _LABEL_TITLES = {
        "結果", "考察", "背景", "方法", "手法", "まとめ", "結論", "議論", "目的",
        "序論", "概要", "課題", "提案手法", "実験", "実験結果", "今後の課題",
        "results", "result", "method", "methods", "methodology", "conclusion",
        "conclusions", "background", "introduction", "discussion", "summary",
        "objective", "overview", "agenda", "outline", "motivation", "approach",
        "experiments", "evaluation", "future work", "related work",
    }
    # density -> (max_elements, max_body_lines, max_line_chars, max_words)
    _LINT_LIMITS = {
        "academic": (6, 6, 42, 75),
        "keynote":  (4, 3, 30, 40),
    }

    def _lint_slide(self, sd, idx: int):
        """Emit stderr warnings for over-stuffed slides / weak titles. Advisory
        only (never blocks). Grounded in PLOS '≤6 elements', 6x6, action-titles."""
        import sys as _sys
        density = getattr(self.theme, "density", "academic")
        max_el, max_lines, max_chars, max_words = self._LINT_LIMITS.get(
            density, self._LINT_LIMITS["academic"])
        warn = lambda m: print(f"[lint] slide {idx}: {m}", file=_sys.stderr)

        def cjk_len(s):  # full-width chars count as 2
            return sum(2 if ord(c) > 0x2E7F else 1 for c in s)

        body = [l.strip() for l in (sd.body_lines or []) if l.strip()
                and not l.strip().startswith(("#", "<"))]
        if len(body) > max_lines:
            warn(f"{len(body)} body lines (>{max_lines}) — split or trim")
        for l in body:
            if cjk_len(l) > max_chars * 2:
                warn(f"a line is long (~{cjk_len(l)//2} chars >{max_chars}) — shorten")
                break
        words = sum(len(l.split()) for l in body)
        if words > max_words:
            warn(f"~{words} words (>{max_words}) — this reads as a document, not a slide")

        # element count (max items across the type's collection + body + media)
        item_fields = ("columns", "kpi_items", "zone_flow_items", "zone_process_items",
                       "steps_items", "card_items", "timeline_items", "agenda_items",
                       "summary_points", "checklist_items", "funnel_items", "stack_items",
                       "multi_result_items", "gallery_items", "history_items",
                       "eq_system", "ref_items", "annotation_notes")
        items = max([len(getattr(sd, f, []) or []) for f in item_fields] + [0])
        elements = max(items, len(body)) + (1 if sd.image_path else 0) + (1 if sd.table_rows else 0)
        if elements > max_el:
            warn(f"{elements} elements (>{max_el}) — one slide, one message")

        # assertion (action-title) check
        h = (sd.h2 or sd.h1 or "").strip()
        h = re.sub(r"^\[[^\]]*\]\s*", "", h)   # drop a "[KPI]" style prefix
        if h and h.strip().rstrip("。.:：").lower() in self._LABEL_TITLES:
            warn(f'title "{h}" is a label — state the takeaway (e.g. add the result/number)')

    def build_all(self, slides: list[SlideData]):
        self._divider_no = 0
        for n, sd in enumerate(slides, 1):
            try:
                self._lint_slide(sd, n)
            except Exception:
                pass
            if sd.slide_class == "divider":
                self._divider_no += 1
            # A specified-but-unknown _class silently fell back to a plain
            # bullet slide before — warn so typos in the type name surface.
            if sd.slide_class and sd.slide_class not in self.BUILDERS:
                self._warn(f"slide {n}: unknown type '{sd.slide_class}', "
                           "rendered as a plain slide (check the _class name)")
            method_name = self.BUILDERS.get(sd.slide_class, "build_default")
            before_n = len(self.prs.slides)
            getattr(self, method_name)(sd)
            # Embed the slide class (+ any speaker note) into the notes of the
            # slide(s) just created so pptx2md can recover the semantic type.
            cls = sd.slide_class or "default"
            user_note = getattr(sd, "speaker_notes", "") or ""
            for i in range(before_n, len(self.prs.slides)):
                self._write_notes(self.prs.slides[i], cls,
                                  user_note if i == before_n else "")
        self._add_global_footer()

    def _write_notes(self, slide, cls: str, user_note: str = ""):
        """Write speaker notes: user note first (what the presenter reads),
        then the machine-readable `_class:` tag last (pptx2md round-trip)."""
        try:
            tf = slide.notes_slide.notes_text_frame
            existing = tf.text or ""
            tag = f"_class: {cls}"
            body = "\n".join(l for l in existing.splitlines()
                             if not l.strip().startswith("_class:")).strip()
            parts = []
            if body:
                parts.append(body)
            if user_note and user_note not in body:
                parts.append(user_note)
            note_text = "\n\n".join(p for p in parts if p.strip())
            tf.text = (note_text + "\n\n" + tag) if note_text else tag
        except Exception:
            pass

    # Backward-compatible alias
    def _write_class_note(self, slide, cls: str):
        self._write_notes(slide, cls, "")

    def _add_global_footer(self):
        n = len(self.prs.slides)
        for i, slide in enumerate(self.prs.slides):
            if i == 0 or i == n - 1:
                continue
            tb = self._add_textbox(slide, int(MARGIN_L), int(SH - Inches(0.42)),
                                   int(CONTENT_W), int(Inches(0.25)))
            p = tb.text_frame.paragraphs[0]
            p.text = f"{i + 1} / {n}"
            p.font.name = self.FONT
            p.font.size = self._fs(SZ_FOOT)
            p.font.color.rgb = self.MUTED
            p.alignment = PP_ALIGN.RIGHT
