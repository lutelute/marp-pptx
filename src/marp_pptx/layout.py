"""Slide dimensions, layout constants, and design tokens.

Refined-minimal / center-balanced design system (16:9, 13.333 x 7.5 in).
All sizes are the *base* values; the builder applies font_scale / margin_scale
via self._fs() / self._ms() so these stay scale-agnostic.
"""

from pptx.util import Inches, Pt

# ── Slide dimensions (16:9 standard) ──
SW = Inches(13.333)
SH = Inches(7.5)

# ── Page frame (symmetric, safe-zone compliant, generous "calm" margins) ──
MARGIN_L = Inches(0.62)
MARGIN_R = Inches(0.62)
MARGIN_T = Inches(0.50)
MARGIN_B = Inches(0.50)
CONTENT_W = SW - MARGIN_L - MARGIN_R          # = 12.093 in

# ── Title band ──
TITLE_H = Inches(0.72)
TITLE_TOP = MARGIN_T
TITLE_GAP = Inches(0.26)                       # title -> body gap
KICKER_H = Inches(0.26)                         # small-caps label above title
BODY_TOP = MARGIN_T + TITLE_H + TITLE_GAP      # = 1.48 in
BODY_H = SH - BODY_TOP - MARGIN_B              # = 5.52 in (footer reserved per-call)
FOOTER_RESERVE = Inches(0.32)                  # page-number band at the bottom

# ── Type scale (≈ Perfect Fourth r=1.333 on base 18pt; title eased to 30pt
#    so two-line academic titles don't crowd the band) ──
SZ_DISPLAY = Pt(40)    # hero title (title / end slides)
SZ_TITLE = Pt(30)      # recurring slide H1   (was 28)
SZ_KICKER = Pt(12)     # small-caps eyebrow label
SZ_H2 = Pt(24)         # subhead              (was 22; 18·1.333)
SZ_H3 = Pt(16)
SZ_BODY = Pt(18)       # base
SZ_COL = Pt(16)
SZ_SMALL = Pt(14)      # caption
SZ_FOOT = Pt(11)
SZ_METRIC = Pt(44)     # KPI value
SZ_BIGNUM = Pt(96)     # big-number single-stat slide
SZ_STATEMENT = Pt(40)  # full-screen statement slide
SZ_EQ = Pt(30)         # canonical display-equation size
SZ_EQ_VAR = Pt(16)     # variable-legend symbol
SZ_ZONE_L = Pt(18)     # card / zone label
SZ_ZONE_B = Pt(15)     # card / zone body

# ── Rhythm / spacing ──
LINE_BODY = 1.30       # lists / default body
LINE_TIGHT = 1.10      # headings / labels
LINE_PROSE = 1.45      # long prose (definition / quote / summary / overview)
LINE_TITLE = 1.06      # multi-line headings
PARA_GAP = Pt(8)       # between list items / paragraphs
BLOCK_GAP = Inches(0.28)   # between stacked content blocks
SECTION_GAP = Inches(0.50) # between major regions (sandwich/overview)

# ── Cards / boxes (shadow-free elevation: surface fill + hairline + accent tick) ──
CARD_GAP = Inches(0.30)
CARD_PAD = Pt(18)
CARD_RADIUS = 0.035
CARD_ACCENT_H = Pt(3)      # top accent bar height
HAIRLINE_W = Pt(0.75)      # all thin rules / card borders
ACCENT_RULE_W = Pt(2.5)    # short accent rule under titles / above heroes
