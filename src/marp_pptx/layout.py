"""Slide dimensions and layout constants."""

from pptx.util import Inches, Pt

# Slide dimensions (16:9 standard)
SW = Inches(13.333)
SH = Inches(7.5)

# Common regions
MARGIN_L = Inches(0.45)
MARGIN_R = Inches(0.45)
MARGIN_T = Inches(0.3)
CONTENT_W = SW - MARGIN_L - MARGIN_R
TITLE_H = Inches(0.55)
TITLE_TOP = MARGIN_T
BODY_TOP = MARGIN_T + TITLE_H + Inches(0.08)
BODY_H = SH - BODY_TOP - Inches(0.35)

# Font size scale
SZ_TITLE = Pt(26)
SZ_H2 = Pt(21)
SZ_H3 = Pt(17)
SZ_BODY = Pt(19)
SZ_COL = Pt(17)
SZ_SMALL = Pt(14)
SZ_FOOT = Pt(11)
SZ_EQ = Pt(32)
SZ_EQ_VAR = Pt(16)
SZ_ZONE_L = Pt(19)
SZ_ZONE_B = Pt(15)
