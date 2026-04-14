"""Slide dimensions and layout constants."""

from pptx.util import Inches, Pt

# Slide dimensions (16:9 standard)
SW = Inches(13.333)
SH = Inches(7.5)

# Common regions
MARGIN_L = Inches(0.6)
MARGIN_R = Inches(0.6)
MARGIN_T = Inches(0.35)
CONTENT_W = SW - MARGIN_L - MARGIN_R
TITLE_H = Inches(0.5)
TITLE_TOP = MARGIN_T
BODY_TOP = MARGIN_T + TITLE_H + Inches(0.1)
BODY_H = SH - BODY_TOP - Inches(0.4)

# Font size scale
SZ_TITLE = Pt(22)
SZ_H2 = Pt(18)
SZ_H3 = Pt(15)
SZ_BODY = Pt(16)
SZ_COL = Pt(15)
SZ_SMALL = Pt(13)
SZ_FOOT = Pt(10)
SZ_EQ = Pt(30)
SZ_EQ_VAR = Pt(15)
SZ_ZONE_L = Pt(17)
SZ_ZONE_B = Pt(14)
