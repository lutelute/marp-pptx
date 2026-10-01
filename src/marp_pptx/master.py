"""Slide master, layouts and theme — the layer PowerPoint itself builds on.

python-pptx starts every deck from its default template: an "Office" theme
(Calibri, ＭＳ Ｐゴシック, Office blue), a master and eleven layouts laid out for
4:3. Every slide marp-pptx drew on top looked right, but the deck underneath
was not the deck on screen: "New Slide" in PowerPoint gave a Calibri slide
with placeholders in 4:3 positions, the Design tab offered Office colors,
text typed into a new box came out in MS P Gothic, and no slide had a title
PowerPoint could see (outline view, navigator, accessibility checker).

`apply_master` rewrites that layer from the deck's own theme:

- theme: color scheme = the palette (ink, background, surface, primary,
  the chart ramp as accent1-6), font scheme = the deck's heading / body /
  East-Asian faces — so PowerPoint's color picker and font menus offer the
  deck's own choices, and inserted charts and shapes come out on-brand;
- master: 16:9 placeholder geometry taken from layout.py, the deck
  background, title / body text styles matching what the builder draws;
- layouts: six (表紙, タイトルとコンテンツ, セクション見出し, 2 つのコンテンツ,
  タイトルのみ, 白紙) instead of eleven 4:3 ones.

The builder then puts each slide's H1 into its layout's title placeholder,
so the drawn title and PowerPoint's idea of the title are the same shape.
"""

from __future__ import annotations

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from marp_pptx.layout import (
    BODY_TOP, CONTENT_W, FOOTER_RESERVE, MARGIN_B, MARGIN_L, SH, SW, TITLE_H,
    TITLE_TOP,
)

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# default-template layout index → (key, PowerPoint's Japanese layout name)
_KEEP = {
    0: ("title", "タイトル スライド"),
    1: ("content", "タイトルとコンテンツ"),
    2: ("section", "セクション見出し"),
    3: ("two", "2 つのコンテンツ"),
    5: ("title_only", "タイトルのみ"),
    6: ("blank", "白紙"),
}


def _hex(rgb) -> str:
    return str(rgb).upper()


def _srgb(parent, rgb):
    for child in list(parent):
        parent.remove(child)
    etree.SubElement(parent, qn("a:srgbClr")).set("val", _hex(rgb))


# ── theme ──────────────────────────────────────────────────────────────────

def _theme(b) -> None:
    part = b.prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(part.blob)
    name = f"marp-pptx {getattr(b.theme, 'palette_name', '') or 'deck'}".strip()
    root.set("name", name)

    clr = root.find(f".//{{{NS_A}}}clrScheme")
    clr.set("name", name)
    accents = b._chart_colors(6)
    scheme = {
        "dk1": b.FG, "lt1": b.theme.bg, "dk2": b.PRIMARY, "lt2": b.LIGHT,
        "accent1": accents[0], "accent2": accents[1], "accent3": accents[2],
        "accent4": accents[3], "accent5": accents[4], "accent6": accents[5],
        "hlink": b.ACCENT_TEXT, "folHlink": b.MUTED,
    }
    for tag, rgb in scheme.items():
        el = clr.find(f"{{{NS_A}}}{tag}")
        if el is not None:
            _srgb(el, rgb)

    fonts = root.find(f".//{{{NS_A}}}fontScheme")
    fonts.set("name", name)
    for which, latin in (("majorFont", b.FONT_HEAD), ("minorFont", b.FONT)):
        f = fonts.find(f"{{{NS_A}}}{which}")
        f.find(f"{{{NS_A}}}latin").set("typeface", latin)
        f.find(f"{{{NS_A}}}ea").set("typeface", b.FONT_EA)
        for script in f.findall(f"{{{NS_A}}}font"):
            if script.get("script") == "Jpan":
                script.set("typeface", b.FONT_EA)
    # Flat by default: the template's effect styles 2/3 are drop shadows that
    # every python-pptx shape references (effectRef idx=2). An explicit empty
    # effectLst hides them in PowerPoint but not in every renderer — so the
    # styles themselves carry no effect; soft shadows are added on purpose.
    for es in root.iter(f"{{{NS_A}}}effectStyle"):
        for child in list(es):
            es.remove(child)
        etree.SubElement(es, f"{{{NS_A}}}effectLst")
    part._blob = etree.tostring(root, xml_declaration=True,
                                encoding="UTF-8", standalone=True)


# ── text styles ────────────────────────────────────────────────────────────

def _rpr(tag, size_pt, color, *, bold=False, major=False):
    r = etree.Element(qn(f"a:{tag}"))
    r.set("sz", str(int(round(size_pt * 100))))
    r.set("b", "1" if bold else "0")
    _srgb(etree.SubElement(r, qn("a:solidFill")), color)
    face = "+mj" if major else "+mn"
    etree.SubElement(r, qn("a:latin")).set("typeface", f"{face}-lt")
    etree.SubElement(r, qn("a:ea")).set("typeface", f"{face}-ea")
    etree.SubElement(r, qn("a:cs")).set("typeface", f"{face}-cs")
    return r


def _ppr(tag, *, algn="l", line_pct=None, before_pt=None, mar_l=0, indent=0,
         bullet=None, bullet_color=None):
    p = etree.Element(qn(f"a:{tag}"))
    p.set("algn", algn)
    if mar_l or indent:
        p.set("marL", str(int(mar_l)))
        p.set("indent", str(int(indent)))
    if line_pct:
        ln = etree.SubElement(p, qn("a:lnSpc"))
        etree.SubElement(ln, qn("a:spcPct")).set("val", str(int(line_pct * 1000)))
    if before_pt is not None:
        sb = etree.SubElement(p, qn("a:spcBef"))
        etree.SubElement(sb, qn("a:spcPts")).set("val", str(int(before_pt * 100)))
    if bullet:
        if bullet_color is not None:
            _srgb(etree.SubElement(p, qn("a:buClr")), bullet_color)
        etree.SubElement(p, qn("a:buChar")).set("char", bullet)
    else:
        etree.SubElement(p, qn("a:buNone"))
    return p


def _set_lvl(style, lvl, ppr, rpr):
    """Replace lvl{n}pPr in a list style with ppr + its defRPr."""
    tag = qn(f"a:lvl{lvl}pPr")
    old = style.find(tag)
    ppr.append(rpr)
    if old is not None:
        old.addprevious(ppr)
        style.remove(old)
    else:
        style.append(ppr)


def _text_styles(b, master_el) -> None:
    tx = master_el.find(qn("p:txStyles"))
    fs = getattr(b.theme, "font_scale", 1.0)
    title = tx.find(qn("p:titleStyle"))
    _set_lvl(title, 1, _ppr("lvl1pPr", line_pct=106),
             _rpr("defRPr", 30 * fs, b.PRIMARY, bold=True, major=True))
    body = tx.find(qn("p:bodyStyle"))
    hang = int(Pt(16))
    for lvl, size, color in ((1, 18, b.FG), (2, 16, b.FG), (3, 14, b.MUTED)):
        _set_lvl(body, lvl,
                 _ppr(f"lvl{lvl}pPr", line_pct=130, before_pt=6,
                      mar_l=hang * lvl, indent=-hang, bullet="•",
                      bullet_color=b.ACCENT),
                 _rpr("defRPr", size * fs, color))
    other = tx.find(qn("p:otherStyle"))
    _set_lvl(other, 1, _ppr("lvl1pPr"), _rpr("defRPr", 18 * fs, b.FG))


def _lst_style(ph_el, ppr, rpr):
    """Give one placeholder its own level-1 style (layout-local override)."""
    body = ph_el.find(qn("p:txBody"))
    if body is None:
        return
    lst = body.find(qn("a:lstStyle"))
    if lst is None:
        lst = etree.Element(qn("a:lstStyle"))
        body.find(qn("a:bodyPr")).addnext(lst)
    for child in list(lst):
        lst.remove(child)
    ppr.append(rpr)
    lst.append(ppr)


def _body_pr(ph_el, *, anchor=None):
    """Flush insets (the builder's geometry is the text's geometry)."""
    body = ph_el.find(qn("p:txBody"))
    if body is None:
        return
    bp = body.find(qn("a:bodyPr"))
    for k in ("lIns", "tIns", "rIns", "bIns"):
        bp.set(k, "0")
    if anchor:
        bp.set("anchor", anchor)


# ── geometry ───────────────────────────────────────────────────────────────

def _geometry(b):
    """(left, top, width, height) per placeholder role, from layout.py."""
    title = (MARGIN_L, TITLE_TOP, CONTENT_W, TITLE_H)
    body_top = BODY_TOP
    if getattr(b.LAYOUT, "h1_deco", "") == "rule2":
        from marp_pptx.builder import R2_RULE_Y, R2_TITLE_H
        title = (MARGIN_L, int(R2_RULE_Y - Inches(0.07) - R2_TITLE_H), CONTENT_W,
                 R2_TITLE_H)
        body_top = int(R2_RULE_Y + Inches(0.22))
    body_h = int(SH - body_top - MARGIN_B - FOOTER_RESERVE)
    foot_y = int(SH - Inches(0.42))
    half = int((CONTENT_W - Inches(0.4)) / 2)
    return {
        "title": title,
        "body": (MARGIN_L, body_top, CONTENT_W, body_h),
        "left": (MARGIN_L, body_top, half, body_h),
        "right": (int(MARGIN_L + half + Inches(0.4)), body_top, half, body_h),
        "ctrTitle": (MARGIN_L, int(Inches(2.35)), CONTENT_W, int(Inches(1.5))),
        "subTitle": (MARGIN_L, int(Inches(4.0)), CONTENT_W, int(Inches(1.0))),
        "secTitle": (MARGIN_L, int(Inches(2.9)), CONTENT_W, int(Inches(1.2))),
        "secBody": (MARGIN_L, int(Inches(4.2)), CONTENT_W, int(Inches(0.8))),
        "dt": (MARGIN_L, foot_y, int(Inches(2.0)), int(Inches(0.25))),
        "ftr": (int(MARGIN_L + Inches(2.2)), foot_y,
                int(CONTENT_W - Inches(4.4)), int(Inches(0.25))),
        "sldNum": (int(MARGIN_L + CONTENT_W - Inches(2.0)), foot_y,
                   int(Inches(2.0)), int(Inches(0.25))),
    }


def _place(shape, box):
    shape.left, shape.top, shape.width, shape.height = (int(v) for v in box)


def _style_footer_ph(b, ph, algn):
    _body_pr(ph._element, anchor="ctr")
    _lst_style(ph._element, _ppr("lvl1pPr", algn=algn),
               _rpr("defRPr", 11 * getattr(b.theme, "font_scale", 1.0), b.MUTED))


def _style_placeholders(b, owner, key, geo):
    """Position + style every placeholder of the master or one layout."""
    fs = getattr(b.theme, "font_scale", 1.0)
    bodies = 0
    for ph in owner.placeholders:
        t = ph.placeholder_format.type
        tn = str(t).split(".")[-1].split(" ")[0]
        el = ph._element
        if tn in ("TITLE", "CENTER_TITLE"):
            if key == "title":
                _place(ph, geo["ctrTitle"])
                _body_pr(el, anchor="b")
                _lst_style(el, _ppr("lvl1pPr", algn="ctr", line_pct=106),
                           _rpr("defRPr", 40 * fs, b.FG, bold=True, major=True))
            elif key == "section":
                _place(ph, geo["secTitle"])
                _body_pr(el, anchor="b")
                _lst_style(el, _ppr("lvl1pPr", line_pct=106),
                           _rpr("defRPr", 34 * fs, b.FG, bold=True, major=True))
            else:
                _place(ph, geo["title"])
                _body_pr(el, anchor="ctr")
        elif tn == "SUBTITLE":
            _place(ph, geo["subTitle"])
            _body_pr(el, anchor="t")
            _lst_style(el, _ppr("lvl1pPr", algn="ctr"),
                       _rpr("defRPr", 18 * fs, b.MUTED))
        elif tn in ("BODY", "OBJECT"):
            if key == "section":
                _place(ph, geo["secBody"])
                _body_pr(el, anchor="t")
                _lst_style(el, _ppr("lvl1pPr"), _rpr("defRPr", 16 * fs, b.MUTED))
            elif key == "two":
                _place(ph, geo["left" if bodies == 0 else "right"])
                _body_pr(el)
            else:
                _place(ph, geo["body"])
                _body_pr(el)
            bodies += 1
        elif tn == "DATE":
            _place(ph, geo["dt"]); _style_footer_ph(b, ph, "l")
        elif tn == "FOOTER":
            _place(ph, geo["ftr"]); _style_footer_ph(b, ph, "ctr")
        elif tn == "SLIDE_NUMBER":
            _place(ph, geo["sldNum"]); _style_footer_ph(b, ph, "r")


def _background(b, cSld) -> None:
    for old in cSld.findall(qn("p:bg")):
        cSld.remove(old)
    bg = etree.Element(qn("p:bg"))
    bgpr = etree.SubElement(bg, qn("p:bgPr"))
    fill = etree.SubElement(bgpr, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:schemeClr")).set("val", "bg1")
    etree.SubElement(bgpr, qn("a:effectLst"))
    cSld.insert(0, bg)


def _layout_rule(b, layout) -> None:
    """rule2 header: a thin full-width rule + a short thick accent rule,
    drawn on the layout itself (every slide inherits it; PowerPoint's own
    "New Slide" gets it too)."""
    from marp_pptx.builder import R2_RULE_Y
    tree = layout._element.find(qn("p:cSld")).find(qn("p:spTree"))
    ids = [int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))]
    nid = max(ids + [1]) + 1
    for name, x, y, w, h, color in (
            ("Rule Base", MARGIN_L, R2_RULE_Y, CONTENT_W, Pt(1), b.HAIRLINE),
            ("Rule Accent", MARGIN_L, R2_RULE_Y - Pt(1.5), Inches(0.81), Pt(3.5),
             b.ACCENT)):
        sp = etree.SubElement(tree, qn("p:sp"))
        nv = etree.SubElement(sp, qn("p:nvSpPr"))
        c = etree.SubElement(nv, qn("p:cNvPr"))
        c.set("id", str(nid)); c.set("name", name)
        nid += 1
        etree.SubElement(nv, qn("p:cNvSpPr"))
        etree.SubElement(nv, qn("p:nvPr")).set("userDrawn", "1")
        pr = etree.SubElement(sp, qn("p:spPr"))
        xf = etree.SubElement(pr, qn("a:xfrm"))
        off = etree.SubElement(xf, qn("a:off"))
        off.set("x", str(int(x))); off.set("y", str(int(y)))
        ext = etree.SubElement(xf, qn("a:ext"))
        ext.set("cx", str(int(w))); ext.set("cy", str(int(h)))
        etree.SubElement(etree.SubElement(pr, qn("a:prstGeom")), qn("a:avLst")) \
            .getparent().set("prst", "rect")
        _srgb(etree.SubElement(pr, qn("a:solidFill")), color)
        etree.SubElement(etree.SubElement(pr, qn("a:ln")), qn("a:noFill"))


def apply_master(b) -> dict:
    """Rewrite theme, master and layouts of `b.prs` from builder `b`'s theme.
    Returns {key: SlideLayout} for the kept layouts (title / content /
    section / two / title_only / blank)."""
    prs = b.prs
    _theme(b)
    master = prs.slide_master
    m_el = master._element
    _background(b, m_el.find(qn("p:cSld")))
    _text_styles(b, m_el)
    geo = _geometry(b)
    _style_placeholders(b, master, "master", geo)

    layouts = list(prs.slide_layouts)
    kept = {}
    for i, layout in enumerate(layouts):
        if i in _KEEP:
            key, name = _KEEP[i]
            layout._element.find(qn("p:cSld")).set("name", name)
            _style_placeholders(b, layout, key, geo)
            if key in ("content", "two", "title_only") and \
                    getattr(b.LAYOUT, "h1_deco", "") == "rule2":
                _layout_rule(b, layout)
            kept[key] = layout
    for i in sorted(set(range(len(layouts))) - set(_KEEP), reverse=True):
        prs.slide_layouts.remove(layouts[i])
    return kept
