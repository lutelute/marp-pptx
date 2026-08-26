"""Marp markdown parser — extracts structured SlideData from .md files."""

from __future__ import annotations

import html
import re
from dataclasses import dataclass, field, replace
from pathlib import Path


def strip_html(text: str) -> str:
    # Unescape AFTER tag removal so entities like &ensp; / &amp; render as
    # real characters instead of leaking into the slide text verbatim.
    return html.unescape(re.sub(r"<[^>]+>", "", text)).strip()


def text_with_breaks(fragment: str) -> str:
    """Extract text from an HTML fragment, preserving <br> as newlines.

    Other runs of whitespace collapse to a single space, so authored
    indentation doesn't leak into the slide while explicit <br> line
    breaks survive (the builders rely on python-pptx turning \\n into
    <a:br/> line breaks).
    """
    s = re.sub(r"<br\s*/?>", "\n", fragment, flags=re.IGNORECASE)
    s = strip_html(s)
    s = re.sub(r"[ \t]+", " ", s)
    return re.sub(r" *\n *", "\n", s).strip()


def html_lists_to_bullets(text: str) -> str:
    """Convert <ul><li>...</li></ul> / <ol> to markdown bullet lines.

    This runs before strip_html so that bulleted content inside boxes
    survives as actionable markdown rather than plain text.
    """
    text = re.sub(r"<li[^>]*>\s*(.*?)\s*</li>", r"- \1", text, flags=re.DOTALL)
    text = re.sub(r"</?[uo]l[^>]*>", "", text)
    return text


def extract_div(text: str, cls: str) -> str | None:
    pattern = rf'<div\s+class="[^"]*{re.escape(cls)}[^"]*">'
    m = re.search(pattern, text)
    if not m:
        return None
    start = m.end()
    depth = 1
    pos = start
    while pos < len(text) and depth > 0:
        no = text.find("<div", pos)
        nc = text.find("</div>", pos)
        if nc == -1:
            break
        if no != -1 and no < nc:
            depth += 1
            pos = no + 4
        else:
            depth -= 1
            if depth == 0:
                return text[start:nc].strip()
            pos = nc + 6
    return text[start:].strip()


def extract_child_divs(text: str) -> list[str]:
    children = []
    pos = 0
    while pos < len(text):
        m = re.search(r"<div[^>]*>", text[pos:])
        if not m:
            break
        ds = pos + m.end()
        depth = 1
        scan = ds
        while scan < len(text) and depth > 0:
            no = text.find("<div", scan)
            nc = text.find("</div>", scan)
            if nc == -1:
                # Unclosed <div>: salvage the rest as this child and stop.
                # Breaking without advancing `pos` here used to loop forever
                # (the outer loop kept re-finding the same opening tag).
                children.append(text[ds:].strip())
                return children
            if no != -1 and no < nc:
                depth += 1
                scan = no + 4
            else:
                depth -= 1
                if depth == 0:
                    children.append(text[ds:nc].strip())
                    pos = nc + 6
                    break
                scan = nc + 6
        else:
            break
    return children


# Column children may contain <div class="box|box-accent|box-primary"> blocks.
# parse_markdown_lines() flattens HTML to text, which silently dropped the box
# styling (the skeletons advertise it). column_lines() converts each box div
# into \x00-sentinel lines that the builder renders as a real card.
_BOX_DIV_RE = re.compile(
    r'<div\s+class="[^"]*\bbox(?:-(accent|primary))?\b[^"]*"\s*>(.*?)</div>',
    re.DOTALL)


def column_lines(text: str) -> list[str]:
    """parse_markdown_lines + box-div sentinels for column content."""
    def repl(m):
        # Blank lines around the sentinels keep the soft-wrap merger from
        # gluing them onto adjacent content lines.
        kind = m.group(1) or "plain"
        return f"\n\n\x00BOX {kind}\n\n{m.group(2)}\n\n\x00END\n\n"
    return parse_markdown_lines(_BOX_DIV_RE.sub(repl, text))


def parse_markdown_lines(text: str) -> list[str]:
    lines = []
    for line in text.split("\n"):
        s = line.strip()
        if s.startswith("<div") or s.startswith("</div>") or s.startswith("<p ") or s.startswith("<ol") or s.startswith("</ol") or s.startswith("<li") or s.startswith("</li"):
            inner = strip_html(s)
            if inner:
                lines.append(inner)
        elif s.startswith("<span"):
            inner = strip_html(s)
            if inner:
                lines.append(inner)
        else:
            lines.append(strip_html(line.rstrip()))
    # Collapse consecutive blank lines
    collapsed = []
    prev_blank = False
    for l in lines:
        if not l.strip():
            if not prev_blank:
                collapsed.append("")
            prev_blank = True
        else:
            collapsed.append(l)
            prev_blank = False

    # Merge Markdown "soft wraps": non-blank lines that are part of the same
    # paragraph. A new paragraph starts on:
    # - blank line
    # - list item (- / * / numbered)
    # - heading (#)
    # - line that follows a list/heading (so we don't glue them together)
    def _is_block_starter(s: str) -> bool:
        stripped = s.strip()
        if not stripped:
            return True
        if stripped.startswith(("# ", "## ", "### ", "#### ")):
            return True
        if stripped.startswith(("- ", "* ")):
            return True
        if re.match(r"^\d+\.\s", stripped):
            return True
        if stripped.startswith(("|", ">")):
            return True
        return False

    merged: list[str] = []
    for line in collapsed:
        if not merged:
            merged.append(line)
            continue
        prev = merged[-1]
        if not line.strip() or not prev.strip():
            merged.append(line)
            continue
        if _is_block_starter(line) or _is_block_starter(prev):
            merged.append(line)
            continue
        # Both are regular text — merge as soft wrap
        merged[-1] = prev + " " + line.strip()
    return merged


@dataclass
class SlideData:
    index: int
    slide_class: str | None
    paginate: bool
    raw: str
    h1: str = ""
    h2: str = ""
    body_lines: list = field(default_factory=list)
    columns: list = field(default_factory=list)
    top_text: str = ""
    bottom_text: str = ""
    table_rows: list = field(default_factory=list)
    image_path: str = ""
    caption: str = ""
    footnote: str = ""
    timeline_items: list = field(default_factory=list)
    eq_main: str = ""
    eq_vars: list = field(default_factory=list)
    eq_system: list = field(default_factory=list)
    ref_items: list = field(default_factory=list)
    zone_flow_items: list = field(default_factory=list)
    zone_compare: dict = field(default_factory=dict)
    zone_matrix: dict = field(default_factory=dict)
    zone_process_items: list = field(default_factory=list)
    agenda_items: list = field(default_factory=list)
    rq_main: str = ""
    rq_sub: str = ""
    summary_points: list = field(default_factory=list)
    result_dual_items: list = field(default_factory=list)
    appendix_label: str = ""
    overview_text: str = ""
    overview_points: list = field(default_factory=list)
    result_text: str = ""
    result_figure: str = ""
    result_caption: str = ""
    result_analysis: list = field(default_factory=list)
    steps_items: list = field(default_factory=list)
    quote_text: str = ""
    quote_source: str = ""
    history_items: list = field(default_factory=list)
    panorama_text: str = ""
    kpi_items: list = field(default_factory=list)
    pros_items: list = field(default_factory=list)
    cons_items: list = field(default_factory=list)
    def_term: str = ""
    def_body: str = ""
    def_note: str = ""
    gallery_items: list = field(default_factory=list)
    highlight_text: str = ""
    checklist_items: list = field(default_factory=list)
    annotation_figure: str = ""
    annotation_notes: list = field(default_factory=list)
    ba_before: dict = field(default_factory=dict)
    ba_after: dict = field(default_factory=dict)
    funnel_items: list = field(default_factory=list)
    stack_items: list = field(default_factory=list)
    card_items: list = field(default_factory=list)
    split_left: dict = field(default_factory=dict)
    split_right: dict = field(default_factory=dict)
    code_text: str = ""
    code_desc: str = ""
    multi_result_items: list = field(default_factory=list)
    takeaway_main: str = ""
    takeaway_points: list = field(default_factory=list)
    profile_name: str = ""
    profile_affiliation: str = ""
    profile_bio: list = field(default_factory=list)
    speaker_notes: str = ""
    source: str = ""              # data-source attribution (<!-- source: ... -->)
    dark: bool = False            # render on a dark (ink) background
    # statement / big-number
    statement_text: str = ""
    bignum_value: str = ""
    bignum_label: str = ""
    bignum_caption: str = ""
    # chart (native, editable)
    chart_kind: str = "bar"       # bar | line | column
    chart_categories: list = field(default_factory=list)
    chart_series: list = field(default_factory=list)  # [(name, [floats]), ...]
    chart_caption: str = ""
    # tmu-cs compatible directives (docs/FEATURE-DESIGN.md)
    eq_annotations: list = field(default_factory=list)  # [(tex, label, note, color)] per display-math line
    code_steps: list = field(default_factory=list)      # [(line_idx, step, action, span)]
    active_step: int | None = None                      # set by expand_step_slides()
    # beamer-style theorem blocks
    block_items: list = field(default_factory=list)     # [(variant, title, body)]


# tmu-cs compatible authoring directives (see docs/FEATURE-DESIGN.md).
# Math: a TeX line may end with  % [!math-annotate note="..." label="..." color="#hex"]
_MATH_ANN_RE = re.compile(r"%\s*\[!math-annotate\s+([^\]]*)\]\s*$")
_ANN_ATTR_RE = re.compile(r'(\w+)="([^"]*)"')
# Code: a code line may end with  // [!step N action[:M]]  (or # / -- comments)
_STEP_RE = re.compile(
    r"(?:(?://|#|--)\s*)?\[!step\s+(\d+)\s+(highlight|focus|warning|error|info)(?::(\d+))?\]\s*$")


def extract_math_annotations(tex_block: str) -> list[tuple[str, str | None, str | None, str | None]]:
    """Split display math into lines, pulling tmu-cs [!math-annotate] directives.

    Returns one (tex, label, note, color) tuple per non-empty line; lines
    without a directive get (tex, None, None, None). `note` is required by the
    source spec — directives without it are treated as plain comments.
    """
    rows = []
    for line in tex_block.split("\n"):
        line = line.strip()
        if not line:
            continue
        m = _MATH_ANN_RE.search(line)
        label = note = color = None
        if m:
            attrs = dict(_ANN_ATTR_RE.findall(m.group(1)))
            note = attrs.get("note")
            if note:
                label = attrs.get("label")
                color = attrs.get("color")
                line = line[:m.start()].rstrip()
        rows.append((line, label, note, color))
    return rows


def extract_code_steps(code_text: str) -> tuple[str, list[tuple[int, int, str, int]]]:
    """Strip tmu-cs [!step N action[:M]] directives from code lines.

    Visible comment text before the directive is kept; a comment that holds
    only the directive is removed entirely (matching the source engine).
    """
    steps = []
    out = []
    for i, line in enumerate(code_text.split("\n")):
        m = _STEP_RE.search(line)
        if m:
            steps.append((i, int(m.group(1)), m.group(2), int(m.group(3) or 1)))
            line = line[:m.start()].rstrip()
            line = re.sub(r"(?://|#|--)\s*$", "", line).rstrip()
        out.append(line)
    return "\n".join(out), steps


def expand_step_slides(slides: list[SlideData]) -> list[SlideData]:
    """Duplicate a slide once per distinct step number (tmu-cs engine parity).

    In PPTX the duplicates read as click-to-advance motion: each copy
    emphasizes only its own step. Slides without step directives pass through
    untouched, and a single-step slide just gets that step activated in place.
    """
    out: list[SlideData] = []
    for sd in slides:
        step_nos = sorted({s for (_, s, _, _) in sd.code_steps})
        if len(step_nos) <= 1:
            if step_nos:
                sd.active_step = step_nos[0]
            out.append(sd)
            continue
        for n in step_nos:
            out.append(replace(sd, active_step=n, index=len(out)))
    for i, sd in enumerate(out):
        sd.index = i
    return out


def parse_slide(index: int, raw: str) -> SlideData:
    """Parse a raw slide chunk into SlideData."""
    directives = {}
    notes_chunks: list[str] = []

    def repl(m):
        directives[m.group(1)] = m.group(2)
        return ""

    def note_repl(m):
        notes_chunks.append(m.group(1).strip())
        return ""

    # Speaker notes: <!-- note: ... --> (multi-line). Extract before the
    # directive pass so they don't get mistaken for _key directives.
    content = re.sub(r"<!--\s*note:\s*(.+?)\s*-->", note_repl, raw,
                     flags=re.DOTALL | re.IGNORECASE)
    # Data-source attribution: <!-- source: ... -->
    source_chunks: list[str] = []
    content = re.sub(r"<!--\s*source:\s*(.+?)\s*-->",
                     lambda m: (source_chunks.append(m.group(1).strip()) or ""),
                     content, flags=re.DOTALL | re.IGNORECASE)
    content = re.sub(r"<!--\s+_(\w+):\s*(.+?)\s*-->", repl, content).strip()

    sd = SlideData(
        index=index,
        slide_class=directives.get("class"),
        paginate=directives.get("paginate", "true") != "false",
        raw=content,
    )
    sd.speaker_notes = "\n\n".join(notes_chunks)
    sd.source = "; ".join(source_chunks)
    sd.dark = directives.get("bg", "").strip().lower() == "dark"

    h1m = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    h2m = re.search(r"^##\s+(.+)$", content, re.MULTILINE)
    if h1m:
        sd.h1 = strip_html(h1m.group(1))
    if h2m:
        sd.h2 = strip_html(h2m.group(1))

    cls = sd.slide_class

    if cls == "equation":
        eq = extract_div(content, "eq-main")
        if eq:
            mm = re.search(r"\$\$(.*?)\$\$", eq, re.DOTALL)
            sd.eq_main = mm.group(1).strip() if mm else strip_html(eq)
            if "[!math-annotate" in sd.eq_main:
                sd.eq_annotations = extract_math_annotations(sd.eq_main)
                sd.eq_main = "\n".join(r[0] for r in sd.eq_annotations)
        desc = extract_div(content, "eq-desc")
        if desc:
            spans = re.findall(r"<span[^>]*>(.*?)</span>", desc, re.DOTALL)
            for i in range(0, len(spans) - 1, 2):
                sym = strip_html(spans[i])
                d = strip_html(spans[i + 1])
                sd.eq_vars.append((sym, d))
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "equations":
        sys_div = extract_div(content, "eq-system")
        if sys_div:
            rows = extract_child_divs(sys_div)
            if rows:
                for row in rows:
                    lm = re.search(r'<span[^>]*class="[^"]*label[^"]*"[^>]*>(.*?)</span>', row, re.DOTALL)
                    label = strip_html(lm.group(1)) if lm else ""
                    mm = re.search(r"\$\$(.*?)\$\$", row, re.DOTALL)
                    if mm:
                        sd.eq_system.append((label, mm.group(1).strip()))
            else:
                pattern = re.compile(
                    r'(?:<span[^>]*class="[^"]*label[^"]*"[^>]*>(.*?)</span>\s*)?\$\$(.*?)\$\$',
                    re.DOTALL,
                )
                for m in pattern.finditer(sys_div):
                    label = strip_html(m.group(1) or "")
                    sd.eq_system.append((label, m.group(2).strip()))
        desc = extract_div(content, "eq-desc")
        if desc:
            spans = re.findall(r"<span[^>]*>(.*?)</span>", desc, re.DOTALL)
            for i in range(0, len(spans) - 1, 2):
                sym = strip_html(spans[i])
                d = strip_html(spans[i + 1])
                sd.eq_vars.append((sym, d))
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls in ("cols-2", "cols-2-wide-l", "cols-2-wide-r", "cols-3"):
        cols = extract_div(content, "columns")
        if cols:
            for child in extract_child_divs(cols):
                sd.columns.append(column_lines(child))
        else:
            # Fallback: no <div class="columns"> wrapper — treat top-level
            # <div>s (that are not known utility classes) as columns directly.
            body = content
            if h1m:
                body = body[:h1m.start()] + body[h1m.end():]
            if h2m:
                body = body[:h2m.start()] + body[h2m.end():]
            # Remove known non-column divs so we don't pick them up as columns
            for tag in ("footnote", "box-accent", "box-primary", "box"):
                pattern = rf'<div\s+class="[^"]*{tag}[^"]*">.*?</div>'
                body = re.sub(pattern, "", body, flags=re.DOTALL)
            children = extract_child_divs(body)
            for child in children:
                sd.columns.append(column_lines(child))
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "sandwich":
        top = extract_div(content, "top")
        if top:
            lead = extract_div(top, "lead")
            sd.top_text = strip_html(lead) if lead else strip_html(top)
        cols = extract_div(content, "columns")
        if cols:
            for child in extract_child_divs(cols):
                sd.columns.append(column_lines(child))
        bottom = extract_div(content, "bottom")
        if bottom:
            conc = extract_div(bottom, "conclusion")
            if conc:
                sd.bottom_text = strip_html(conc)
            else:
                box = extract_div(bottom, "box")
                sd.bottom_text = strip_html(box) if box else strip_html(bottom)

    elif cls == "figure":
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)
        cap = extract_div(content, "caption")
        if cap:
            sd.caption = strip_html(cap)
        desc = extract_div(content, "description")
        if desc:
            sd.body_lines = parse_markdown_lines(desc)

    elif cls == "table-slide":
        rows = []
        for line in content.split("\n"):
            s = line.strip()
            if s.startswith("|") and not re.match(r"^\|[-:|]+\|$", s):
                cells = [c.strip() for c in s.strip("|").split("|")]
                rows.append(cells)
        sd.table_rows = rows
        ba = extract_div(content, "box-accent")
        if ba:
            sd.bottom_text = strip_html(ba)
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "references":
        lis = re.findall(r"<li>(.*?)</li>", content, re.DOTALL)
        for li in lis:
            am = re.search(r'class="author"[^>]*>(.*?)</span>', li)
            tm = re.search(r'class="title"[^>]*>(.*?)</span>', li)
            vm = re.search(r'class="venue"[^>]*>(.*?)</span>', li)
            sd.ref_items.append((
                am.group(1).strip() if am else "",
                tm.group(1).strip() if tm else "",
                vm.group(1).strip() if vm else "",
            ))

    elif cls == "timeline-h":
        container = extract_div(content, "tl-h-container")
        if container:
            items = extract_child_divs(container)
            # extract_child_divs returns each item's INNER html, dropping the
            # `tl-h-item highlight` class on the item's own div — so capture the
            # opening-tag classes separately (same order) to recover highlight.
            item_classes = re.findall(r'<div\s+class="([^"]*tl-h-item[^"]*)"', container)
            for i, item in enumerate(items):
                block = extract_child_divs(item)
                inner = block[0] if block else item
                # Match the class token anywhere in the attribute so the
                # highlighted item's `class="tl-h-text bold"` (two classes) is
                # still captured — the old exact-quote regex dropped it, losing
                # the most important entry (usually "our work / 本研究").
                ym = re.search(r'class="[^"]*tl-h-year[^"]*"[^>]*>(.*?)</span>', inner, re.DOTALL)
                tm = re.search(r'class="[^"]*tl-h-text[^"]*"[^>]*>(.*?)</span>', inner, re.DOTALL)
                dm = re.search(r'class="[^"]*tl-h-detail[^"]*"[^>]*>(.*?)</div>', inner, re.DOTALL)
                hl = "highlight" in (item_classes[i] if i < len(item_classes) else "")
                sd.timeline_items.append({
                    "year": strip_html(ym.group(1)) if ym else "",
                    "text": strip_html(tm.group(1)) if tm else "",
                    "detail": text_with_breaks(dm.group(1)) if dm else "",
                    "highlight": hl,
                })

    elif cls == "timeline":
        container = extract_div(content, "tl-container")
        if container:
            items = extract_child_divs(container)
            item_classes = re.findall(r'<div\s+class="([^"]*tl-item[^"]*)"', container)
            for i, item in enumerate(items):
                ym = re.search(r'class="[^"]*tl-year[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                tm = re.search(r'class="[^"]*tl-text[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                dm = re.search(r'class="[^"]*tl-detail[^"]*"[^>]*>(.*?)</div>', item, re.DOTALL)
                hl = "highlight" in (item_classes[i] if i < len(item_classes) else "")
                sd.timeline_items.append({
                    "year": strip_html(ym.group(1)) if ym else "",
                    "text": strip_html(tm.group(1)) if tm else "",
                    "detail": text_with_breaks(dm.group(1)) if dm else "",
                    "highlight": hl,
                })

    elif cls == "zone-flow":
        container = extract_div(content, "zf-container")
        if container:
            for box in extract_child_divs(container):
                lbl = re.search(r'class="[^"]*zf-label[^"]*"[^>]*>(.*?)</span>', box, re.DOTALL)
                bod = re.search(r'class="[^"]*zf-body[^"]*"[^>]*>(.*?)</span>', box, re.DOTALL)
                sd.zone_flow_items.append({
                    "label": strip_html(lbl.group(1)) if lbl else "",
                    "body": strip_html(bod.group(1)) if bod else "",
                })
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "zone-compare":
        for side in ("zc-left", "zc-right"):
            div = extract_div(content, side)
            prefix = "left" if "left" in side else "right"
            if div:
                lbl = re.search(r'class="[^"]*zc-label[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                bod = re.search(r'class="[^"]*zc-body[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                sd.zone_compare[f"{prefix}_label"] = strip_html(lbl.group(1)) if lbl else ""
                sd.zone_compare[f"{prefix}_body"] = strip_html(bod.group(1)) if bod else ""
        vs = extract_div(content, "zc-vs")
        sd.zone_compare["vs_text"] = strip_html(vs) if vs else "VS"
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "zone-matrix":
        extract_div(content, "zm-container")
        xl = extract_div(content, "zm-xlabel")
        yl = extract_div(content, "zm-ylabel")
        sd.zone_matrix["x_label"] = strip_html(xl) if xl else ""
        sd.zone_matrix["y_label"] = strip_html(yl) if yl else ""
        cells = []
        for pos in ("zm-tl", "zm-tr", "zm-bl", "zm-br"):
            cell = extract_div(content, pos)
            if cell:
                lbl = re.search(r'class="[^"]*zm-label[^"]*"[^>]*>(.*?)</span>', cell, re.DOTALL)
                bod = re.search(r'class="[^"]*zm-body[^"]*"[^>]*>(.*?)</span>', cell, re.DOTALL)
                cells.append({
                    "label": strip_html(lbl.group(1)) if lbl else "",
                    "body": strip_html(bod.group(1)) if bod else "",
                })
            else:
                cells.append({"label": "", "body": ""})
        sd.zone_matrix["cells"] = cells
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "zone-process":
        container = extract_div(content, "zp-container")
        if container:
            for step_div in extract_child_divs(container):
                nm = re.search(r'class="[^"]*zp-num[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                ti = re.search(r'class="[^"]*zp-title[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                bo = re.search(r'class="[^"]*zp-body[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                sd.zone_process_items.append({
                    "step": strip_html(nm.group(1)) if nm else "",
                    "title": strip_html(ti.group(1)) if ti else "",
                    "body": strip_html(bo.group(1)) if bo else "",
                })
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "agenda":
        agenda = extract_div(content, "agenda-list")
        if agenda:
            for m in re.finditer(r"\d+\.\s*(.+)", agenda):
                sd.agenda_items.append(strip_html(m.group(1).strip()))

    elif cls == "rq":
        main = extract_div(content, "rq-main")
        if main:
            sd.rq_main = strip_html(main)
        sub = extract_div(content, "rq-sub")
        if sub:
            sd.rq_sub = strip_html(sub)

    elif cls == "result-dual":
        results = extract_div(content, "results")
        if results:
            items = extract_child_divs(results)
            for item in items:
                img_m = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", item)
                cap = extract_div(item, "caption")
                sd.result_dual_items.append({
                    "image": img_m.group(1) if img_m else "",
                    "caption": strip_html(cap) if cap else "",
                })

    elif cls == "summary":
        sp = extract_div(content, "summary-points")
        if not sp:
            sp_m = re.search(r'<ol[^>]*class="[^"]*summary-points[^"]*"[^>]*>(.*?)</ol>',
                             content, re.DOTALL)
            sp = sp_m.group(1) if sp_m else ""
        if sp:
            for li_m in re.finditer(r"<li>(.*?)</li>", sp, re.DOTALL):
                sd.summary_points.append(strip_html(li_m.group(1)))

    elif cls == "appendix":
        lbl = re.search(r'class="[^"]*appendix-label[^"]*"[^>]*>(.*?)</span>', content, re.DOTALL)
        if lbl:
            sd.appendix_label = strip_html(lbl.group(1))
        body = content
        if h1m:
            body = body[:h1m.start()] + body[h1m.end():]
        if h2m:
            body = body[:h2m.start()] + body[h2m.end():]
        for tag in ("appendix-label",):
            pattern = rf'<span\s+class="[^"]*{tag}[^"]*"[^>]*>.*?</span>'
            body = re.sub(pattern, "", body, flags=re.DOTALL)
        rows = []
        for line in body.split("\n"):
            s = line.strip()
            if s.startswith("|") and not re.match(r"^\|[-:|]+\|$", s):
                cells = [c.strip() for c in s.strip("|").split("|")]
                rows.append(cells)
        if rows:
            sd.table_rows = rows
        else:
            sd.body_lines = parse_markdown_lines(body)

    elif cls == "overview":
        lead = extract_div(content, "ov-lead")
        if lead:
            sd.overview_text = strip_html(lead)
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)
        cap = extract_div(content, "caption")
        if cap:
            sd.caption = strip_html(cap)
        pts = extract_div(content, "ov-points")
        if pts:
            for li in re.finditer(r"<li>(.*?)</li>", pts, re.DOTALL):
                sd.overview_points.append(strip_html(li.group(1)))
            if not sd.overview_points:
                for line in pts.split("\n"):
                    s = line.strip()
                    if s.startswith("- ") or s.startswith("* "):
                        sd.overview_points.append(s[2:].strip())
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "result":
        lead = extract_div(content, "rs-lead")
        if lead:
            sd.result_text = strip_html(lead)
        fig = extract_div(content, "rs-figure")
        if fig:
            img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", fig)
            if img:
                sd.result_figure = img.group(1)
            cap = extract_div(fig, "caption")
            if cap:
                sd.result_caption = strip_html(cap)
        analysis = extract_div(content, "rs-analysis")
        if analysis:
            for li in re.finditer(r"<li>(.*?)</li>", analysis, re.DOTALL):
                sd.result_analysis.append(strip_html(li.group(1)))
            if not sd.result_analysis:
                for line in analysis.split("\n"):
                    s = line.strip()
                    if s.startswith("- ") or s.startswith("* "):
                        sd.result_analysis.append(s[2:].strip())
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "steps":
        container = extract_div(content, "st-container")
        if container:
            for step_div in extract_child_divs(container):
                nm = re.search(r'class="[^"]*st-num[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                ti = re.search(r'class="[^"]*st-title[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                bo = re.search(r'class="[^"]*st-body[^"]*"[^>]*>(.*?)</span>', step_div, re.DOTALL)
                sd.steps_items.append({
                    "num": strip_html(nm.group(1)) if nm else "",
                    "title": strip_html(ti.group(1)) if ti else "",
                    "body": strip_html(bo.group(1)) if bo else "",
                })
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "quote":
        qt = extract_div(content, "qt-text")
        if qt:
            sd.quote_text = strip_html(qt)
        qs = extract_div(content, "qt-source")
        if qs:
            sd.quote_source = strip_html(qs)

    elif cls == "history":
        container = extract_div(content, "hs-container")
        if container:
            for item in extract_child_divs(container):
                ym = re.search(r'class="[^"]*hs-year[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                em = re.search(r'class="[^"]*hs-event[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.history_items.append({
                    "year": strip_html(ym.group(1)) if ym else "",
                    "event": strip_html(em.group(1)) if em else "",
                })

    elif cls == "panorama":
        pn = extract_div(content, "pn-text")
        if pn:
            sd.panorama_text = strip_html(pn)
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)

    elif cls == "kpi":
        container = extract_div(content, "kpi-container")
        if container:
            for item in extract_child_divs(container):
                vm = re.search(r'class="[^"]*kpi-value[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                lm = re.search(r'class="[^"]*kpi-label[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.kpi_items.append({
                    "value": strip_html(vm.group(1)) if vm else "",
                    "label": strip_html(lm.group(1)) if lm else "",
                })

    elif cls == "pros-cons":
        pros = extract_div(content, "pc-pros")
        if pros:
            for li in re.finditer(r"<li>(.*?)</li>", pros, re.DOTALL):
                sd.pros_items.append(strip_html(li.group(1)))
        cons = extract_div(content, "pc-cons")
        if cons:
            for li in re.finditer(r"<li>(.*?)</li>", cons, re.DOTALL):
                sd.cons_items.append(strip_html(li.group(1)))

    elif cls == "definition":
        dt = extract_div(content, "df-term")
        if dt:
            sd.def_term = strip_html(dt)
        db = extract_div(content, "df-body")
        if db:
            sd.def_body = strip_html(db)
        dn = extract_div(content, "df-note")
        if dn:
            sd.def_note = strip_html(dn)

    elif cls == "diagram":
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)
        cap = extract_div(content, "caption")
        if cap:
            sd.caption = strip_html(cap)

    elif cls == "gallery-img":
        container = extract_div(content, "gi-container")
        if container:
            for item in extract_child_divs(container):
                img_m = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", item)
                cap = extract_div(item, "gi-caption")
                sd.gallery_items.append({
                    "image": img_m.group(1) if img_m else "",
                    "caption": strip_html(cap) if cap else "",
                })

    elif cls == "highlight":
        hl = extract_div(content, "hl-text")
        if hl:
            sd.highlight_text = strip_html(hl)

    elif cls == "checklist":
        container = extract_div(content, "cl-container")
        if container:
            for li in re.finditer(r'<li(\s+class="done")?>(.*?)</li>', container, re.DOTALL):
                sd.checklist_items.append({
                    "text": strip_html(li.group(2)),
                    "done": li.group(1) is not None,
                })

    elif cls == "annotation":
        fig = extract_div(content, "an-figure")
        if fig:
            img_m = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", fig)
            if img_m:
                sd.annotation_figure = img_m.group(1)
        notes = extract_div(content, "an-notes")
        if notes:
            for li in re.finditer(r"<li>(.*?)</li>", notes, re.DOTALL):
                sd.annotation_notes.append(strip_html(li.group(1)))

    elif cls == "before-after":
        for prefix, div_cls in [("ba_before", "ba-before"), ("ba_after", "ba-after")]:
            div = extract_div(content, div_cls)
            if div:
                lm = re.search(r'class="[^"]*ba-label[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                bm = re.search(r'class="[^"]*ba-body[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                setattr(sd, prefix, {
                    "label": strip_html(lm.group(1)) if lm else "",
                    "body": strip_html(bm.group(1)) if bm else "",
                })

    elif cls == "funnel":
        container = extract_div(content, "fn-container")
        if container:
            for item in extract_child_divs(container):
                lm = re.search(r'class="[^"]*fn-label[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                vm = re.search(r'class="[^"]*fn-value[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.funnel_items.append({
                    "label": strip_html(lm.group(1)) if lm else "",
                    "value": strip_html(vm.group(1)) if vm else "",
                })

    elif cls == "stack":
        container = extract_div(content, "sk-container")
        if container:
            for item in extract_child_divs(container):
                nm = re.search(r'class="[^"]*sk-name[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                dm = re.search(r'class="[^"]*sk-desc[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.stack_items.append({
                    "name": strip_html(nm.group(1)) if nm else "",
                    "desc": strip_html(dm.group(1)) if dm else "",
                })

    elif cls == "card-grid":
        container = extract_div(content, "cg-container")
        if container:
            for item in extract_child_divs(container):
                tm = re.search(r'class="[^"]*cg-title[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                bm = re.search(r'class="[^"]*cg-body[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.card_items.append({
                    "title": strip_html(tm.group(1)) if tm else "",
                    "body": strip_html(bm.group(1)) if bm else "",
                })

    elif cls == "split-text":
        for prefix, div_cls in [("split_left", "sp-left"), ("split_right", "sp-right")]:
            div = extract_div(content, div_cls)
            if div:
                lm = re.search(r'class="[^"]*sp-label[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                bm = re.search(r'class="[^"]*sp-body[^"]*"[^>]*>(.*?)</span>', div, re.DOTALL)
                setattr(sd, prefix, {
                    "label": strip_html(lm.group(1)) if lm else "",
                    "body": strip_html(bm.group(1)) if bm else "",
                })

    elif cls == "blocks":
        container = extract_div(content, "bk-container")
        if container:
            items = extract_child_divs(container)
            # extract_child_divs drops the opening tag's class — recover the
            # variants by scanning the opening tags in order (same lesson as
            # the timeline highlight bug).
            variants = re.findall(r'<div\s+class="bk(?:\s+([\w-]+))?\s*"', container)
            for i, item in enumerate(items):
                tm = re.search(r'class="[^"]*bk-title[^"]*"[^>]*>(.*?)</span>',
                               item, re.DOTALL)
                bm = re.search(r'class="[^"]*bk-body[^"]*"[^>]*>(.*?)</span>',
                               item, re.DOTALL)
                variant = (variants[i] if i < len(variants) else "") or "theorem"
                sd.block_items.append((
                    variant,
                    strip_html(tm.group(1)) if tm else "",
                    strip_html(bm.group(1)) if bm else strip_html(item),
                ))
        fn = extract_div(content, "footnote")
        if fn:
            sd.footnote = strip_html(fn)

    elif cls == "code":
        cd = extract_div(content, "cd-code")
        if cd:
            code_m = re.search(r"```[\w]*\n(.*?)```", cd, re.DOTALL)
            if code_m:
                sd.code_text = code_m.group(1).rstrip()
            else:
                sd.code_text = strip_html(cd)
            if "[!step" in sd.code_text:
                sd.code_text, sd.code_steps = extract_code_steps(sd.code_text)
        desc = extract_div(content, "cd-desc")
        if desc:
            sd.code_desc = strip_html(desc)

    elif cls == "multi-result":
        container = extract_div(content, "mr-container")
        if container:
            for item in extract_child_divs(container):
                mm = re.search(r'class="[^"]*mr-metric[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                vm = re.search(r'class="[^"]*mr-value[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                dm = re.search(r'class="[^"]*mr-desc[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
                sd.multi_result_items.append({
                    "metric": strip_html(mm.group(1)) if mm else "",
                    "value": strip_html(vm.group(1)) if vm else "",
                    "desc": strip_html(dm.group(1)) if dm else "",
                })

    elif cls == "takeaway":
        ta = extract_div(content, "ta-main")
        if ta:
            sd.takeaway_main = strip_html(ta)
        pts = extract_div(content, "ta-points")
        if pts:
            for li in re.finditer(r"<li>(.*?)</li>", pts, re.DOTALL):
                sd.takeaway_points.append(strip_html(li.group(1)))
        # Fallback: H2 as the message, markdown bullets as points. Without
        # this, a takeaway written without the div skeleton (the natural way
        # to author it) rendered a silently empty body.
        if not sd.takeaway_main and sd.h2:
            sd.takeaway_main = sd.h2
        if not sd.takeaway_points:
            for m in re.finditer(r"^\s*[-*]\s+(.+)$", content, re.MULTILINE):
                sd.takeaway_points.append(strip_html(m.group(1)))

    elif cls == "profile":
        container = extract_div(content, "pf-container")
        if container:
            nm = extract_div(container, "pf-name")
            if nm:
                sd.profile_name = strip_html(nm)
            af = extract_div(container, "pf-affiliation")
            if af:
                sd.profile_affiliation = strip_html(af)
            bio = extract_div(container, "pf-bio")
            if bio:
                for li in re.finditer(r"<li>(.*?)</li>", bio, re.DOTALL):
                    sd.profile_bio.append(strip_html(li.group(1)))
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)

    elif cls in ("statement", "dark", "big-statement"):
        if cls == "dark":
            sd.dark = True
        st = extract_div(content, "statement")
        if not st:
            # first non-heading, non-comment line(s)
            lines = [l.strip() for l in content.split("\n")
                     if l.strip() and not l.strip().startswith(("#", "<"))]
            st = " ".join(lines)
        sd.statement_text = strip_html(st)

    elif cls in ("big-number", "big-number-dark"):
        if cls.endswith("-dark"):
            sd.dark = True
        bn = extract_div(content, "big-number") or content
        vm = re.search(r'class="[^"]*bn-value[^"]*"[^>]*>(.*?)</span>', bn, re.DOTALL)
        lm = re.search(r'class="[^"]*bn-label[^"]*"[^>]*>(.*?)</span>', bn, re.DOTALL)
        cm = re.search(r'class="[^"]*bn-caption[^"]*"[^>]*>(.*?)</span>', bn, re.DOTALL)
        if vm:
            sd.bignum_value = strip_html(vm.group(1))
            sd.bignum_label = strip_html(lm.group(1)) if lm else ""
            sd.bignum_caption = strip_html(cm.group(1)) if cm else ""
        else:
            lines = [l.strip() for l in content.split("\n")
                     if l.strip() and not l.strip().startswith(("#", "<", "-"))]
            if lines:
                sd.bignum_value = strip_html(lines[0])
                sd.bignum_label = strip_html(lines[1]) if len(lines) > 1 else (sd.h2 or "")
                sd.bignum_caption = strip_html(lines[2]) if len(lines) > 2 else ""

    elif cls == "chart":
        sd.chart_kind = directives.get("chart", "bar").strip().lower()
        rows = []
        for line in content.split("\n"):
            s = line.strip()
            if s.startswith("|") and not re.match(r"^\|[-:|\s]+\|$", s):
                rows.append([c.strip() for c in s.strip("|").split("|")])
        if len(rows) >= 2:
            header = rows[0]
            series_names = [strip_html(h) for h in header[1:]]
            cats, series_vals = [], [[] for _ in series_names]
            for r in rows[1:]:
                cats.append(strip_html(r[0]) if r else "")
                for si in range(len(series_names)):
                    raw = r[si + 1] if si + 1 < len(r) else ""
                    num = re.sub(r"[^\d.\-]", "", strip_html(raw))
                    try:
                        series_vals[si].append(float(num))
                    except ValueError:
                        series_vals[si].append(0.0)
            sd.chart_categories = cats
            sd.chart_series = list(zip(series_names, series_vals))
        cap = extract_div(content, "chart-caption")
        if cap:
            sd.chart_caption = strip_html(cap)

    else:
        body = content
        if h1m:
            body = body[:h1m.start()] + body[h1m.end():]
        if h2m:
            body = body[:h2m.start()] + body[h2m.end():]
        ba = extract_div(body, "box-accent")
        bp = extract_div(body, "box-primary")
        fn = extract_div(body, "footnote")
        for tag in ("box-accent", "box-primary", "box", "footnote"):
            div = extract_div(body, tag)
            if div:
                pattern = rf'<div\s+class="[^"]*{tag}[^"]*">.*?</div>'
                body = re.sub(pattern, "", body, flags=re.DOTALL)

        # Detect Markdown tables in remaining body
        table_rows: list[list[str]] = []
        non_table_lines: list[str] = []
        for line in body.split("\n"):
            s = line.strip()
            if s.startswith("|") and not re.match(r"^\|[-:|\s]+\|$", s):
                cells = [c.strip() for c in s.strip("|").split("|")]
                table_rows.append(cells)
            elif re.match(r"^\|[-:|\s]+\|$", s):
                continue  # table separator row
            else:
                non_table_lines.append(line)
        if table_rows:
            sd.table_rows = table_rows

        sd.body_lines = parse_markdown_lines("\n".join(non_table_lines))

        # Boxes: convert HTML lists to markdown bullets BEFORE strip_html so
        # that `_fill_multiline_box` sees `- item` lines instead of plain text.
        if ba:
            sd.bottom_text = strip_html(html_lists_to_bullets(ba))
        elif bp:
            sd.bottom_text = strip_html(html_lists_to_bullets(bp))
        if fn:
            sd.footnote = strip_html(fn)
        img = re.search(r"!\[(?:w:\d+)?\]\(([^)]+)\)", content)
        if img:
            sd.image_path = img.group(1)

    # A data source falls back to the footnote slot so existing footnote
    # rendering surfaces it (small-caps, bottom of slide).
    if sd.source and not sd.footnote:
        sd.footnote = sd.source

    return sd


def parse_marp(path: str | Path) -> list[SlideData]:
    """Parse a Marp markdown file into a list of SlideData."""
    text = Path(path).read_text(encoding="utf-8")
    if text.startswith("---"):
        end = text.find("---", 3)
        if end != -1:
            text = text[end + 3:]
    chunks = re.split(r"\n---\n", text)
    slides = []
    for i, chunk in enumerate(chunks):
        chunk = chunk.strip()
        if chunk:
            slides.append(parse_slide(i, chunk))
    return expand_step_slides(slides)
