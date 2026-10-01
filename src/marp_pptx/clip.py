"""Cut a quoted passage out of the source PDF as an image (論文の切り抜き).

An excerpt that shows the paper's own typesetting — its column, its font,
the words either side — reads as evidence in a way retyped text cannot.
Given a quote, find its words on a page, crop the lines they sit on (the
column's width, one line of context above and below), dim the context and
lay a marker behind the quoted words.

Matching folds case, width, ligatures, punctuation and hyphenated line
breaks (the same key as ingest.check_quotes), so a quote copied from the
extracted text finds its place in the PDF. A quote that is not in the PDF
returns None — the caller falls back to typed text and warns.
"""

from __future__ import annotations

import hashlib
import re
import tempfile
import unicodedata
from pathlib import Path

_CACHE_DIR = Path(tempfile.gettempdir()) / "marp_pptx_clips"
_KEEP = re.compile(r"[^0-9a-z぀-ヿ㐀-鿿]+")


def _key(s: str) -> str:
    return _KEEP.sub("", unicodedata.normalize("NFKC", s).lower())


def _fragments(quote: str) -> list[str]:
    """Keys of the quote's pieces between ellipses (10+ significant chars)."""
    parts = re.split(r"\[?(?:…|\.\.\.)\]?", quote)
    return [k for k in (_key(p) for p in parts) if len(k) >= 10]


def _locate(words, frags):
    """Indices of the words covering every fragment, in order, or None."""
    stream, owner = [], []
    for i, w in enumerate(words):
        k = _key(w[4])
        stream.append(k)
        owner.extend([i] * len(k))
    text = "".join(stream)
    hit, pos = [], 0
    for f in frags:
        at = text.find(f, pos)
        if at < 0:
            return None
        hit.extend(range(owner[at], owner[at + len(f) - 1] + 1))
        pos = at + len(f)
    return sorted(set(hit))


def _lines(words):
    """Group words into visual lines: [(y0, y1, [word index…])], top-down."""
    rows: list[list] = []
    for i in sorted(range(len(words)), key=lambda j: (words[j][1], words[j][0])):
        x0, y0, x1, y1 = words[i][:4]
        mid = (y0 + y1) / 2
        for r in rows:
            if r[0] <= mid <= r[1]:
                r[0], r[1] = min(r[0], y0), max(r[1], y1)
                r[2].append(i)
                break
        else:
            rows.append([y0, y1, [i]])
    rows.sort(key=lambda r: r[0])
    return rows


def clip_quote(pdf_path, quote: str, *, marker=(0xF4, 0xD6, 0xCF),
               dpi: int = 300, context: int = 1) -> dict | None:
    """Crop `quote` out of `pdf_path`. Returns {png, page, body_pt, width_pt,
    height_pt} (PNG path, 1-based page, the paper's body size, the crop's
    size in PDF points) or None when the quote is not on any page."""
    frags = _fragments(quote)
    if not frags:
        return None
    pdf_path = Path(pdf_path)
    tag = hashlib.sha1(
        f"{pdf_path.resolve()}|{pdf_path.stat().st_mtime}|{quote}|{marker}|"
        f"{dpi}|{context}".encode()).hexdigest()[:16]
    import fitz  # PyMuPDF — marp-pptx[ingest]
    from PIL import Image, ImageDraw

    with fitz.open(pdf_path) as doc:
        for pno, page in enumerate(doc):
            words = page.get_text("words")
            hit = _locate(words, frags)
            if not hit:
                continue
            # The column the quote sits in: the x-range of its words, widened
            # to every line that shares it (so the crop is a clean column).
            cx0 = min(words[i][0] for i in hit)
            cx1 = max(words[i][2] for i in hit)
            col = [i for i, w in enumerate(words)
                   if w[2] > cx0 - 2 and w[0] < cx1 + 2]
            cx0 = min(words[i][0] for i in col)
            cx1 = max(words[i][2] for i in col)
            col = [i for i in col if words[i][0] >= cx0 - 1 and words[i][2] <= cx1 + 1]
            rows = _lines([words[i] for i in col])
            rows = [(y0, y1, [col[j] for j in idx]) for y0, y1, idx in rows]
            hit_set = set(hit)
            first = next(n for n, r in enumerate(rows) if hit_set & set(r[2]))
            last = max(n for n, r in enumerate(rows) if hit_set & set(r[2]))
            lo, hi = max(0, first - context), min(len(rows) - 1, last + context)
            pad = 3.0

            def mid(a, b):   # halfway between two lines' centres: clear of both
                return ((rows[a][0] + rows[a][1]) + (rows[b][0] + rows[b][1])) / 4

            top = mid(lo - 1, lo) if lo > 0 else rows[lo][0] - pad
            bot = mid(hi, hi + 1) if hi + 1 < len(rows) else rows[hi][1] + pad
            clip = fitz.Rect(cx0 - pad, top, cx1 + pad, bot)

            _CACHE_DIR.mkdir(parents=True, exist_ok=True)
            out = _CACHE_DIR / f"{tag}.png"
            # a word box spans ascender to descender — about one em
            body_pt = sorted(words[i][3] - words[i][1] for i in hit)[len(hit) // 2]
            if not out.exists():
                pix = page.get_pixmap(clip=clip, dpi=dpi)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                s = dpi / 72.0

                def box(x0, y0, x1, y1):
                    return (round((x0 - clip.x0) * s), round((y0 - clip.y0) * s),
                            round((x1 - clip.x0) * s), round((y1 - clip.y0) * s))

                # one marker bar per line, from the first to the last quoted word
                bars = []
                for y0, y1, idx in rows[first:last + 1]:
                    on = [i for i in idx if i in hit_set]
                    if on:
                        bars.append(box(min(words[i][0] for i in on) - 1, y0 - 0.5,
                                        max(words[i][2] for i in on) + 1, y1 + 0.5))
                # dim everything that is context, keep the quote at full ink
                veil = Image.new("L", img.size, 150)
                keep = ImageDraw.Draw(veil)
                for b in bars:
                    keep.rectangle(b, fill=0)
                img = Image.composite(Image.new("RGB", img.size, (255, 255, 255)),
                                      img, veil.point(lambda v: v * 0.62))
                mark = Image.new("RGB", img.size, (255, 255, 255))
                md = ImageDraw.Draw(mark)
                for b in bars:
                    md.rectangle(b, fill=tuple(marker))
                from PIL import ImageChops
                img = ImageChops.multiply(img, mark)
                img.save(out)
            return {"png": str(out), "page": pno + 1, "body_pt": body_pt,
                    "width_pt": clip.width, "height_pt": clip.height}
    return None
