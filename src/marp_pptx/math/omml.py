"""LaTeX -> OMML (Office Math Markup Language) conversion via Pandoc.

Output is a lxml <a14:m> element ready to append to a python-pptx paragraph,
producing equations that are natively editable in PowerPoint.
"""
from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_CACHE_DIR = Path(tempfile.gettempdir()) / "marp_omml_cache"
_CACHE_DIR.mkdir(exist_ok=True)

_PANDOC = shutil.which("pandoc")


class OmmlError(RuntimeError):
    pass


def _cache_key(tex: str, display: bool) -> str:
    return hashlib.md5(f"{int(display)}:{tex}".encode()).hexdigest()


def _run_pandoc(latex: str, display: bool) -> bytes:
    if _PANDOC is None:
        raise OmmlError("pandoc not found in PATH")
    wrapped = f"$${latex}$$" if display else f"${latex}$"
    md = f"---\ntitle: x\n---\n\n# s\n\n{wrapped}\n"
    with tempfile.TemporaryDirectory() as d:
        dp = Path(d)
        (dp / "in.md").write_text(md, encoding="utf-8")
        pp = dp / "out.pptx"
        subprocess.run(
            [_PANDOC, "-t", "pptx", str(dp / "in.md"), "-o", str(pp)],
            check=True,
            capture_output=True,
            timeout=20,
        )
        with zipfile.ZipFile(pp) as z:
            return z.read("ppt/slides/slide2.xml")


def latex_to_omml_element(latex: str, display: bool = False) -> etree._Element:
    """Convert a LaTeX math string into an <a14:m> lxml element.

    Raises OmmlError on any failure.
    """
    if _PANDOC is None:
        raise OmmlError("pandoc not found in PATH")

    key = _cache_key(latex, display)
    cached = _CACHE_DIR / f"{key}.xml"
    if cached.exists():
        return etree.fromstring(cached.read_bytes())

    try:
        slide_xml = _run_pandoc(latex, display)
    except subprocess.CalledProcessError as e:
        raise OmmlError(f"pandoc failed: {e.stderr.decode(errors='ignore')[:200]}") from e
    except subprocess.TimeoutExpired as e:
        raise OmmlError("pandoc timeout") from e
    except KeyError as e:
        raise OmmlError(f"unexpected pandoc pptx layout: {e}") from e

    root = etree.fromstring(slide_xml)
    a14_elems = root.findall(f".//{{{NS_A14}}}m")
    if not a14_elems:
        raise OmmlError("no <a14:m> element in pandoc output")

    chosen = None
    for el in a14_elems:
        has_para = el.find(f"{{{NS_M}}}oMathPara") is not None
        if display and has_para:
            chosen = el
            break
        if not display and not has_para:
            chosen = el
            break
    if chosen is None:
        chosen = a14_elems[0]

    serialized = etree.tostring(chosen)
    cached.write_bytes(serialized)
    return etree.fromstring(serialized)


def set_omml_size(el: etree._Element, half_points: int) -> etree._Element:
    """Pin the run size of every <m:r> by injecting <w:rPr><w:sz/><w:szCs/>.

    half_points = pt * 2 (e.g. 30pt -> 60). In OOXML math the run size lives in
    the wordprocessingml run properties (<w:rPr>), NOT in <m:rPr> (which only
    holds math styling). Schema order inside <m:r> is: m:rPr?, w:rPr?, m:t.
    This makes the equation render at a deterministic size across viewers,
    fixing the "M_{ij} balloons huge" auto-scaling inconsistency.
    """
    m = f"{{{NS_M}}}"
    w = f"{{{NS_W}}}"
    val = str(int(half_points))
    for r in el.iter(f"{m}r"):
        wrpr = r.find(f"{w}rPr")
        if wrpr is None:
            wrpr = etree.Element(f"{w}rPr")
            mrpr = r.find(f"{m}rPr")
            if mrpr is not None:
                mrpr.addnext(wrpr)      # w:rPr follows m:rPr
            else:
                r.insert(0, wrpr)       # else first child (before m:t)
        for tag in (f"{w}sz", f"{w}szCs"):
            node = wrpr.find(tag)
            if node is None:
                node = etree.SubElement(wrpr, tag)
            node.set(f"{w}val", val)
    return el


def omml_glyph_count(el: etree._Element) -> int:
    """Rough width proxy: total <m:t> characters plus a structural penalty
    for fractions/radicals (which render wide). Used to auto-shrink only the
    equations that would overflow their box, keeping the rest uniform."""
    m = f"{{{NS_M}}}"
    chars = sum(len(t.text or "") for t in el.iter(f"{m}t"))
    nest = len(list(el.iter(f"{m}f"))) + len(list(el.iter(f"{m}rad")))
    return chars + nest * 2
