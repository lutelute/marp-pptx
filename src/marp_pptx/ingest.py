"""Ingestion + grounding for "hand over a paper + repo, make slides".

This is the front half of the deck pipeline: turn a real PDF paper and a code
repository into structured material an agent can draft *from* (instead of from
memory), plus a fidelity gate that flags deck claims/numbers that don't trace
back to the source.

    read_paper(pdf)            -> {title, abstract, sections, figures, numbers, full_text}
    read_repo(path)            -> {name, readme, tree, languages, key_files, summary}
    check_fidelity(md, source) -> {score, supported, unsupported}

PDF reading needs PyMuPDF: ``pip install "marp-pptx[ingest]"``.
"""
from __future__ import annotations

import re
import urllib.request
from pathlib import Path

# Section headings commonly found in papers (case-insensitive, start of line).
_SECTION_WORDS = (
    "abstract", "introduction", "background", "related work", "preliminaries",
    "method", "methods", "methodology", "approach", "model", "architecture",
    "experiments", "experimental setup", "evaluation", "results", "analysis",
    "discussion", "ablation", "limitations", "conclusion", "conclusions",
    "future work", "acknowledg", "references", "appendix",
)
_HEADING_RE = re.compile(
    r"^\s*(?:(\d+(?:\.\d+)*)\s+)?([A-Z][A-Za-z][^\n]{0,60})\s*$"
)
# numbers worth grounding: percentages, decimals, multipliers, ints with units
_NUMBER_RE = re.compile(
    r"(?<![\w.])(\d+(?:\.\d+)?\s?%|\d+\.\d+[a-zA-Z]*|\d+(?:\.\d+)?\s?[×x]\b|\d{2,})"
)


def _looks_like_heading(line: str) -> bool:
    s = line.strip()
    if not s or len(s) > 64:
        return False
    low = s.lower()
    if any(low == w or low.startswith(w + " ") or low.lstrip("0123456789. ").startswith(w)
           for w in _SECTION_WORDS):
        return True
    # "3 Method" / "3.1 Scaled Dot-Product Attention"
    if re.match(r"^\d+(\.\d+)*\s+[A-Z]", s) and len(s.split()) <= 9:
        return True
    return False


def _split_sections(full_text: str) -> list[dict]:
    sections: list[dict] = []
    cur = {"heading": "(preamble)", "text": []}
    for line in full_text.splitlines():
        if _looks_like_heading(line):
            if cur["text"]:
                sections.append({"heading": cur["heading"], "text": "\n".join(cur["text"]).strip()})
            cur = {"heading": line.strip(), "text": []}
        else:
            cur["text"].append(line)
    if cur["text"]:
        sections.append({"heading": cur["heading"], "text": "\n".join(cur["text"]).strip()})
    return [s for s in sections if s["text"]]


def _extract_numbers(full_text: str, limit: int = 80) -> list[dict]:
    seen: set[str] = set()
    out: list[dict] = []
    for m in _NUMBER_RE.finditer(full_text):
        val = m.group(1).strip()
        if val in seen:
            continue
        seen.add(val)
        start, end = max(0, m.start() - 50), min(len(full_text), m.end() + 50)
        ctx = " ".join(full_text[start:end].split())
        out.append({"value": val, "context": ctx})
        if len(out) >= limit:
            break
    return out


def _maybe_download(path_or_url: str) -> Path:
    if path_or_url.startswith(("http://", "https://")):
        url = path_or_url
        # arXiv abstract page -> pdf
        m = re.search(r"arxiv\.org/abs/([\w.]+)", url)
        if m:
            url = f"https://arxiv.org/pdf/{m.group(1)}.pdf"
        dest = Path("/tmp") / (re.sub(r"\W+", "_", url)[-60:] + ".pdf")
        urllib.request.urlretrieve(url, dest)  # noqa: S310 (user-supplied source)
        return dest
    return Path(path_or_url).expanduser()


def read_paper(path: str, extract_figures: bool = True, figures_dir: str | None = None) -> dict:
    """Extract structured content from a paper PDF (or arXiv URL).

    Returns {title, abstract, sections[{heading,text}], figures[{caption,path}],
    numbers[{value,context}], full_text, page_count}. Draft the deck *from*
    this text — then ground it with check_fidelity().
    """
    try:
        import fitz  # PyMuPDF
    except ImportError as e:
        raise RuntimeError('read_paper needs PyMuPDF: pip install "marp-pptx[ingest]"') from e

    pdf_path = _maybe_download(path)
    if not pdf_path.is_file():
        raise FileNotFoundError(f"paper not found: {pdf_path}")

    doc = fitz.open(str(pdf_path))
    pages = [doc.load_page(i).get_text("text") for i in range(doc.page_count)]
    full_text = "\n".join(pages)

    # Title: prefer the largest-font text near the top of page 1 (robust to
    # license/attribution headers); fall back to metadata, then first line.
    _bad = ("attribution", "google", "grants", "license", "copyright",
            "microsoft word", "permission", "reproduce", "tables and figures",
            "arxiv:", "preprint", "under review", "scholarly works")

    def _clean(s: str) -> bool:
        return 8 < len(s) < 140 and not any(b in s.lower() for b in _bad)

    title = ""
    if doc.page_count:
        spans = []  # (font_size, y_top, text)
        d = doc.load_page(0).get_text("dict")
        for block in d.get("blocks", []):
            for line in block.get("lines", []):
                txt = "".join(sp.get("text", "") for sp in line.get("spans", [])).strip()
                if not txt or not line.get("spans"):
                    continue
                size = max(sp.get("size", 0) for sp in line["spans"])
                y = line["spans"][0].get("bbox", [0, 0, 0, 0])[1]
                spans.append((size, y, txt))
        if spans:
            top_size = max(s[0] for s in spans)
            big = [t for sz, y, t in sorted(spans, key=lambda s: s[1])
                   if sz >= top_size - 0.5 and y < 400]
            cand = " ".join(big).strip()
            if _clean(cand):
                title = cand
    if not title:
        meta = (doc.metadata or {}).get("title") or ""
        title = meta if _clean(meta) else ""
    if not title:
        for line in (pages[0].splitlines() if pages else []):
            if _clean(line.strip()):
                title = line.strip()
                break

    sections = _split_sections(full_text)
    abstract = ""
    for s in sections:
        if s["heading"].lower().startswith("abstract"):
            abstract = s["text"][:1500]
            break

    figures: list[dict] = []
    if extract_figures:
        fdir = Path(figures_dir) if figures_dir else (pdf_path.parent / f"{pdf_path.stem}_figures")
        fdir.mkdir(parents=True, exist_ok=True)
        cap_re = re.compile(r"(Figure|Fig\.?)\s*(\d+)[.:]?\s*([^\n]{0,120})", re.IGNORECASE)
        for pno in range(doc.page_count):
            page = doc.load_page(pno)
            caps = {m.group(2): m.group(0).strip() for m in cap_re.finditer(pages[pno])}
            for idx, img in enumerate(page.get_images(full=True)):
                xref = img[0]
                try:
                    info = doc.extract_image(xref)
                except Exception:
                    continue
                if len(info.get("image", b"")) < 4000:  # skip tiny logos/rules
                    continue
                ext = info.get("ext", "png")
                out = fdir / f"p{pno + 1}_{idx}.{ext}"
                out.write_bytes(info["image"])
                cap = next(iter(caps.values()), "") if caps else ""
                figures.append({"caption": cap, "path": str(out), "page": pno + 1})

    doc.close()
    return {
        "title": title.strip(),
        "abstract": abstract.strip(),
        "sections": sections,
        "figures": figures,
        "numbers": _extract_numbers(full_text),
        "full_text": full_text,
        "page_count": len(pages),
    }


_SKIP_DIRS = {".git", "node_modules", ".venv", "venv", "__pycache__", "dist",
              "build", ".pytest_cache", ".mypy_cache", ".ruff_cache", ".idea",
              ".tox", "site-packages", ".egg-info"}
_LANG_EXT = {
    ".py": "Python", ".js": "JavaScript", ".ts": "TypeScript", ".tsx": "TypeScript",
    ".jsx": "JavaScript", ".java": "Java", ".go": "Go", ".rs": "Rust", ".c": "C",
    ".cpp": "C++", ".h": "C/C++", ".rb": "Ruby", ".php": "PHP", ".cs": "C#",
    ".swift": "Swift", ".kt": "Kotlin", ".m": "MATLAB/ObjC", ".jl": "Julia",
    ".r": "R", ".sh": "Shell", ".html": "HTML", ".css": "CSS", ".md": "Markdown",
}
_KEY_FILES = ("pyproject.toml", "setup.py", "package.json", "Cargo.toml",
              "go.mod", "pom.xml", "requirements.txt", "Makefile", "Dockerfile")


def read_repo(path: str, max_files: int = 4000, tree_limit: int = 200) -> dict:
    """Summarize a local code repository for slide drafting.

    Returns {name, readme, tree, languages, key_files, file_count, summary}.
    The README + language mix + key manifests give an agent enough to describe
    what the project is without inventing it.
    """
    root = Path(path).expanduser().resolve()
    if not root.is_dir():
        raise NotADirectoryError(f"repo not found: {root}")

    readme = ""
    for name in ("README.md", "README.rst", "README.txt", "README", "readme.md"):
        p = root / name
        if p.is_file():
            readme = p.read_text(encoding="utf-8", errors="ignore")
            break

    languages: dict[str, int] = {}
    key_files: list[str] = []
    tree: list[str] = []
    file_count = 0
    for cur, dirs, files in __import__("os").walk(root):
        dirs[:] = [d for d in dirs if d not in _SKIP_DIRS and not d.endswith(".egg-info")]
        rel = Path(cur).relative_to(root)
        depth = len(rel.parts)
        if depth <= 2 and len(tree) < tree_limit:
            for d in sorted(dirs):
                tree.append(f"{'  ' * depth}{(rel / d)}/")
        for f in files:
            file_count += 1
            if file_count > max_files:
                break
            ext = Path(f).suffix.lower()
            if ext in _LANG_EXT:
                languages[_LANG_EXT[ext]] = languages.get(_LANG_EXT[ext], 0) + 1
            if f in _KEY_FILES:
                key_files.append(str(rel / f) if str(rel) != "." else f)
            if depth <= 2 and len(tree) < tree_limit and ext in _LANG_EXT:
                tree.append(f"{'  ' * depth}{(rel / f) if str(rel) != '.' else f}")

    langs_sorted = sorted(languages.items(), key=lambda kv: -kv[1])
    top_langs = ", ".join(f"{k} ({v})" for k, v in langs_sorted[:5])
    summary = (
        f"{root.name}: {file_count} files. Languages: {top_langs or 'n/a'}. "
        f"Manifests: {', '.join(key_files) or 'none'}."
    )
    return {
        "name": root.name,
        "readme": readme[:8000],
        "tree": tree[:tree_limit],
        "languages": dict(langs_sorted),
        "key_files": key_files,
        "file_count": file_count,
        "summary": summary,
    }


def _slides_of(markdown: str) -> list[str]:
    parts, cur = [], []
    for line in markdown.splitlines():
        if line.strip() == "---":
            if cur:
                parts.append("\n".join(cur))
            cur = []
        else:
            cur.append(line)
    if cur:
        parts.append("\n".join(cur))
    return parts


# Result-like numbers worth grounding: decimals, percentages, multipliers.
# Bare integers / 4-digit years are too noisy (bullet counts, dates) to gate on.
_CLAIM_NUM_RE = re.compile(r"(\d+\.\d+\s?%?|\d+\s?%|\d+(?:\.\d+)?\s?[×x](?![a-zA-Z]))")


def check_fidelity(markdown: str, source_text: str) -> dict:
    """Flag deck result-numbers that do NOT appear in the source text.

    The content-fidelity gate the renderer's lint lacks: every quantitative
    claim (BLEU/accuracy/speedup/…) in the deck should trace back to the paper.
    Returns {score, supported[], unsupported[{value, slide, context}]}. Any
    unsupported number is a likely hallucination — re-check it against the paper
    before shipping. Matches whole numeric tokens (no substring false hits);
    bare integers and years are ignored as too noisy. Heuristic: verifies
    numeric presence, not full semantics.
    """
    # Whole numeric tokens present in the source.
    src_nums = {m.group(0) for m in re.finditer(r"\d+(?:\.\d+)?", source_text)}

    def _present(val: str) -> bool:
        core = val.replace(" ", "").rstrip("%×x")
        return core in src_nums

    supported, unsupported = [], []
    for idx, slide in enumerate(_slides_of(markdown), 1):
        for m in _CLAIM_NUM_RE.finditer(slide):
            val = m.group(1).strip()
            entry = {"value": val, "slide": idx}
            if _present(val):
                supported.append(entry)
            else:
                ctx = " ".join(slide[max(0, m.start() - 30):m.end() + 30].split())
                unsupported.append({**entry, "context": ctx})

    total = len(supported) + len(unsupported)
    score = round(100 * len(supported) / total) if total else 100
    return {"score": score, "supported": supported, "unsupported": unsupported}

