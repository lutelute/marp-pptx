"""MCP server for marp-pptx.

Exposes the marp-pptx engine as Model-Context-Protocol tools so an agent can
turn a request into an **editable** PowerPoint deck and visually self-check it:

    slide_types / slide_template   — learn the 57-type vocabulary + exact HTML
    list_themes                    — pick a theme by topic (claude/research/…)
    authoring_guide                — the condensed writing rules (markup, leads)
    list_presets / get_preset      — start from a curated deck (incl. paper-review)
    build_pptx                     — Marp markdown -> editable .pptx (+ lint + audit)
    check_deck                     — measured defects (overflow/overlap/contrast)
    preview_png                    — render slides to PNG so the agent can SEE them

Run:  marp-pptx-mcp        (stdio transport; configure it in your MCP client)
Install:  pip install "marp-pptx[mcp]"
"""
from __future__ import annotations

import contextlib
import io
import tempfile
from pathlib import Path

from mcp.server.fastmcp import FastMCP, Image

from marp_pptx.types import TYPE_REGISTRY, CATEGORIES

mcp = FastMCP("marp-pptx")

# Guard against a runaway deck exhausting the local process.
_MAX_MARKDOWN = 5_000_000  # ~5 MB of Marp markdown


def _check_markdown(markdown: str) -> None:
    if len(markdown) > _MAX_MARKDOWN:
        raise RuntimeError(
            f"markdown too large: {len(markdown)} bytes (limit {_MAX_MARKDOWN})"
        )


def _data_dir() -> Path:
    return Path(__file__).parent / "data"


def _type_by_name(name: str):
    for t in TYPE_REGISTRY:
        if t.name == name or t.css_class == name:
            return t
    return None


def _strip_frontmatter(text: str) -> str:
    if text.startswith("---"):
        end = text.find("---", 3)
        if end != -1:
            return text[end + 3:].strip()
    return text.strip()


def _themed_config(palette: str, math: str, density: str):
    """Build a ThemeConfig matching the CLI's `convert` defaults (claude theme)."""
    from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path

    tc = ThemeConfig.from_css(get_default_theme_path())
    tc.density = density
    if density == "keynote":
        tc.font_scale = 1.22
        tc.margin_scale = 1.12
    elif density == "dense":
        tc.font_scale = 0.88      # hearing-deck information load
    tc.math_mode = math
    chosen = get_palette_path(palette) if palette else get_palette_path("claude")
    if chosen:
        tc.apply_palette(chosen)
    return tc


def _build(markdown: str, base_path: Path, tc):
    """Parse + build a deck, returning (PptxBuilder, slides, lint_warnings)."""
    from marp_pptx.parser import parse_marp
    from marp_pptx.builder import PptxBuilder

    md_file = base_path / "deck.md"
    md_file.write_text(markdown, encoding="utf-8")
    slides = parse_marp(str(md_file))
    builder = PptxBuilder(base_path=base_path, theme=tc)
    lint = io.StringIO()
    with contextlib.redirect_stderr(lint):
        builder.build_all(slides)
    warnings = [ln[len("[lint] "):] for ln in lint.getvalue().splitlines() if ln.startswith("[lint] ")]
    return builder, slides, warnings


@mcp.tool()
def slide_types(category: str = "") -> list[dict]:
    """List the semantic slide types (the deck-building vocabulary).

    Pass a category key (structure/temporal/convergence/evaluation/knowledge/
    flow/narrative/meta) to filter, or "" for all types. Use this first to pick the
    right type for each slide, then `slide_template` for its exact HTML.
    """
    types = TYPE_REGISTRY
    if category:
        types = [t for t in types if t.category == category]
    return [
        {
            "name": t.name,
            "category": t.category,
            "category_ja": CATEGORIES.get(t.category, t.category),
            "geometry": t.geometry,
            "meaning": t.meaning,
            "use_when": t.use_when,
        }
        for t in types
    ]


@mcp.tool()
def slide_template(type_name: str) -> str:
    """Return the exact HTML skeleton for a slide type (fill it without breaking
    the `<div class>` structure). `type_name` is a type name like 'kpi' or
    'equation'. Use the result verbatim as one slide's body."""
    t = _type_by_name(type_name)
    if t is None:
        valid = ", ".join(sorted(x.name for x in TYPE_REGISTRY))
        return f"unknown type '{type_name}'. valid types: {valid}"
    body = _strip_frontmatter((_data_dir() / "templates" / t.template_file).read_text(encoding="utf-8"))
    return body


@mcp.tool()
def list_themes() -> list[dict]:
    """List every theme/palette with its intended topic, so the deck's look is
    chosen by subject (a palette that would fit any topic is a weak choice).

    Full themes carry layout behavior; academic-* entries recolor only.
    Notable: 'claude' (default; dark sandwich heroes + card shadows),
    'research' (doctoral-hearing spec: section crumb + n／m chrome, auto
    図N captions with 出典, 6-color chart ramp), 'beamer' (LaTeX Madrid),
    'midnight/teal/terracotta/cherry' (executive / public-trust / warm /
    keynote). Pass the name as build_pptx's `palette`.
    """
    import json
    import re as _re
    out: list[dict] = []
    pal_dir = _data_dir() / "themes" / "palettes"
    for cfg in sorted(pal_dir.glob("config-*.yaml")):
        text = cfg.read_text(encoding="utf-8")
        def _field(key: str) -> str:
            m = _re.search(rf"^{key}:\s*(.+)$", text, _re.MULTILINE)
            return m.group(1).strip().strip("'\"") if m else ""
        name = _field("name") or cfg.stem[len("config-"):]
        out.append({"name": name, "kind": "theme",
                    "title": _field("title"), "description": _field("description")})
    try:
        for pjt in json.loads((pal_dir / "palettes.json").read_text(encoding="utf-8")):
            out.append({"name": pjt.get("name", "?"), "kind": "palette",
                        "title": pjt.get("desc", ""), "description": pjt.get("desc", "")})
    except (OSError, ValueError):
        pass
    return out


@mcp.tool()
def authoring_guide() -> str:
    """The condensed marp-pptx writing rules — read once before drafting a deck
    (this is the skill knowledge for MCP clients that don't load SKILL.md)."""
    return (
        "# marp-pptx 執筆規範（要約）\n"
        "1. スライドは `---` 区切り、frontmatter は先頭に1つ（`marp: true`）。\n"
        "2. 各スライド冒頭に `<!-- _class: 型名 -->`。型は slide_types で選び、\n"
        "   HTML 骨組みは slide_template の出力を崩さず埋める（崩すと無言で空になる）。\n"
        "3. インライン: **太字** / `code` / $x^2$ / ==マーカー==（1枚1〜2個）。*斜体*は無効。\n"
        "4. H1 はラベルでなく要約・主張文。`##` は多くの型で色付きリード行\n"
        "   「トピック｜結論」として描かれる（title/divider/table/cols 系は別用途）。\n"
        "5. 1枚1メッセージ＋視覚アンカー1つ。箇条書きだけの連続は避け、比較は\n"
        "   cols-2/before-after、数値は kpi/big-number/chart、構造は flow/zone-* に。\n"
        "6. 高密度は sections 型（リード行＋本文の帯×2-4）＋ density='dense'。\n"
        "7. 図: `![w:800](path)` 相対パス。無い画像はプレースホルダ描画＋警告。\n"
        "   アニメGIFは埋め込まれ自動再生。ブロック図/ループ図は flow 型の\n"
        "   ```mermaid フェンス（flowchart LR、`-.->` は破線の戻り辺）。\n"
        "8. 表: Markdown 表。◎○△× セルは自動で緑/琥珀/赤に着色（サーベイ比較向け）。\n"
        "9. 文献報告は preset 'paper-review'（paper 型=書誌カードから始める）。\n"
        "10. 段階開示: sections/flow スライドに <!-- build --> で帯/ノードが1つずつ\n"
        "    増える連作に自動展開（未来分ゴースト）。導出・学習過程を見せる要。\n"
        "11. 説得の3点セット: 主張(H1)→根拠(数値+出典)→解釈(h2リード/結論箱)。\n"
        "    表紙直後に graphical-abstract（課題▶提案▶成果の一枚絵）を1枚。\n"
        "    導出は build で積み上げ、限界(言えること/言えないこと)を1枚明示。\n"
        "12. 出す前に必ず check_deck（error は修正必須）→ 必要なら preview_png で目視。\n"
        "13. 発表者ノート `<!-- note: ... -->`、出典 `<!-- source: ... -->`（図の出典｜行）。\n"
        "14. agenda の項目は同順の divider へ自動ハイパーリンクされる。\n"
    )


@mcp.tool()
def list_presets() -> list[dict]:
    """List curated starter decks (id/title/description). Use `get_preset` to
    fetch one as a ready-to-edit Marp markdown skeleton."""
    import json
    manifest = _data_dir() / "presets" / "manifest.json"
    try:
        return json.loads(manifest.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []


@mcp.tool()
def get_preset(preset_id: str) -> str:
    """Return a preset deck's full Marp markdown (e.g. 'academic-talk',
    'product', 'lecture', 'minimal'). A good starting point to then edit."""
    if "/" in preset_id or ".." in preset_id:
        return "invalid preset id"
    f = _data_dir() / "presets" / f"{preset_id}.md"
    if not f.is_file():
        ids = [p.stem for p in (_data_dir() / "presets").glob("*.md")]
        return f"unknown preset '{preset_id}'. available: {', '.join(sorted(ids))}"
    return f.read_text(encoding="utf-8")


@mcp.tool()
def build_pptx(
    markdown: str,
    output_path: str = "",
    palette: str = "",
    math: str = "omml",
    density: str = "academic",
) -> dict:
    """Convert Marp markdown to an EDITABLE .pptx and return its path.

    - markdown: the full deck (frontmatter `marp: true`, slides split by `---`,
      each with `<!-- _class: type -->` and the type's HTML).
    - output_path: where to write; default a temp file (absolute path returned).
    - palette: '' = claude (default); see `list_themes` for the full set and
      which theme fits which topic (research = doctoral-hearing spec with
      section crumbs + auto figure numbers).
    - math: 'omml' (PowerPoint-editable, needs pandoc) or 'png' (image; for
      LibreOffice/Keynote).
    - density: 'academic' (default), 'keynote' (bigger for projection), or
      'dense' (hearing-deck information load; looser authoring-lint limits).

    Authoring features the markdown can use (details: `authoring_guide`):
    **bold**, `code`, $math$, ==marker highlight==, `##` after the H1 becomes
    a colored lead line on most content types, ```mermaid fences on a `flow`
    slide draw editable block/loop diagrams, animated GIFs embed and autoplay,
    and ◎○△× table cells render color-coded.

    Returns {output_path, slide_count, lint_warnings, authoring_warnings,
    defects}. Heed lint_warnings (weak titles, missing structure, …) AND
    authoring_warnings — the latter flag content the tool could not honour:
    an unknown `_class` (rendered as a plain slide → likely a typo), a missing
    image (silently skipped), or an SVG that needs cairosvg. A non-empty
    authoring_warnings almost always means the deck is not what you intended.

    `defects` is the measured audit (see `check_deck`): text that overflows its
    box, overlaps, off-slide shapes, AA contrast failures, package faults. Any
    entry at severity "error" is user-visible — fix the content and rebuild
    rather than shipping it.
    """
    _check_markdown(markdown)
    if output_path:
        out = Path(output_path).expanduser().resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
    else:
        out = Path(tempfile.mkdtemp(prefix="marp_mcp_")) / "deck.pptx"
    tc = _themed_config(palette, math, density)
    with tempfile.TemporaryDirectory() as td:
        builder, slides, warnings = _build(markdown, Path(td), tc)
        builder.save(str(out))
    return {"output_path": str(out), "slide_count": len(slides),
            "lint_warnings": warnings, "authoring_warnings": list(builder.warnings),
            "defects": _audit(out)}


def _audit(pptx_path: Path, limit: int = 60) -> list[dict]:
    from dataclasses import asdict
    from marp_pptx.audit import audit_pptx
    from marp_pptx.pkgcheck import check_package

    findings = check_package(pptx_path) + audit_pptx(pptx_path)
    findings.sort(key=lambda f: ({"error": 0, "warn": 1, "info": 2}.get(f.severity, 3),
                                 f.slide))
    return [asdict(f) for f in findings[:limit]]


@mcp.tool()
def check_deck(markdown: str = "", pptx_path: str = "", palette: str = "",
               density: str = "academic") -> dict:
    """Measure a deck for the defects a reader would notice — no renderer needed.

    Give it `markdown` (built first, exactly as build_pptx would) or the
    `pptx_path` of a deck already on disk. Every text box is measured against
    the real advance widths of the fonts it names, so the answers are exact:

        overflow  the text needs more room than its box has (the #1 defect —
                  it clips in PowerPoint), with the size that would fit
        overlap   two text blocks cover each other, or nearly touch
        offslide  a shape leaves the slide or crowds the edge
        contrast  a colour pair below WCAG AA for its size
        font      a typeface missing here, or one whose preview lies
        package   a fault that makes PowerPoint call the file damaged
        deck      layout monotony, text-only slides

    Prefer this over `visual_lint`/`preview_png` as the first check: it needs
    nothing installed, runs in milliseconds, and names the exact overshoot.
    Use preview_png afterwards to look at what survived.

    Returns {ok, counts, findings[{kind, severity, slide, message, fix, shape}]}.
    """
    if not markdown and not pptx_path:
        raise RuntimeError("provide markdown or pptx_path")
    if markdown:
        _check_markdown(markdown)
        tc = _themed_config(palette, "omml", density)
        with tempfile.TemporaryDirectory() as td:
            builder, slides, _ = _build(markdown, Path(td), tc)
            out = Path(td) / "check.pptx"
            builder.save(str(out))
            findings = _audit(out)
    else:
        findings = _audit(Path(pptx_path).expanduser())
    counts: dict[str, int] = {}
    for f in findings:
        counts[f["severity"]] = counts.get(f["severity"], 0) + 1
    return {"ok": not counts.get("error"), "counts": counts, "findings": findings}


@mcp.tool(structured_output=False)
def preview_png(markdown: str, palette: str = "", max_slides: int = 12) -> list[Image]:
    """Render the deck to per-slide PNGs and return them as images, so you can
    SEE the result and fix layout/overflow before finalizing. Math is baked as
    PNG for fidelity. Requires LibreOffice (soffice) + pdftoppm; returns a short
    text note as a single image-less result if those are missing."""
    from marp_pptx.render import pptx_to_pngs, tools_available

    _check_markdown(markdown)
    if not tools_available():
        raise RuntimeError("preview_png needs LibreOffice (soffice) + pdftoppm installed")
    tc = _themed_config(palette, "png", "academic")
    with tempfile.TemporaryDirectory() as td:
        tdp = Path(td)
        builder, slides, _ = _build(markdown, tdp, tc)
        pptx = tdp / "deck.pptx"
        builder.save(str(pptx))
        pngs = pptx_to_pngs(pptx, tdp, dpi=96)
        if not pngs:
            raise RuntimeError("rendering failed (LibreOffice could not convert the deck)")
        return [Image(data=p.read_bytes(), format="png") for p in pngs[:max_slides]]


@mcp.tool()
def read_paper(path: str, extract_figures: bool = True) -> dict:
    """Ingest a paper PDF (local path or arXiv URL) into structured material to
    draft a deck FROM (not from memory).

    Returns {title, abstract, sections[{heading,text}], figures[{caption,path}],
    numbers[{value,context}], page_count}. Draft each slide from these sections,
    feature the extracted `numbers` as KPIs (never invent values), and reference
    extracted figure `path`s with `![w:..](path)`. Then ground the draft with
    check_deck_against_source(). full_text is omitted here to keep the response
    small — pass paper_path to check_deck_against_source for grounding.
    """
    from marp_pptx.ingest import read_paper as _rp

    r = _rp(path, extract_figures=extract_figures)
    r.pop("full_text", None)  # large; grounding re-reads from paper_path
    # trim section bodies so the response stays manageable
    for s in r["sections"]:
        s["text"] = s["text"][:2500]
    return r


@mcp.tool()
def read_repo(path: str) -> dict:
    """Summarize a local code repository for slide drafting.

    Returns {name, readme, tree, languages, key_files, file_count, summary} so
    an agent can describe what the project is/does without inventing it.
    """
    from marp_pptx.ingest import read_repo as _rr

    return _rr(path)


@mcp.tool()
def check_deck_against_source(markdown: str, paper_path: str = "", source_text: str = "") -> dict:
    """Ground a drafted deck: flag every result-number that does NOT appear in
    the source paper.

    Provide the paper via `paper_path` (a local PDF path; re-read server-side) or
    inline `source_text`. Returns {score, supported[], unsupported[{value,slide,
    context}]}. Treat each `unsupported` number as a likely hallucination — fix
    it against the paper before finalizing. Always run this before build_pptx.
    """
    _check_markdown(markdown)
    src = source_text
    if paper_path:
        from marp_pptx.ingest import read_paper as _rp
        src = _rp(paper_path, extract_figures=False)["full_text"]
    if not src:
        raise RuntimeError("provide paper_path or source_text")
    from marp_pptx.ingest import check_fidelity
    return check_fidelity(markdown, src)


@mcp.tool()
def visual_lint(markdown: str, palette: str = "claude") -> list:
    """Render the deck and return deterministic LAYOUT warnings per slide
    (overflow / very-sparse / top-or-bottom skew) — no AI. Use with preview_png:
    this tells you WHICH slides look off; preview_png lets you SEE them. Fix the
    flagged slides, then re-run. Requires LibreOffice + pdftoppm."""
    _check_markdown(markdown)
    from marp_pptx.visuallint import lint_deck

    return lint_deck(markdown, palette=palette)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
