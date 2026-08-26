"""CLI entry point for marp-pptx."""

from __future__ import annotations

import sys
from pathlib import Path

import click

from marp_pptx import __version__


@click.group(invoke_without_command=True)
@click.version_option(__version__)
@click.pass_context
def main(ctx):
    """marp-pptx: Convert Marp markdown to editable PowerPoint with 50+ semantic slide types."""
    if ctx.invoked_subcommand is None:
        click.echo(ctx.get_help())


@main.command()
@click.argument("input_file", type=click.Path(exists=True))
@click.option("-o", "--output", help="Output .pptx path")
@click.option("-p", "--palette", help="Palette name (e.g. navy, copper, earth)")
@click.option("-t", "--theme", help="Custom palette CSS path")
@click.option("--font-scale", type=float, default=1.0, help="Font size multiplier (0.7-1.3)")
@click.option("--math", type=click.Choice(["omml", "png"]), default="omml",
              help="Math rendering: omml (editable, default) or png (matplotlib image — "
                   "use for LibreOffice/Keynote where OMML renders poorly)")
@click.option("--density", type=click.Choice(["academic", "keynote"]), default="academic",
              help="Content density: academic (dense, default) or keynote (sparse, "
                   "bigger type for projection)")
def convert(input_file: str, output: str | None, palette: str | None, theme: str | None,
            font_scale: float, math: str, density: str):
    """Convert a Marp markdown file to editable PPTX."""
    from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
    from marp_pptx.parser import parse_marp
    from marp_pptx.builder import PptxBuilder

    input_path = Path(input_file)

    # Load theme
    tc = ThemeConfig.from_css(get_default_theme_path())
    tc.density = density
    # Keynote density: bigger type + more generous margins (projection-first).
    if density == "keynote":
        font_scale *= 1.22
        tc.margin_scale = 1.12
    tc.font_scale = font_scale
    tc.math_mode = math

    # Apply palette. With no explicit palette/theme, fall back to the "claude"
    # default (Anthropic cream + clay) so a bare `convert deck.md` looks polished.
    applied = "claude (default)"
    if palette:
        palette_path = get_palette_path(palette)
        if palette_path:
            tc.apply_palette(palette_path)
            applied = palette
        else:
            click.echo(f"Warning: palette '{palette}' not found, using claude", err=True)
            mp = get_palette_path("claude")
            if mp:
                tc.apply_palette(mp)
    elif theme:
        tc.apply_palette(Path(theme))
        applied = theme
    else:
        mp = get_palette_path("claude")
        if mp:
            tc.apply_palette(mp)

    print(f"[theme] {applied}  latin={tc.font}  ea={tc.font_ea}  head={tc.font_head}  font_scale={tc.font_scale}", file=sys.stderr)

    # Parse
    slides = parse_marp(str(input_path))
    click.echo(f"Parsed {len(slides)} slides", err=True)
    if not slides:
        click.echo("Warning: no slides found — is this a Marp deck? "
                   "(slides are separated by '---'; the first '---' block is "
                   "front matter). Writing an empty presentation.", err=True)

    # Build
    builder = PptxBuilder(base_path=input_path.parent, theme=tc)
    builder.build_all(slides)

    # Save
    output_path = output or str(input_path.with_name(input_path.stem + "_editable.pptx"))
    builder.save(output_path)

    click.echo(f"Saved: {output_path}", err=True)
    click.echo(f"  {len(slides)} slides, all editable text boxes", err=True)
    if builder.warnings:
        click.echo(f"  {len(builder.warnings)} warning(s) — see [warn] lines above",
                   err=True)


@main.command("types")
@click.option("--category", "-c", help="Filter by category")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
def list_types(category: str | None, as_json: bool):
    """List all available slide types with their semantic meanings."""
    from marp_pptx.types import TYPE_REGISTRY, CATEGORIES

    types = TYPE_REGISTRY
    if category:
        types = [t for t in types if t.category == category]

    if as_json:
        import json
        data = [
            {
                "name": t.name,
                "css_class": t.css_class,
                "category": t.category,
                "category_ja": CATEGORIES.get(t.category, t.category),
                "geometry": t.geometry,
                "meaning": t.meaning,
                "use_when": t.use_when,
                "template": t.template_file,
            }
            for t in types
        ]
        click.echo(json.dumps(data, ensure_ascii=False, indent=2))
        return

    # Table output
    click.echo(f"{'Type':<18} {'Category':<12} {'Geometry':<16} {'Meaning':<24} Use When")
    click.echo("-" * 100)

    current_cat = None
    for t in types:
        cat_ja = CATEGORIES.get(t.category, t.category)
        if t.category != current_cat:
            current_cat = t.category
            if t != types[0]:
                click.echo()
        click.echo(f"{t.name:<18} {cat_ja:<12} {t.geometry:<16} {t.meaning:<24} {t.use_when}")

    click.echo(f"\nTotal: {len(types)} types in {len(CATEGORIES)} categories")
    if not category:
        click.echo(f"Categories: {', '.join(f'{v} ({k})' for k, v in CATEGORIES.items())}")


@main.command("themes")
def list_themes():
    """List available themes / color palettes."""
    import json

    click.echo("Themes (color + layout):")
    click.echo(f"  {'claude':<14} Anthropic cream + clay・中央バランス  accent #d97757   (default)")
    click.echo(f"  {'minimal':<14} 洗練ミニマル白基調・中央バランス      accent #c2410c")
    click.echo(f"  {'tmu-cs':<14} TMU グリーン白地・学術 (緑見出し+細下線) accent #006543")
    click.echo(f"  {'research':<14} PowerPoint マスタ灰・研究審査 (高密度・上詰め) accent #dd5400")
    click.echo(f"  {'beamer':<14} LaTeX beamer 紺・定理ブロック+フッターバー    accent #b8860b")
    click.echo(f"  {'midnight':<14} 紺ヒーロー×ロイヤルブルー・経営/戦略       accent #3d52c4")
    click.echo(f"  {'terracotta':<14} テラコッタ×セージ・講義/教育の温かさ      accent #b85042")
    click.echo(f"  {'teal':<14} 深ティール×ミント・医療/公共の信頼感       accent #00907f")
    click.echo(f"  {'cherry':<14} チェリー×ネイビー・キーノート/強い主張      accent #c31126")

    palettes_dir = Path(__file__).parent / "data" / "themes" / "palettes"
    pj = palettes_dir / "palettes.json"
    if pj.exists():
        click.echo("\nColor palettes (academic family):")
        try:
            for p in json.loads(pj.read_text(encoding="utf-8")):
                name = p.get("name", "?")
                desc = p.get("desc", "")
                accent = p.get("accent", "")
                click.echo(f"  {name:<14} {desc:<40} accent {accent}")
        except Exception:
            pass
    click.echo("\nUsage:  marp-pptx convert deck.md -p <name>")
    click.echo("        marp-pptx convert deck.md            # claude (default)")


def _build_catalog(output_path, tc):
    """Build a catalog PPTX with one slide per registered type.

    Returns the ordered list of types actually included (those whose template
    file exists), so callers can map per-slide output back to each type.
    """
    import tempfile

    from marp_pptx.parser import parse_marp
    from marp_pptx.builder import PptxBuilder
    from marp_pptx.types import TYPE_REGISTRY

    templates_dir = Path(__file__).parent / "data" / "templates"
    included = []
    all_md = "---\nmarp: true\n---\n"
    for t in TYPE_REGISTRY:
        template_path = templates_dir / t.template_file
        if not template_path.exists():
            continue
        text = template_path.read_text(encoding="utf-8")
        if text.startswith("---"):
            end = text.find("---", 3)
            if end != -1:
                text = text[end + 3:]
        all_md += f"\n---\n{text.strip()}\n"
        included.append(t)

    with tempfile.NamedTemporaryFile(mode="w", suffix=".md", delete=False, encoding="utf-8") as f:
        f.write(all_md)
        tmp_path = Path(f.name)
    try:
        slides = parse_marp(str(tmp_path))
        builder = PptxBuilder(base_path=templates_dir, theme=tc)
        builder.build_all(slides)
        builder.save(str(output_path))
    finally:
        tmp_path.unlink(missing_ok=True)
    return included


@main.command()
@click.option("-o", "--output", default="type_catalog.pptx", help="Output catalog PPTX")
@click.option("-p", "--palette", help="Palette name")
def preview(output: str, palette: str | None):
    """Generate a visual catalog PPTX showing all slide types."""
    from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path

    tc = ThemeConfig.from_css(get_default_theme_path())
    if palette:
        palette_path = get_palette_path(palette)
        if palette_path:
            tc.apply_palette(palette_path)

    included = _build_catalog(Path(output), tc)
    click.echo(f"Generated catalog: {output} ({len(included)} slides)")


@main.command("render-gallery")
@click.option("-o", "--output-dir", default=None,
              help="Output dir (default: the web UI's static/type-gallery)")
@click.option("-p", "--palette", help="Palette name (default: claude theme)")
@click.option("--width", default=640, type=int, help="Thumbnail width in px")
@click.option("--dpi", default=110, type=int, help="Render DPI before downscaling")
def render_gallery(output_dir: str | None, palette: str | None, width: int, dpi: int):
    """Render a thumbnail PNG for every slide type into the web gallery dir."""
    import tempfile

    from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
    from marp_pptx.render import pptx_batch_to_pdf, pdf_to_pngs, tools_available

    if not tools_available():
        raise SystemExit("render-gallery requires LibreOffice (soffice) + pdftoppm")
    try:
        from PIL import Image
    except ImportError:
        raise SystemExit("render-gallery requires Pillow")

    out = Path(output_dir) if output_dir else Path(__file__).parent / "web" / "static" / "type-gallery"
    out.mkdir(parents=True, exist_ok=True)

    from marp_pptx.parser import parse_marp
    from marp_pptx.builder import PptxBuilder
    from marp_pptx.types import TYPE_REGISTRY

    tc = ThemeConfig.from_css(get_default_theme_path())
    # Thumbnails are raster previews via LibreOffice, which ignores OMML sizing —
    # bake math as PNG so equations look right.
    tc.math_mode = "png"
    # No -p means the shipped default (claude), same as `convert` — the old
    # bare-academic fallback made the gallery show a theme the CLI never ships.
    pp = get_palette_path(palette or "claude")
    if pp:
        tc.apply_palette(pp)

    templates_dir = Path(__file__).parent / "data" / "templates"
    with tempfile.TemporaryDirectory() as tmp:
        td = Path(tmp)
        # Build one isolated single-slide PPTX per type (so a content overflow
        # in LibreOffice can't shift another type's page mapping).
        included, pptx_paths = [], []
        for t in TYPE_REGISTRY:
            tpl = templates_dir / t.template_file
            if not tpl.exists():
                continue
            text = tpl.read_text(encoding="utf-8")
            if text.startswith("---"):
                end = text.find("---", 3)
                if end != -1:
                    text = text[end + 3:]
            md_path = td / f"{t.name}.md"
            md_path.write_text("---\nmarp: true\n---\n\n" + text.strip() + "\n", encoding="utf-8")
            builder = PptxBuilder(base_path=templates_dir, theme=tc)
            builder.build_all(parse_marp(str(md_path)))
            pptx = td / f"{t.name}.pptx"
            builder.save(str(pptx))
            included.append(t)
            pptx_paths.append(pptx)

        # One LibreOffice invocation for all decks, then first page of each.
        pdfs = pptx_batch_to_pdf(pptx_paths, td)
        rendered = 0
        for pdf, t in zip(pdfs, included):
            pngs = pdf_to_pngs(pdf, td, dpi=dpi, prefix=t.name)
            if not pngs:
                click.echo(f"  warning: failed to render {t.name}", err=True)
                continue
            im = Image.open(pngs[0]).convert("RGB")  # page 1 only
            w, h = im.size
            im = im.resize((width, round(h * width / w)), Image.LANCZOS)
            im.save(out / f"{t.name}.png", optimize=True)
            rendered += 1
    click.echo(f"Rendered {rendered}/{len(included)} thumbnails ({width}px) to {out}")


@main.command()
@click.argument("deck", type=click.Path(exists=True))
@click.option("-p", "--palette", help="Palette to build with, when DECK is a .md")
@click.option("--strict", is_flag=True,
              help="Exit non-zero when anything at warn level or worse is found")
@click.option("--json", "as_json", is_flag=True, help="Machine-readable output")
@click.option("--limit", default=40, type=int, help="Max findings to print")
def doctor(deck: str, palette: str | None, strict: bool, as_json: bool, limit: int):
    """Check a deck for the defects a reader would notice.

    Takes a .pptx, or a .md which it builds first. Measures every text box
    against real font metrics — no LibreOffice, no API key — and reports text
    that overflows its box, shapes that overlap or leave the slide, colour
    combinations below WCAG AA, fonts that will not render as measured, and
    package faults that make PowerPoint call the file damaged.
    """
    from marp_pptx.audit import audit_pptx, format_findings
    from marp_pptx.pkgcheck import check_package

    path = Path(deck)
    if path.suffix.lower() == ".md":
        pptx_path = Path(_build_for_check(path, palette))
        click.echo(f"built {pptx_path}", err=True)
    else:
        pptx_path = path

    findings = check_package(pptx_path) + audit_pptx(pptx_path)
    findings.sort(key=lambda f: ({"error": 0, "warn": 1, "info": 2}.get(f.severity, 3),
                                 f.slide))

    if as_json:
        import json
        from dataclasses import asdict
        click.echo(json.dumps([asdict(f) for f in findings], ensure_ascii=False,
                              indent=2))
    else:
        click.echo(format_findings(findings, limit=limit))

    worst = min((f.severity for f in findings), key=lambda s:
                {"error": 0, "warn": 1, "info": 2}.get(s, 3), default="none")
    if worst == "error" or (strict and worst in ("error", "warn")):
        raise SystemExit(1)


def _build_for_check(md_path: Path, palette: str | None) -> str:
    """Build `md_path` into a temp .pptx so doctor can inspect the real output."""
    import tempfile
    from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
    from marp_pptx.parser import parse_marp
    from marp_pptx.builder import PptxBuilder

    tc = ThemeConfig.from_css(get_default_theme_path())
    pal = get_palette_path(palette or "claude")
    if pal:
        tc.apply_palette(pal)
    slides = parse_marp(str(md_path))
    builder = PptxBuilder(base_path=md_path.parent, theme=tc)
    builder.build_all(slides)
    out = Path(tempfile.mkdtemp(prefix="marp-doctor-")) / (md_path.stem + ".pptx")
    builder.save(str(out))
    return str(out)


@main.command()
@click.option("--host", default="127.0.0.1")
@click.option("--port", default=8080, type=int)
def serve(host: str, port: int):
    """Start the web UI for shared/lab use (requires marp-pptx[web])."""
    try:
        from marp_pptx.web.app import create_app
    except ImportError:
        click.echo("Web UI requires Flask. Install with: pip install marp-pptx[web]", err=True)
        raise SystemExit(1)

    app = create_app()
    click.echo(f"Starting marp-pptx web UI on http://{host}:{port}")
    app.run(host=host, port=port)


@main.command("from-paper")
@click.argument("paper")
@click.option("--repo", default=None, help="Local repo path to summarize alongside the paper")
@click.option("-o", "--output", default=None, help="Output .pptx (default: <paper>_deck.pptx)")
@click.option("-p", "--palette", default="claude", help="Palette (default: claude)")
@click.option("--math", type=click.Choice(["omml", "png"]), default="omml")
@click.option("--slides", default=10, type=int, help="Target slide count")
@click.option("--model", default="claude-sonnet-4-6", help="Anthropic model")
@click.option("--no-review", is_flag=True, help="Skip the semantic faithfulness pass")
def from_paper(paper, repo, output, palette, math, slides, model, no_review):
    """Turnkey: a paper (PDF or arXiv URL) [+ repo] -> a grounded editable PPTX.

    Ingests the source, drafts a deck grounded in the extracted text, auto-repairs
    any number that doesn't trace back to the paper, then builds the PPTX.
    Requires: pip install "marp-pptx[ai]" and ANTHROPIC_API_KEY.
    """
    from marp_pptx.autodeck import build_deck_from_paper

    try:
        res = build_deck_from_paper(
            paper, repo_path=repo, palette=palette, math=math, out=output,
            n_slides=slides, semantic_review=not no_review, model=model,
        )
    except RuntimeError as e:
        raise SystemExit(str(e))
    click.echo(f"Built {res['slide_count']} slides -> {res['output_path']}")
    click.echo(f"  fidelity {res['fidelity']['score']}/100 (repair rounds: {res['rounds']})")
    if res["fidelity"]["unsupported"]:
        click.echo("  ⚠ still-unsupported numbers (verify against the paper): "
                   + ", ".join(u["value"] for u in res["fidelity"]["unsupported"]), err=True)
    if res["fidelity"].get("mislabeled"):
        click.echo("  ⚠ possibly-mislabelled numbers: "
                   + ", ".join(f"{m['value']}({m['label']})" for m in res["fidelity"]["mislabeled"]), err=True)
    left = [d for d in res.get("defects", []) if d["severity"] == "error"]
    if left:
        click.echo("  ⚠ measured defects remain on slides: "
                   + ", ".join(str(d["slide"]) for d in left)
                   + "  (run `marp-pptx doctor` for detail)", err=True)
    if res.get("visual"):
        click.echo("  ⚠ layout warnings remain on slides: "
                   + ", ".join(str(v["slide"]) for v in res["visual"]), err=True)
    if res["review"] and res["review"].upper() != "OK":
        click.echo("  review notes:\n" + res["review"], err=True)
