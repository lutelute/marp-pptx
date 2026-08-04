"""Turnkey: a paper (+ repo) -> a grounded, editable .pptx in one call.

`build_deck_from_paper` is the single entry point the toolkit was missing:
ingest the source, have an LLM draft the deck *grounded* in the extracted text,
auto-repair any number that doesn't trace back to the paper (and, optionally,
any claim a reviewer LLM finds unsupported), then build the editable PPTX.

The LLM is injectable (`llm=` is a ``str -> str`` callable) so the orchestration
is testable without a network. The default backend uses the Anthropic API and
needs ``pip install "marp-pptx[ai]"`` + ``ANTHROPIC_API_KEY``.
"""
from __future__ import annotations

import re
from pathlib import Path

from marp_pptx.ingest import read_paper, read_repo, check_fidelity

_DEFAULT_MODEL = "claude-sonnet-4-6"
_SKILL_DIR = Path(__file__).parent.parent.parent / "skills" / "marp-pptx"


def _skill_context() -> str:
    """The authoring guide + per-type HTML skeletons, used as the LLM's rules."""
    parts = []
    for rel in ("SKILL.md", "references/type-skeletons.md"):
        p = _SKILL_DIR / rel
        if p.is_file():
            parts.append(p.read_text(encoding="utf-8"))
    if not parts:  # installed without the skill tree → fall back to the registry
        from marp_pptx.types import TYPE_REGISTRY
        parts.append("Slide types:\n" + "\n".join(
            f"- {t.name}: {t.meaning} ({t.use_when})" for t in TYPE_REGISTRY))
    return "\n\n".join(parts)


def _anthropic_llm(model: str = _DEFAULT_MODEL):
    try:
        from anthropic import Anthropic
    except ImportError as e:  # pragma: no cover - env dependent
        raise RuntimeError('default LLM needs: pip install "marp-pptx[ai]" + ANTHROPIC_API_KEY') from e
    client = Anthropic()

    def _call(prompt: str) -> str:  # pragma: no cover - network
        msg = client.messages.create(
            model=model, max_tokens=8000,
            messages=[{"role": "user", "content": prompt}],
        )
        return "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")

    return _call


def _strip_fences(text: str) -> str:
    """Pull the Marp markdown out of an LLM reply (handle ```markdown fences)."""
    m = re.search(r"```(?:markdown|md)?\s*\n(.*?)```", text, re.S)
    body = m.group(1) if m else text
    body = body.strip()
    # ensure a frontmatter marker exists
    if not body.lstrip().startswith("---"):
        body = "---\nmarp: true\n---\n\n" + body
    return body


def _draft_prompt(paper: dict, repo: dict | None, n_slides: int) -> str:
    nums = ", ".join(n["value"] for n in paper["numbers"][:40])
    secs = "\n\n".join(
        f"## {s['heading']}\n{s['text'][:1200]}" for s in paper["sections"][:12])
    figs = "\n".join(
        f"- {f['caption'][:80] or 'figure'} -> {f['path']}" for f in paper["figures"][:6])
    repo_block = ""
    if repo:
        repo_block = (
            f"\n\nCODE REPOSITORY (describe what it does from THIS, do not invent):\n"
            f"name: {repo['name']}\nlanguages: {repo['languages']}\n"
            f"key files: {repo['key_files']}\nREADME (excerpt):\n{repo['readme'][:1500]}")
    return f"""Write a {n_slides}-slide conference-talk deck in **Marp markdown for marp-pptx**, following the authoring guide below EXACTLY (frontmatter `marp: true`, each slide `<!-- _class: type -->` + the type's exact HTML structure).

CRITICAL GROUNDING RULE: use ONLY facts, claims, and numbers from the SOURCE below. Never invent a number. Feature the listed KEY NUMBERS as KPIs where appropriate. If you reference a figure, use its extracted path.

=== SOURCE PAPER ===
TITLE: {paper['title']}
ABSTRACT: {paper['abstract'][:1200]}
KEY NUMBERS (use these exact values only): {nums}
FIGURES (path you may embed with ![w:760](path)):
{figs or '(none extracted)'}
SECTIONS:
{secs}{repo_block}

Output ONLY the Marp markdown (no commentary)."""


def _fix_prompt(markdown: str, fidelity: dict) -> str:
    uns = "\n".join(f"- slide {u['slide']}: '{u['value']}' NOT in paper ({u['context']})"
                    for u in fidelity["unsupported"])
    mis = "\n".join(f"- slide {m['slide']}: '{m['value']}' labelled '{m['label']}' but the"
                    f" paper does not pair that value with that metric ({m['context']})"
                    for m in fidelity.get("mislabeled", []))
    return f"""The deck below has quantitative claims that don't trace back to the source paper. Fix EACH against the paper (use the correct value/metric, or drop the claim if no source value exists). Change nothing else.

HALLUCINATED NUMBERS (not in the paper):
{uns or '(none)'}

MISLABELLED NUMBERS (value present but wrong metric/pairing):
{mis or '(none)'}

DECK:
{markdown}

Output ONLY the corrected Marp markdown."""


def _review_prompt(markdown: str, paper: dict) -> str:
    secs = "\n\n".join(f"## {s['heading']}\n{s['text'][:900]}" for s in paper["sections"][:12])
    return f"""Review this deck for FAITHFULNESS to the source paper. List any slide whose claim/method/result is NOT supported by the source sections (misstatements, wrong attributions, invented contributions). If everything is supported, reply exactly "OK".

SOURCE SECTIONS:
{secs}

DECK:
{markdown}

Reply with a short bullet list of unsupported claims, or "OK"."""


def _visual_fix_prompt(markdown: str, visual: list[dict]) -> str:
    items = "\n".join(f"- slide {v['slide']} ({v['class']}): {'; '.join(v['warnings'])}"
                      for v in visual)
    return f"""A render of this deck flagged these LAYOUT problems. Revise the markdown to fix them — for SPARSE slides add substance or merge with a neighbour or pick a more fitting type; for OVERFLOW split or trim; for SKEW rebalance. Keep the content faithful and the type HTML valid. Change only what's needed.

LAYOUT WARNINGS:
{items}

DECK:
{markdown}

Output ONLY the revised Marp markdown."""


def _defect_fix_prompt(markdown: str, defects: list[dict]) -> str:
    items = "\n".join(
        f"- slide {d['slide']} [{d['kind']}] {d['message']}" for d in defects)
    return f"""These slides were MEASURED against the real font metrics and do not fit. Revise the markdown so the content fits: cut words, split a slide in two, or move detail into the speaker notes (`<!-- note: ... -->`). Do not change any number or claim, and keep the type HTML valid. Change only the slides listed.

MEASURED DEFECTS:
{items}

DECK:
{markdown}

Output ONLY the revised Marp markdown."""


def build_deck_from_paper(
    paper_path: str,
    repo_path: str | None = None,
    *,
    palette: str = "claude",
    math: str = "omml",
    out: str | None = None,
    n_slides: int = 10,
    max_rounds: int = 3,
    semantic_review: bool = True,
    visual_polish: bool = True,
    llm=None,
    model: str = _DEFAULT_MODEL,
) -> dict:
    """Ingest a paper (+ repo), draft a grounded deck, auto-repair, and build it.

    Content grounding (numeric + semantic), a measured layout audit, and — when
    LibreOffice is present — a deterministic visual pass (sparse / skew) drive
    automatic LLM repairs. Returns {output_path, slide_count, markdown,
    fidelity, rounds, review, visual, defects, lint_warnings}.
    """
    from marp_pptx.mcp import build_pptx  # reuse the themed build

    call = llm or _anthropic_llm(model)
    paper = read_paper(paper_path, extract_figures=True)
    repo = read_repo(repo_path) if repo_path else None
    source = paper["full_text"]
    skill = _skill_context()

    markdown = _strip_fences(call(skill + "\n\n" + _draft_prompt(paper, repo, n_slides)))

    def _issues(f: dict) -> list:
        return f["unsupported"] + f.get("mislabeled", [])

    rounds = 0
    fidelity = check_fidelity(markdown, source)
    while _issues(fidelity) and rounds < max_rounds:
        markdown = _strip_fences(call(_fix_prompt(markdown, fidelity)))
        fidelity = check_fidelity(markdown, source)
        rounds += 1

    review = ""
    if semantic_review:
        review = call(_review_prompt(markdown, paper)).strip()
        if review and review.upper() != "OK" and rounds < max_rounds:
            markdown = _strip_fences(call(
                f"Fix these faithfulness issues against the paper, changing nothing else:\n{review}\n\nDECK:\n{markdown}\n\nOutput ONLY the corrected Marp markdown."))
            fidelity = check_fidelity(markdown, source)
            rounds += 1

    out = out or str(Path(paper_path).with_suffix("").name + "_deck.pptx")
    res = build_pptx(markdown, output_path=out, palette=palette, math=math)

    # Measured repair: the audit reads the built deck's geometry against real
    # font metrics, so it runs everywhere (no LibreOffice, no vision model) and
    # says exactly which slide holds more text than its boxes can show.
    defects = [d for d in res.get("defects", []) if d["severity"] == "error"]
    while defects and rounds < max_rounds:
        markdown = _strip_fences(call(_defect_fix_prompt(markdown, defects)))
        res = build_pptx(markdown, output_path=out, palette=palette, math=math)
        fidelity = check_fidelity(markdown, source)
        rounds += 1
        defects = [d for d in res.get("defects", []) if d["severity"] == "error"]

    # Visual polish: deterministic lint is the "eyes"; the text LLM fixes the
    # markdown from its warnings (no vision model needed). Skips if no soffice.
    visual: list = []
    if visual_polish:
        try:
            from marp_pptx.render import tools_available
            from marp_pptx.visuallint import lint_deck
            if tools_available():
                for _ in range(2):
                    visual = lint_deck(markdown, palette=palette)
                    if not visual or rounds >= max_rounds:
                        break
                    markdown = _strip_fences(call(_visual_fix_prompt(markdown, visual)))
                    res = build_pptx(markdown, output_path=out, palette=palette, math=math)
                    fidelity = check_fidelity(markdown, source)
                    rounds += 1
                visual = lint_deck(markdown, palette=palette)
        except Exception:  # noqa: BLE001 - polish is best-effort, never fatal
            visual = []

    return {
        "output_path": res["output_path"],
        "slide_count": res["slide_count"],
        "markdown": markdown,
        "fidelity": fidelity,
        "rounds": rounds,
        "review": review,
        "visual": visual,
        "defects": res.get("defects", []),
        "lint_warnings": res["lint_warnings"],
    }
