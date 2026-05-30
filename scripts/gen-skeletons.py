#!/usr/bin/env python3
"""Regenerate skills/marp-pptx/references/type-skeletons.md from the templates.

Run:  PYTHONPATH=src python scripts/gen-skeletons.py
Keeps the skill's HTML skeletons in sync with src/marp_pptx/data/templates/.
"""
from pathlib import Path

from marp_pptx.types import TYPE_REGISTRY, CATEGORIES

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = ROOT / "src" / "marp_pptx" / "data" / "templates"
OUT = ROOT / "skills" / "marp-pptx" / "references" / "type-skeletons.md"


def _body(path: Path) -> str:
    text = path.read_text(encoding="utf-8")
    if text.startswith("---"):
        end = text.find("---", 3)
        if end != -1:
            text = text[end + 3:]
    return text.strip()


def main() -> None:
    lines = [
        "# 型ごとの HTML 骨組みリファレンス",
        "",
        "marp-pptx の全 52 型の正確な構造。`<!-- _class: -->` とその下の HTML を**崩さずに**埋める。",
        "（このファイルは `scripts/gen-skeletons.py` が `src/marp_pptx/data/templates/` から自動生成）",
        "",
    ]
    by_cat: dict[str, list] = {}
    for t in TYPE_REGISTRY:
        by_cat.setdefault(t.category, []).append(t)

    for cat_key, cat_label in CATEGORIES.items():
        types = by_cat.get(cat_key, [])
        if not types:
            continue
        lines.append(f"## {cat_label}（{cat_key}）\n")
        for t in types:
            lines += [
                f"### `{t.name}` — {t.meaning}",
                f"<!-- {t.geometry} · {t.use_when} -->",
                "",
                "```markdown",
                _body(TEMPLATES / t.template_file),
                "```",
                "",
            ]

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {OUT} ({len(TYPE_REGISTRY)} types)")


if __name__ == "__main__":
    main()
