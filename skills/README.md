# Claude skill: marp-pptx

`marp-pptx/` is a [Claude Code skill](https://docs.claude.com/en/docs/claude-code/skills)
that teaches Claude to author **marp-pptx** decks: pick semantic slide types,
write correct Marp Markdown, and convert to an editable `.pptx`.

- `marp-pptx/SKILL.md` — the authoring guide (workflow, type-selection flow, rules, gotchas)
- `marp-pptx/references/type-skeletons.md` — exact HTML skeleton for all 52 types
  (auto-generated from the templates by `scripts/gen-skeletons.py`)

## Install (make it available in every project)

A project-local skill only triggers inside this repo. To make slides from anywhere,
symlink it into your personal skills dir:

```bash
ln -s "$(pwd)/skills/marp-pptx" ~/.claude/skills/marp-pptx
```

Then in any session, asking Claude to "marp-pptx でスライドを作って" (or to build a
deck / presentation) loads the skill. Requires `marp-pptx` installed
(`pip install marp-pptx`, or `marp-pptx[web]` for the browser UI).

## Keep the skeletons in sync

After changing slide types or templates:

```bash
PYTHONPATH=src python scripts/gen-skeletons.py
```
