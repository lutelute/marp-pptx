"""OOXML package integrity — will PowerPoint actually open this file?

A .pptx is a ZIP of XML parts wired together by relationship files.  Most tools
in the chain are forgiving: python-pptx will happily read a deck whose image
relationship points at nothing, and LibreOffice renders it.  PowerPoint is not
forgiving — it shows "PowerPoint found a problem with content" and repairs the
deck by throwing the broken part away.

These checks are the package-level half of the audit: structure, not layout.
They are cheap (a few ms), need nothing installed, and every finding names the
part at fault.

    findings = check_package("deck.pptx")     # list[audit.Finding]
"""

from __future__ import annotations

import posixpath
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .audit import Finding

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# Leading bytes → the extensions PowerPoint accepts for that format.  A file
# named .png that is really an SVG is the classic "repaired deck" cause.
_MAGIC = [
    (b"\x89PNG\r\n\x1a\n", {"png"}),
    (b"\xff\xd8\xff", {"jpg", "jpeg"}),
    (b"GIF87a", {"gif"}),
    (b"GIF89a", {"gif"}),
    (b"<?xml", {"svg", "xml", "rels", "emf", "wmf"}),
    (b"<svg", {"svg"}),
    (b"%PDF", {"pdf"}),
]


def check_package(path: str | Path) -> list[Finding]:
    p = Path(path)
    out: list[Finding] = []
    try:
        zf = zipfile.ZipFile(p)
    except zipfile.BadZipFile as exc:
        return [Finding("package", "error", 0, f"not a readable .pptx: {exc}",
                        fix="rebuild the deck")]
    with zf:
        bad = zf.testzip()
        if bad:
            out.append(Finding("package", "error", 0,
                               f"corrupt entry in the archive: {bad}",
                               fix="rebuild the deck"))
        names = set(zf.namelist())
        out += _check_content_types(zf, names)
        out += _check_rels(zf, names)
        out += _check_media(zf, names)
        out += _check_presentation(zf, names)
    return out


def _parse(zf, name):
    try:
        return ET.fromstring(zf.read(name))
    except ET.ParseError as exc:
        raise ValueError(f"{name}: {exc}") from None


def _check_content_types(zf, names) -> list[Finding]:
    out: list[Finding] = []
    if "[Content_Types].xml" not in names:
        return [Finding("package", "error", 0, "[Content_Types].xml is missing",
                        fix="rebuild the deck — the package has no part map")]
    try:
        root = _parse(zf, "[Content_Types].xml")
    except ValueError as exc:
        return [Finding("package", "error", 0, str(exc), fix="rebuild the deck")]
    defaults = {e.get("Extension", "").lower()
                for e in root.findall(f"{{{_CT_NS}}}Default")}
    overrides = {e.get("PartName", "").lstrip("/")
                 for e in root.findall(f"{{{_CT_NS}}}Override")}
    for name in sorted(names):
        if name.endswith("/") or name == "[Content_Types].xml":
            continue
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if ext in defaults or name in overrides:
            continue
        out.append(Finding(
            "package", "error", 0,
            f"{name} has no content type — PowerPoint will refuse the file",
            fix="add a Default for its extension or an Override for the part"))
    return out


def _rels_for(part: str) -> str:
    d, base = posixpath.split(part)
    return posixpath.join(d, "_rels", base + ".rels")


def _check_rels(zf, names) -> list[Finding]:
    """Every relationship target resolves, and every r:id used in a part's XML
    is declared in that part's .rels."""
    out: list[Finding] = []
    rels_files = [n for n in names if n.endswith(".rels")]
    declared: dict[str, set[str]] = {}
    for rf in sorted(rels_files):
        try:
            root = _parse(zf, rf)
        except ValueError as exc:
            out.append(Finding("package", "error", 0, str(exc),
                               fix="rebuild the deck"))
            continue
        owner_dir = posixpath.dirname(posixpath.dirname(rf))
        ids: set[str] = set()
        for rel in root.findall(f"{{{_REL_NS}}}Relationship"):
            rid = rel.get("Id", "")
            if rid in ids:
                out.append(Finding("package", "error", 0,
                                   f"{rf} declares {rid} twice",
                                   fix="give each relationship a unique Id"))
            ids.add(rid)
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "")
            resolved = posixpath.normpath(posixpath.join(owner_dir, target))
            if resolved.lstrip("/") not in names:
                out.append(Finding(
                    "package", "error", 0,
                    f"{rf}: {rid} points at “{target}”, which is not in the package",
                    fix="drop the relationship or add the missing part"))
        declared[rf] = ids

    for name in sorted(names):
        if not name.endswith(".xml") or name.endswith(".rels"):
            continue
        if not (name.startswith("ppt/") or name == "docProps/app.xml"):
            continue
        try:
            data = zf.read(name).decode("utf-8", "replace")
        except Exception:
            continue
        used = set()
        for attr in ("r:id", "r:embed", "r:link", "r:pict", "r:dm", "r:lo", "r:qs", "r:cs"):
            marker = f'{attr}="'
            start = 0
            while True:
                i = data.find(marker, start)
                if i < 0:
                    break
                j = data.find('"', i + len(marker))
                used.add(data[i + len(marker):j])
                start = j
        if not used:
            continue
        have = declared.get(_rels_for(name), set())
        for rid in sorted(used - have):
            out.append(Finding(
                "package", "error", 0,
                f"{name} references {rid}, which its .rels does not declare",
                fix="add the relationship, or remove the reference"))
    return out


def _check_media(zf, names) -> list[Finding]:
    """Media parts: referenced by something, and actually the format their
    extension claims."""
    out: list[Finding] = []
    referenced: set[str] = set()
    for rf in [n for n in names if n.endswith(".rels")]:
        try:
            root = _parse(zf, rf)
        except ValueError:
            continue
        owner_dir = posixpath.dirname(posixpath.dirname(rf))
        for rel in root.findall(f"{{{_REL_NS}}}Relationship"):
            if rel.get("TargetMode") == "External":
                continue
            referenced.add(posixpath.normpath(
                posixpath.join(owner_dir, rel.get("Target", ""))).lstrip("/"))
    for name in sorted(n for n in names if n.startswith("ppt/media/")):
        if name not in referenced:
            out.append(Finding("package", "info", 0,
                               f"{name} is embedded but never used ("
                               f"{zf.getinfo(name).file_size // 1024} KB of dead weight)",
                               fix="drop the unused media to keep the file small"))
        ext = name.rsplit(".", 1)[-1].lower()
        head = zf.read(name)[:8]
        for magic, exts in _MAGIC:
            if head.startswith(magic):
                if ext not in exts:
                    out.append(Finding(
                        "package", "error", 0,
                        f"{name} is a {sorted(exts)[0].upper()} file with a .{ext} "
                        f"extension — PowerPoint will report the deck as damaged",
                        fix="convert the image, or give it its true extension"))
                break
    return out


def _check_presentation(zf, names) -> list[Finding]:
    out: list[Finding] = []
    pres = "ppt/presentation.xml"
    if pres not in names:
        return [Finding("package", "error", 0, "ppt/presentation.xml is missing",
                        fix="rebuild the deck")]
    try:
        root = _parse(zf, pres)
    except ValueError as exc:
        return [Finding("package", "error", 0, str(exc), fix="rebuild the deck")]
    try:
        rels = _parse(zf, _rels_for(pres))
    except (KeyError, ValueError):
        return [Finding("package", "error", 0,
                        "ppt/_rels/presentation.xml.rels is missing or unreadable",
                        fix="rebuild the deck")]
    targets = {r.get("Id"): r.get("Target", "") for r in
               rels.findall(f"{{{_REL_NS}}}Relationship")}
    p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    lst = root.find(f"{{{p_ns}}}sldIdLst")
    ids = [] if lst is None else [e.get(f"{{{_R_NS}}}id") for e in lst]
    if not ids:
        out.append(Finding("package", "warn", 0, "the deck has no slides",
                           fix="check that the markdown parsed into slides"))
    for rid in ids:
        target = targets.get(rid)
        if not target:
            out.append(Finding("package", "error", 0,
                               f"slide list references {rid}, undeclared in the rels",
                               fix="rebuild the deck"))
            continue
        resolved = posixpath.normpath(posixpath.join("ppt", target))
        if resolved not in names:
            out.append(Finding("package", "error", 0,
                               f"slide list points at missing part {resolved}",
                               fix="rebuild the deck"))
    # Slides present in the package but absent from the ordering never render.
    listed = {posixpath.normpath(posixpath.join("ppt", targets[i]))
              for i in ids if targets.get(i)}
    for name in sorted(n for n in names if n.startswith("ppt/slides/slide")):
        if name.endswith(".xml") and name not in listed:
            out.append(Finding("package", "warn", 0,
                               f"{name} exists but is not in the slide order — "
                               f"it will not appear in the deck",
                               fix="add it to <p:sldIdLst> or delete the part"))
    return out
