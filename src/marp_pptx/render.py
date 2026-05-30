"""Shared PPTX → PNG rendering via LibreOffice + pdftoppm.

Used by the web live-preview and the `render-gallery` CLI command. Returns the
sorted list of per-slide PNG paths, or [] if the tools are missing or the
conversion fails.
"""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
PDFTOPPM = shutil.which("pdftoppm")


def tools_available() -> bool:
    return SOFFICE is not None and PDFTOPPM is not None


def pptx_to_pngs(pptx_path: Path, out_dir: Path, dpi: int = 100) -> list[Path]:
    """Render a PPTX to one PNG per slide (``slide-1.png`` …).

    Returns the sorted PNG paths, or [] if tools are unavailable or rendering
    fails.
    """
    if not tools_available():
        return []
    out_dir.mkdir(parents=True, exist_ok=True)
    if not pptx_batch_to_pdf([pptx_path], out_dir):
        return []
    pdf = out_dir / (pptx_path.stem + ".pdf")
    if not pdf.exists():
        return []
    return pdf_to_pngs(pdf, out_dir, dpi=dpi, prefix="slide")


def pptx_batch_to_pdf(pptx_paths: list[Path], out_dir: Path) -> list[Path]:
    """Convert several PPTX files to PDF in a single LibreOffice invocation.

    One soffice startup is amortized across all files, and each PDF stays
    isolated to its own deck (so a content overflow can't shift a neighbour's
    page mapping). Returns the expected PDF paths, or [] on failure.
    """
    if SOFFICE is None or not pptx_paths:
        return []
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        subprocess.run(
            [SOFFICE, "--headless", "--convert-to", "pdf",
             "--outdir", str(out_dir), *(str(p) for p in pptx_paths)],
            check=True, capture_output=True, timeout=300,
        )
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return []
    return [out_dir / (p.stem + ".pdf") for p in pptx_paths]


def pdf_to_pngs(pdf: Path, out_dir: Path, dpi: int = 100, prefix: str = "slide") -> list[Path]:
    """Render every page of a PDF to ``{prefix}-N.png`` via pdftoppm."""
    if PDFTOPPM is None or not pdf.exists():
        return []
    try:
        subprocess.run(
            [PDFTOPPM, "-png", "-r", str(dpi), str(pdf), str(out_dir / prefix)],
            check=True, capture_output=True, timeout=120,
        )
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return []
    return sorted(out_dir.glob(f"{prefix}-*.png"))
