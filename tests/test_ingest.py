"""Tests for the ingestion + grounding layer."""
from pathlib import Path

import pytest

from marp_pptx.ingest import read_repo, check_fidelity

fitz = pytest.importorskip("fitz")  # PyMuPDF; skip paper tests if absent

ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture
def sample_pdf(tmp_path):
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "A Tiny Test Paper", fontsize=22)
    page.insert_text(
        (72, 140),
        "Abstract\nWe report 28.4 BLEU and 91.2% accuracy with a 3.5x speedup.",
        fontsize=11,
    )
    out = tmp_path / "tiny.pdf"
    doc.save(str(out))
    doc.close()
    return str(out)


def test_read_paper_extracts_structure_and_numbers(sample_pdf):
    from marp_pptx.ingest import read_paper
    r = read_paper(sample_pdf, extract_figures=False)
    assert r["title"] == "A Tiny Test Paper"
    assert r["page_count"] == 1
    vals = {n["value"] for n in r["numbers"]}
    assert "28.4" in vals
    assert any(v.startswith("91.2") for v in vals)
    assert "BLEU" in r["full_text"]


def test_read_repo_summarizes_this_project():
    r = read_repo(str(ROOT))
    assert r["name"] == "marp-pptx"
    assert "Python" in r["languages"]
    assert "pyproject.toml" in r["key_files"]
    assert r["file_count"] > 10


def test_check_fidelity_flags_unsupported_numbers():
    source = "The model reached 28.4 BLEU and 41.8 BLEU on the test sets."
    deck = (
        "---\nmarp: true\n---\n\n# Results\n"
        "<span>28.4</span> BLEU\n<span>99.9%</span> accuracy\n<span>500x</span> faster\n"
    )
    r = check_fidelity(deck, source)
    unsupported = {u["value"] for u in r["unsupported"]}
    assert "99.9%" in unsupported
    assert "500x" in unsupported
    assert r["score"] < 100
    # the real number is not flagged
    assert all(u["value"] != "28.4" for u in r["unsupported"])


def test_check_fidelity_clean_deck_scores_100():
    source = "BLEU was 28.4 and 41.8."
    deck = "---\nmarp: true\n---\n\n# R\n<span>28.4</span>\n<span>41.8</span>\n"
    assert check_fidelity(deck, source)["score"] == 100
