"""Tests for the turnkey build_deck_from_paper orchestration (stub LLM)."""
import pytest

from marp_pptx.autodeck import build_deck_from_paper, _strip_fences

fitz = pytest.importorskip("fitz")  # needed only to synthesize a test PDF


@pytest.fixture
def paper_pdf(tmp_path):
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "Grounded Test Paper", fontsize=22)
    page.insert_text((72, 140), "Abstract\nWe report 28.4 BLEU on the benchmark.", fontsize=11)
    out = tmp_path / "paper.pdf"
    doc.save(str(out))
    doc.close()
    return str(out)


def test_strip_fences_handles_code_block():
    assert _strip_fences("```markdown\n---\nmarp: true\n---\n# Hi\n```").startswith("---")
    # adds frontmatter when missing
    assert _strip_fences("# just a heading").lstrip().startswith("---")


def test_loop_repairs_hallucinated_number(paper_pdf, tmp_path):
    calls = {"draft": 0, "fix": 0, "review": 0}

    def stub_llm(prompt: str) -> str:
        if "Write a" in prompt:  # initial draft — plant a hallucinated 99.9%
            calls["draft"] += 1
            return ("```markdown\n---\nmarp: true\n---\n\n"
                    "<!-- _class: kpi -->\n# Results\n"
                    '<div class="kpi-container">'
                    '<div><span class="kpi-value">99.9%</span><span class="kpi-label">acc</span></div>'
                    '<div><span class="kpi-value">28.4</span><span class="kpi-label">BLEU</span></div>'
                    "</div>\n```")
        if "UNSUPPORTED" in prompt:  # repair — drop the unsupported number
            calls["fix"] += 1
            return ("---\nmarp: true\n---\n\n<!-- _class: kpi -->\n# Results\n"
                    '<div class="kpi-container">'
                    '<div><span class="kpi-value">28.4</span><span class="kpi-label">BLEU</span></div>'
                    "</div>\n")
        calls["review"] += 1
        return "OK"

    out = tmp_path / "deck.pptx"
    res = build_deck_from_paper(paper_pdf, out=str(out), llm=stub_llm, n_slides=3)

    assert calls["draft"] == 1
    assert calls["fix"] == 1          # the loop repaired exactly once
    assert res["rounds"] >= 1
    assert res["fidelity"]["score"] == 100   # clean after repair
    assert res["fidelity"]["unsupported"] == []
    assert "99.9%" not in res["markdown"]     # hallucination removed
    assert "28.4" in res["markdown"]          # grounded number kept
    assert out.is_file() and res["slide_count"] >= 1
