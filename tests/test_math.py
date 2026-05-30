"""Tests for math rendering."""
import shutil
from pathlib import Path

from marp_pptx.math.renderer import _sanitize_for_mathtext, render_latex_png


def test_sanitize_strips_colorbox():
    out = _sanitize_for_mathtext(r"\frac{\colorbox{#fff3cd}{$QK^\top$}}{\colorbox{#cce}{$\sqrt{d_k}$}}")
    assert "colorbox" not in out
    assert "QK^\\top" in out
    assert "\\sqrt{d_k}" in out  # nested braces preserved


def test_sanitize_strips_brace_annotations():
    out = _sanitize_for_mathtext(r"\hat{y} = \underbrace{\sigma}_{\text{活性化}}(\overbrace{W}^{\text{重み}} x)")
    assert "underbrace" not in out and "overbrace" not in out
    assert "活性化" not in out and "重み" not in out  # CJK annotation dropped
    assert "\\sigma" in out and "W" in out  # base terms kept


def test_render_advanced_equation_after_sanitize():
    # the equation-highlight template — must render once colorbox is stripped
    tex = r"\text{softmax}\!\left(\frac{\colorbox{#fff3cd}{$QK^\top$}}{\sqrt{d_k}}\right)V"
    assert render_latex_png(tex, display=True) is not None


def test_render_simple_latex():
    result = render_latex_png(r"x^2 + y^2 = z^2", fontsize=20)
    assert result is not None
    assert Path(result).exists()
    assert Path(result).stat().st_size > 0


def test_render_fraction():
    result = render_latex_png(r"\frac{a}{b}", fontsize=24, display=True)
    assert result is not None


def test_render_cache():
    # Same expression should return cached path
    r1 = render_latex_png(r"e^{i\pi} + 1 = 0", fontsize=20)
    r2 = render_latex_png(r"e^{i\pi} + 1 = 0", fontsize=20)
    assert r1 == r2


def test_omml_with_pandoc():
    if not shutil.which("pandoc"):
        return  # skip if pandoc not installed
    from marp_pptx.math.omml import latex_to_omml_element
    el = latex_to_omml_element(r"\frac{a}{b}", display=True)
    assert el is not None
