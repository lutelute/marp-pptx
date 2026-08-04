"""Real font metrics: widths, wrapping, 禁則処理, shrink-to-fit.

These are the numbers every layout decision downstream is made from, so the
assertions are about *properties* that must hold on any machine — a box never
gets a line wider than itself, punctuation never opens a line — rather than
exact point values, which depend on which fonts are installed.
"""
import pytest

from marp_pptx import metrics as M

LATIN = "Helvetica Neue"
EA = "Hiragino Sans"


def test_narrow_and_wide_latin_are_not_the_same_width():
    """The old heuristic charged 0.55 em for every Latin glyph, so "llll" and
    "WWWW" measured identically — the root cause of both spurious wraps and
    missed overflow."""
    narrow = M.measure_em("llll", LATIN)
    wide = M.measure_em("WWWW", LATIN)
    assert wide > narrow * 2.5


def test_full_width_cjk_is_one_em_each():
    w = M.measure_em("あいうえお", LATIN, ea_font=EA)
    assert 4.5 <= w <= 5.5


def test_missing_font_falls_back_to_the_width_table():
    """A font absent from this machine still measures per glyph, not per class."""
    narrow = M.measure_em("iiii", "No Such Font ZZZ")
    wide = M.measure_em("WWWW", "No Such Font ZZZ")
    assert 0 < narrow < wide


@pytest.mark.parametrize("text,width", [
    ("The quick brown fox jumps over the lazy dog near the riverbank", 200.0),
    ("近接勾配法の収束解析について、仮定と記号を整理する。ここでは強凸性を仮定する。", 300.0),
    ("mixed 日本語 and English text that has to wrap somewhere sensible", 180.0),
])
def test_wrapped_lines_fit_the_box(text, width):
    lines = M.wrap_text(text, LATIN, 18, width, ea_font=EA)
    for line in lines:
        # 6% tolerance: a single trailing punctuation mark is allowed to hang
        # (ぶら下げ), which is what PowerPoint does too.
        assert M.measure_pt(line, LATIN, 18, ea_font=EA) <= width * 1.07


def test_wrapping_preserves_every_character():
    text = "近接勾配法の収束解析について、仮定と記号を整理する。"
    assert "".join(M.wrap_text(text, LATIN, 18, 120, ea_font=EA)) == text


def test_kinsoku_never_opens_a_line_with_closing_punctuation():
    text = "これは折り返しの試験です。句読点が行頭に来てはいけません。" * 3
    lines = M.wrap_text(text, LATIN, 18, 140, ea_font=EA)
    assert len(lines) > 3
    for line in lines[1:]:
        assert line[0] not in "、。」）"


def test_latin_wrapping_breaks_at_word_boundaries():
    lines = M.wrap_text("alpha beta gamma delta epsilon", LATIN, 18, 120)
    assert len(lines) > 1
    for line in lines:
        assert not line.startswith(" ")
        for word in line.split():
            assert word in "alpha beta gamma delta epsilon".split()


def test_explicit_newlines_are_kept():
    assert M.wrap_text("one\ntwo", LATIN, 18, 500) == ["one", "two"]


def test_fit_size_shrinks_only_when_needed():
    box_w, box_h = 300.0, 60.0
    short = M.fit_size("Short", LATIN, box_w, box_h, max_size=32, ea_font=EA)
    assert short == 32
    long_ja = "とても長いタイトルがここに入り二行以上になります"
    shrunk = M.fit_size(long_ja, LATIN, box_w, box_h, max_size=32, ea_font=EA)
    assert shrunk < 32
    lines = M.wrap_text(long_ja, LATIN, shrunk, box_w, ea_font=EA)
    assert M.block_height_pt(lines, shrunk) <= box_h


def test_fit_size_stops_at_the_floor_instead_of_going_unreadable():
    huge = "字" * 400
    size = M.fit_size(huge, LATIN, 200.0, 40.0, max_size=24, min_size=12, ea_font=EA)
    assert size == 12


def test_overflow_pt_is_zero_when_the_text_fits():
    assert M.overflow_pt("Short", LATIN, 14, 300.0, 100.0) == 0.0
    assert M.overflow_pt("字" * 200, LATIN, 14, 300.0, 40.0, ea_font=EA) > 0


def test_font_status_flags_the_previewer_traps():
    assert M.font_status("Aptos") == "unreliable"
    assert M.font_status("No Such Font ZZZ") == "missing"


def test_line_height_is_a_sane_multiple():
    assert 1.0 <= M.line_height_em(LATIN) <= 1.7
    assert M.DEFAULT_LINE_FACTOR == 1.25
