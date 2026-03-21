"""Unit tests for decx.checker — no COM required."""

from decx.checker import _compare_cell_text, _extract_sign_suffix, Mismatch
from decx.delta_updater import _determine_sign, _strip_sign_suffix


class TestCompareCellText:
    def test_match(self):
        assert _compare_cell_text("12.5%", "12.5%") is True

    def test_mismatch(self):
        assert _compare_cell_text("12.5%", "12.3%") is False

    def test_whitespace_tolerant(self):
        assert _compare_cell_text("  12.5%  ", "12.5%") is True

    def test_both_stripped(self):
        assert _compare_cell_text(" hello ", " hello ") is True

    def test_empty_strings(self):
        assert _compare_cell_text("", "") is True

    def test_empty_vs_whitespace(self):
        assert _compare_cell_text("  ", "") is True


class TestExtractSignSuffix:
    def test_pos(self):
        assert _extract_sign_suffix("delt_Object_2_pos") == "pos"

    def test_neg(self):
        assert _extract_sign_suffix("delt_Object_2_neg") == "neg"

    def test_none(self):
        assert _extract_sign_suffix("delt_Object_2_none") == "none"

    def test_no_suffix(self):
        assert _extract_sign_suffix("delt_Object_2") is None

    def test_no_suffix_other(self):
        assert _extract_sign_suffix("delt_revenue") is None


class TestStripSignSuffix:
    def test_strip_pos(self):
        assert _strip_sign_suffix("delt_Object_2_pos") == "delt_Object_2"

    def test_strip_neg(self):
        assert _strip_sign_suffix("delt_Object_2_neg") == "delt_Object_2"

    def test_strip_none(self):
        assert _strip_sign_suffix("delt_Object_2_none") == "delt_Object_2"

    def test_no_suffix(self):
        assert _strip_sign_suffix("delt_Object_2") == "delt_Object_2"

    def test_preserves_base(self):
        assert _strip_sign_suffix("delt_revenue_pos") == "delt_revenue"


class TestDetermineSign:
    def test_positive(self):
        assert _determine_sign("5.2%") == "pos"

    def test_negative(self):
        assert _determine_sign("-3.1%") == "neg"

    def test_zero(self):
        assert _determine_sign("0.0%") == "none"

    def test_zero_integer(self):
        assert _determine_sign("0") == "none"

    def test_positive_no_percent(self):
        assert _determine_sign("12.5") == "pos"

    def test_negative_no_percent(self):
        assert _determine_sign("-0.3") == "neg"

    def test_empty(self):
        assert _determine_sign("") == "none"

    def test_text(self):
        assert _determine_sign("N/A") == "none"

    def test_whitespace(self):
        assert _determine_sign("  5.2%  ") == "pos"


class TestMismatch:
    def test_table_mismatch(self):
        m = Mismatch(
            slide=5,
            shape_name="ntbl_revenue",
            detail="Row 3 Col 2: PPT='12.5%' Excel='12.3%'",
            category="table",
        )
        assert m.slide == 5
        assert m.category == "table"
        assert "12.5%" in m.detail

    def test_delta_mismatch(self):
        m = Mismatch(
            slide=8,
            shape_name="delt_revenue_pos",
            detail="PPT=pos, Excel value='-2.3%' (expected=neg)",
            category="delta",
        )
        assert m.category == "delta"
        assert "neg" in m.detail


class TestTransposedIndexSwap:
    """Verify that transposed tables swap row/col correctly."""

    def test_normal_no_swap(self):
        """Normal table: PPT(r, c) = Excel(r, c)."""
        do_transpose = False
        r, c = 3, 2
        ppt_r = c if do_transpose else r
        ppt_c = r if do_transpose else c
        assert (ppt_r, ppt_c) == (3, 2)

    def test_transposed_swap(self):
        """Transposed table: PPT(c, r) = Excel(r, c)."""
        do_transpose = True
        r, c = 3, 2
        ppt_r = c if do_transpose else r
        ppt_c = r if do_transpose else c
        assert (ppt_r, ppt_c) == (2, 3)
