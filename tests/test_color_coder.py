"""Unit tests for color_coder logic — no COM required.

Tests the internal helper functions that don't require COM objects.
"""

from decx.color_coder import _is_numeric, _validate_symbol_removal


class TestIsNumeric:
    def test_positive_integer(self):
        assert _is_numeric("5") == (True, 5.0)

    def test_negative_integer(self):
        assert _is_numeric("-3") == (True, -3.0)

    def test_positive_float(self):
        assert _is_numeric("5.2") == (True, 5.2)

    def test_negative_float(self):
        assert _is_numeric("-3.7") == (True, -3.7)

    def test_zero(self):
        assert _is_numeric("0") == (True, 0.0)

    def test_non_numeric_text(self):
        is_num, _ = _is_numeric("abc")
        assert is_num is False

    def test_empty_string(self):
        is_num, _ = _is_numeric("")
        assert is_num is False

    def test_plus_prefix(self):
        assert _is_numeric("+5.2") == (True, 5.2)

    def test_none(self):
        is_num, _ = _is_numeric(None)
        assert is_num is False


class TestValidateSymbolRemoval:
    def test_percent_only(self):
        assert _validate_symbol_removal("%") is True

    def test_plus_only(self):
        assert _validate_symbol_removal("+") is True

    def test_minus_only(self):
        assert _validate_symbol_removal("-") is True

    def test_all_three(self):
        assert _validate_symbol_removal("%+-") is True

    def test_empty(self):
        assert _validate_symbol_removal("") is True

    def test_invalid_char(self):
        assert _validate_symbol_removal("x") is False

    def test_mixed_valid_invalid(self):
        assert _validate_symbol_removal("%a") is False
