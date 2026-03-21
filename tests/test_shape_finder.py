"""Unit tests for decx.shape_finder — no COM required."""

from decx.shape_finder import is_exact_token_match


class TestIsExactTokenMatch:
    def test_basic_match(self):
        assert is_exact_token_match("ntbl_Object_5", "Object_5") is True

    def test_no_match_substring(self):
        # Object_5 should NOT match Object_55 (5 followed by digit)
        assert is_exact_token_match("ntbl_Object_55", "Object_5") is False

    def test_match_at_start(self):
        assert is_exact_token_match("Object_5_ntbl", "Object_5") is True

    def test_match_at_end(self):
        assert is_exact_token_match("ntbl_Object_5", "Object_5") is True

    def test_match_exact(self):
        assert is_exact_token_match("Object_5", "Object_5") is True

    def test_no_match_not_found(self):
        assert is_exact_token_match("ntbl_SomethingElse", "Object_5") is False

    def test_underscore_boundary(self):
        assert is_exact_token_match("delt_NetSent_ccst", "NetSent") is True

    def test_space_boundary(self):
        assert is_exact_token_match("ntbl Object 5", "Object 5") is True

    def test_hyphen_boundary(self):
        assert is_exact_token_match("ntbl-Object_5", "Object_5") is True

    def test_digit_not_boundary(self):
        # Digit is alphanumeric, not a boundary
        assert is_exact_token_match("ntbl_2Object_5", "Object_5") is False

    def test_letter_not_boundary(self):
        assert is_exact_token_match("ntbl_XObject_5", "Object_5") is False

    def test_multiple_occurrences_second_matches(self):
        # First "ab" at position 0 is followed by "c" (alphanumeric, not boundary)
        # But if we have "x_ab_abc_ab" and search for "ab":
        # occurrence at idx=2 has _ before and _ after -> match
        assert is_exact_token_match("x_ab_abc_ab", "ab") is True

    def test_ccst_suffix(self):
        assert is_exact_token_match("ntbl_Object5_ccst", "Object5") is True

    def test_linked_name_with_spaces(self):
        # Real PowerPoint shape names can have spaces: "Object 5"
        assert is_exact_token_match("ntbl_Object 5", "Object 5") is True

    def test_empty_linked_name(self):
        # Empty string matches everywhere in Python's str.find, but
        # boundary checks on "" are degenerate — just verify no crash
        # VBA would not call this with empty linkedName in practice
        is_exact_token_match("ntbl_something", "")  # no crash

    def test_linked_longer_than_shape(self):
        assert is_exact_token_match("ab", "abcdef") is False
