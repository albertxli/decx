"""Unit tests for decx.utils — no COM required."""

from decx.utils import (
    hex_to_rgb,
    convert_r1c1_to_a1,
    extract_link_parts,
    get_contrast_font_color,
)


class TestHexToRgb:
    def test_basic_red(self):
        # Red: R=255, G=0, B=0 -> 255 + 0 + 0 = 255
        assert hex_to_rgb("#FF0000") == 255

    def test_basic_green(self):
        # Green: R=0, G=255, B=0 -> 0 + 255*256 + 0 = 65280
        assert hex_to_rgb("#00FF00") == 65280

    def test_basic_blue(self):
        # Blue: R=0, G=0, B=255 -> 0 + 0 + 255*65536 = 16711680
        assert hex_to_rgb("#0000FF") == 16711680

    def test_white(self):
        assert hex_to_rgb("#FFFFFF") == 255 + 255 * 256 + 255 * 65536

    def test_black(self):
        assert hex_to_rgb("#000000") == 0

    def test_no_hash(self):
        assert hex_to_rgb("FF0000") == 255

    def test_invalid_length(self):
        assert hex_to_rgb("#FFF") == 0  # fallback to black

    def test_heatmap_red(self):
        # #F8696B -> R=248, G=105, B=107
        expected = 248 + 105 * 256 + 107 * 65536
        assert hex_to_rgb("#F8696B") == expected

    def test_ccst_green(self):
        # #33CC33 -> R=51, G=204, B=51
        expected = 51 + 204 * 256 + 51 * 65536
        assert hex_to_rgb("#33CC33") == expected


class TestConvertR1C1ToA1:
    def test_single_cell(self):
        assert convert_r1c1_to_a1("R1C1") == "A1"

    def test_single_cell_multi_digit(self):
        assert convert_r1c1_to_a1("R10C3") == "C10"

    def test_range(self):
        assert convert_r1c1_to_a1("R1C1:R5C5") == "A1:E5"

    def test_multi_letter_column(self):
        # Column 27 = AA
        assert convert_r1c1_to_a1("R1C27") == "AA1"

    def test_column_26(self):
        assert convert_r1c1_to_a1("R1C26") == "Z1"

    def test_column_28(self):
        assert convert_r1c1_to_a1("R1C28") == "AB1"

    def test_large_range(self):
        assert convert_r1c1_to_a1("R1C1:R100C52") == "A1:AZ100"


class TestExtractLinkParts:
    def test_full_link(self):
        link = r"C:\data\file.xlsx!Sheet1!R1C1:R5C5"
        fp, sn, ra = extract_link_parts(link)
        assert fp == r"C:\data\file.xlsx"
        assert sn == "Sheet1"
        assert ra == "A1:E5"

    def test_missing_range(self):
        link = r"C:\data\file.xlsx!Sheet1"
        fp, sn, ra = extract_link_parts(link)
        assert fp == r"C:\data\file.xlsx"
        assert sn == "Sheet1"
        assert ra == "Not Specified"

    def test_file_only(self):
        link = r"C:\data\file.xlsx"
        fp, sn, ra = extract_link_parts(link)
        assert fp == r"C:\data\file.xlsx"
        assert sn == "Not Specified"
        assert ra == "Not Specified"

    def test_sheet_with_spaces(self):
        link = r"C:\path\data.xlsx!Market Data!R1C1:R10C5"
        fp, sn, ra = extract_link_parts(link)
        assert fp == r"C:\path\data.xlsx"
        assert sn == "Market Data"
        assert ra == "A1:E10"


class TestGetContrastFontColor:
    def test_dark_background_returns_light(self):
        black = 0  # R=0, G=0, B=0
        dark_font = 100
        light_font = 200
        assert get_contrast_font_color(black, dark_font, light_font) == light_font

    def test_light_background_returns_dark(self):
        white = 255 + 255 * 256 + 255 * 65536
        dark_font = 100
        light_font = 200
        assert get_contrast_font_color(white, dark_font, light_font) == dark_font

    def test_mid_brightness(self):
        # R=128, G=128, B=128 -> brightness ~ 128 (floating point: 127.999...)
        # Due to floating point, this is just under 128 threshold -> light font
        # Matches VBA behavior (same floating point)
        mid = 128 + 128 * 256 + 128 * 65536
        assert get_contrast_font_color(mid, 0, 0xFFFFFF) == 0xFFFFFF  # light font

    def test_above_threshold(self):
        # R=130, G=130, B=130 -> brightness ~ 130 > 128 -> dark font
        val = 130 + 130 * 256 + 130 * 65536
        assert get_contrast_font_color(val, 0, 0xFFFFFF) == 0  # dark font

    def test_red_is_relatively_dark(self):
        # Pure red: R=255 -> brightness = 0.299*255 = 76.2 < 128 -> light font
        red = 255
        assert get_contrast_font_color(red, 0, 0xFFFFFF) == 0xFFFFFF  # light font

    def test_green_is_bright(self):
        # Pure green: G=255 -> brightness = 0.587*255 = 149.7 > 128 -> dark font
        green = 255 * 256
        assert get_contrast_font_color(green, 0, 0xFFFFFF) == 0  # dark font
