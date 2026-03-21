"""Shared utility functions: hex conversion, R1C1->A1, link parsing, contrast color."""


def hex_to_rgb(hex_color: str) -> int:
    """Convert '#RRGGBB' hex string to a win32com-compatible RGB Long value.

    VBA's RGB() returns B * 65536 + G * 256 + R (BGR order as a Long).
    """
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return 0  # black fallback
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return r + (g * 256) + (b * 65536)


def _convert_single_r1c1(cell: str) -> str:
    """Convert a single R1C1 cell reference like 'R5C3' to 'C5'."""
    cell = cell.strip()
    c_pos = cell.upper().index("C", 1)  # skip first char which is 'R'
    row_num = int(cell[1:c_pos])
    col_num = int(cell[c_pos + 1:])

    col_letter = ""
    n = col_num
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        col_letter = chr(65 + remainder) + col_letter

    return f"{col_letter}{row_num}"


def convert_r1c1_to_a1(r1c1_range: str) -> str:
    """Convert an R1C1 range string to A1 notation.

    Examples:
        'R1C1' -> 'A1'
        'R1C1:R5C5' -> 'A1:E5'
        'R1C27' -> 'AA1'
    """
    parts = r1c1_range.split(":")
    if len(parts) == 1:
        return _convert_single_r1c1(parts[0])
    return f"{_convert_single_r1c1(parts[0])}:{_convert_single_r1c1(parts[1])}"


def extract_link_parts(source_full_name: str) -> tuple[str, str, str]:
    """Parse an OLE SourceFullName into (file_path, sheet_name, range_address).

    The format is: 'filepath!sheetname!R1C1:R5C5'
    Returns ('Not Specified', ...) for missing parts.
    The range is converted from R1C1 to A1 notation.
    """
    parts = source_full_name.split("!")
    file_path = parts[0] if len(parts) >= 1 else "Not Specified"
    sheet_name = parts[1] if len(parts) >= 2 else "Not Specified"
    if len(parts) >= 3:
        range_address = convert_r1c1_to_a1(parts[2])
    else:
        range_address = "Not Specified"
    return file_path, sheet_name, range_address


def get_contrast_font_color(bg_color: int, dark_rgb: int, light_rgb: int) -> int:
    """Return dark or light font color based on background brightness.

    bg_color is a win32com Long (BGR order): B*65536 + G*256 + R.
    Uses weighted luminance formula: 0.299*R + 0.587*G + 0.114*B.
    """
    r = bg_color & 0xFF
    g = (bg_color >> 8) & 0xFF
    b = (bg_color >> 16) & 0xFF
    brightness = 0.299 * r + 0.587 * g + 0.114 * b
    return light_rgb if brightness < 128 else dark_rgb
