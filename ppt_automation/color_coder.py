"""Step 1d: Apply color coding to _ccst tables based on numeric sign."""

import logging

from ppt_automation.utils import hex_to_rgb

log = logging.getLogger(__name__)


def _is_numeric(value: str) -> tuple[bool, float]:
    """Check if a string is numeric (after stripping %). Returns (is_num, parsed_value)."""
    try:
        return True, float(value)
    except (ValueError, TypeError):
        return False, 0.0


def _validate_symbol_removal(symbols: str) -> bool:
    """Validate that symbol_removal only contains +, -, %."""
    return all(ch in "+-%" for ch in symbols)


def apply_color_coding(session, config: dict) -> int:
    """Apply color coding to all tables with '_ccst' in their name.

    Positive values -> green, negative -> red, zero/text -> grey.
    Optionally adds '+' prefix and strips symbols.
    Returns the count of tables processed.
    """
    ccst_cfg = config.get("ccst", {})
    positive_color = hex_to_rgb(ccst_cfg.get("positive_color", "#33CC33"))
    negative_color = hex_to_rgb(ccst_cfg.get("negative_color", "#ED0590"))
    neutral_color = hex_to_rgb(ccst_cfg.get("neutral_color", "#595959"))
    positive_prefix = ccst_cfg.get("positive_prefix", "+")
    symbol_removal = ccst_cfg.get("symbol_removal", "%")

    if not _validate_symbol_removal(symbol_removal):
        log.error("Invalid symbol_removal setting: only '+', '-', '%%' allowed")
        return 0

    total_tables = 0

    for slide in session.presentation.Slides:
        for shp in slide.Shapes:
            if "_ccst" not in shp.Name:
                continue
            if not shp.HasTable:
                continue

            total_tables += 1
            tbl = shp.Table

            for row_idx in range(1, tbl.Rows.Count + 1):
                for col_idx in range(1, tbl.Columns.Count + 1):
                    cell_text = tbl.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Text.strip()
                    test_val = cell_text
                    had_percent = False

                    # Strip trailing % for numeric test
                    if test_val.endswith("%"):
                        had_percent = True
                        test_val = test_val[:-1].strip()

                    is_num, cell_value = _is_numeric(test_val)

                    if is_num:
                        # Add prefix for positives
                        if cell_value > 0 and positive_prefix:
                            if not cell_text.startswith(positive_prefix):
                                if had_percent:
                                    cell_text = f"{positive_prefix}{test_val.strip()}%"
                                else:
                                    cell_text = f"{positive_prefix}{test_val.strip()}"
                                tbl.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Text = cell_text

                        # Color by sign
                        if cell_value > 0:
                            tbl.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Font.Color.RGB = positive_color
                        elif cell_value < 0:
                            tbl.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Font.Color.RGB = negative_color
                        else:
                            tbl.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Font.Color.RGB = neutral_color

                        # Symbol removal
                        if symbol_removal:
                            if "%" in symbol_removal and cell_text.endswith("%"):
                                cell_text = cell_text[:-1]
                            if "+" in symbol_removal and cell_text.startswith("+"):
                                cell_text = cell_text[1:]
                            if "-" in symbol_removal and cell_text.startswith("-"):
                                cell_text = cell_text[1:]
                            tbl.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Text = cell_text
                    else:
                        # Non-numeric: neutral color
                        tbl.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Font.Color.RGB = neutral_color

    log.info("Color-coded %d _ccst table(s)", total_tables)
    return total_tables
