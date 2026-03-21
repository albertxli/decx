"""Step 1b: Read Excel data and populate PowerPoint tables."""

import logging
import os

from decx.formatting import extract_formatting, extract_formatting_minimal, apply_formatting
from decx.shape_finder import (
    MSO_LINKED_OLE_OBJECT,
    find_table_shape,
    find_delt_shape,
)
from decx.utils import (
    extract_link_parts,
    hex_to_rgb,
    get_contrast_font_color,
)

log = logging.getLogger(__name__)

# Excel conditional formatting constants
XL_CONDITION_VALUE_LOWEST = 1
XL_CONDITION_VALUE_HIGHEST = 2
XL_CONDITION_VALUE_PERCENTILE = 5


def _apply_color_scale(cell_range, config: dict):
    """Apply a 3-color scale to an Excel range (for htmp_ tables)."""
    heatmap = config.get("heatmap", {})
    color_min = hex_to_rgb(heatmap.get("color_minimum", "#F8696B"))
    color_mid = hex_to_rgb(heatmap.get("color_midpoint", "#FFEB84"))
    color_max = hex_to_rgb(heatmap.get("color_maximum", "#63BE7B"))

    cell_range.FormatConditions.Delete()
    cs = cell_range.FormatConditions.AddColorScale(ColorScaleType=3)

    cs.ColorScaleCriteria(1).Type = XL_CONDITION_VALUE_LOWEST
    cs.ColorScaleCriteria(1).FormatColor.Color = color_min

    cs.ColorScaleCriteria(2).Type = XL_CONDITION_VALUE_PERCENTILE
    cs.ColorScaleCriteria(2).Value = 50
    cs.ColorScaleCriteria(2).FormatColor.Color = color_mid

    cs.ColorScaleCriteria(3).Type = XL_CONDITION_VALUE_HIGHEST
    cs.ColorScaleCriteria(3).FormatColor.Color = color_max


def _process_linked_shape(session, slide, ole_shape, config: dict,
                          inventory=None) -> bool:
    """Process a single linked OLE shape: populate its associated PPT table.

    When inventory is provided, uses O(1) dict lookups instead of slide scans.
    For ntbl_ and trns_ tables (preserve_fill=True), skips extract/apply formatting
    entirely -- only updates cell text values. This eliminates ~232k COM calls.

    Returns True if a table was created/updated, False if skipped.
    """
    file_path, sheet_name, range_address = extract_link_parts(
        ole_shape.LinkFormat.SourceFullName
    )

    if range_address == "Not Specified":
        return False
    if not os.path.exists(file_path):
        return False

    ole_name = ole_shape.Name

    # Find existing special table shape (O(1) with inventory, O(n) without)
    if inventory is not None:
        tbl_entry = inventory.tables.get(ole_name)
        if tbl_entry is not None:
            existing_table, table_type = tbl_entry
        else:
            existing_table, table_type = None, None
    else:
        existing_table, table_type = find_table_shape(slide, ole_name)

    if existing_table is None:
        # If delt-only OLE (no table but has delt_ shape), skip
        if inventory is not None:
            has_delt = ole_name in inventory.delts
        else:
            has_delt = find_delt_shape(slide, ole_name) is not None
        if has_delt:
            return False

    # Determine behavior by prefix
    do_transpose = False
    local_preserve = True

    if existing_table is not None:
        if table_type == "ntbl":
            local_preserve = True
        elif table_type == "htmp":
            local_preserve = False
        elif table_type == "trns":
            local_preserve = True
            do_transpose = True
    else:
        local_preserve = True  # brand new table

    # For preserve_fill tables (ntbl_, trns_), skip formatting entirely.
    # The table already has correct formatting from the PPT template.
    # We only need to update cell text values.
    skip_formatting = (existing_table is not None and local_preserve)

    # Extract old formatting only when needed
    old_fmt = None
    if existing_table is not None and not skip_formatting:
        # For htmp_ tables, use minimal extraction (skip borders/margins/alignment)
        old_fmt = extract_formatting_minimal(existing_table)

    # Open Excel workbook and get range
    wb = session.get_or_open_workbook(file_path)
    excel_sheet = wb.Sheets(sheet_name)
    cell_range = excel_sheet.Range(range_address)

    # For heatmap (htmp_) or brand new without old formatting: apply color scale
    if not skip_formatting and (old_fmt is None or not local_preserve):
        _apply_color_scale(cell_range, config)
        session.excel_app.Calculate()
        excel_sheet.Calculate()

        # Apply contrast font colors
        heatmap = config.get("heatmap", {})
        dark_rgb = hex_to_rgb(heatmap.get("dark_font", "#000000"))
        light_rgb = hex_to_rgb(heatmap.get("light_font", "#FFFFFF"))

        for target_cell in cell_range:
            try:
                cell_color = target_cell.DisplayFormat.Interior.Color
                target_cell.Font.Color = get_contrast_font_color(
                    cell_color, dark_rgb, light_rgb
                )
            except Exception:
                pass

    # Create or reuse table shape
    max_rows = cell_range.Rows.Count
    max_cols = cell_range.Columns.Count

    if existing_table is not None:
        table_shape = existing_table
        if not skip_formatting and old_fmt is not None:
            apply_formatting(table_shape, old_fmt, preserve_fill=local_preserve)
    elif old_fmt is not None:
        table_shape = existing_table
        apply_formatting(table_shape, old_fmt, preserve_fill=local_preserve)
    else:
        table_shape = slide.Shapes.AddTable(max_rows, max_cols, 100, 100, 400, 200)
        table_shape.Name = f"ntbl_{ole_name}"
        table_shape.AlternativeText = (
            f"Linked to Excel File: {file_path}\n"
            f"Sheet: {sheet_name}\n"
            f"Range: {range_address}"
        )

    # Fill table cells from Excel
    tbl = table_shape.Table
    total_rows = tbl.Rows.Count
    total_cols = tbl.Columns.Count

    # Bulk-read all cell values in ONE COM call instead of per-cell .Text reads
    # Range.Value2 returns a 2D tuple: ((r1c1, r1c2, ...), (r2c1, ...), ...)
    # For single-cell ranges it returns a scalar, for single-row/col a 1D tuple
    raw_values = cell_range.Value2
    if raw_values is None:
        return True

    # Normalize to 2D list for uniform access
    if not isinstance(raw_values, tuple):
        # Single cell
        values = [[_format_value(raw_values)]]
    elif not isinstance(raw_values[0], tuple):
        # Single row
        values = [[_format_value(v) for v in raw_values]]
    else:
        # Normal 2D
        values = [[_format_value(v) for v in row] for row in raw_values]

    for row_idx in range(max_rows):
        for col_idx in range(max_cols):
            if do_transpose:
                ppt_row, ppt_col = col_idx + 1, row_idx + 1
            else:
                ppt_row, ppt_col = row_idx + 1, col_idx + 1

            if ppt_row > total_rows or ppt_col > total_cols:
                continue

            cell_shape = tbl.Cell(ppt_row, ppt_col).Shape

            # Write text from bulk-read values (no per-cell Excel COM call)
            cell_shape.TextFrame.TextRange.Text = values[row_idx][col_idx]

            # If not preserving formatting (htmp_ or brand new): pull fill & font from Excel
            # These still need per-cell COM reads for DisplayFormat colors
            if not skip_formatting and (old_fmt is None or not local_preserve):
                try:
                    cell_color = (
                        cell_range.Cells(row_idx + 1, col_idx + 1)
                        .DisplayFormat.Interior.Color
                    )
                    cell_shape.Fill.ForeColor.RGB = cell_color
                except Exception:
                    pass

                try:
                    excel_font = cell_range.Cells(row_idx + 1, col_idx + 1).Font
                    ppt_font = cell_shape.TextFrame.TextRange.Font
                    ppt_font.Name = excel_font.Name
                    ppt_font.Size = excel_font.Size
                    ppt_font.Bold = excel_font.Bold
                    ppt_font.Italic = excel_font.Italic
                    ppt_font.Color = excel_font.Color
                except Exception:
                    pass

    return True


def _format_value(value) -> str:
    """Convert a raw Excel Value2 to display string.

    Value2 returns raw values (numbers as floats, dates as serial numbers).
    We format them as strings matching what .Text would show.
    """
    if value is None:
        return ""
    if isinstance(value, float):
        # Remove trailing .0 for integers
        if value == int(value):
            return str(int(value))
        return str(value)
    return str(value)


def update_tables(session, config: dict, inventory=None) -> int:
    """Process all linked OLE shapes and update/create their PPT tables.

    When inventory is provided, iterates over pre-collected OLE shapes
    instead of scanning all slides again.

    Returns the count of tables updated.
    """
    if inventory is not None:
        ole_shapes = inventory.ole_shapes
    else:
        # Fallback: scan slides (backward compatibility)
        ole_shapes = []
        for slide in session.presentation.Slides:
            for shp in slide.Shapes:
                if shp.Type != MSO_LINKED_OLE_OBJECT:
                    continue
                try:
                    prog_id = shp.OLEFormat.ProgID
                except Exception:
                    continue
                if "Excel.Sheet" not in prog_id:
                    continue
                ole_shapes.append((slide, shp))

    count = 0
    for slide, shp in ole_shapes:
        try:
            if _process_linked_shape(session, slide, shp, config,
                                     inventory=inventory):
                count += 1
        except Exception as e:
            log.warning("Error processing shape '%s': %s", shp.Name, e)

    log.info("Updated %d table(s)", count)
    return count
