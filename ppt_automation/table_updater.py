"""Step 1b: Read Excel data and populate PowerPoint tables."""

import logging
import os

from ppt_automation.formatting import extract_formatting, apply_formatting
from ppt_automation.shape_finder import (
    MSO_LINKED_OLE_OBJECT,
    find_table_shape,
    find_delt_shape,
)
from ppt_automation.utils import (
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


def _process_linked_shape(session, slide, ole_shape, config: dict) -> bool:
    """Process a single linked OLE shape: populate its associated PPT table.

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

    # Find existing special table shape
    existing_table, table_type = find_table_shape(slide, ole_name)

    if existing_table is None:
        # If delt-only OLE (no table but has delt_ shape), skip
        if find_delt_shape(slide, ole_name) is not None:
            return False

    # Extract old formatting if table exists
    old_fmt = None
    if existing_table is not None:
        old_fmt = extract_formatting(existing_table)

    # Open Excel workbook and get range
    wb = session.get_or_open_workbook(file_path)
    excel_sheet = wb.Sheets(sheet_name)
    cell_range = excel_sheet.Range(range_address)

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

    # For heatmap (htmp_) or brand new without old formatting: apply color scale
    if old_fmt is None or not local_preserve:
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

    if old_fmt is not None:
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

    for row_idx in range(1, max_rows + 1):
        for col_idx in range(1, max_cols + 1):
            if do_transpose:
                ppt_row, ppt_col = col_idx, row_idx
            else:
                ppt_row, ppt_col = row_idx, col_idx

            if ppt_row > total_rows or ppt_col > total_cols:
                continue

            cell_shape = tbl.Cell(ppt_row, ppt_col).Shape

            # Always copy text
            cell_shape.TextFrame.TextRange.Text = (
                cell_range.Cells(row_idx, col_idx).Text
            )

            # If no old formatting or htmp_ (not preserving): pull fill & font from Excel
            if old_fmt is None or not local_preserve:
                try:
                    cell_color = (
                        cell_range.Cells(row_idx, col_idx)
                        .DisplayFormat.Interior.Color
                    )
                    cell_shape.Fill.ForeColor.RGB = cell_color
                except Exception:
                    pass

                try:
                    excel_font = cell_range.Cells(row_idx, col_idx).Font
                    ppt_font = cell_shape.TextFrame.TextRange.Font
                    ppt_font.Name = excel_font.Name
                    ppt_font.Size = excel_font.Size
                    ppt_font.Bold = excel_font.Bold
                    ppt_font.Italic = excel_font.Italic
                    ppt_font.Color = excel_font.Color
                except Exception:
                    pass

    return True


def update_tables(session, config: dict) -> int:
    """Process all linked OLE shapes and update/create their PPT tables.

    Returns the count of tables updated.
    """
    count = 0
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

            try:
                if _process_linked_shape(session, slide, shp, config):
                    count += 1
            except Exception as e:
                log.warning("Error processing shape '%s': %s", shp.Name, e)

    log.info("Updated %d table(s)", count)
    return count
