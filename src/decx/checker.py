"""Post-pipeline data validation — compare PPT values against Excel source."""

import logging
import re
from dataclasses import dataclass

from decx.delta_updater import _determine_sign, SIGN_SUFFIXES
from decx.utils import extract_link_parts

log = logging.getLogger(__name__)


@dataclass
class Mismatch:
    """A single mismatch between PPT and Excel."""

    slide: int
    shape_name: str
    detail: str
    category: str  # "table" or "delta"


def _extract_sign_suffix(name: str) -> str | None:
    """Extract the sign suffix from a delta shape name, or None if missing."""
    for suffix in SIGN_SUFFIXES:
        if name.endswith(suffix):
            return suffix[1:]  # "_pos" -> "pos"
    return None


def _parse_a1_top_left(range_address: str) -> tuple[int, int]:
    """Parse the top-left cell of an A1 range like 'B3:E10' into (row, col).

    Returns 1-based (row, col). E.g. 'B3' -> (3, 2), 'AA1' -> (1, 27).
    """
    top_left = range_address.split(":")[0]
    match = re.match(r"([A-Z]+)(\d+)", top_left.upper())
    if not match:
        return 1, 1
    col_str, row_str = match.groups()
    col = 0
    for ch in col_str:
        col = col * 26 + (ord(ch) - 64)
    return int(row_str), col


def _cell_ref(base_row: int, base_col: int, row_offset: int, col_offset: int) -> str:
    """Compute an A1 cell reference from a base position + 1-based offsets.

    E.g. base (3, 2) + offset (1, 1) -> 'B3', offset (2, 3) -> 'D4'.
    """
    row = base_row + row_offset - 1
    col = base_col + col_offset - 1
    col_letter = ""
    n = col
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        col_letter = chr(65 + remainder) + col_letter
    return f"{col_letter}{row}"


def _compare_cell_text(ppt_text: str, excel_text: str) -> bool:
    """Compare PPT and Excel cell text, tolerant of whitespace."""
    return ppt_text.strip() == excel_text.strip()


def check_tables(session, config, inventory, excel_override=None):
    """Check all table values against Excel source.

    Returns (checked_count, mismatches).
    """
    checked = 0
    mismatches = []

    for slide, ole_shp in inventory.ole_shapes:
        ole_name = ole_shp.Name
        sld_idx = slide.SlideIndex

        if sld_idx <= 1:
            continue  # skip template slide

        tbl_entry = inventory.tables.get((sld_idx, ole_name))
        if tbl_entry is None:
            continue

        table_shape, table_type = tbl_entry
        do_transpose = table_type == "trns"

        # Get Excel source
        try:
            source_full = ole_shp.LinkFormat.SourceFullName
            file_path, sheet_name, range_address = extract_link_parts(source_full)
        except Exception:
            log.warning("Slide %d, %s: cannot read OLE link", sld_idx, ole_name)
            continue

        if excel_override:
            file_path = excel_override

        # Open workbook and get range
        try:
            wb = session.get_or_open_workbook(file_path)
            excel_sheet = wb.Sheets(sheet_name)
            cell_range = excel_sheet.Range(range_address)
        except Exception as e:
            log.warning(
                "Slide %d, %s: cannot open Excel range: %s", sld_idx, ole_name, e
            )
            continue

        # Get PPT table
        try:
            ppt_table = table_shape.Table
        except Exception:
            log.warning("Slide %d, %s: shape is not a table", sld_idx, table_shape.Name)
            continue

        # Compare cell by cell
        excel_rows = cell_range.Rows.Count
        excel_cols = cell_range.Columns.Count
        base_row, base_col = _parse_a1_top_left(range_address)

        for r in range(1, excel_rows + 1):
            for c in range(1, excel_cols + 1):
                # Handle transposition
                ppt_r = c if do_transpose else r
                ppt_c = r if do_transpose else c

                try:
                    excel_text = cell_range.Cells(r, c).Text
                except Exception:
                    continue

                try:
                    ppt_text = ppt_table.Cell(
                        ppt_r, ppt_c
                    ).Shape.TextFrame.TextRange.Text
                except Exception:
                    continue

                excel_ref = f"{sheet_name}!{_cell_ref(base_row, base_col, r, c)}"

                checked += 1
                if not _compare_cell_text(ppt_text, excel_text):
                    detail = (
                        f"Row {ppt_r} Col {ppt_c}: "
                        f"PPT={ppt_text!r} Excel={excel_text!r} [{excel_ref}]"
                    )
                    mismatches.append(
                        Mismatch(
                            slide=sld_idx,
                            shape_name=table_shape.Name,
                            detail=detail,
                            category="table",
                        )
                    )
                else:
                    log.debug(
                        "Slide %d, %s — Row %d Col %d: OK (%s) [%s]",
                        sld_idx,
                        table_shape.Name,
                        ppt_r,
                        ppt_c,
                        ppt_text.strip(),
                        excel_ref,
                    )

    return checked, mismatches


def check_deltas(session, config, inventory, excel_override=None):
    """Check all delta indicators against Excel values.

    Returns (checked_count, mismatches).
    """
    checked = 0
    mismatches = []

    for slide, ole_shp in inventory.ole_shapes:
        ole_name = ole_shp.Name
        sld_idx = slide.SlideIndex

        if sld_idx <= 1:
            continue

        delt_shape = inventory.delts.get((sld_idx, ole_name))
        if delt_shape is None:
            continue

        # Extract actual sign from shape name suffix
        actual_sign = _extract_sign_suffix(delt_shape.Name)
        if actual_sign is None:
            log.warning(
                "Slide %d, %s: delta shape has no sign suffix (run decx update first)",
                sld_idx,
                delt_shape.Name,
            )
            continue

        # Get expected sign from Excel
        try:
            source_full = ole_shp.LinkFormat.SourceFullName
            file_path, sheet_name, range_address = extract_link_parts(source_full)
        except Exception:
            log.warning("Slide %d, %s: cannot read OLE link", sld_idx, ole_name)
            continue

        if excel_override:
            file_path = excel_override

        try:
            wb = session.get_or_open_workbook(file_path)
            cell_value = wb.Sheets(sheet_name).Range(range_address).Cells(1, 1).Text
        except Exception as e:
            log.warning(
                "Slide %d, %s: cannot read Excel value: %s", sld_idx, ole_name, e
            )
            continue

        base_row, base_col = _parse_a1_top_left(range_address)
        excel_ref = f"{sheet_name}!{_cell_ref(base_row, base_col, 1, 1)}"

        expected_sign = _determine_sign(cell_value.strip())
        checked += 1

        if actual_sign != expected_sign:
            detail = (
                f"PPT={actual_sign}, Excel value={cell_value.strip()!r} "
                f"(expected={expected_sign}) [{excel_ref}]"
            )
            mismatches.append(
                Mismatch(
                    slide=sld_idx,
                    shape_name=delt_shape.Name,
                    detail=detail,
                    category="delta",
                )
            )
        else:
            log.debug(
                "Slide %d, %s — sign=%s: OK (%s) [%s]",
                sld_idx,
                delt_shape.Name,
                actual_sign,
                cell_value.strip(),
                excel_ref,
            )

    return checked, mismatches
