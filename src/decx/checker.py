"""Post-pipeline data validation — compare PPT values against Excel source.

Future improvements:
- _ccst color coding check: verify font color matches expected color for each cell value
  (positive → green, negative → pink, neutral → gray per config). Pure PPT check, no Excel needed.
"""

import logging
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field

from decx.delta_updater import _determine_sign, SIGN_SUFFIXES
from decx.utils import extract_link_parts

log = logging.getLogger(__name__)


@dataclass
class CheckResult:
    """Aggregated results from a single check run."""

    tbl_checked: int = 0
    tbl_mismatches: list = field(default_factory=list)
    delt_checked: int = 0
    delt_mismatches: list = field(default_factory=list)
    num_charts: int = 0
    chart_series_checked: int = 0
    chart_mismatches: list = field(default_factory=list)

    @property
    def total_checked(self):
        return self.tbl_checked + self.delt_checked + self.chart_series_checked

    @property
    def all_mismatches(self):
        return self.tbl_mismatches + self.delt_mismatches + self.chart_mismatches

    @property
    def passed(self):
        return len(self.all_mismatches) == 0


def run_check(session, config, inventory, excel_override=None):
    """Run all checks (tables, deltas, charts) and return a CheckResult."""
    result = CheckResult()

    result.tbl_checked, result.tbl_mismatches = check_tables(
        session, config, inventory, excel_override=excel_override
    )
    result.delt_checked, result.delt_mismatches = check_deltas(
        session, config, inventory, excel_override=excel_override
    )
    result.num_charts, result.chart_series_checked, result.chart_mismatches = (
        check_charts(session, config, inventory, excel_override=excel_override)
    )

    return result


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

        # Skip _ccst tables — color coder transforms their text (strips %,
        # adds +/- prefix), so raw Excel comparison would always mismatch.
        if "_ccst" in table_shape.Name:
            continue

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


def _is_empty_or_zero(val, tol: float = 1e-9) -> bool:
    """Check if a value is None, empty string, or zero."""
    if val is None or val == "":
        return True
    try:
        return abs(float(val)) < tol
    except (TypeError, ValueError):
        return False


def _values_match(before: tuple, after: tuple, tol: float = 1e-9) -> bool:
    """Compare two tuples of numeric values with float tolerance.

    Treats None, empty string, and 0.0 as equivalent (charts plot empty cells as zero).
    """
    if len(before) != len(after):
        return False
    for a, b in zip(before, after):
        # Both empty/zero → match
        if _is_empty_or_zero(a, tol) and _is_empty_or_zero(b, tol):
            continue
        try:
            if abs(float(a) - float(b)) > tol:
                return False
        except (TypeError, ValueError):
            if a != b:
                return False
    return True


def _build_chart_ref_map(pptx_path: str) -> dict[tuple[int, int], list[str]]:
    """Parse PPTX zip to build a map of (slide_num, chart_position) -> [range_refs].

    Uses the slide XML → slide rels → chart XML chain to correctly map
    each chart's position on each slide to its series range references.
    """
    ns_p = "http://schemas.openxmlformats.org/presentationml/2006/main"
    ns_c = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    result = {}

    try:
        with zipfile.ZipFile(pptx_path, "r") as z:
            # Build presentation-order slide list from presentation.xml
            # (XML slide filenames may not match COM slide indices)
            ns_pres = "http://schemas.openxmlformats.org/presentationml/2006/main"
            ns_rel = (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            )

            pres_root = ET.fromstring(z.read("ppt/presentation.xml"))
            pres_rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
            rid_to_target = {r.get("Id"): r.get("Target") for r in pres_rels}

            slide_files_ordered = []
            sld_id_lst = pres_root.find(f"{{{ns_pres}}}sldIdLst")
            if sld_id_lst is not None:
                for sld_id in sld_id_lst:
                    rid = sld_id.get(f"{{{ns_rel}}}id")
                    target = rid_to_target.get(rid, "")
                    if target:
                        slide_files_ordered.append(f"ppt/{target}")

            for com_index, sf in enumerate(slide_files_ordered, 1):
                # Extract slide filename number for rels lookup
                slide_file_num = re.search(r"slide(\d+)", sf)
                if not slide_file_num:
                    continue
                file_num = slide_file_num.group(1)

                # Parse slide rels to map rId -> chart file
                rels_path = f"ppt/slides/_rels/slide{file_num}.xml.rels"
                try:
                    rels_root = ET.fromstring(z.read(rels_path))
                    rid_map = {r.get("Id"): r.get("Target") for r in rels_root}
                except Exception:
                    continue

                # Find chart graphicFrames in document order
                slide_root = ET.fromstring(z.read(sf))
                chart_pos = 0

                for gf in slide_root.iter(f"{{{ns_p}}}graphicFrame"):
                    chart_elem = gf.find(f".//{{{ns_c}}}chart")
                    if chart_elem is None:
                        continue

                    rid = chart_elem.get(f"{{{ns_r}}}id")
                    target = rid_map.get(rid)
                    if not target:
                        chart_pos += 1
                        continue

                    # Resolve relative path: ../charts/chart5.xml -> ppt/charts/chart5.xml
                    chart_path = target.replace("../", "ppt/")

                    # Skip unlinked charts (no external .rels entry).
                    # COM's IsLinked check excludes these, so we must too
                    # to keep position alignment.
                    chart_rels_path = (
                        chart_path.replace("ppt/charts/", "ppt/charts/_rels/") + ".rels"
                    )
                    try:
                        chart_rels = ET.fromstring(z.read(chart_rels_path))
                        has_external = any(
                            r.get("TargetMode") == "External" for r in chart_rels
                        )
                    except Exception:
                        has_external = False
                    if not has_external:
                        continue  # skip — not counted by COM either

                    # Parse chart XML for series Y-value ranges ONLY
                    # (skip cat/numRef which are category/X-axis references)
                    try:
                        chart_root = ET.fromstring(z.read(chart_path))
                        refs = []
                        for ser in chart_root.iter(f"{{{ns_c}}}ser"):
                            val_elem = ser.find(f"{{{ns_c}}}val")
                            if val_elem is not None:
                                num_ref = val_elem.find(f"{{{ns_c}}}numRef")
                                if num_ref is not None:
                                    f_elem = num_ref.find(f"{{{ns_c}}}f")
                                    if f_elem is not None and f_elem.text:
                                        refs.append(f_elem.text)
                        result[(com_index, chart_pos)] = refs
                    except Exception:
                        pass

                    chart_pos += 1

    except Exception as e:
        log.warning("Cannot parse chart XML from %s: %s", pptx_path, e)

    return result


def _flatten_range_value(raw) -> list:
    """Flatten the result of Range.Value2 into a flat list of values.

    Range.Value2 returns:
    - A single value (int/float/str/None) for a 1-cell range
    - A tuple of tuples ((row1_vals,), (row2_vals,), ...) for multi-cell ranges
    """
    if raw is None:
        return [None]
    if not isinstance(raw, tuple):
        return [raw]
    # Tuple of tuples (2D) — flatten row by row
    values = []
    for item in raw:
        if isinstance(item, tuple):
            values.extend(item)
        else:
            values.append(item)
    return values


# Cache sheet objects to avoid repeated COM lookups
_sheet_cache: dict[tuple, object] = {}


def _get_sheet(wb, sheet_name: str):
    """Get a worksheet, caching by (workbook, sheet_name)."""
    key = (id(wb), sheet_name)
    if key not in _sheet_cache:
        _sheet_cache[key] = wb.Sheets(sheet_name)
    return _sheet_cache[key]


def _read_chart_range(wb, range_ref: str) -> list:
    """Read values from a chart range reference, handling multi-area ranges.

    Uses bulk Range.Value2 (one COM call per sub-range) instead of per-cell reads.
    GOTCHAS #16 does not apply here — chart check compares raw floats, not formatted text.

    Supports both simple ranges ('Tables!$B$9:$B$13') and non-contiguous
    ranges ('(Tables!$C$810,Tables!$F$810)').
    """
    # Strip outer parentheses if present: (Tables!$C$810,Tables!$F$810)
    ref = range_ref.strip()
    if ref.startswith("(") and ref.endswith(")"):
        ref = ref[1:-1]

    # Split on comma to handle multi-area ranges
    sub_refs = [s.strip() for s in ref.split(",")]

    values = []
    for sub_ref in sub_refs:
        clean = sub_ref.replace("$", "")
        if "!" in clean:
            sheet_name, range_addr = clean.split("!", 1)
        else:
            sheet_name, range_addr = "Sheet1", clean

        sheet = _get_sheet(wb, sheet_name)
        raw = sheet.Range(range_addr).Value2
        values.extend(_flatten_range_value(raw))

    return values


def check_charts(session, config, inventory, excel_override=None):
    """Check linked chart data against Excel source.

    Reads Series.Values from the chart (via COM) and compares against
    the actual Excel range values (parsed from the chart XML inside the PPTX zip).
    Charts are matched by (slide_number, position_within_slide).

    Returns (chart_count, series_checked, mismatches).
    """
    series_checked = 0
    mismatches = []

    if not inventory.charts:
        return 0, series_checked, mismatches

    # Clear sheet cache from any prior run
    _sheet_cache.clear()

    # Build chart ref map keyed by (slide_num, position_on_slide)
    pptx_path = session.pptx_path
    chart_ref_map = _build_chart_ref_map(pptx_path)

    # Group COM charts by slide, preserving order
    from collections import defaultdict

    charts_by_slide = defaultdict(list)
    for slide, shp in inventory.charts:
        charts_by_slide[slide.SlideIndex].append(shp)

    chart_count = 0
    for slide_idx in sorted(charts_by_slide.keys()):
        for pos, shp in enumerate(charts_by_slide[slide_idx]):
            chart_count += 1
            chart_name = shp.Name
            try:
                chart = shp.Chart
                sc = chart.SeriesCollection()
                series_count = sc.Count
            except Exception as e:
                log.warning("Chart %s: cannot read series: %s", chart_name, e)
                continue

            # Get the Excel file path for this chart
            try:
                excel_file = shp.LinkFormat.SourceFullName
            except Exception:
                excel_file = None

            if excel_override:
                excel_file = excel_override

            # Get range refs from XML using (slide_num, position)
            series_refs = chart_ref_map.get((slide_idx, pos), [])

            for i in range(1, series_count + 1):
                try:
                    s = sc.Item(i)
                    series_name = s.Name
                    ppt_values = tuple(s.Values)
                except Exception as e:
                    log.warning(
                        "Chart %s, Series %d: cannot read values: %s",
                        chart_name,
                        i,
                        e,
                    )
                    continue

                ref_idx = i - 1
                if ref_idx < len(series_refs) and excel_file:
                    range_ref = series_refs[ref_idx]

                    try:
                        wb = session.get_or_open_workbook(excel_file)
                        excel_values = _read_chart_range(wb, range_ref)

                        series_checked += 1

                        if not _values_match(tuple(ppt_values), tuple(excel_values)):
                            changes = []
                            for idx, (pv, ev) in enumerate(
                                zip(ppt_values, excel_values)
                            ):
                                try:
                                    if (
                                        ev is not None
                                        and abs(float(pv) - float(ev)) > 1e-9
                                    ):
                                        changes.append(
                                            f"[{idx + 1}]: PPT={pv:.6g} Excel={ev:.6g}"
                                        )
                                except (TypeError, ValueError):
                                    if pv != ev:
                                        changes.append(
                                            f"[{idx + 1}]: PPT={pv!r} Excel={ev!r}"
                                        )
                            detail = (
                                f"Series '{series_name}': "
                                f"{', '.join(changes[:3])}"
                                f"{' ...' if len(changes) > 3 else ''}"
                                f" [{range_ref}]"
                            )
                            mismatches.append(
                                Mismatch(
                                    slide=slide_idx,
                                    shape_name=chart_name,
                                    detail=detail,
                                    category="chart",
                                )
                            )
                        else:
                            log.debug(
                                "Slide %d, Chart %s, Series '%s': OK (%d values) [%s]",
                                slide_idx,
                                chart_name,
                                series_name,
                                len(ppt_values),
                                range_ref,
                            )
                    except Exception as e:
                        log.warning(
                            "Chart %s, Series '%s': cannot read Excel range %s: %s",
                            chart_name,
                            series_name,
                            range_ref,
                            e,
                        )
                else:
                    series_checked += 1
                    log.debug(
                        "Slide %d, Chart %s, Series '%s': %d values (no range ref)",
                        slide_idx,
                        chart_name,
                        series_name,
                        len(ppt_values),
                    )

    return chart_count, series_checked, mismatches
