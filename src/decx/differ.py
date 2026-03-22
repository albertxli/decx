"""Compare two PPTX files and report value differences.

Extracts table cell text, delta sign states, and chart series values
from both files via COM (read-only), then diffs them.
No Excel needed — pure PPT-to-PPT comparison.
"""

import logging
from dataclasses import dataclass, field

from decx.checker import _extract_sign_suffix
from decx.shape_finder import build_presentation_inventory

log = logging.getLogger(__name__)


@dataclass
class Diff:
    """A single value difference between two files."""

    slide: int
    shape: str
    category: str  # "table", "delta", "chart"
    detail: str


@dataclass
class DiffResult:
    """Aggregated diff results."""

    diffs: list[Diff] = field(default_factory=list)
    only_in_a: list[str] = field(default_factory=list)
    only_in_b: list[str] = field(default_factory=list)
    tables_compared: int = 0
    cells_compared: int = 0
    deltas_compared: int = 0
    charts_compared: int = 0
    series_compared: int = 0

    @property
    def has_differences(self) -> bool:
        return bool(self.diffs or self.only_in_a or self.only_in_b)


def _extract_table_cells(table_shape) -> list[list[str]]:
    """Extract all cell text from a PPT table shape as a 2D grid."""
    tbl = table_shape.Table
    rows = tbl.Rows.Count
    cols = tbl.Columns.Count
    grid = []
    for r in range(1, rows + 1):
        row = []
        for c in range(1, cols + 1):
            try:
                text = tbl.Cell(r, c).Shape.TextFrame.TextRange.Text
            except Exception:
                text = ""
            row.append(text.strip())
        grid.append(row)
    return grid


def _extract_chart_series(chart_shape) -> list[tuple[str, tuple]]:
    """Extract series names and values from a chart shape."""
    series_data = []
    try:
        chart = chart_shape.Chart
        sc = chart.SeriesCollection()
        for i in range(1, sc.Count + 1):
            s = sc.Item(i)
            try:
                name = s.Name
            except Exception:
                name = f"Series {i}"
            try:
                values = tuple(s.Values)
            except Exception:
                values = ()
            series_data.append((name, values))
    except Exception:
        pass
    return series_data


def run_diff(session_a, session_b) -> DiffResult:
    """Compare two open presentations and return differences.

    Both sessions must be open with read_only=True.
    """
    result = DiffResult()

    inv_a = build_presentation_inventory(session_a.presentation)
    inv_b = build_presentation_inventory(session_b.presentation)

    _diff_tables(inv_a, inv_b, result)
    _diff_deltas(inv_a, inv_b, result)
    _diff_charts(inv_a, inv_b, result)

    # Clean up COM refs before sessions exit
    del inv_a, inv_b

    return result


def _diff_tables(inv_a, inv_b, result: DiffResult):
    """Compare tables across two inventories."""
    keys_a = set(inv_a.tables.keys())
    keys_b = set(inv_b.tables.keys())

    for key in sorted(keys_a - keys_b):
        result.only_in_a.append(f"slide {key[0]}: {key[1]} (table)")
    for key in sorted(keys_b - keys_a):
        result.only_in_b.append(f"slide {key[0]}: {key[1]} (table)")

    for key in sorted(keys_a & keys_b):
        sld_idx, ole_name = key
        shape_a, _ = inv_a.tables[key]
        shape_b, _ = inv_b.tables[key]

        try:
            grid_a = _extract_table_cells(shape_a)
            grid_b = _extract_table_cells(shape_b)
        except Exception as e:
            log.warning("Slide %d, %s: cannot read table: %s", sld_idx, ole_name, e)
            continue

        result.tables_compared += 1
        rows = max(len(grid_a), len(grid_b))
        cols = max(
            (max(len(r) for r in grid_a) if grid_a else 0),
            (max(len(r) for r in grid_b) if grid_b else 0),
        )

        for r in range(rows):
            for c in range(cols):
                val_a = grid_a[r][c] if r < len(grid_a) and c < len(grid_a[r]) else ""
                val_b = grid_b[r][c] if r < len(grid_b) and c < len(grid_b[r]) else ""
                result.cells_compared += 1
                if val_a != val_b:
                    result.diffs.append(
                        Diff(
                            slide=sld_idx,
                            shape=shape_a.Name,
                            category="table",
                            detail=f"[{r+1},{c+1}] {val_a!r} → {val_b!r}",
                        )
                    )


def _diff_deltas(inv_a, inv_b, result: DiffResult):
    """Compare delta sign states across two inventories."""
    keys_a = set(inv_a.delts.keys())
    keys_b = set(inv_b.delts.keys())

    for key in sorted(keys_a - keys_b):
        result.only_in_a.append(f"slide {key[0]}: {key[1]} (delta)")
    for key in sorted(keys_b - keys_a):
        result.only_in_b.append(f"slide {key[0]}: {key[1]} (delta)")

    for key in sorted(keys_a & keys_b):
        sld_idx, ole_name = key
        shape_a = inv_a.delts[key]
        shape_b = inv_b.delts[key]

        sign_a = _extract_sign_suffix(shape_a.Name) or "unknown"
        sign_b = _extract_sign_suffix(shape_b.Name) or "unknown"

        result.deltas_compared += 1
        if sign_a != sign_b:
            result.diffs.append(
                Diff(
                    slide=sld_idx,
                    shape=ole_name,
                    category="delta",
                    detail=f"{sign_a} → {sign_b}",
                )
            )


def _diff_charts(inv_a, inv_b, result: DiffResult):
    """Compare chart series values across two inventories."""
    # Group charts by slide index for position-based matching
    by_slide_a: dict[int, list] = {}
    for slide, shp in inv_a.charts:
        by_slide_a.setdefault(slide.SlideIndex, []).append(shp)

    by_slide_b: dict[int, list] = {}
    for slide, shp in inv_b.charts:
        by_slide_b.setdefault(slide.SlideIndex, []).append(shp)

    all_slides = sorted(set(by_slide_a) | set(by_slide_b))

    for sld_idx in all_slides:
        list_a = by_slide_a.get(sld_idx, [])
        list_b = by_slide_b.get(sld_idx, [])

        for i in range(max(len(list_a), len(list_b))):
            if i >= len(list_a):
                result.only_in_b.append(f"slide {sld_idx}: chart #{i+1}")
                continue
            if i >= len(list_b):
                result.only_in_a.append(f"slide {sld_idx}: chart #{i+1}")
                continue

            shp_a = list_a[i]
            shp_b = list_b[i]

            series_a = _extract_chart_series(shp_a)
            series_b = _extract_chart_series(shp_b)

            result.charts_compared += 1
            chart_name = shp_a.Name

            for j in range(max(len(series_a), len(series_b))):
                if j >= len(series_a) or j >= len(series_b):
                    result.diffs.append(
                        Diff(
                            slide=sld_idx,
                            shape=chart_name,
                            category="chart",
                            detail=f"series count: {len(series_a)} → {len(series_b)}",
                        )
                    )
                    break

                name_a, vals_a = series_a[j]
                name_b, vals_b = series_b[j]
                result.series_compared += 1

                if vals_a != vals_b:
                    result.diffs.append(
                        Diff(
                            slide=sld_idx,
                            shape=chart_name,
                            category="chart",
                            detail=f"series {name_a!r}: values differ",
                        )
                    )
