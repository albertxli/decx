"""Unit tests for decx.differ — no COM required."""

from unittest.mock import MagicMock

from decx.differ import (
    Diff,
    DiffResult,
    _diff_deltas,
    _diff_tables,
    _extract_table_cells,
)


def _mock_table_shape(name, cells):
    """Create a mock table shape with given cell text grid.

    cells: list[list[str]] — 2D grid of cell text values.
    """
    shape = MagicMock()
    shape.Name = name
    tbl = MagicMock()
    tbl.Rows.Count = len(cells)
    tbl.Columns.Count = len(cells[0]) if cells else 0

    def cell_fn(r, c):
        cell = MagicMock()
        cell.Shape.TextFrame.TextRange.Text = cells[r - 1][c - 1]
        return cell

    tbl.Cell = cell_fn
    shape.Table = tbl
    return shape


def _mock_delta_shape(name):
    """Create a mock delta shape with the given name (includes sign suffix)."""
    shape = MagicMock()
    shape.Name = name
    return shape


def _mock_inventory(tables=None, delts=None, charts=None):
    """Create a mock SlideInventory."""
    inv = MagicMock()
    inv.tables = tables or {}
    inv.delts = delts or {}
    inv.charts = charts or []
    return inv


class TestExtractTableCells:
    def test_basic_grid(self):
        shape = _mock_table_shape("tbl1", [["a", "b"], ["c", "d"]])
        grid = _extract_table_cells(shape)
        assert grid == [["a", "b"], ["c", "d"]]

    def test_single_cell(self):
        shape = _mock_table_shape("tbl1", [["hello"]])
        grid = _extract_table_cells(shape)
        assert grid == [["hello"]]

    def test_strips_whitespace(self):
        shape = _mock_table_shape("tbl1", [["  x  ", " y"]])
        grid = _extract_table_cells(shape)
        assert grid == [["x", "y"]]


class TestDiffTables:
    def test_identical_tables(self):
        shape_a = _mock_table_shape("ntbl_Obj1", [["10%", "20%"], ["30%", "40%"]])
        shape_b = _mock_table_shape("ntbl_Obj1", [["10%", "20%"], ["30%", "40%"]])
        inv_a = _mock_inventory(tables={(2, "Object 1"): (shape_a, "ntbl")})
        inv_b = _mock_inventory(tables={(2, "Object 1"): (shape_b, "ntbl")})

        result = DiffResult()
        _diff_tables(inv_a, inv_b, result)

        assert result.tables_compared == 1
        assert result.cells_compared == 4
        assert result.diffs == []

    def test_different_values(self):
        shape_a = _mock_table_shape("ntbl_Obj1", [["10%", "20%"]])
        shape_b = _mock_table_shape("ntbl_Obj1", [["10%", "25%"]])
        inv_a = _mock_inventory(tables={(2, "Object 1"): (shape_a, "ntbl")})
        inv_b = _mock_inventory(tables={(2, "Object 1"): (shape_b, "ntbl")})

        result = DiffResult()
        _diff_tables(inv_a, inv_b, result)

        assert result.tables_compared == 1
        assert len(result.diffs) == 1
        assert result.diffs[0].category == "table"
        assert "'20%'" in result.diffs[0].detail
        assert "'25%'" in result.diffs[0].detail

    def test_table_only_in_a(self):
        shape_a = _mock_table_shape("ntbl_Obj1", [["10%"]])
        inv_a = _mock_inventory(tables={(2, "Object 1"): (shape_a, "ntbl")})
        inv_b = _mock_inventory(tables={})

        result = DiffResult()
        _diff_tables(inv_a, inv_b, result)

        assert result.tables_compared == 0
        assert len(result.only_in_a) == 1
        assert "Object 1" in result.only_in_a[0]

    def test_table_only_in_b(self):
        shape_b = _mock_table_shape("ntbl_Obj1", [["10%"]])
        inv_a = _mock_inventory(tables={})
        inv_b = _mock_inventory(tables={(2, "Object 1"): (shape_b, "ntbl")})

        result = DiffResult()
        _diff_tables(inv_a, inv_b, result)

        assert result.tables_compared == 0
        assert len(result.only_in_b) == 1

    def test_multiple_diffs_same_table(self):
        shape_a = _mock_table_shape("ntbl_Obj1", [["a", "b"], ["c", "d"]])
        shape_b = _mock_table_shape("ntbl_Obj1", [["x", "b"], ["c", "y"]])
        inv_a = _mock_inventory(tables={(2, "Object 1"): (shape_a, "ntbl")})
        inv_b = _mock_inventory(tables={(2, "Object 1"): (shape_b, "ntbl")})

        result = DiffResult()
        _diff_tables(inv_a, inv_b, result)

        assert len(result.diffs) == 2


class TestDiffDeltas:
    def test_same_sign(self):
        inv_a = _mock_inventory(
            delts={(2, "Object 1"): _mock_delta_shape("delt_Object_1_pos")}
        )
        inv_b = _mock_inventory(
            delts={(2, "Object 1"): _mock_delta_shape("delt_Object_1_pos")}
        )

        result = DiffResult()
        _diff_deltas(inv_a, inv_b, result)

        assert result.deltas_compared == 1
        assert result.diffs == []

    def test_different_sign(self):
        inv_a = _mock_inventory(
            delts={(2, "Object 1"): _mock_delta_shape("delt_Object_1_pos")}
        )
        inv_b = _mock_inventory(
            delts={(2, "Object 1"): _mock_delta_shape("delt_Object_1_neg")}
        )

        result = DiffResult()
        _diff_deltas(inv_a, inv_b, result)

        assert len(result.diffs) == 1
        assert result.diffs[0].category == "delta"
        assert "pos" in result.diffs[0].detail
        assert "neg" in result.diffs[0].detail

    def test_delta_only_in_a(self):
        inv_a = _mock_inventory(
            delts={(2, "Object 1"): _mock_delta_shape("delt_Object_1_pos")}
        )
        inv_b = _mock_inventory(delts={})

        result = DiffResult()
        _diff_deltas(inv_a, inv_b, result)

        assert result.deltas_compared == 0
        assert len(result.only_in_a) == 1

    def test_no_suffix_treated_as_unknown(self):
        inv_a = _mock_inventory(
            delts={(2, "Object 1"): _mock_delta_shape("delt_Object_1")}
        )
        inv_b = _mock_inventory(
            delts={(2, "Object 1"): _mock_delta_shape("delt_Object_1_pos")}
        )

        result = DiffResult()
        _diff_deltas(inv_a, inv_b, result)

        assert len(result.diffs) == 1
        assert "unknown" in result.diffs[0].detail


class TestDiffResult:
    def test_has_differences_empty(self):
        r = DiffResult()
        assert r.has_differences is False

    def test_has_differences_with_diffs(self):
        r = DiffResult(diffs=[Diff(1, "shape", "table", "detail")])
        assert r.has_differences is True

    def test_has_differences_with_only_in_a(self):
        r = DiffResult(only_in_a=["slide 2: Object 1"])
        assert r.has_differences is True
