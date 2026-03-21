"""Extract and apply table formatting (fonts, borders, fills, margins)."""

import logging
from dataclasses import dataclass, field

log = logging.getLogger(__name__)

# PowerPoint border type constants
PP_BORDER_LEFT = 1
PP_BORDER_RIGHT = 2
PP_BORDER_TOP = 3
PP_BORDER_BOTTOM = 4
BORDER_TYPES = (PP_BORDER_LEFT, PP_BORDER_RIGHT, PP_BORDER_TOP, PP_BORDER_BOTTOM)

# MSO constants
MSO_FALSE = 0
MSO_TRUE = -1
MSO_FILL_SOLID = 1
MSO_FILL_GRADIENT = 7
MSO_LINE_SOLID = 1


@dataclass
class CellFormatting:
    """Formatting data for a single table cell."""

    # Font
    font_name: str = ""
    font_size: float = 0
    font_bold: int = 0
    font_italic: int = 0
    font_color: int = 0
    font_underline: int = 0
    font_shadow: int = 0
    # Paragraph / alignment
    h_alignment: int = 0
    v_alignment: int = 0
    margin_left: float = 0
    margin_right: float = 0
    margin_top: float = 0
    margin_bottom: float = 0
    # Fill
    fill_visible: bool = True
    fill_type: int = 0
    fill_color: int = 0
    fill_transparency: float = 0
    # Borders: list of 4 tuples (visible, weight, dash_style, color)
    borders: list = field(
        default_factory=lambda: [
            (MSO_FALSE, 0, MSO_LINE_SOLID, 0xFFFFFF) for _ in range(4)
        ]
    )


@dataclass
class TableFormatting:
    """Complete formatting snapshot for a PowerPoint table shape."""

    shape_left: float = 0
    shape_top: float = 0
    shape_width: float = 0
    shape_height: float = 0
    row_heights: list[float] = field(default_factory=list)
    column_widths: list[float] = field(default_factory=list)
    cells: list[list[CellFormatting]] = field(default_factory=list)


def extract_formatting(table_shape) -> TableFormatting:
    """Extract all formatting data from a PowerPoint table shape."""
    tbl = table_shape.Table
    num_rows = tbl.Rows.Count
    num_cols = tbl.Columns.Count

    fmt = TableFormatting(
        shape_left=table_shape.Left,
        shape_top=table_shape.Top,
        shape_width=table_shape.Width,
        shape_height=table_shape.Height,
    )

    for i in range(1, num_rows + 1):
        fmt.row_heights.append(tbl.Rows(i).Height)

    for j in range(1, num_cols + 1):
        fmt.column_widths.append(tbl.Columns(j).Width)

    for i in range(1, num_rows + 1):
        row_cells = []
        for j in range(1, num_cols + 1):
            cell = tbl.Cell(i, j)
            cell_shape = cell.Shape
            tf = cell_shape.TextFrame
            font = tf.TextRange.Font
            fill = cell_shape.Fill

            cf = CellFormatting(
                font_name=font.Name,
                font_size=font.Size,
                font_bold=font.Bold,
                font_italic=font.Italic,
                font_color=font.Color.RGB,
                font_underline=font.Underline,
                font_shadow=font.Shadow,
                h_alignment=tf.TextRange.ParagraphFormat.Alignment,
                v_alignment=tf.VerticalAnchor,
                margin_left=tf.MarginLeft,
                margin_right=tf.MarginRight,
                margin_top=tf.MarginTop,
                margin_bottom=tf.MarginBottom,
                fill_visible=(fill.Visible != MSO_FALSE),
                fill_type=fill.Type,
                fill_color=fill.ForeColor.RGB,
                fill_transparency=fill.Transparency,
            )

            # Borders
            borders = []
            for bt in BORDER_TYPES:
                try:
                    border = cell.Borders(bt)
                    borders.append(
                        (
                            border.Visible,
                            border.Weight,
                            border.DashStyle,
                            border.ForeColor.RGB,
                        )
                    )
                except Exception:
                    borders.append((MSO_FALSE, 0, MSO_LINE_SOLID, 0xFFFFFF))
            cf.borders = borders

            row_cells.append(cf)
        fmt.cells.append(row_cells)

    return fmt


def extract_formatting_minimal(table_shape) -> TableFormatting:
    """Extract only geometry + fill formatting (skip borders/margins/alignment/font).

    Used for htmp_ tables where we only need fill colors from the existing table
    to avoid re-extracting them from Excel. Borders, margins, alignment, and font
    don't change so we skip them entirely (~75% fewer COM calls than full extract).
    """
    tbl = table_shape.Table
    num_rows = tbl.Rows.Count
    num_cols = tbl.Columns.Count

    fmt = TableFormatting(
        shape_left=table_shape.Left,
        shape_top=table_shape.Top,
        shape_width=table_shape.Width,
        shape_height=table_shape.Height,
    )

    for i in range(1, num_rows + 1):
        fmt.row_heights.append(tbl.Rows(i).Height)

    for j in range(1, num_cols + 1):
        fmt.column_widths.append(tbl.Columns(j).Width)

    for i in range(1, num_rows + 1):
        row_cells = []
        for j in range(1, num_cols + 1):
            cell = tbl.Cell(i, j)
            cell_shape = cell.Shape
            fill = cell_shape.Fill

            cf = CellFormatting(
                fill_visible=(fill.Visible != MSO_FALSE),
                fill_type=fill.Type,
                fill_color=fill.ForeColor.RGB,
                fill_transparency=fill.Transparency,
            )
            # Leave borders/font/margins at defaults - they won't be applied
            row_cells.append(cf)
        fmt.cells.append(row_cells)

    return fmt


def apply_formatting(table_shape, fmt: TableFormatting, preserve_fill: bool = True):
    """Apply saved formatting data back to a PowerPoint table shape.

    Args:
        table_shape: The PPT table shape to format.
        fmt: The saved formatting data.
        preserve_fill: If True, restore cell fill colors. If False, skip fill.
    """
    tbl = table_shape.Table
    num_rows = tbl.Rows.Count
    num_cols = tbl.Columns.Count

    # Row heights
    for i, height in enumerate(fmt.row_heights):
        if i < num_rows:
            try:
                tbl.Rows(i + 1).Height = height
            except Exception:
                pass

    # Column widths
    for j, width in enumerate(fmt.column_widths):
        if j < num_cols:
            try:
                tbl.Columns(j + 1).Width = width
            except Exception:
                pass

    # Per-cell formatting
    for i in range(min(num_rows, len(fmt.cells))):
        for j in range(min(num_cols, len(fmt.cells[i]))):
            cf = fmt.cells[i][j]
            cell = tbl.Cell(i + 1, j + 1)
            cell_shape = cell.Shape

            # Fill
            if preserve_fill:
                if not cf.fill_visible:
                    cell_shape.Fill.Visible = MSO_FALSE
                else:
                    cell_shape.Fill.Visible = MSO_TRUE
                    if cf.fill_type == MSO_FILL_GRADIENT:
                        cell_shape.Fill.TwoColorGradient(4, 1)  # msoGradientDiagonalUp
                    else:
                        cell_shape.Fill.Solid()
                    cell_shape.Fill.ForeColor.RGB = cf.fill_color
                    cell_shape.Fill.Transparency = cf.fill_transparency

            # Font
            try:
                font = cell_shape.TextFrame.TextRange.Font
                font.Name = cf.font_name
                font.Size = cf.font_size
                font.Bold = cf.font_bold
                font.Italic = cf.font_italic
                font.Color.RGB = cf.font_color
                font.Underline = cf.font_underline
                font.Shadow = cf.font_shadow
            except Exception:
                pass

            # Paragraph / margins
            try:
                tf = cell_shape.TextFrame
                tf.MarginLeft = cf.margin_left
                tf.MarginRight = cf.margin_right
                tf.MarginTop = cf.margin_top
                tf.MarginBottom = cf.margin_bottom
                tf.TextRange.ParagraphFormat.Alignment = cf.h_alignment
                tf.VerticalAnchor = cf.v_alignment
            except Exception:
                pass

            # Borders
            for k, bt in enumerate(BORDER_TYPES):
                try:
                    border = cell.Borders(bt)
                    visible, weight, dash, color = cf.borders[k]
                    if visible == MSO_FALSE:
                        border.Visible = MSO_FALSE
                        border.Weight = 0
                        border.DashStyle = MSO_LINE_SOLID
                        border.ForeColor.RGB = 0xFFFFFF
                    else:
                        border.Visible = visible
                        border.Weight = weight
                        border.DashStyle = dash
                        border.ForeColor.RGB = color
                except Exception:
                    pass
