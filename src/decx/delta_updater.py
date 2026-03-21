"""Step 1c: Swap delta indicator arrows based on linked values."""

import logging
import os
from dataclasses import dataclass

from decx.shape_finder import (
    MSO_LINKED_OLE_OBJECT,
    find_delt_shape,
    find_table_shape,
    find_template_shape,
)
from decx.utils import extract_link_parts

log = logging.getLogger(__name__)


@dataclass
class DeltaItem:
    """Collected data for a single delta indicator to process."""

    slide_index: int
    ole_source_full: str
    ole_name: str
    delt_name: str
    delt_left: float
    delt_top: float
    delt_width: float
    delt_height: float


def _determine_sign(value: str) -> str:
    """Determine the sign of a cell value. Returns 'pos', 'neg', or 'none'."""
    test_val = value.strip()
    if test_val.endswith("%"):
        test_val = test_val[:-1].strip()
    try:
        num = float(test_val)
        if num > 0:
            return "pos"
        elif num < 0:
            return "neg"
        else:
            return "none"
    except (ValueError, TypeError):
        return "none"


def update_deltas(session, config: dict, inventory=None) -> int:
    """Apply delta indicator arrows based on linked values.

    Uses a two-pass approach:
      Pass 1: Collect OLE shapes with matching delt_ shapes (no modifications)
      Pass 2: Process collected items (delete old, paste new template)

    When inventory is provided, Pass 1 uses the pre-built delts dict
    for O(1) lookups instead of scanning slides.

    Returns the count of deltas updated.
    """
    delta_cfg = config.get("delta", {})
    tmpl_slide = delta_cfg.get("template_slide", 1)

    # Locate template shapes
    tmpl_pos = find_template_shape(
        session.presentation,
        delta_cfg.get("template_positive", "tmpl_delta_pos"),
        tmpl_slide,
    )
    tmpl_neg = find_template_shape(
        session.presentation,
        delta_cfg.get("template_negative", "tmpl_delta_neg"),
        tmpl_slide,
    )
    tmpl_none = find_template_shape(
        session.presentation,
        delta_cfg.get("template_none", "tmpl_delta_none"),
        tmpl_slide,
    )

    if not all([tmpl_pos, tmpl_neg, tmpl_none]):
        log.error(
            "Missing template shapes on slide %d. Expected: %s, %s, %s",
            tmpl_slide,
            delta_cfg.get("template_positive"),
            delta_cfg.get("template_negative"),
            delta_cfg.get("template_none"),
        )
        return 0

    pres = session.presentation

    # === PASS 1: Collect ===
    items: list[DeltaItem] = []

    if inventory is not None:
        # Use inventory: iterate only OLE shapes that have a delt_ match
        for slide, ole_shp in inventory.ole_shapes:
            ole_name = ole_shp.Name
            sld_idx = slide.SlideIndex
            delt = inventory.delts.get((sld_idx, ole_name))
            if delt is not None:
                if sld_idx <= 1:
                    continue  # skip template slide
                items.append(
                    DeltaItem(
                        slide_index=sld_idx,
                        ole_source_full=ole_shp.LinkFormat.SourceFullName,
                        ole_name=ole_name,
                        delt_name=delt.Name,
                        delt_left=delt.Left,
                        delt_top=delt.Top,
                        delt_width=delt.Width,
                        delt_height=delt.Height,
                    )
                )
    else:
        # Fallback: scan slides (backward compatibility)
        for sld_idx in range(2, pres.Slides.Count + 1):  # skip template slide
            slide = pres.Slides(sld_idx)
            for shp in slide.Shapes:
                if shp.Type != MSO_LINKED_OLE_OBJECT:
                    continue
                try:
                    if "Excel.Sheet" not in shp.OLEFormat.ProgID:
                        continue
                except Exception:
                    continue

                delt = find_delt_shape(slide, shp.Name)
                if delt is not None:
                    items.append(
                        DeltaItem(
                            slide_index=sld_idx,
                            ole_source_full=shp.LinkFormat.SourceFullName,
                            ole_name=shp.Name,
                            delt_name=delt.Name,
                            delt_left=delt.Left,
                            delt_top=delt.Top,
                            delt_width=delt.Width,
                            delt_height=delt.Height,
                        )
                    )

    # === PASS 2: Process ===
    count = 0
    for item in items:
        slide = pres.Slides(item.slide_index)
        cell_value = ""
        got_value = False

        # Primary: read from existing ntbl_/htmp_/trns_ table
        if inventory is not None:
            tbl_entry = inventory.tables.get((item.slide_index, item.ole_name))
            tbl_shape = tbl_entry[0] if tbl_entry is not None else None
        else:
            tbl_shape, _ = find_table_shape(slide, item.ole_name)

        if tbl_shape is not None and tbl_shape.HasTable:
            try:
                cell_value = tbl_shape.Table.Cell(
                    1, 1
                ).Shape.TextFrame.TextRange.Text.strip()
                if cell_value:
                    got_value = True
            except Exception:
                pass

        # Fallback: read from Excel
        if not got_value:
            file_path, sheet_name, range_address = extract_link_parts(
                item.ole_source_full
            )
            if range_address == "Not Specified" or not os.path.exists(file_path):
                continue

            try:
                wb = session.get_or_open_workbook(file_path)
                cell_value = wb.Sheets(sheet_name).Range(range_address).Text.strip()
                if cell_value:
                    got_value = True
            except Exception as e:
                log.warning(
                    "Failed to read Excel for delta '%s': %s", item.delt_name, e
                )

        if not got_value:
            continue

        # Determine sign and pick template
        sign = _determine_sign(cell_value)
        if sign == "pos":
            src_template = tmpl_pos
        elif sign == "neg":
            src_template = tmpl_neg
        else:
            src_template = tmpl_none

        # Delete old delt_ shape
        for shp in slide.Shapes:
            if shp.Name == item.delt_name:
                shp.Delete()
                break

        # Copy template to slide
        src_template.Copy()
        slide.Shapes.Paste()

        # The pasted shape is the last one on the slide
        new_shape = slide.Shapes(slide.Shapes.Count)
        new_shape.Left = item.delt_left
        new_shape.Top = item.delt_top
        new_shape.Width = item.delt_width
        new_shape.Height = item.delt_height
        new_shape.Name = item.delt_name

        count += 1
        log.debug("Slide %d | Delta '%s': %s -> %s", item.slide_index, item.delt_name, cell_value, sign)

    log.info("Updated %d delta indicator(s)", count)
    return count
