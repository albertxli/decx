"""Shape discovery: word-boundary matching, find special shapes by prefix."""

from dataclasses import dataclass, field

# COM constants
MSO_LINKED_OLE_OBJECT = 10  # msoLinkedOLEObject
MSO_GROUP = 6  # msoGroup


def is_exact_token_match(shape_name: str, linked_name: str) -> bool:
    """Check if linked_name appears as a complete token in shape_name.

    A token boundary is: start/end of string, underscore, space, hyphen,
    or any non-alphanumeric character.
    """
    name_len = len(shape_name)
    linked_len = len(linked_name)
    pos = 0

    while True:
        pos = shape_name.find(linked_name, pos)
        if pos == -1:
            return False

        # Check character before
        if pos == 0:
            before_ok = True
        else:
            before_ok = _is_word_boundary(shape_name[pos - 1])

        # Check character after
        end_pos = pos + linked_len
        if end_pos >= name_len:
            after_ok = True
        else:
            after_ok = _is_word_boundary(shape_name[end_pos])

        if before_ok and after_ok:
            return True

        pos += 1


def _is_word_boundary(char: str) -> bool:
    """Return True if char is a word boundary (non-alphanumeric)."""
    return not char.isalnum()


@dataclass
class SlideInventory:
    """Pre-scanned index of all interesting shapes in a presentation.

    Built once by build_presentation_inventory(), used by all pipeline steps
    for O(1) lookups instead of repeated slide enumeration.
    """

    # All linked OLE Excel.Sheet shapes: list of (slide, shape)
    ole_shapes: list = field(default_factory=list)
    # ole_name -> (shape, table_type) for ntbl_/htmp_/trns_ shapes
    tables: dict = field(default_factory=dict)
    # ole_name -> shape for delt_ shapes
    delts: dict = field(default_factory=dict)
    # Shapes with _ccst in name and HasTable
    ccst_tables: list = field(default_factory=list)
    # Linked chart shapes (HasChart + IsLinked)
    charts: list = field(default_factory=list)
    # Raw counts of ALL special shapes (for decx info, includes unmatched)
    count_ntbl: int = 0
    count_htmp: int = 0
    count_trns: int = 0
    count_delt: int = 0
    count_ccst: int = 0


def _scan_shape_recursive(
    shape,
    slide,
    inventory: SlideInventory,
    all_table_shapes: list,
    all_delt_shapes: list,
):
    """Recursively scan a shape (and groups) to populate inventory lists."""
    name = shape.Name

    if shape.Type == MSO_GROUP:
        # Check the group shape itself for delt_ prefix BEFORE recursing.
        # delt_ shapes are often groups (arrows with multiple sub-shapes).
        if "delt_" in name:
            all_delt_shapes.append((shape, name))
            inventory.count_delt += 1
        # Recurse into group items for OLE objects, charts, etc.
        for sub_shp in shape.GroupItems:
            _scan_shape_recursive(
                sub_shp, slide, inventory, all_table_shapes, all_delt_shapes
            )
        return

    # Linked OLE Excel.Sheet
    if shape.Type == MSO_LINKED_OLE_OBJECT:
        if shape.LinkFormat is not None:
            try:
                prog_id = shape.OLEFormat.ProgID
                if "Excel.Sheet" in prog_id:
                    inventory.ole_shapes.append((slide, shape))
            except Exception:
                pass

    # Linked charts
    if shape.HasChart:
        try:
            if shape.Chart.ChartData.IsLinked:
                inventory.charts.append(shape)
        except Exception:
            pass

    # Table shapes: collect for later OLE-name matching
    if shape.HasTable:
        # _ccst tables
        if "_ccst" in name:
            inventory.ccst_tables.append(shape)
            inventory.count_ccst += 1

        # ntbl_/htmp_/trns_ candidates
        for prefix in ("ntbl_", "htmp_", "trns_"):
            if prefix in name:
                all_table_shapes.append((shape, prefix.rstrip("_"), name))
                if prefix == "ntbl_":
                    inventory.count_ntbl += 1
                elif prefix == "htmp_":
                    inventory.count_htmp += 1
                elif prefix == "trns_":
                    inventory.count_trns += 1
                break  # a shape matches at most one prefix

    # delt_ candidates (non-group shapes with delt_ in name)
    if "delt_" in name:
        all_delt_shapes.append((shape, name))
        inventory.count_delt += 1


def build_presentation_inventory(presentation) -> SlideInventory:
    """Scan all slides/shapes ONCE and build a complete inventory.

    This replaces multiple per-step slide enumerations with a single pass.
    Table and delt shapes are indexed by OLE name for O(1) lookup.
    """
    inventory = SlideInventory()
    all_table_shapes: list = []  # (shape, table_type, shape_name)
    all_delt_shapes: list = []  # (shape, shape_name)

    for slide in presentation.Slides:
        for shp in slide.Shapes:
            _scan_shape_recursive(
                shp, slide, inventory, all_table_shapes, all_delt_shapes
            )

    # Now index tables and delts by OLE name.
    # For each OLE shape, find matching table/delt shapes.
    for _slide, ole_shp in inventory.ole_shapes:
        ole_name = ole_shp.Name

        # Tables: search priority ntbl -> htmp -> trns
        if ole_name not in inventory.tables:
            for tbl_shape, tbl_type, tbl_name in all_table_shapes:
                if tbl_type == "ntbl" and "ntbl_" in tbl_name:
                    if is_exact_token_match(tbl_name, ole_name):
                        inventory.tables[ole_name] = (tbl_shape, tbl_type)
                        break
            else:
                # Try htmp
                for tbl_shape, tbl_type, tbl_name in all_table_shapes:
                    if tbl_type == "htmp" and "htmp_" in tbl_name:
                        if is_exact_token_match(tbl_name, ole_name):
                            inventory.tables[ole_name] = (tbl_shape, tbl_type)
                            break
                else:
                    # Try trns
                    for tbl_shape, tbl_type, tbl_name in all_table_shapes:
                        if tbl_type == "trns" and "trns_" in tbl_name:
                            if is_exact_token_match(tbl_name, ole_name):
                                inventory.tables[ole_name] = (tbl_shape, tbl_type)
                                break

        # Delts
        if ole_name not in inventory.delts:
            for delt_shape, delt_name in all_delt_shapes:
                if is_exact_token_match(delt_name, ole_name):
                    inventory.delts[ole_name] = delt_shape
                    break

    return inventory


# === Backward-compatible functions (deprecated, use inventory instead) ===


def find_table_shape(slide, ole_name: str):
    """Find an existing ntbl_/htmp_/trns_ table shape associated with an OLE object.

    Search priority: ntbl -> htmp -> trns (matches VBA).
    Returns (shape, table_type) or (None, None).

    Deprecated: use build_presentation_inventory() and inventory.tables instead.
    """
    # Collect candidates in priority order
    for prefix in ("ntbl_", "htmp_", "trns_"):
        for shp in slide.Shapes:
            if shp.HasTable:
                name = shp.Name
                if prefix in name and is_exact_token_match(name, ole_name):
                    return shp, prefix.rstrip("_")
    return None, None


def find_delt_shape(slide, ole_name: str):
    """Find a delt_ shape associated with an OLE object. Returns shape or None.

    Deprecated: use build_presentation_inventory() and inventory.delts instead.
    """
    for shp in slide.Shapes:
        name = shp.Name
        if "delt_" in name and is_exact_token_match(name, ole_name):
            return shp
    return None


def find_template_shape(presentation, template_name: str, slide_index: int = 1):
    """Find a template shape by exact name on the specified slide.

    Returns the shape or None.
    """
    slide = presentation.Slides(slide_index)
    for shp in slide.Shapes:
        if shp.Name == template_name:
            return shp
    return None


def _collect_ole_recursive(shape, results, slide):
    """Recursively collect linked OLE Excel.Sheet shapes, including inside groups.

    Deprecated: use build_presentation_inventory() instead.
    """
    if shape.Type == MSO_GROUP:
        for sub_shp in shape.GroupItems:
            _collect_ole_recursive(sub_shp, results, slide)
    elif shape.Type == MSO_LINKED_OLE_OBJECT:
        if shape.LinkFormat is not None:
            try:
                prog_id = shape.OLEFormat.ProgID
                if "Excel.Sheet" in prog_id:
                    results.append((slide, shape))
            except Exception:
                pass


def collect_linked_ole_shapes(presentation) -> list:
    """Collect all linked OLE Excel.Sheet shapes across all slides.

    Returns list of (slide, shape) tuples. Recurses into groups.

    Deprecated: use build_presentation_inventory() instead.
    """
    results = []
    for slide in presentation.Slides:
        for shp in slide.Shapes:
            _collect_ole_recursive(shp, results, slide)
    return results


def _collect_charts_recursive(shape, results):
    """Recursively collect linked chart shapes, including inside groups.

    Deprecated: use build_presentation_inventory() instead.
    """
    if shape.Type == MSO_GROUP:
        for sub_shp in shape.GroupItems:
            _collect_charts_recursive(sub_shp, results)
    elif shape.HasChart:
        try:
            if shape.Chart.ChartData.IsLinked:
                results.append(shape)
        except Exception:
            pass


def collect_linked_charts(presentation) -> list:
    """Collect all linked chart shapes across all slides.

    Returns list of shapes. Recurses into groups.

    Deprecated: use build_presentation_inventory() instead.
    """
    results = []
    for slide in presentation.Slides:
        for shp in slide.Shapes:
            _collect_charts_recursive(shp, results)
    return results
