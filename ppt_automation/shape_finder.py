"""Shape discovery: word-boundary matching, find special shapes by prefix."""

# COM constants
MSO_LINKED_OLE_OBJECT = 10  # msoLinkedOLEObject
MSO_GROUP = 6               # msoGroup


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


def find_table_shape(slide, ole_name: str):
    """Find an existing ntbl_/htmp_/trns_ table shape associated with an OLE object.

    Search priority: ntbl -> htmp -> trns (matches VBA).
    Returns (shape, table_type) or (None, None).
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
    """Find a delt_ shape associated with an OLE object. Returns shape or None."""
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
    """Recursively collect linked OLE Excel.Sheet shapes, including inside groups."""
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
    """
    results = []
    for slide in presentation.Slides:
        for shp in slide.Shapes:
            _collect_ole_recursive(shp, results, slide)
    return results


def _collect_charts_recursive(shape, results):
    """Recursively collect linked chart shapes, including inside groups."""
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
    """
    results = []
    for slide in presentation.Slides:
        for shp in slide.Shapes:
            _collect_charts_recursive(shp, results)
    return results
