"""Step 1a: Re-point OLE links to a new Excel file."""

import logging

from ppt_automation.shape_finder import collect_linked_ole_shapes

log = logging.getLogger(__name__)

# ppUpdateOptionManual = 1 (verified via PowerPoint type library)
# ppUpdateOptionAutomatic = 2
PP_UPDATE_OPTION_MANUAL = 1


def update_links(session, excel_path: str, config: dict, inventory=None) -> int:
    """Re-point all linked OLE objects to a new Excel file.

    Preserves sheet name and range from the original link.
    Optionally sets links to manual update mode.

    Performance optimization: repoints all SourceFullName paths first,
    sets AutoUpdate, then calls presentation.UpdateLinks() ONCE at the end
    instead of per-shape LinkFormat.Update().

    Args:
        session: COM session with presentation and ppt_app.
        excel_path: Path to the new Excel file.
        config: Configuration dict.
        inventory: Optional SlideInventory for O(1) shape lookup.

    Returns the count of updated links.
    """
    set_manual = config.get("links", {}).get("set_manual", True)

    if inventory is not None:
        ole_shapes = inventory.ole_shapes
    else:
        ole_shapes = collect_linked_ole_shapes(session.presentation)

    if not ole_shapes:
        log.info("No linked OLE shapes found")
        return 0

    updated = 0
    for _slide, shp in ole_shapes:
        try:
            old_link = shp.LinkFormat.SourceFullName
            bang_pos = old_link.find("!")
            if bang_pos < 0:
                continue

            # Everything after first '!' is 'sheet!range'
            link_tail = old_link[bang_pos + 1:]
            shp.LinkFormat.SourceFullName = f"{excel_path}!{link_tail}"

            if set_manual:
                shp.LinkFormat.AutoUpdate = PP_UPDATE_OPTION_MANUAL

            # Per-shape refresh — faster than bulk UpdateLinks() for this workload
            shp.LinkFormat.Update()

            updated += 1
            log.debug("Updated link: %s -> %s", shp.Name, excel_path)
        except Exception as e:
            log.warning("Failed to update link for shape '%s': %s", shp.Name, e)

    log.info("Updated %d OLE link(s)", updated)
    return updated
