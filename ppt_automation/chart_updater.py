"""Step 2: Update linked chart data sources."""

import logging

from ppt_automation.shape_finder import collect_linked_charts

log = logging.getLogger(__name__)

# ppUpdateOptionManual = 1 (verified via PowerPoint type library)
PP_UPDATE_OPTION_MANUAL = 1


def update_charts(session, excel_path: str) -> int:
    """Re-link all embedded charts to the specified Excel file.

    Sets chart links to manual update mode after updating.
    Returns the count of updated charts.
    """
    charts = collect_linked_charts(session.presentation)

    if not charts:
        log.info("No linked charts found")
        return 0

    updated = 0
    for chart_shape in charts:
        try:
            chart_shape.LinkFormat.SourceFullName = excel_path
            chart_shape.LinkFormat.Update()
            chart_shape.LinkFormat.AutoUpdate = PP_UPDATE_OPTION_MANUAL
            updated += 1
            log.debug("Updated chart: %s", chart_shape.Name)
        except Exception as e:
            log.warning("Failed to update chart '%s': %s", chart_shape.Name, e)

    log.info("Updated %d chart(s)", updated)
    return updated
