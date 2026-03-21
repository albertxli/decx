"""Integration tests using real test files — requires COM (Windows + PowerPoint + Excel).

Run with: uv run pytest tests/test_integration.py -k integration
Skip with: uv run pytest tests/ -k "not integration"
"""

import os
import shutil
import tempfile

import pytest

# Mark all tests in this module as integration
pytestmark = pytest.mark.integration

TEST_FILES_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "test_files",
)
TEMPLATE_PPTX = os.path.join(TEST_FILES_DIR, "rpm_2024_market_report_template.pptx")
EXCEL_ARGENTINA = os.path.join(TEST_FILES_DIR, "rpm_tracking_Argentina_(05_07).xlsx")
EXCEL_MEXICO = os.path.join(TEST_FILES_DIR, "rpm_tracking_Mexico_(05_07).xlsx")
EXCEL_US = os.path.join(TEST_FILES_DIR, "rpm_tracking_United_States_(05_07).xlsx")

# Check if test files exist
HAS_TEST_FILES = all(
    os.path.exists(p) for p in [TEMPLATE_PPTX, EXCEL_ARGENTINA, EXCEL_MEXICO, EXCEL_US]
)

# Check if COM is available
try:
    import win32com.client  # noqa: F401

    HAS_COM = True
except ImportError:
    HAS_COM = False

skip_no_files = pytest.mark.skipif(not HAS_TEST_FILES, reason="Test files not found")
skip_no_com = pytest.mark.skipif(not HAS_COM, reason="win32com not available")


def _make_temp_copy(src: str) -> str:
    """Create a temporary copy of a file so we don't modify the original."""
    suffix = os.path.splitext(src)[1]
    fd, dst = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    shutil.copy2(src, dst)
    return dst


@skip_no_files
@skip_no_com
class TestFullPipeline:
    """Test the complete pipeline on real files."""

    def test_process_single_presentation(self):
        """Run full pipeline on template + Argentina data, verify no errors."""
        from decx.config import get_config
        from decx.session import Session
        from decx.shape_finder import build_presentation_inventory
        from decx import (
            linker,
            table_updater,
            delta_updater,
            color_coder,
            chart_updater,
        )

        config = get_config()

        # Work on a temp copy
        pptx_copy = _make_temp_copy(TEMPLATE_PPTX)
        excel_path = os.path.abspath(EXCEL_ARGENTINA)

        try:
            with Session(pptx_copy, excel_path) as session:
                inventory = build_presentation_inventory(session.presentation)
                links = linker.update_links(
                    session, excel_path, config, inventory=inventory
                )
                tables = table_updater.update_tables(
                    session, config, inventory=inventory
                )
                deltas = delta_updater.update_deltas(
                    session, config, inventory=inventory
                )
                colors = color_coder.apply_color_coding(
                    session, config, inventory=inventory
                )
                charts = chart_updater.update_charts(
                    session, excel_path, inventory=inventory
                )
                session.save()

            # Basic sanity: no exceptions raised, counts are non-negative
            assert links >= 0
            assert tables >= 0
            assert deltas >= 0
            assert colors >= 0
            assert charts >= 0
        finally:
            os.unlink(pptx_copy)

    def test_link_update_changes_source(self):
        """Verify OLE links point to new Excel file after Step 1a."""
        from decx.config import get_config
        from decx.session import Session
        from decx.shape_finder import build_presentation_inventory
        from decx import linker

        config = get_config()

        pptx_copy = _make_temp_copy(TEMPLATE_PPTX)
        excel_path = os.path.abspath(EXCEL_MEXICO)

        try:
            with Session(pptx_copy, excel_path) as session:
                inventory = build_presentation_inventory(session.presentation)
                updated = linker.update_links(
                    session, excel_path, config, inventory=inventory
                )

                if updated > 0:
                    ole_shapes = inventory.ole_shapes
                    for _slide, shp in ole_shapes:
                        source = shp.LinkFormat.SourceFullName
                        # The file path portion should now reference the Mexico file
                        assert "Mexico" in source or excel_path in source

                session.save()
        finally:
            os.unlink(pptx_copy)

    def test_batch_mode_three_countries(self):
        """Process template with all 3 Excel files sequentially."""
        from decx.config import get_config
        from decx.session import Session
        from decx.shape_finder import build_presentation_inventory
        from decx import linker, table_updater

        config = get_config()

        excel_files = [EXCEL_ARGENTINA, EXCEL_MEXICO, EXCEL_US]
        temp_copies = []

        try:
            for excel_path in excel_files:
                pptx_copy = _make_temp_copy(TEMPLATE_PPTX)
                temp_copies.append(pptx_copy)
                excel_abs = os.path.abspath(excel_path)

                with Session(pptx_copy, excel_abs) as session:
                    inventory = build_presentation_inventory(session.presentation)
                    linker.update_links(session, excel_abs, config, inventory=inventory)
                    table_updater.update_tables(session, config, inventory=inventory)
                    session.save()

            # All 3 ran without errors
            assert len(temp_copies) == 3
        finally:
            for tc in temp_copies:
                if os.path.exists(tc):
                    os.unlink(tc)

    def test_table_preserves_number_formatting(self):
        """Verify that Excel number formatting (%, decimals) survives the pipeline.

        Regression test for GOTCHAS #16: Range.Value2 strips formatting.
        Slide 3 should contain percentage values in ntbl_ tables, not raw decimals.
        """
        from decx.config import get_config
        from decx.session import Session
        from decx.shape_finder import build_presentation_inventory
        from decx import linker, table_updater

        config = get_config()
        pptx_copy = _make_temp_copy(TEMPLATE_PPTX)
        excel_path = os.path.abspath(EXCEL_ARGENTINA)

        try:
            with Session(pptx_copy, excel_path) as session:
                inventory = build_presentation_inventory(session.presentation)
                linker.update_links(session, excel_path, config, inventory=inventory)
                table_updater.update_tables(session, config, inventory=inventory)

                # Check ntbl_ tables on slide 3 for percentage formatting
                slide3 = session.presentation.Slides(3)
                found_percent = False
                for shp in slide3.Shapes:
                    if shp.HasTable and "ntbl_" in shp.Name:
                        tbl = shp.Table
                        for row in range(1, tbl.Rows.Count + 1):
                            for col in range(1, tbl.Columns.Count + 1):
                                text = tbl.Cell(row, col).Shape.TextFrame.TextRange.Text
                                if "%" in text:
                                    found_percent = True
                                # Raw decimals like "0.05" should NOT appear
                                # (they indicate Value2 was used instead of .Text)
                                if text.startswith("0.0") and "%" not in text:
                                    assert False, (
                                        f"Raw decimal found in ntbl_ table: '{text}'. "
                                        f"This means .Value2 is being used instead of .Text"
                                    )

                # At least some cells should have % formatting
                assert found_percent, (
                    "No percentage values found in slide 3 ntbl_ tables. "
                    "Expected formatted text like '5%', got raw values instead."
                )

                session.save()
        finally:
            os.unlink(pptx_copy)
