"""COM session manager for PowerPoint and Excel lifecycle."""

import gc
import logging
import threading
import time

import win32com.client as win32
import win32con
import win32gui

log = logging.getLogger(__name__)

# Excel calculation modes
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105

# Dialog title to auto-dismiss
_SECURITY_DIALOG_TITLE = "Microsoft PowerPoint Security Notice"


def _auto_dismiss_security_dialog(timeout: float = 15.0):
    """Background thread: find and close the 'Update Links' security dialog.

    PowerPoint shows a blocking 'Security Notice' dialog when opening files
    with OLE links via COM. Neither AutomationSecurity nor DisplayAlerts
    suppress it. This thread polls for the dialog and sends WM_CLOSE.
    """
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        time.sleep(0.3)
        found = []

        def _enum_callback(hwnd, results):
            title = win32gui.GetWindowText(hwnd)
            if _SECURITY_DIALOG_TITLE in title:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                results.append(hwnd)

        win32gui.EnumWindows(_enum_callback, found)
        if found:
            log.debug("Auto-dismissed PowerPoint security dialog")
            return

    log.debug("No security dialog found within %.0fs", timeout)


class Session:
    """Context manager for PowerPoint + Excel COM instances.

    Uses DispatchEx for isolated COM processes — never touches the
    user's already-open PowerPoint or Excel.

    Automatically dismisses the 'Update Links' security dialog that
    PowerPoint shows when opening files with OLE links via COM.

    Usage:
        with Session(pptx_path, excel_path) as s:
            # s.presentation, s.workbook, s.ppt_app, s.excel_app
            ...
    """

    def __init__(
        self, pptx_path: str, excel_path: str | None = None, *, read_only: bool = False
    ):
        self.pptx_path = pptx_path
        self.excel_path = excel_path
        self.read_only = read_only
        self.ppt_app = None
        self.excel_app = None
        self.presentation = None
        self.workbook = None
        self._workbook_cache: dict[str, object] = {}
        self._prev_calc = None

    def __enter__(self):
        # DispatchEx creates a NEW process — won't touch user's open PPT/Excel
        self.ppt_app = win32.DispatchEx("PowerPoint.Application")
        self.ppt_app.DisplayAlerts = 0  # ppAlertsNone

        # Start background thread to auto-dismiss the security dialog
        # that blocks Presentations.Open on files with OLE links
        dismisser = threading.Thread(
            target=_auto_dismiss_security_dialog,
            daemon=True,
        )
        dismisser.start()

        self.presentation = self.ppt_app.Presentations.Open(
            self.pptx_path, ReadOnly=self.read_only, Untitled=False, WithWindow=False
        )
        log.info(
            "Opened presentation: %s (%d slides)",
            self.pptx_path,
            self.presentation.Slides.Count,
        )

        # Excel (if path provided and not read-only mode)
        if self.excel_path and not self.read_only:
            self._init_excel()
            self.workbook = self.get_or_open_workbook(self.excel_path)

        return self

    def _init_excel(self):
        """Initialize Excel COM instance if not already done."""
        if self.excel_app is not None:
            return
        # DispatchEx = new Excel process, isolated from user's Excel
        self.excel_app = win32.DispatchEx("Excel.Application")
        self.excel_app.Visible = False
        self.excel_app.ScreenUpdating = False
        self.excel_app.EnableEvents = False
        self.excel_app.DisplayAlerts = False

        # Set manual calculation for performance
        try:
            self._prev_calc = self.excel_app.Calculation
            self.excel_app.Calculation = XL_CALCULATION_MANUAL
        except Exception:
            self._prev_calc = None

        log.info("Started Excel COM instance")

    def get_or_open_workbook(self, file_path: str):
        """Open a workbook or return a cached one. Lazy-inits Excel if needed."""
        if file_path in self._workbook_cache:
            return self._workbook_cache[file_path]

        self._init_excel()
        # UpdateLinks=0 prevents Excel from auto-refreshing links on open
        wb = self.excel_app.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=False)
        self._workbook_cache[file_path] = wb
        log.info("Opened workbook: %s", file_path)
        return wb

    def save(self):
        """Save the presentation."""
        self.presentation.Save()
        log.info("Saved presentation: %s", self.pptx_path)

    def __exit__(self, exc_type, exc_val, exc_tb):
        # --- Excel cleanup ---
        for path, wb in list(self._workbook_cache.items()):
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        self._workbook_cache.clear()
        self.workbook = None

        if self.excel_app is not None:
            # Restore calculation mode
            if self._prev_calc is not None:
                try:
                    self.excel_app.Calculation = self._prev_calc
                except Exception:
                    pass
            try:
                self.excel_app.Quit()
            except Exception:
                pass
            self.excel_app = None

        # --- PowerPoint cleanup ---
        if self.presentation is not None:
            try:
                self.presentation.Close()
            except Exception:
                pass
            self.presentation = None

        if self.ppt_app is not None:
            try:
                self.ppt_app.Quit()
            except Exception:
                pass
            self.ppt_app = None

        # Release COM pointers and give OS time to clean up processes
        gc.collect()
        time.sleep(0.3)

        return False  # don't suppress exceptions
