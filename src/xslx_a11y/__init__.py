"""xslx-a11y: Headless, deterministic WCAG 2.1 AA audit and remediation for Excel .xlsx workbooks."""

__version__ = "0.1.0"

# Install accessible drawing hooks for openpyxl so charts preserve alt text
try:
    import openpyxl.drawing.spreadsheet_drawing as _sd

    _orig_chart_frame = getattr(_sd.SpreadsheetDrawing, "_chart_frame")

    def _accessible_chart_frame(self, idx):
        frame = _orig_chart_frame(self, idx)
        if idx <= len(self.charts):
            c = self.charts[idx - 1]
            descr = getattr(c, "_alt_text", None) or getattr(c, "title", None) or ""
            title = getattr(c, "_alt_title", None) or (c.title if isinstance(c.title, str) else "") or "Chart"
            if descr:
                frame.nvGraphicFramePr.cNvPr.descr = str(descr)
            if title:
                frame.nvGraphicFramePr.cNvPr.title = str(title)
        return frame

    setattr(_sd.SpreadsheetDrawing, "_chart_frame", _accessible_chart_frame)
except Exception:
    pass
