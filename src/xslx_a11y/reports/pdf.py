"""Tagged Accessible PDF report engine derived strictly from HTML5.

Delegates to engine_a11y.reports.pdf.
"""
from pathlib import Path
from typing import Optional, Union
from engine_a11y.reports.pdf import render_pdf as _engine_render_pdf


def render_pdf(
    html_doc: str,
    out_path: Optional[Union[str, Path]] = None,
    lang: str = "en",
    title: Optional[str] = None,
) -> Path:
    """Render HTML report into an accessible, searchable PDF with document structure."""
    return _engine_render_pdf(
        html_doc=html_doc,
        out_path=out_path,
        lang=lang,
        title=title,
    )


__all__ = ["render_pdf"]
