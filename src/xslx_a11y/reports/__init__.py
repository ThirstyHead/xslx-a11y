"""Report generation package for xslx-a11y."""
from xslx_a11y.reports.html import render_html
from xslx_a11y.reports.md import render_md
from xslx_a11y.reports.pdf import render_pdf
from xslx_a11y.reports.theme import available_themes, theme_css

__all__ = [
    "render_md",
    "render_html",
    "render_pdf",
    "theme_css",
    "available_themes",
]
