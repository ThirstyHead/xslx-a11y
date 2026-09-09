"""Canonical Markdown report renderer for xslx-a11y.

Delegates to engine_a11y.reports.md with xlsx document profile and Excel assistant notes.
"""
from typing import Any, Dict, Optional, Set
from engine_a11y.profile import get_xlsx_profile
from engine_a11y.reports.md import render_md as _engine_render_md
from engine_a11y.reports.tone import WORD_ASSISTANT_NOTES
from xslx_a11y import __version__
from xslx_a11y.reports.tone import EXCEL_ASSISTANT_NOTES


def render_md(
    result: Dict[str, Any],
    after_result: Optional[Dict[str, Any]] = None,
    source_path: Optional[str] = None,
    profile: Optional[Any] = None,
    excluded_sc: Optional[Set[str]] = None,
) -> str:
    """Render canonical Markdown accessibility audit report."""
    active_profile = profile or get_xlsx_profile()
    result_copy = dict(result)
    result_copy.setdefault("tool", f"xslx-a11y/{__version__}")
    md = _engine_render_md(
        result=result_copy,
        after_result=after_result,
        source_path=source_path,
        profile=active_profile,
        excluded_sc=excluded_sc,
    )
    for rule_id, word_note in WORD_ASSISTANT_NOTES.items():
        if rule_id in EXCEL_ASSISTANT_NOTES:
            excel_note = EXCEL_ASSISTANT_NOTES[rule_id]
            md = md.replace(
                f"- **Word Accessibility Assistant Note:** {word_note}",
                f"- **Excel Accessibility Assistant Note:** {excel_note}",
            )
    md = md.replace("- **Word Accessibility Assistant Note:**", "- **Excel Accessibility Assistant Note:**")
    return md


__all__ = ["render_md"]
