"""Tests for canonical Markdown report generation."""
from xslx_a11y.reports.md import render_md


def test_render_md_basic():
    result = {
        "file": "/path/to/Finances.xlsx",
        "sha256": "abc1234567890def",
        "audited_at": "2026-09-08T12:00:00Z",
        "summary": {"total": 2, "blocking": 1, "pass": False},
        "findings": [
            {
                "rule_id": "title-missing",
                "sc": "2.4.2",
                "severity": "critical",
                "location": "core.xml",
                "description": "Workbook title is missing",
                "evidence": "title=None",
                "fix": "Add title",
            },
            {
                "rule_id": "merged-cell",
                "sc": "1.3.1",
                "severity": "serious",
                "location": "Sheet1!B2:D2",
                "description": "Merged cell range found",
                "evidence": "B2:D2",
                "fix": "Unmerge",
            },
        ],
    }
    md = render_md(result)
    assert "# Accessibility Audit Report" in md
    assert "Finances.xlsx" in md
    assert "POUR" not in md or "Perceivable" in md
    assert "Operable" in md
    assert "Excel Accessibility Assistant Note" in md
    assert "abc1234567890def" in md
