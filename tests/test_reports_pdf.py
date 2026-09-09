"""Tests for accessible PDF report generation."""
from pathlib import Path
import pikepdf
from xlsx_a11y.reports.html import render_html
from xlsx_a11y.reports.pdf import render_pdf


def test_render_pdf_tags_and_lang(tmp_path: Path):
    md = (
        "# Accessibility Audit Report: Financial-Summary.xlsx\n\n"
        "## Executive Summary\n\n"
        "> Evaluation completed with 0 blocking barriers.\n\n"
        "## 1. Perceivable\n\n"
        "_No accessibility barriers detected under Principle Perceivable._\n"
    )
    html_doc = render_html(md, theme="print")
    pdf_out = tmp_path / "report.pdf"

    generated = render_pdf(html_doc, out_path=pdf_out, lang="en", title="Financial Summary Accessibility Report")
    assert generated.exists()
    assert generated.stat().st_size > 0

    # Verify PDF/UA tags with pikepdf
    with pikepdf.open(generated) as pdf:
        assert str(pdf.Root.Lang) == "en"
        assert "/MarkInfo" in pdf.Root
        assert bool(pdf.Root.MarkInfo.Marked) is True
        meta = pdf.open_metadata()
        assert meta["dc:title"] == "Financial Summary Accessibility Report"
