"""Tests for HTML report rendering."""
from xslx_a11y.reports.html import render_html


def test_render_html_structure():
    md = (
        "# Accessibility Audit Report: Test.xlsx\n\n"
        "## Executive Summary\n\n> Clean report.\n\n"
        "## 1. Perceivable\n\n### 1. [CRITICAL] Image missing alt\n\n"
        "- Location: Sheet1!image[0]\n"
    )
    html = render_html(md, theme="ocean")
    assert "<!doctype html>" in html
    assert '<html lang="en">' in html
    assert '<a class="skip" href="#main">' in html
    assert '<nav class="toc"' in html
    assert '<main id="main"' in html
    assert 'class="finding"' in html
    assert "xslx-a11y theme: ocean" in html
