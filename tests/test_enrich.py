"""Tests for report enrichment and WCAG guidance."""
from xlsx_a11y.enrich import get_criterion_text, enrich_finding
from xlsx_a11y.findings import Finding


def test_get_criterion_text():
    text = get_criterion_text("1.1.1")
    assert text is not None
    assert "Non-text Content" in text

    unknown = get_criterion_text("9.9.9")
    assert unknown is None


def test_enrich_finding_subjective():
    f = Finding(
        rule_id="image-alt-missing",
        sc="1.1.1",
        severity="serious",
        location="Sheet1!image[0]",
        description="Embedded image lacks alt text",
    )
    enriched = enrich_finding(f)
    assert "why_unfixable" in enriched
    assert "manual_steps" in enriched
    assert len(enriched["manual_steps"]) > 0
    assert "sc_info" in enriched
