"""Tests for Finding data structure and summary calculations."""
from xlsx_a11y.findings import Finding, SEVERITY_ORDER, BLOCKING, summarize, findings_sorted, finding_to_jsonable


def test_finding_instantiation_and_dict():
    f = Finding(
        rule_id="title-missing",
        sc="2.4.2",
        severity="critical",
        location="core.xml",
        description="Workbook title is missing",
        evidence="wb.properties.title is empty",
        fixable=True,
        fix="Set title from filename",
    )
    assert f.rule_id == "title-missing"
    assert f.sc == "2.4.2"
    assert f.severity == "critical"
    assert f.location == "core.xml"
    assert f.blocking is True
    assert f.fixable is True

    d = f.to_dict()
    assert d["rule_id"] == "title-missing"
    assert d["sc"] == "2.4.2"
    assert d["severity"] == "critical"
    assert d["fixable"] is True
    assert finding_to_jsonable(f) == d


def test_finding_defaults():
    f = Finding(
        rule_id="sheet-name-default",
        sc="2.4.6",
        severity="moderate",
        location="Sheet1",
        description="Sheet has default name",
    )
    assert f.evidence == ""
    assert f.fixable is False
    assert f.fix == ""
    assert f.blocking is False


def test_findings_sorted_and_summary():
    f1 = Finding(
        rule_id="sheet-name-default",
        sc="2.4.6",
        severity="moderate",
        location="Sheet1",
        description="Sheet has default name",
    )
    f2 = Finding(
        rule_id="title-missing",
        sc="2.4.2",
        severity="critical",
        location="core.xml",
        description="Workbook title is missing",
    )
    f3 = Finding(
        rule_id="merged-cell",
        sc="1.3.1",
        severity="serious",
        location="Sheet1!B2:D2",
        description="Merged cells present",
    )
    sorted_f = findings_sorted([f1, f2, f3])
    # Critical should come first, then serious, then moderate
    assert sorted_f[0].rule_id == "title-missing"
    assert sorted_f[1].rule_id == "merged-cell"
    assert sorted_f[2].rule_id == "sheet-name-default"

    summary = summarize([f1, f2, f3])
    assert summary["total"] == 3
    assert summary["blocking"] == 2
    assert summary["pass"] is False
    assert summary["by_severity"]["critical"] == 1
    assert summary["by_severity"]["serious"] == 1
    assert summary["by_severity"]["moderate"] == 1
    assert summary["by_severity"]["minor"] == 0

    clean_summary = summarize([])
    assert clean_summary["total"] == 0
    assert clean_summary["blocking"] == 0
    assert clean_summary["pass"] is True
