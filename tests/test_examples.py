"""Regression tests for example workbooks."""
from pathlib import Path
from xslx_a11y.audit import audit_file
from xslx_a11y.immutability import calculate_sha256
from xslx_a11y.remediate import remediate_file

EXAMPLES_DIR = Path(__file__).resolve().parent.parent / "examples"


def test_clean_financial_summary_passes():
    clean_p = EXAMPLES_DIR / "Financial-Summary.xlsx"
    assert clean_p.exists(), f"Fixture missing: {clean_p}"

    result = audit_file(clean_p)
    summary = result["summary"]
    assert summary["pass"] is True
    assert summary["total"] == 0
    assert len(result["findings"]) == 0


def test_barrier_financial_summary_audit_and_remediate(tmp_path: Path):
    test_p = EXAMPLES_DIR / "Financial-Summary-test.xlsx"
    assert test_p.exists(), f"Fixture missing: {test_p}"

    # Audit initial non-compliant state
    result = audit_file(test_p)
    summary = result["summary"]
    assert summary["pass"] is False
    assert summary["total"] > 0

    rule_ids = {f.rule_id for f in result["findings"]}
    assert "title-missing" in rule_ids
    assert "sheet-name-default" in rule_ids
    assert "merged-cell" in rule_ids
    assert "chart-alt-missing" in rule_ids

    # Remediate to temporary destination
    sha_before = calculate_sha256(test_p)
    out_p = tmp_path / "Financial-Summary-remediated.xlsx"
    rem_res = remediate_file(test_p, out_p)

    # Verify original file integrity
    assert calculate_sha256(test_p) == sha_before
    assert rem_res["original_file_immutable"] is True

    # Audit remediated file
    rem_audit = audit_file(out_p)
    rem_rule_ids = {f.rule_id for f in rem_audit["findings"]}

    # Verify deterministic barriers resolved
    assert "title-missing" not in rem_rule_ids
    assert "merged-cell" not in rem_rule_ids
