"""Audit orchestrator for Excel workbooks."""
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional, Union
import openpyxl

from xlsx_a11y.findings import Finding, findings_sorted, summarize
from xlsx_a11y.immutability import calculate_sha256
from xlsx_a11y.rules import audit_rules


def audit_workbook(
    wb: openpyxl.Workbook,
    file_path: Optional[Union[str, Path]] = None,
) -> Dict[str, Any]:
    """Audits an openpyxl Workbook instance and returns structured findings and summary."""
    findings = audit_rules(wb)
    sorted_findings = findings_sorted(findings)

    sha256 = ""
    resolved_path_str = str(file_path) if file_path else "in-memory"
    if file_path:
        p = Path(file_path)
        if p.exists() and p.is_file():
            sha256 = calculate_sha256(p)
            resolved_path_str = str(p.resolve())

    return {
        "file": resolved_path_str,
        "audited_at": datetime.now(timezone.utc).isoformat(),
        "sha256": sha256,
        "summary": summarize(sorted_findings),
        "findings": sorted_findings,
    }


def audit_file(path: Union[str, Path]) -> Dict[str, Any]:
    """Audits an .xlsx file path without modifying it."""
    p = Path(path).resolve()
    if not p.exists() or not p.is_file():
        raise FileNotFoundError(f"Excel file not found: {p}")

    wb = openpyxl.load_workbook(p, data_only=False)
    return audit_workbook(wb, file_path=p)


def audit_result_to_json(result: Dict[str, Any]) -> str:
    """Serializes audit result dictionary to formatted JSON string."""
    import json
    data = dict(result)
    data["findings"] = [
        f.to_dict() if hasattr(f, "to_dict") else f for f in result.get("findings", [])
    ]
    return json.dumps(data, indent=2)
