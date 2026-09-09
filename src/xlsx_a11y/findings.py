"""Finding data structures, sorting, and summary metrics.

Delegates core finding structures to engine_a11y.findings.
"""
from typing import Any, Dict
from engine_a11y.findings import (
    Finding,
    Severity,
    SEVERITY_ORDER,
    findings_sorted,
    summarize,
)

BLOCKING = ("critical", "serious")


def finding_to_jsonable(f: Finding) -> Dict[str, Any]:
    return f.to_dict()


__all__ = [
    "Finding",
    "Severity",
    "SEVERITY_ORDER",
    "BLOCKING",
    "findings_sorted",
    "summarize",
    "finding_to_jsonable",
]
