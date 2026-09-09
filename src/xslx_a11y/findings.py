"""Finding data model and summary utilities for xslx-a11y.

Every rule violation is a Finding. Findings are the currency between
audit (producer), remediate (consumer), report (renderer), and GUI.
"""
from dataclasses import asdict, dataclass
from typing import Any, Dict, List, Optional

SEVERITY_ORDER = {"critical": 0, "serious": 1, "moderate": 2, "minor": 3}
BLOCKING = ("critical", "serious")


@dataclass
class Finding:
    """One accessibility barrier or violation in an Excel workbook.

    Attributes:
        rule_id:     Stable rule identifier (e.g., 'title-missing', 'merged-cell').
        sc:          WCAG success criterion (e.g., '2.4.2', '1.3.1').
        severity:    critical | serious | moderate | minor.
        location:    Human-readable element location (e.g. 'Sheet1!B2:D4', 'core.xml').
        description: Plain text explanation of the issue.
        evidence:    Raw evidence or attribute snippet.
        fixable:     Whether a deterministic programmatic fix exists.
        fix:         Description of the fix applied or suggested.
    """

    rule_id: str
    sc: str
    severity: str
    location: str
    description: str
    evidence: str = ""
    fixable: bool = False
    fix: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @property
    def blocking(self) -> bool:
        return self.severity in BLOCKING

    def sort_key(self):
        return (SEVERITY_ORDER.get(self.severity, 9), self.location, self.rule_id)


def finding_to_jsonable(f: Finding) -> Dict[str, Any]:
    return f.to_dict()


def findings_sorted(findings: List[Finding]) -> List[Finding]:
    return sorted(findings, key=lambda f: f.sort_key())


def summarize(findings: List[Finding]) -> Dict[str, Any]:
    return {
        "total": len(findings),
        "by_severity": {
            s: sum(1 for f in findings if f.severity == s)
            for s in ("critical", "serious", "moderate", "minor")
        },
        "blocking": sum(1 for f in findings if f.blocking),
        "pass": not any(f.blocking for f in findings),
    }
