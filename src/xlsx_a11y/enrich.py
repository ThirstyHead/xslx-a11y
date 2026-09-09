"""Report enrichment with official WCAG criterion text and human triage guidance."""
import json
from pathlib import Path
import re
from typing import Any, Dict, List, Optional

from xlsx_a11y.findings import Finding

BUNDLED_CACHE = Path(__file__).resolve().parent / "sc_cache.json"

RULE_HUMAN_GUIDANCE = {
    "image-alt-missing": {
        "why_unfixable": "Alternative text requires author intent to accurately convey the visual information.",
        "manual_steps": [
            "Right-click the image or shape in Excel.",
            "Select 'View Alt Text' from the context menu.",
            "Enter a clear, concise description (1-2 sentences) or check 'Mark as decorative' if non-informative.",
        ],
    },
    "chart-alt-missing": {
        "why_unfixable": "Data visualizations require author intent to articulate trends, outliers, and key takeaways.",
        "manual_steps": [
            "Right-click the chart boundary.",
            "Select 'Format Chart Area' -> 'Alt Text'.",
            "Provide a descriptive title and detailed summary of the data insights.",
        ],
    },
    "sheet-name-default": {
        "why_unfixable": "Worksheet tab names must describe the specific content or purpose of the sheet.",
        "manual_steps": [
            "Double-click the worksheet tab at the bottom of Excel (or right-click -> Rename).",
            "Enter a descriptive title up to 31 characters without special characters (* / : ? [ ]).",
            "Press Enter to save.",
        ],
    },
    "sheet-empty": {
        "why_unfixable": "Unused sheets clutter screen reader navigation and should be verified by the author.",
        "manual_steps": [
            "Right-click the empty sheet tab at the bottom of the workbook.",
            "Select 'Delete' to remove it, or add relevant tabular content.",
        ],
    },
    "color-contrast": {
        "why_unfixable": "Adjusting brand color palettes and cell highlighting requires designer or author review.",
        "manual_steps": [
            "Select the flagged cell range.",
            "Adjust the font color or cell fill to achieve at least a 4.5:1 contrast ratio (3:1 for large/bold text).",
            "Verify readability using an accessible contrast tool.",
        ],
    },
    "link-text-vague": {
        "why_unfixable": "Link text must explain the destination resource out of context.",
        "manual_steps": [
            "Select the cell containing the hyperlink.",
            "Replace ambiguous phrases (e.g. 'Click here', bare URLs) with descriptive text indicating the target.",
        ],
    },
}

_CACHE_STORE: Optional[Dict[str, str]] = None


def _load_cache() -> Dict[str, str]:
    global _CACHE_STORE
    if _CACHE_STORE is None:
        if BUNDLED_CACHE.exists():
            try:
                with open(BUNDLED_CACHE, "r", encoding="utf-8") as f:
                    _CACHE_STORE = json.load(f)
            except Exception:
                _CACHE_STORE = {}
        else:
            _CACHE_STORE = {}
    return _CACHE_STORE or {}


def get_criterion_text(sc: str) -> Optional[str]:
    """Returns markdown documentation for a given WCAG success criterion number."""
    cache = _load_cache()
    return cache.get(sc)


def parse_criterion(text: Optional[str]) -> Dict[str, Any]:
    """Parses criterion markdown into structured fields."""
    out = {
        "num": "",
        "handle": "",
        "level": "",
        "principle": "",
        "guideline": "",
        "in_brief": "",
        "description": "",
        "intent": "",
    }
    if not text:
        return out

    m = re.match(r"^# (\d+\.\d+\.\d+) (.+)$", text, re.M)
    if m:
        out["num"], out["handle"] = m.group(1), m.group(2).strip()

    for key, pat in (
        ("level", r"^\*\*Level:\*\* (\S+)"),
        ("principle", r"^\*\*Principle:\*\* (.+)$"),
        ("guideline", r"^\*\*Guideline:\*\* (.+)$"),
    ):
        match = re.search(pat, text, re.M)
        if match:
            out[key] = match.group(1).strip()

    def section(name: str) -> str:
        m_sec = re.search(rf"^## {re.escape(name)}\s*\n(.*?)(?=^## |\Z)", text, re.M | re.S)
        return m_sec.group(1).strip() if m_sec else ""

    out["in_brief"] = section("In Brief")
    out["description"] = section("Description")
    out["intent"] = section("Intent")
    return out


def enrich_finding(finding: Finding) -> Dict[str, Any]:
    """Enriches a Finding with official WCAG guidance and human triage instructions."""
    data = finding.to_dict()
    guidance = RULE_HUMAN_GUIDANCE.get(finding.rule_id, {})
    data["why_unfixable"] = guidance.get("why_unfixable", "")
    data["manual_steps"] = guidance.get("manual_steps", [])

    sc_raw = get_criterion_text(finding.sc)
    data["sc_info"] = parse_criterion(sc_raw) if sc_raw else {}
    return data
