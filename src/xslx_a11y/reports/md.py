"""Markdown report generator: audit result JSON -> canonical accessibility report (source of truth)."""
from typing import Any, Dict, List, Optional
from xslx_a11y import __version__
from xslx_a11y.reports.meta import PRINCIPLES, POUR_INTROS, SC_META, W3C_QUICKREF
from xslx_a11y.reports.stats import compute_progress_stats
from xslx_a11y.reports.tone import (
    EXCEL_ASSISTANT_NOTES,
    RULE_BARRIER_EXPLANATIONS,
    WHO_MAP,
    assert_social_model_language,
)


def render_md(
    result: Dict[str, Any],
    after_result: Optional[Dict[str, Any]] = None,
    source_path: Optional[str] = None,
) -> str:
    """Renders canonical Markdown accessibility audit report."""
    src = source_path or result.get("file", "workbook.xlsx")
    summary = result.get("summary", {})
    findings = result.get("findings", [])
    after_summary = after_result.get("summary") if after_result else None
    stats = compute_progress_stats(summary, after_summary)

    lines: List[str] = []
    lines.append(f"# Accessibility Audit Report: {src}")
    lines.append("")
    lines.append(f"- **Workbook Evaluated:** `{src}`")
    lines.append(f"- **Original File SHA-256:** `{result.get('sha256', 'n/a')}`")
    lines.append("- **Integrity Verification:** Immutable (original workbook is strictly read-only and never modified in place)")
    lines.append(f"- **Audit Standard:** [WCAG 2.1 Levels A & AA]({W3C_QUICKREF})")
    lines.append(f"- **Evaluated At:** {result.get('audited_at', 'n/a')}")
    lines.append(f"- **Audit Tool:** `xslx-a11y/{__version__}`")
    lines.append(f"- **Compliance Status:** **{stats['compliance_verdict']}**")
    lines.append("")

    # Summary Progress Banner
    lines.append("## Executive Summary")
    lines.append("")
    if stats["mode"] == "remediated":
        lines.append(
            f"> **Remediation Progress:** Resolved **{stats['resolved_blocking']}** of "
            f"**{stats['before_blocking']}** blocking accessibility barriers "
            f"(**{stats['improvement_rate_pct']}% improvement**). "
            f"Remaining barriers: **{stats['after_blocking']}**."
        )
        lines.append(">")
        lines.append(
            f"> **Integrity Guarantee:** The original source workbook was verified untouched and preserved byte-for-byte (SHA-256: `{result.get('sha256')}`)."
        )
    else:
        status_word = "clean and passes" if summary.get("pass") else "contains accessibility barriers that need attention"
        lines.append(
            f"> **Assessment:** This workbook {status_word}. "
            f"A total of **{summary.get('total', 0)}** items were cataloged "
            f"(**{summary.get('blocking', 0)}** blocking WCAG 2.1 AA conformance)."
        )
    lines.append("")

    # Group findings by Principle (POUR)
    findings_by_principle: Dict[str, List[Dict[str, Any]]] = {"1": [], "2": [], "3": [], "4": []}
    for f in findings:
        f_dict = f.to_dict() if hasattr(f, "to_dict") else dict(f)
        sc = f_dict.get("sc", "1.1.1")
        principle_key = sc.split(".")[0]
        if principle_key in findings_by_principle:
            findings_by_principle[principle_key].append(f_dict)

    # Render each POUR section
    for p_key, p_name in PRINCIPLES:
        p_findings = findings_by_principle[p_key]
        lines.append(f"## {p_key}. {p_name}")
        lines.append("")
        lines.append(POUR_INTROS[p_key])
        lines.append("")

        if not p_findings:
            lines.append(f"_No accessibility barriers detected under Principle {p_name}._")
            lines.append("")
            continue

        for idx, f in enumerate(p_findings, start=1):
            sc = f["sc"]
            sc_info = SC_META.get(sc, ("Accessibility Requirement", "A", f"{p_key} {p_name}", W3C_QUICKREF))
            sc_title, sc_level, _, sc_url = sc_info
            rule_id = f.get("rule_id", "custom-rule")
            sev = f.get("severity", "moderate").upper()

            lines.append(f"### {idx}. [{sev}] {f['description']}")
            lines.append("")
            lines.append(f"- **Success Criterion:** [WCAG 2.1 SC {sc}: {sc_title} (Level {sc_level})]({sc_url})")
            lines.append(f"- **Location:** `{f.get('location', 'Unknown')}`")
            lines.append(f"- **Barrier Detected:** {RULE_BARRIER_EXPLANATIONS.get(rule_id, f.get('description'))}")
            lines.append(f"- **Recommended Remediation:** {f.get('fix', 'Inspect and resolve.')}")
            lines.append(f"- **Who Benefits:** {WHO_MAP.get(sc, 'All readers gain improved access.')}")
            lines.append(f"- **Technical Evidence:** `{f.get('evidence', '')}`")

            note = EXCEL_ASSISTANT_NOTES.get(rule_id)
            if note:
                lines.append(f"- **Excel Accessibility Assistant Note:** {note}")

            why_unfixable = f.get("why_unfixable")
            if why_unfixable:
                lines.append(f"- **Why Software Cannot Automatically Fix This:** {why_unfixable}")

            manual_steps = f.get("manual_steps")
            if manual_steps:
                lines.append("- **Manual Remediation Steps:**")
                for step in manual_steps:
                    lines.append(f"  1. {step}")

            lines.append("")

    rendered = "\n".join(lines)
    assert_social_model_language(rendered)
    return rendered
