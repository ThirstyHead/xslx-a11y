"""Social model tone phrasing, who-benefits mappings, and language guards for Excel.

Delegates core tone checks to engine_a11y.reports.tone.
"""
from engine_a11y.reports.tone import (
    BANNED_PHRASES,
    assert_social_model_language,
)

DISALLOWED_MEDICAL_TERMS = list(BANNED_PHRASES)

WHO_MAP = {
    "1.1.1": (
        "People who are blind, have low vision, or process information through audio receive the complete "
        "message through spoken screen reader narration or braille displays."
    ),
    "1.3.1": (
        "People navigating by screen reader, or individuals who benefit from clean, predictable structure, "
        "can follow table headers, sheet tabs, and cell ranges without disorientation."
    ),
    "1.3.2": (
        "People using screen readers or keyboard grid navigation experience cell values in logical reading "
        "sequence rather than being halted by irregular cell merges."
    ),
    "1.4.3": (
        "People with low vision, color vision differences, or anyone reading under bright ambient light "
        "can comfortably distinguish cell text from its fill background."
    ),
    "2.4.2": (
        "People who navigate documents using title landmarks—such as screen reader users and audiences "
        "scanning lists or window titles—can immediately identify and differentiate workbooks."
    ),
    "2.4.4": (
        "People listening to links or scanning quickly can instantly understand where each hyperlink leads "
        "without needing to decipher ambiguous phrases."
    ),
    "2.4.6": (
        "People skimming sheet tabs using screen reader rotor or hotkeys can immediately locate relevant "
        "sheets rather than navigating generic default names like 'Sheet1'."
    ),
}

RULE_BARRIER_EXPLANATIONS = {
    "title-missing": (
        "The workbook lacks an embedded title in its core metadata properties, meaning file managers and screen "
        "readers identify it only by its raw filename."
    ),
    "sheet-name-default": (
        "The worksheet tab uses a default label (e.g. 'Sheet1'), forcing screen reader users to open and explore "
        "the sheet to deduce its contents."
    ),
    "sheet-empty": (
        "The worksheet contains no data or labels, creating confusing blank stops during screen reader navigation."
    ),
    "merged-cell": (
        "The sheet contains merged cells, disrupting two-dimensional grid navigation and causing screen readers "
        "to misreport row/column headers or skip cell coordinates."
    ),
    "table-header-missing": (
        "The tabular data range is not defined as an official Excel Table with a designated header row, preventing "
        "screen readers from associating data cells with their respective column labels."
    ),
    "image-alt-missing": (
        "This visual asset has no alternative text description, leaving screen reader users unaware of its purpose."
    ),
    "chart-alt-missing": (
        "The embedded chart lacks alternative text or a descriptive title, leaving readers unable to perceive "
        "the visualized trends or key data takeaways."
    ),
    "color-contrast": (
        "The cell font contrast ratio is below the required 4.5:1 threshold (or 3:1 for large/bold text), creating "
        "a visual barrier under typical lighting conditions."
    ),
    "link-text-vague": (
        "The hyperlink uses generic text (such as 'click here') or a raw URL rather than describing the target destination."
    ),
}

EXCEL_ASSISTANT_NOTES = {
    "merged-cell": (
        "Microsoft Excel's built-in Accessibility Assistant routinely overlooks merged cells in title banners or "
        "standalone ranges outside formal tables unless they disrupt active calculation ranges or table objects. "
        "Regardless of whether Excel warns, merged cells break the expected two-dimensional coordinate system in "
        "screen readers, causing assistive technology to skip or misread row/column associations (WCAG 2.1 SC 1.3.1)."
    ),
    "table-header-missing": (
        "Microsoft Excel's built-in Accessibility Assistant only validates formal Table objects (ListObject), "
        "silently passing unformatted tabular cell ranges on single-sheet or smaller workbooks even when they "
        "represent complex data grids. Under WCAG 2.1 SC 1.3.1, every data table requires explicit programmatic "
        "column headers so screen readers can announce header context as users navigate between cells."
    ),
    "sheet-name-default": (
        "Microsoft Excel's built-in Accessibility Assistant checks for default sheet names like 'Sheet1' in recent "
        "Microsoft 365 versions, but silently passes them in older desktop versions and template files. WCAG 2.1 "
        "SC 2.4.6 requires descriptive headings and labels across all user environments."
    ),
    "color-contrast": (
        "Microsoft Excel's built-in Accessibility Assistant inspects color contrast against expected cell fill layers, "
        "and can report persistent warnings when white text is placed on dark table styles if cell-level fills override "
        "theme defaults. Under WCAG 2.1 SC 1.4.3, high-contrast dark text on light backgrounds or explicit high-contrast "
        "cell pairings guarantee readability across all assistive viewers."
    ),
}

__all__ = [
    "WHO_MAP",
    "RULE_BARRIER_EXPLANATIONS",
    "EXCEL_ASSISTANT_NOTES",
    "DISALLOWED_MEDICAL_TERMS",
    "assert_social_model_language",
]
