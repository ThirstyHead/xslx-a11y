"""Canonical W3C WCAG 2.1 metadata, Understanding URLs, and POUR taxonomy for Excel."""

W3C_UNDERSTANDING_BASE = "https://www.w3.org/WAI/WCAG21/Understanding/"
W3C_QUICKREF = "https://www.w3.org/WAI/WCAG21/quickref/?currentsidebar=%23col_customize&levels=aaa"

SC_META = {
    "1.1.1": (
        "Non-text Content",
        "A",
        "1 Perceivable",
        W3C_UNDERSTANDING_BASE + "non-text-content.html",
    ),
    "1.3.1": (
        "Info and Relationships",
        "A",
        "1 Perceivable",
        W3C_UNDERSTANDING_BASE + "info-and-relationships.html",
    ),
    "1.3.2": (
        "Meaningful Sequence",
        "A",
        "1 Perceivable",
        W3C_UNDERSTANDING_BASE + "meaningful-sequence.html",
    ),
    "1.4.3": (
        "Contrast (Minimum)",
        "AA",
        "1 Perceivable",
        W3C_UNDERSTANDING_BASE + "contrast-minimum.html",
    ),
    "2.4.2": (
        "Page Titled",
        "A",
        "2 Operable",
        W3C_UNDERSTANDING_BASE + "page-titled.html",
    ),
    "2.4.4": (
        "Link Purpose (In Context)",
        "A",
        "2 Operable",
        W3C_UNDERSTANDING_BASE + "link-purpose-in-context.html",
    ),
    "2.4.6": (
        "Headings and Labels",
        "AA",
        "2 Operable",
        W3C_UNDERSTANDING_BASE + "headings-and-labels.html",
    ),
}

PRINCIPLES = [
    ("1", "Perceivable"),
    ("2", "Operable"),
    ("3", "Understandable"),
    ("4", "Robust"),
]

POUR_INTROS = {
    "1": (
        "Perceivable content ensures that spreadsheet information can be received by everyone's "
        "senses. Images and charts have text descriptions, cell fonts maintain sufficient contrast against "
        "background fills, and tables have explicit header rows."
    ),
    "2": (
        "Operable spreadsheets let every user navigate efficiently: meaningful document titles identify "
        "the workbook, descriptive sheet tabs organize sections, and hyperlinks clarify destinations."
    ),
    "3": (
        "Understandable content is predictable and clear: sheet layouts avoid confusing blank tabs, "
        "and data structures follow consistent formats."
    ),
    "4": (
        "Robust workbooks use standard SpreadsheetML markup (official Excel Tables, standard drawing metadata) "
        "ensuring compatibility across assistive technologies and spreadsheet viewers."
    ),
}
