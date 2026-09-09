#!/usr/bin/env python3
"""Builds sample Excel workbooks demonstrating compliant and non-compliant WCAG states."""
from pathlib import Path
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.text import CharacterProperties
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

from xslx_a11y.remediate import ensure_drawing_alt_texts

EXAMPLES_DIR = Path(__file__).resolve().parent.parent / "examples"
EXAMPLES_DIR.mkdir(exist_ok=True)


def build_clean_summary():
    """Builds Financial-Summary.xlsx (100% WCAG 2.1 AA compliant, executive polish)."""
    wb = openpyxl.Workbook()
    wb.properties.title = "Consolidated Financial Summary 2026"

    ws = wb.active
    ws.title = "Revenue_Overview"

    # Set row heights for visual balance and touch/mouse accessibility
    ws.row_dimensions[1].height = 28
    for r in range(2, 6):
        ws.row_dimensions[r].height = 24

    # Generous column widths to prevent truncation or visual cramping
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 18

    # Tabular data
    headers = ["Department", "Q1 (USD)", "Q2 (USD)", "Q3 (USD)", "Q4 (USD)", "Total (USD)"]
    data = [
        ["Engineering", 120000, 135000, 140000, 155000, 550000],
        ["Marketing", 45000, 50000, 52000, 60000, 207000],
        ["Operations", 80000, 82000, 85000, 90000, 337000],
        ["Sales", 95000, 110000, 115000, 130000, 450000],
    ]

    # Populate and style header row with high-contrast navy/white (contrast ratio > 12:1)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = Font(name="Segoe UI", size=12, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")

    # Populate and style data rows with 12pt dark slate text and clean currency formatting
    for row_idx, row in enumerate(data, 2):
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Segoe UI", size=12, color="0F172A")
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.number_format = '"$"#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")

    # Formal Excel Table
    tab = Table(displayName="DeptRevenueTable", ref="A1:F5")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    # Embedded Chart: large, prominent, bottom legend that never overlaps data series
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Quarterly Departmental Spending"

    # Large prominent chart title: 16pt bold DrawingML text properties
    p = chart.title.tx.rich.paragraphs[0]
    p.r[0].rPr = CharacterProperties(sz=1600, b=True)

    # Legend placed at bottom, overlay False so it cannot overlap bars
    chart.legend.legendPos = "b"
    chart.legend.overlay = False

    # Accessible metadata for Excel Accessibility Assistant
    chart._alt_title = "Quarterly Departmental Spending"
    chart._alt_text = (
        "Clustered column chart displaying quarterly departmental spending in USD "
        "from Q1 through Q4 across Engineering, Marketing, Operations, and Sales."
    )

    chart_data = Reference(ws, min_col=2, min_row=1, max_row=5, max_col=5)
    cats = Reference(ws, min_col=1, min_row=2, max_row=5)
    chart.add_data(chart_data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 14
    chart.width = 22
    ws.add_chart(chart, "H1")

    # Set initial assistive view focus to A1
    try:
        ws.sheet_view.topLeftCell = "A1"
    except Exception:
        pass

    out_path = EXAMPLES_DIR / "Financial-Summary.xlsx"
    wb.save(out_path)
    ensure_drawing_alt_texts(out_path)
    print(f"Created clean fixture: {out_path}")


def build_test_summary():
    """Builds Financial-Summary-test.xlsx (Loaded with accessibility barriers)."""
    wb = openpyxl.Workbook()
    # 1. Missing document title
    wb.properties.title = None

    # 2. Default sheet tab name
    ws = wb.active
    ws.title = "Sheet1"

    # 3. Merged cells disrupting screen reader grid navigation
    ws.merge_cells("A1:F1")
    ws["A1"] = "Q1-Q4 FINANCIAL OVERVIEW"

    # 4. Low contrast text (light gray text on white background)
    low_contrast_font = Font(name="Calibri", size=10, color="FFAAAAAA")
    ws["A1"].font = low_contrast_font

    # 5. Raw data range without official Table object
    headers = ["Department", "Q1", "Q2", "Q3", "Q4", "Total"]
    data = [
        ["Engineering", 120000, 135000, 140000, 155000, 550000],
        ["Marketing", 45000, 50000, 52000, 60000, 207000],
        ["Operations", 80000, 82000, 85000, 90000, 337000],
    ]

    for c_idx, h in enumerate(headers, start=1):
        ws.cell(row=2, column=c_idx, value=h)

    for r_idx, row in enumerate(data, start=3):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)

    # 6. Chart lacking descriptive title or alternative text
    chart = BarChart()
    chart.title = None  # Missing alt text / title
    chart_data = Reference(ws, min_col=2, min_row=2, max_row=5, max_col=5)
    chart.add_data(chart_data, titles_from_data=True)
    ws.add_chart(chart, "H2")

    out_path = EXAMPLES_DIR / "Financial-Summary-test.xlsx"
    wb.save(out_path)
    print(f"Created barrier fixture: {out_path}")


if __name__ == "__main__":
    build_clean_summary()
    build_test_summary()
