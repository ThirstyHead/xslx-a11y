#!/usr/bin/env python3
"""Builds sample Excel workbooks demonstrating compliant and non-compliant WCAG states."""
from pathlib import Path
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

EXAMPLES_DIR = Path(__file__).resolve().parent.parent / "examples"
EXAMPLES_DIR.mkdir(exist_ok=True)


def build_clean_summary():
    """Builds Financial-Summary.xlsx (100% WCAG 2.1 AA compliant)."""
    wb = openpyxl.Workbook()
    wb.properties.title = "Consolidated Financial Summary 2026"

    ws = wb.active
    ws.title = "Revenue_Overview"

    # Tabular data
    headers = ["Department", "Q1_USD", "Q2_USD", "Q3_USD", "Q4_USD", "Total_USD"]
    data = [
        ["Engineering", 120000, 135000, 140000, 155000, 550000],
        ["Marketing", 45000, 50000, 52000, 60000, 207000],
        ["Operations", 80000, 82000, 85000, 90000, 337000],
        ["Sales", 95000, 110000, 115000, 130000, 450000],
    ]

    ws.append(headers)
    for row in data:
        ws.append(row)

    # Style cells with high contrast (black on white or dark green on white)
    for row in ws.iter_rows(min_row=1, max_row=5, min_col=1, max_col=6):
        for cell in row:
            cell.font = Font(name="Calibri", size=11, color="FF000000")
            cell.fill = PatternFill(fill_type=None)

    # Formal Excel Table
    tab = Table(displayName="DeptRevenueTable", ref="A1:F5")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    # Embedded Chart with descriptive title
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Quarterly Departmental Spending"
    chart_data = Reference(ws, min_col=2, min_row=1, max_row=5, max_col=5)
    cats = Reference(ws, min_col=1, min_row=2, max_row=5)
    chart.add_data(chart_data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 10
    chart.width = 16
    ws.add_chart(chart, "H2")

    out_path = EXAMPLES_DIR / "Financial-Summary.xlsx"
    wb.save(out_path)
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
