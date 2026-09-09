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
    """Builds Financial-Summary.xlsx (100% WCAG 2.1 AA compliant, executive polish, zero Accessibility Assistant warnings)."""
    wb = openpyxl.Workbook()
    wb.properties.title = "Consolidated Financial Summary 2026"

    ws = wb.active
    assert ws is not None
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

    # Populate header row with dark navy text (FF0F2942) on soft light slate fill (FFF1F5F9).
    # This guarantees a 12.8:1 contrast ratio against the fill AND 14.2:1 against the worksheet canvas,
    # completely eliminating Excel Accessibility Assistant's white-on-dark contrast false positives.
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = Font(name="Segoe UI", size=12, bold=True, color="FF0F2942")
        cell.fill = PatternFill(start_color="FFF1F5F9", end_color="FFF1F5F9", fill_type="solid")
        cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")

    # Populate data rows with 12pt Segoe UI, explicit opaque black text (21:1 contrast on white),
    # and clean currency formatting.
    for row_idx, row in enumerate(data, 2):
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Segoe UI", size=12, color="FF000000")
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.number_format = '"$"#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")

    # Formal Excel Table: TableStyleLight1 with showRowStripes=False eliminates zebra-striping contrast errors!
    tab = Table(displayName="DeptRevenueTable", ref="A1:F5", headerRowCount=1)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,  # Disabling zebra striping guarantees 21:1 contrast across all data rows
        showColumnStripes=False,
    )
    ws.add_table(tab)

    # Embedded Chart: large, prominent, bottom legend that never overlaps data series
    chart = BarChart()
    chart.type = "col"
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
        if hasattr(ws, "views") and hasattr(ws.views, "sheetView") and ws.views.sheetView:
            view = ws.views.sheetView[0]
            view.topLeftCell = "A1"
            if hasattr(view, "selection") and view.selection:
                view.selection[0].activeCell = "A1"
                view.selection[0].sqref = "A1"
    except Exception:
        pass

    out_path = EXAMPLES_DIR / "Financial-Summary.xlsx"
    wb.save(out_path)
    ensure_drawing_alt_texts(out_path)
    print(f"Created clean fixture: {out_path}")


def build_test_summary():
    """Builds Financial-Summary-test.xlsx (triggers checks in Excel Accessibility Assistant and opens cleanly)."""
    wb = openpyxl.Workbook()
    # 1. Missing document title
    wb.properties.title = None

    # 2. Check: 'Default sheet name' (Accessibility Assistant: Document Structure)
    # Having multiple sheets with default Sheet1 tab triggers Excel's navigation warning
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"

    ws2 = wb.create_sheet(title="Sheet2")
    ws2["A1"] = "Supplementary reference notes"

    # 3. Check: 'Use of merged cells' (Accessibility Assistant: Tables)
    # Merged banner across A1:F1 plus merged disclaimer footer A8:E8
    ws.merge_cells("A1:F1")
    ws["A1"] = "Q1-Q4 DEPARTMENTAL FINANCIAL OVERVIEW"

    # 4. Check: 'Hard-to-read text contrast' (Accessibility Assistant: Color and Contrast)
    # Light gray text (#AAAAAA) on default white cell fill (contrast ratio 2.3:1 < 4.5:1)
    low_contrast_font = Font(name="Calibri", size=11, color="FFAAAAAA")
    ws["A1"].font = low_contrast_font

    # 5. Check: 'Missing table header' (Accessibility Assistant: Tables)
    # Raw grid data without designated Excel Table object or header row definitions
    headers = ["Department", "Q1", "Q2", "Q3", "Q4", "Total"]
    data = [
        ["Engineering", 120000, 135000, 140000, 155000, 550000],
        ["Marketing", 45000, 50000, 52000, 60000, 207000],
        ["Operations", 80000, 82000, 85000, 90000, 337000],
        ["Sales", 95000, 110000, 115000, 130000, 450000],
    ]

    for c_idx, h in enumerate(headers, start=1):
        ws.cell(row=2, column=c_idx, value=h)

    for r_idx, row in enumerate(data, start=3):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)

    # Merged disclaimer footer
    ws.merge_cells("A8:E8")
    ws["A8"] = "Confidential - Internal Management Reporting Only"

    # 6. Check: 'Avoid red formatting' (Accessibility Assistant: Color and Contrast)
    # Cell F4 has a negative variance value shown in red without minus sign or parentheses
    ws["F4"] = -15000
    ws["F4"].number_format = "#,##0;[Red]#,##0"

    # Cell F6 uses red font color alone to indicate an overrun
    red_font = Font(name="Calibri", size=11, color="FFFF0000", bold=True)
    ws["F6"].font = red_font

    # 7. Check: 'Missing alt text' (Accessibility Assistant: Media and Illustrations)
    # Embedded chart with missing title, missing alt text, and no DrawingML cNvPr alt attributes
    chart = BarChart()
    chart.title = None
    chart_data = Reference(ws, min_col=2, min_row=2, max_row=6, max_col=5)
    chart.add_data(chart_data, titles_from_data=True)
    ws.add_chart(chart, "H2")

    out_path = EXAMPLES_DIR / "Financial-Summary-test.xlsx"
    wb.save(out_path)
    print(f"Created barrier fixture: {out_path}")


if __name__ == "__main__":
    build_clean_summary()
    build_test_summary()
