"""Deterministic remediation engine for Excel workbooks."""
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Tuple, Union
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from xslx_a11y.immutability import (
    assert_not_same_path,
    calculate_sha256,
    get_remediated_path,
    verify_immutability,
)
from xslx_a11y.rules import _is_coord_in_table


def _sanitize_table_name(name: str) -> str:
    """Sanitizes a string to form a valid Excel table displayName (alphanumeric and underscore)."""
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", name)
    if not cleaned or not cleaned[0].isalpha():
        cleaned = f"T_{cleaned}"
    return cleaned[:30]


def remediate_workbook(
    wb: openpyxl.Workbook,
    fallback_title: Optional[str] = None,
) -> Tuple[openpyxl.Workbook, List[str]]:
    """Applies deterministic WCAG 2.1 AA remediations to an in-memory openpyxl Workbook."""
    fixes: List[str] = []

    # 1. Fix missing document title
    current_title = getattr(wb.properties, "title", None)
    if not current_title or not str(current_title).strip():
        title_to_set = fallback_title or "Accessible Workbook"
        wb.properties.title = title_to_set
        fixes.append(f"Set workbook title to '{title_to_set}' in document properties")

    for ws in wb.worksheets:
        # 2. Fix merged cells (unmerge and replicate top-left parent value)
        merged_ranges = list(ws.merged_cells.ranges)
        for rng in merged_ranges:
            coord_str = str(rng.coord)
            min_col, min_row, max_col, max_row = rng.bounds
            top_left_val = ws.cell(row=min_row, column=min_col).value

            ws.unmerge_cells(coord_str)
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    cell = ws.cell(row=r, column=c)
                    if cell.value is None:
                        cell.value = top_left_val
            fixes.append(f"Unmerged range '{coord_str}' on sheet '{ws.title}' and replicated values")

        # 3. Fix missing table headers & convert raw data blocks into official Tables
        tables = list(ws.tables.values())
        for tab in tables:
            if tab.headerRowCount is not None and tab.headerRowCount == 0:
                tab.headerRowCount = 1
                fixes.append(f"Enabled header row on table '{tab.displayName}' on sheet '{ws.title}'")

        # Detect raw rectangular data block not covered by any Table
        non_empty_rows = []
        non_empty_cols = []
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value is not None and str(cell.value).strip() != "":
                    non_empty_rows.append(cell.row)
                    non_empty_cols.append(cell.column)

        if non_empty_rows and non_empty_cols:
            min_r, max_r = min(non_empty_rows), max(non_empty_rows)
            min_c, max_c = min(non_empty_cols), max(non_empty_cols)

            if (max_r - min_r + 1 >= 2) and (max_c - min_c + 1 >= 2):
                covered = any(_is_coord_in_table(min_r, min_c, max_r, max_c, t) for t in tables)
                if not covered:
                    ref_str = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r}"
                    table_name = _sanitize_table_name(f"{ws.title}_Table_{len(tables) + 1}")
                    new_tab = Table(displayName=table_name, ref=ref_str)
                    new_tab.headerRowCount = 1
                    style = TableStyleInfo(
                        name="TableStyleLight1",
                        showFirstColumn=False,
                        showLastColumn=False,
                        showRowStripes=True,
                        showColumnStripes=False,
                    )
                    new_tab.tableStyleInfo = style
                    ws.add_table(new_tab)
                    fixes.append(
                        f"Converted data range '{ref_str}' into official Table '{table_name}' on sheet '{ws.title}'"
                    )

        # 4. Reset initial cell focus to A1 for assistive technologies
        try:
            if hasattr(ws, "views") and hasattr(ws.views, "sheetView") and ws.views.sheetView:
                view = ws.views.sheetView[0]
                view.topLeftCell = "A1"
                if hasattr(view, "selection") and view.selection:
                    view.selection[0].activeCell = "A1"
                    view.selection[0].sqref = "A1"
        except Exception:
            pass

    return wb, fixes


def remediate_file(
    input_path: Union[str, Path],
    out_path: Optional[Union[str, Path]] = None,
) -> Dict[str, Any]:
    """Remediates an Excel file, guaranteeing non-destructive processing of the original."""
    in_p = Path(input_path).resolve()
    if not in_p.exists() or not in_p.is_file():
        raise FileNotFoundError(f"Source file not found: {in_p}")

    out_p = get_remediated_path(in_p, out_path)
    assert_not_same_path(in_p, out_p)

    sha_before = calculate_sha256(in_p)

    # Derive fallback title from sanitized filename stem
    fallback_title = in_p.stem.replace("-", " ").replace("_", " ").title()

    wb = openpyxl.load_workbook(in_p, data_only=False)
    remediated_wb, fixes = remediate_workbook(wb, fallback_title=fallback_title)

    out_p.parent.mkdir(parents=True, exist_ok=True)
    remediated_wb.save(out_p)

    verify_immutability(in_p, sha_before)

    return {
        "input_file": str(in_p),
        "output_file": str(out_p),
        "original_sha256": sha_before,
        "remediated_sha256": calculate_sha256(out_p),
        "original_file_immutable": True,
        "remediations_applied": fixes,
    }
