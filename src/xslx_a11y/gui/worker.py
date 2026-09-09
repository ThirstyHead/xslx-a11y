"""Background QThread worker for executing Excel workbook audits and batch remediation."""
from pathlib import Path
from typing import List, Union
import openpyxl
from PySide6.QtCore import QThread, Signal

from xslx_a11y.audit import audit_file, audit_result_to_json
from xslx_a11y.gui.models import BatchItem
from xslx_a11y.remediate import remediate_file
from xslx_a11y.reports.html import render_html
from xslx_a11y.reports.md import render_md
from xslx_a11y.reports.pdf import render_pdf


class BatchWorker(QThread):
    item_started = Signal(int, int, str)  # (item_idx, total_items, file_name)
    item_progress = Signal(int, str)  # (item_idx, step_name)
    item_finished = Signal(int, str, int, float)  # (item_idx, status, findings_count, score)
    all_completed = Signal(int, int)  # (total_processed, total_errors)
    error_occurred = Signal(int, str)  # (item_idx, error_message)

    def __init__(
        self,
        items: List[BatchItem],
        out_dir: Union[Path, str],
        formats: List[str],
        theme: str = "light",
        auto_fix: bool = False,
    ):
        super().__init__()
        self.items = items
        self.out_dir = Path(out_dir)
        self.formats = formats
        self.theme = theme
        self.auto_fix = auto_fix
        self._stop_requested = False

    def request_stop(self):
        self._stop_requested = True

    def run(self):
        try:
            self.out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            self.error_occurred.emit(0, f"Cannot create output directory '{self.out_dir}': {exc}")
            self.all_completed.emit(0, len(self.items))
            return

        total = len(self.items)
        processed = 0
        errors = 0

        for idx, item in enumerate(self.items):
            if self._stop_requested:
                break

            item.status = "Auditing"
            self.item_started.emit(idx, total, item.path.name)

            try:
                # 1. Count sheets
                try:
                    wb = openpyxl.load_workbook(item.path, read_only=True)
                    item.sheet_count = len(wb.sheetnames)
                    wb.close()
                except Exception:
                    item.sheet_count = 0

                # 2. Audit workbook
                self.item_progress.emit(idx, "Auditing workbook...")
                audit_before = audit_file(item.path)
                item.findings_count = len(audit_before.get("findings", []))
                by_sev = audit_before.get("summary", {}).get("by_severity", {})
                item.critical_count = by_sev.get("critical", 0)

                if item.findings_count == 0:
                    item.score = 100.0
                else:
                    penalties = (
                        item.critical_count * 20.0
                        + by_sev.get("serious", 0) * 10.0
                        + by_sev.get("moderate", 0) * 5.0
                        + by_sev.get("minor", 0) * 2.0
                    )
                    item.score = max(0.0, round(100.0 - penalties, 1))

                # 3. Optional deterministic remediation
                audit_after = None
                if self.auto_fix:
                    self.item_progress.emit(idx, "Applying deterministic remediation...")
                    remediated_p = self.out_dir / f"{item.path.stem}-remediated.xlsx"
                    remediate_file(item.path, remediated_p)
                    item.remediated_path = remediated_p
                    audit_after = audit_file(remediated_p)

                # 4. Generate reports
                stem = item.path.stem
                md_text = render_md(audit_before, after_result=audit_after, source_path=str(item.path))

                if "json" in self.formats:
                    self.item_progress.emit(idx, "Writing JSON report...")
                    json_p = self.out_dir / f"{stem}-audit.json"
                    json_p.write_text(audit_result_to_json(audit_before), encoding="utf-8")
                    item.reports["json"] = json_p

                if "md" in self.formats:
                    self.item_progress.emit(idx, "Writing Markdown report...")
                    md_p = self.out_dir / f"{stem}-a11y-report.md"
                    md_p.write_text(md_text, encoding="utf-8")
                    item.reports["md"] = md_p

                if "html" in self.formats:
                    self.item_progress.emit(idx, "Rendering HTML report...")
                    html_doc = render_html(md_text, theme=self.theme)
                    html_p = self.out_dir / f"{stem}-a11y-report.html"
                    html_p.write_text(html_doc, encoding="utf-8")
                    item.reports["html"] = html_p

                if "pdf" in self.formats:
                    self.item_progress.emit(idx, "Rendering accessible PDF report...")
                    html_doc = render_html(md_text, theme=self.theme)
                    pdf_p = self.out_dir / f"{stem}-a11y-report.pdf"
                    render_pdf(html_doc, out_path=pdf_p)
                    item.reports["pdf"] = pdf_p

                item.status = "Completed"
                processed += 1
                self.item_finished.emit(idx, item.status, item.findings_count, item.score or 0.0)

            except Exception as e:
                item.status = "Failed"
                item.error_message = str(e)
                errors += 1
                self.error_occurred.emit(idx, str(e))
                self.item_finished.emit(idx, "Failed", 0, 0.0)

        self.all_completed.emit(processed, errors)
