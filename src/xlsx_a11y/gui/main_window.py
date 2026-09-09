"""Main dashboard window for xlsx-a11y desktop application."""
from pathlib import Path
from typing import List, Optional
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from xlsx_a11y.audit import audit_file
from xlsx_a11y.gui.models import BatchItem, BatchQueue
from xlsx_a11y.gui.theme import APP_STYLESHEET
from xlsx_a11y.gui.triage_dialog import TriageDialog
from xlsx_a11y.gui.worker import BatchWorker
from xlsx_a11y.reports.theme import available_themes


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("xlsx-a11y: Excel WCAG Accessibility Remediation")
        self.resize(960, 680)
        self.setStyleSheet(APP_STYLESHEET)
        self.setAcceptDrops(True)

        self.queue = BatchQueue()
        self.worker: Optional[BatchWorker] = None

        self._init_ui()

    def _init_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(16, 16, 16, 16)

        # 1. Top Action Toolbar
        toolbar_layout = QHBoxLayout()
        self.btn_add_folder = QPushButton("📁 Add Folder...")
        self.btn_add_folder.clicked.connect(self.select_folder)

        self.btn_add_files = QPushButton("📄 Add Files...")
        self.btn_add_files.clicked.connect(self.select_files)

        self.btn_clear = QPushButton("🗑️ Clear List")
        self.btn_clear.clicked.connect(self.clear_files)

        self.btn_triage = QPushButton("🔍 Interactive Triage")
        self.btn_triage.clicked.connect(self.open_triage)

        toolbar_layout.addWidget(self.btn_add_folder)
        toolbar_layout.addWidget(self.btn_add_files)
        toolbar_layout.addWidget(self.btn_clear)
        toolbar_layout.addWidget(self.btn_triage)
        toolbar_layout.addStretch()

        main_layout.addLayout(toolbar_layout)

        # 2. Batch Queue Table
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["File Name", "Sheets", "Status", "Findings", "Score"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)

        main_layout.addWidget(self.table, stretch=1)

        # 3. Settings Box
        settings_group = QGroupBox("Configuration & Output Settings")
        settings_layout = QVBoxLayout(settings_group)
        settings_layout.setSpacing(10)

        # Output directory selector
        out_dir_layout = QHBoxLayout()
        out_dir_layout.addWidget(QLabel("Output Directory:"))
        docs_dir = Path.home() / "Documents"
        default_out = (docs_dir if docs_dir.is_dir() else Path.home()) / "xlsx-a11y-output"
        self.txt_out_dir = QLineEdit(str(default_out))
        self.btn_browse_out = QPushButton("Browse...")
        self.btn_browse_out.clicked.connect(self.browse_output_dir)
        out_dir_layout.addWidget(self.txt_out_dir, stretch=1)
        out_dir_layout.addWidget(self.btn_browse_out)
        settings_layout.addLayout(out_dir_layout)

        # Options row: Formats, Theme, Auto-fix
        opts_layout = QHBoxLayout()
        opts_layout.addWidget(QLabel("Export Formats:"))
        self.chk_md = QCheckBox("Markdown (md)")
        self.chk_md.setChecked(True)
        self.chk_html = QCheckBox("HTML")
        self.chk_html.setChecked(True)
        self.chk_pdf = QCheckBox("PDF")
        self.chk_pdf.setChecked(True)
        self.chk_json = QCheckBox("JSON")
        self.chk_json.setChecked(True)

        opts_layout.addWidget(self.chk_md)
        opts_layout.addWidget(self.chk_html)
        opts_layout.addWidget(self.chk_pdf)
        opts_layout.addWidget(self.chk_json)
        opts_layout.addSpacing(16)

        opts_layout.addWidget(QLabel("Theme:"))
        self.cmb_theme = QComboBox()
        for t in available_themes():
            self.cmb_theme.addItem(t.get("label", t["name"]), t["name"])
        opts_layout.addWidget(self.cmb_theme)
        opts_layout.addSpacing(16)

        self.chk_autofix = QCheckBox("Deterministic Auto-Remediation")
        self.chk_autofix.setChecked(True)
        opts_layout.addWidget(self.chk_autofix)
        opts_layout.addStretch()

        settings_layout.addLayout(opts_layout)
        main_layout.addWidget(settings_group)

        # 4. Progress bar and controls
        progress_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.lbl_status = QLabel("Ready")

        self.btn_run = QPushButton("🚀 Run Batch Audit & Remediation")
        self.btn_run.setObjectName("btn_primary")
        self.btn_run.clicked.connect(self.start_batch)

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.cancel_batch)

        progress_layout.addWidget(self.lbl_status, stretch=1)
        progress_layout.addWidget(self.progress_bar, stretch=2)
        progress_layout.addWidget(self.btn_cancel)
        progress_layout.addWidget(self.btn_run)

        main_layout.addLayout(progress_layout)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        added_count = 0
        for url in urls:
            path = Path(url.toLocalFile())
            if path.is_dir():
                added_count += self.queue.add_directory(path)
            elif path.is_file() and path.suffix.lower() == ".xlsx":
                if self.queue.add_file(path):
                    added_count += 1
        if added_count > 0:
            self.refresh_table()

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Excel Workbooks", "", "Excel Workbooks (*.xlsx)"
        )
        if files:
            for f in files:
                self.queue.add_file(f)
            self.refresh_table()

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder of Excel Workbooks")
        if folder:
            self.queue.add_directory(folder)
            self.refresh_table()

    def clear_files(self):
        self.queue.clear()
        self.refresh_table()

    def browse_output_dir(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if folder:
            self.txt_out_dir.setText(folder)

    def refresh_table(self):
        self.table.setRowCount(len(self.queue.items))
        for row, item in enumerate(self.queue.items):
            self.table.setItem(row, 0, QTableWidgetItem(item.path.name))
            self.table.setItem(row, 1, QTableWidgetItem(str(item.sheet_count) if item.sheet_count else "-"))
            self.table.setItem(row, 2, QTableWidgetItem(item.status))
            self.table.setItem(row, 3, QTableWidgetItem(str(item.findings_count) if item.findings_count else "-"))
            score_text = f"{item.score:.1f}%" if item.score is not None else "-"
            self.table.setItem(row, 4, QTableWidgetItem(score_text))

    def open_triage(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.information(self, "Triage", "Please select a workbook from the list to triage.")
            return

        row = selected_rows[0].row()
        item = self.queue.items[row]
        audit_res = audit_file(item.path)
        findings = audit_res.get("findings", [])

        dialog = TriageDialog(item.path, findings, parent=self)
        dialog.exec()
        self.refresh_table()

    def start_batch(self):
        if len(self.queue.items) == 0:
            QMessageBox.warning(self, "No Files", "Please add at least one .xlsx workbook to process.")
            return

        out_dir = Path(self.txt_out_dir.text().strip())
        formats = []
        if self.chk_md.isChecked():
            formats.append("md")
        if self.chk_html.isChecked():
            formats.append("html")
        if self.chk_pdf.isChecked():
            formats.append("pdf")
        if self.chk_json.isChecked():
            formats.append("json")

        theme = self.cmb_theme.currentData() or "light"
        auto_fix = self.chk_autofix.isChecked()

        self.btn_run.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.progress_bar.setValue(0)
        self.lbl_status.setText("Processing batch...")

        self.worker = BatchWorker(
            items=self.queue.items,
            out_dir=out_dir,
            formats=formats,
            theme=theme,
            auto_fix=auto_fix,
        )
        self.worker.item_started.connect(self._on_item_started)
        self.worker.item_progress.connect(self._on_item_progress)
        self.worker.item_finished.connect(self._on_item_finished)
        self.worker.all_completed.connect(self._on_all_completed)
        self.worker.start()

    def cancel_batch(self):
        if self.worker and self.worker.isRunning():
            self.worker.request_stop()
            self.lbl_status.setText("Stopping...")

    def _on_item_started(self, idx: int, total: int, filename: str):
        pct = int((idx / total) * 100)
        self.progress_bar.setValue(pct)
        self.lbl_status.setText(f"Processing ({idx + 1}/{total}): {filename}")
        self.refresh_table()

    def _on_item_progress(self, idx: int, step_name: str):
        self.lbl_status.setText(f"Item {idx + 1}: {step_name}")

    def _on_item_finished(self, idx: int, status: str, findings: int, score: float):
        self.refresh_table()

    def _on_all_completed(self, processed: int, errors: int):
        self.progress_bar.setValue(100)
        self.btn_run.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.lbl_status.setText(f"Complete! Processed: {processed}, Errors: {errors}")
        self.refresh_table()
        QMessageBox.information(
            self,
            "Batch Complete",
            f"Finished processing {processed} workbook(s) with {errors} error(s).\n\n"
            f"Reports saved to: {self.txt_out_dir.text()}",
        )
