"""Tests for xlsx-a11y main window two-pane storytelling and report actions."""
from pathlib import Path
from xlsx_a11y.gui.main_window import MainWindow
from xlsx_a11y.gui.models import BatchItem


def test_main_window_after_tree_population(qtbot, tmp_path: Path):
    win = MainWindow()
    qtbot.addWidget(win)

    win.txt_out_dir.setText(str(tmp_path))

    # Create dummy output files
    fixed_xlsx = tmp_path / "sample.fixed.xlsx"
    fixed_xlsx.write_text("dummy fixed xlsx", encoding="utf-8")
    md_report = tmp_path / "sample-a11y-report.md"
    md_report.write_text("# Excel Report\nScore: 92%", encoding="utf-8")

    dummy_src = tmp_path / "sample.xlsx"
    dummy_src.touch()
    item = BatchItem(path=dummy_src)
    item.score = 92.0
    item.status = "Remediated"

    win._populate_after_tree_for_item(item)

    assert win.tree_after.topLevelItemCount() == 1
    doc_node = win.tree_after.topLevelItem(0)
    assert doc_node is not None
    assert "sample.xlsx" in doc_node.text(0)

    # Check child nodes: fixed xlsx and md report
    assert doc_node.childCount() == 2
    child_texts = [c.text(0) for i in range(doc_node.childCount()) if (c := doc_node.child(i)) is not None]
    assert any("Fixed XLSX" in t for t in child_texts)
    assert any("Audit Report" in t for t in child_texts)
