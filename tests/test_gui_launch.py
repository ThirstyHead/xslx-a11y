"""Tests for GUI main window initialization."""
from xlsx_a11y.gui.main_window import MainWindow


def test_main_window_init(qtbot):
    window = MainWindow()
    qtbot.addWidget(window)
    assert "xlsx-a11y" in window.windowTitle()
    assert window.table.columnCount() == 5
