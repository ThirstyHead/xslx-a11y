"""Tests for GUI main window initialization."""
from xlsx_a11y.gui.main_window import MainWindow


def test_main_window_init(qtbot):
    window = MainWindow()
    qtbot.addWidget(window)
    assert "xlsx-a11y" in window.windowTitle()
    assert window.table.columnCount() == 5
    assert hasattr(window, "pane_before")
    assert hasattr(window, "pane_after")
    assert hasattr(window, "center_bridge")
    assert hasattr(window, "tree_after")
    assert hasattr(window, "btn_remediate_bridge")
    assert hasattr(window, "guide_banner")

