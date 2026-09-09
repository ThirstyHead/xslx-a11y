"""Application launcher for xslx-a11y desktop GUI."""
import sys
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from xslx_a11y.gui.main_window import MainWindow


def main():
    # Enable High DPI scaling
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName("xslx-a11y")
    app.setOrganizationName("ThirstyHead")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
