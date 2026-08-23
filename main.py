import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("MATTUpdater")
    app.setOrganizationName("MATTUpdater")

    # Force consistent cross-platform style, so system accent colors
    # (e.g. purple) do not leak into the UI.
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
