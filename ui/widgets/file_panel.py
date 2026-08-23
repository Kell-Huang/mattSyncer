import os

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class FilePanel(QWidget):
    """File selection panel for source and target files.

    Supports multiple source files with adjustable priority order.
    Excel files are read with all sheets merged automatically.
    """

    files_selected = Signal(str, str)  # backward-compatible signal
    source_files_updated = Signal(list, str)  # (source_files, target_path)

    def __init__(self):
        super().__init__()
        self.source_files = []  # List of dicts: {'path': str, 'sheets': 'all'}
        self.target_path = ""
        self._output_format = "xlsx"  # default output format
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(4)  # Keep vertical spacing compact

        # ---- Source files section ----
        src_label = QLabel("Source Files (multiple allowed)")
        src_label.setObjectName("sectionLabel")  # Style handled by QSS
        main_layout.addWidget(src_label)

        # Source file list
        self.src_list = QListWidget()
        self.src_list.setObjectName("sourceFileList")
        self.src_list.setMinimumHeight(50)  # Slightly reduced min height
        self.src_list.setMaximumHeight(100)  # Slightly reduced max height
        main_layout.addWidget(self.src_list)

        # Button row: Add Source Files stays left, other buttons grouped on the right.
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(6)  # Add breathing room between button groups

        btn_add = QPushButton("Add Source Files")
        btn_add.setFixedHeight(24)  # Compact height
        btn_add.clicked.connect(self.add_source_files)

        btn_up = QPushButton("Up")
        btn_up.setProperty("secondary", True)
        btn_up.setFixedHeight(24)
        btn_up.clicked.connect(self.move_up)

        btn_down = QPushButton("Down")
        btn_down.setProperty("secondary", True)
        btn_down.setFixedHeight(24)
        btn_down.clicked.connect(self.move_down)

        btn_remove = QPushButton("Remove")
        btn_remove.setProperty("secondary", True)
        btn_remove.setFixedHeight(24)
        btn_remove.clicked.connect(self.remove_selected)

        btn_layout.addWidget(btn_add)
        btn_layout.addStretch()  # Separate Add Source Files from the rest
        btn_layout.addWidget(btn_up)
        btn_layout.addWidget(btn_down)
        btn_layout.addWidget(btn_remove)
        main_layout.addLayout(btn_layout)

        # ---- Target file row ----
        tgt_label = QLabel("Target File")
        tgt_label.setObjectName("sectionLabel")
        main_layout.addWidget(tgt_label)

        tgt_layout = QHBoxLayout()
        self.tgt_path_label = QLabel("No file selected")
        self.tgt_path_label.setObjectName("pathLabel")
        btn_tgt = QPushButton("Browse Target")
        btn_tgt.setProperty("secondary", True)
        btn_tgt.setFixedHeight(24)  # Compact height
        btn_tgt.clicked.connect(self.browse_target)
        tgt_layout.addWidget(self.tgt_path_label, 1)
        tgt_layout.addWidget(btn_tgt)
        main_layout.addLayout(tgt_layout)

        # ---- Output format row ----
        fmt_layout = QHBoxLayout()
        fmt_label = QLabel("Output")
        fmt_label.setFixedWidth(55)
        fmt_label.setObjectName("sectionLabel")
        self.format_combo = QComboBox()
        self.format_combo.addItems(
            ["XLSX (default)", "CSV (no formulas/format)"]
        )
        self.format_combo.setFixedHeight(26)  # Slightly reduced height
        self.format_combo.setObjectName("formatCombo")
        self.format_combo.currentIndexChanged.connect(self._on_format_changed)
        fmt_layout.addWidget(fmt_label)
        fmt_layout.addWidget(self.format_combo, stretch=1)
        main_layout.addLayout(fmt_layout)

    # ---- Source file management ----
    def add_source_files(self):
        """Open file dialog to select one or more source files."""
        filter_str = "All Supported Files (*.csv *.xlsx *.xls);;CSV Files (*.csv);;Excel Files (*.xlsx *.xls)"
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Source Files", "", filter_str
        )
        if not files:
            return

        self.src_list.setUpdatesEnabled(False)
        added = False
        for f in files:
            if not any(d["path"] == f for d in self.source_files):
                self.source_files.append({"path": f, "sheets": "all"})
                added = True
        if added:
            self._refresh_list()
        self.src_list.setUpdatesEnabled(True)
        if added:
            self._notify_update()

    def move_up(self):
        row = self.src_list.currentRow()
        if row > 0:
            self.source_files[row], self.source_files[row - 1] = (
                self.source_files[row - 1],
                self.source_files[row],
            )
            self._refresh_list()
            self._notify_update()

    def move_down(self):
        row = self.src_list.currentRow()
        if row < len(self.source_files) - 1:
            self.source_files[row], self.source_files[row + 1] = (
                self.source_files[row + 1],
                self.source_files[row],
            )
            self._refresh_list()
            self._notify_update()

    def remove_selected(self):
        row = self.src_list.currentRow()
        if row >= 0:
            del self.source_files[row]
            self._refresh_list()
            self._notify_update()

    def _refresh_list(self):
        self.src_list.clear()
        for f in self.source_files:
            item = QListWidgetItem(os.path.basename(f["path"]))
            self.src_list.addItem(item)

    # ---- Target file ----
    def browse_target(self):
        """Open file dialog to select the target file."""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Target File",
            "",
            "Excel/CSV Files (*.xlsx *.xls *.csv);;All Files (*.*)",
        )
        if path:
            self.target_path = path
            self.tgt_path_label.setText(os.path.basename(path))
            self._notify_update()

    # ---- Output format ----
    def _on_format_changed(self, index):
        if index == 1:
            QMessageBox.information(
                self,
                "CSV Output",
                "CSV format does NOT preserve:\n"
                "  - Excel formulas (will be saved as values)\n"
                "  - Cell formatting (colors, fonts, borders)\n"
                "  - Column widths\n\n"
                "Use CSV only if you need fast processing and don't require formatting.",
            )

    # ---- Signals & helpers ----
    def _notify_update(self):
        if self.source_files and self.target_path:
            self.source_files_updated.emit(self.source_files, self.target_path)
            # Also emit the old signal for backward compatibility
            self.files_selected.emit(
                self.source_files[0]["path"], self.target_path
            )

    @property
    def source_path(self):
        """Return the first source file path for backward compatibility."""
        if self.source_files:
            return self.source_files[0]["path"]
        return ""

    def get_source_files(self):
        return self.source_files

    def get_output_path(self):
        if not self.target_path:
            return ""
        d = os.path.dirname(self.target_path)
        base = os.path.splitext(os.path.basename(self.target_path))[0]
        if self.format_combo.currentIndex() == 1:
            return os.path.join(d, f"{base}_updated.csv")
        return os.path.join(d, f"{base}_updated.xlsx")

    def get_output_format(self):
        return "csv" if self.format_combo.currentIndex() == 1 else "xlsx"

    def validate(self):
        if not self.source_files:
            QMessageBox.warning(
                self, "Error", "Please add at least one source file."
            )
            return False
        if not self.target_path:
            QMessageBox.warning(self, "Error", "Please select a target file.")
            return False
        return True
