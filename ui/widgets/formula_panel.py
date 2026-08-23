from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)


class FormulaPanel(QWidget):
    """
    Panel for selecting formula columns from the target file.

    Provides a searchable checkbox list of all target columns.
    Selected columns will be passed to Writer via session_data['formula_columns'].
    If no columns are selected, Writer will auto-detect formula columns.
    """

    formula_columns_changed = Signal(
        list
    )  # emits list of selected column names

    def __init__(self):
        super().__init__()
        self._checkboxes = {}  # col_name -> QCheckBox
        self._all_columns = []
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)  # Reduced from 6 to tighten vertical spacing

        # Section title
        title = QLabel("Formula Columns (optional)")
        title.setObjectName("sectionLabel")  # Style handled by QSS
        layout.addWidget(title)

        hint = QLabel(
            "Select columns to preserve formulas. Leave empty for auto-detect."
        )
        hint.setObjectName("hintLabel")  # Style handled by QSS
        hint.setWordWrap(True)
        layout.addWidget(hint)

        # Search box
        self.search_input = QLineEdit()
        self.search_input.setObjectName("searchInput")  # Style handled by QSS
        self.search_input.setPlaceholderText("Search columns...")
        self.search_input.setFixedHeight(24)  # Compact height for input
        self.search_input.textChanged.connect(self._on_search)
        layout.addWidget(self.search_input)

        # Scrollable checkbox list
        scroll = QScrollArea()
        scroll.setObjectName("formulaScrollArea")  # Style handled by QSS
        scroll.setWidgetResizable(True)

        self.checkbox_container = QWidget()
        self.checkbox_layout = QVBoxLayout(self.checkbox_container)
        self.checkbox_layout.setContentsMargins(4, 2, 4, 2)  # Compact margins
        self.checkbox_layout.setSpacing(1)  # Tight vertical spacing
        self.checkbox_layout.addStretch()

        scroll.setWidget(self.checkbox_container)
        layout.addWidget(scroll)

        # Select all / Deselect all buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(4)  # Reduced spacing between buttons

        self.select_all_btn = QPushButton("Select All")
        self.select_all_btn.setProperty("flat", True)
        self.select_all_btn.setFixedHeight(24)  # Compact button height
        self.select_all_btn.clicked.connect(self._select_all)
        btn_row.addWidget(self.select_all_btn)

        self.deselect_all_btn = QPushButton("Deselect All")
        self.deselect_all_btn.setProperty("flat", True)
        self.deselect_all_btn.setFixedHeight(24)  # Compact button height
        self.deselect_all_btn.clicked.connect(self._deselect_all)
        btn_row.addWidget(self.deselect_all_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)

    def set_columns(self, columns: list):
        """Populate checkbox list with target file column names."""
        self._all_columns = columns

        # Clear existing checkboxes
        for cb in self._checkboxes.values():
            self.checkbox_layout.removeWidget(cb)
            cb.deleteLater()
        self._checkboxes.clear()

        # Remove stretch if present
        if self.checkbox_layout.count() > 0:
            item = self.checkbox_layout.takeAt(
                self.checkbox_layout.count() - 1
            )
            if item.spacerItem():
                del item

        # Add checkboxes
        for col_name in columns:
            cb = QCheckBox(col_name)
            cb.toggled.connect(self._on_checkbox_toggled)
            self._checkboxes[col_name] = cb
            self.checkbox_layout.addWidget(cb)

        # Re-add stretch
        self.checkbox_layout.addStretch()

    def get_selected_columns(self) -> list:
        """Return list of selected column names."""
        return [
            name for name, cb in self._checkboxes.items() if cb.isChecked()
        ]

    def _on_search(self, text: str):
        """Filter checkboxes by search text."""
        search_lower = text.lower().strip()
        for name, cb in self._checkboxes.items():
            if not search_lower or search_lower in name.lower():
                cb.setVisible(True)
            else:
                cb.setVisible(False)

    def _on_checkbox_toggled(self):
        """Emit signal when selection changes."""
        self.formula_columns_changed.emit(self.get_selected_columns())

    def _select_all(self):
        """Select all visible checkboxes."""
        for cb in self._checkboxes.values():
            if cb.isVisible():
                cb.blockSignals(True)
                cb.setChecked(True)
                cb.blockSignals(False)
        self.formula_columns_changed.emit(self.get_selected_columns())

    def _deselect_all(self):
        """Deselect all checkboxes."""
        for cb in self._checkboxes.values():
            cb.blockSignals(True)
            cb.setChecked(False)
            cb.blockSignals(False)
        self.formula_columns_changed.emit(self.get_selected_columns())
