from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class ActionPanel(QWidget):
    preview_clicked = Signal()
    execute_clicked = Signal()

    def __init__(self):
        super().__init__()
        # Single horizontal layout to keep the panel height minimal.
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(6)

        # Push all controls to the right side.
        main_layout.addStretch()

        self.preview_btn = QPushButton("Preview Changes")
        self.preview_btn.setProperty(
            "secondary", True
        )  # Secondary button style
        self.preview_btn.setFixedHeight(24)  # Compact height
        self.preview_btn.clicked.connect(self.preview_clicked.emit)

        self.execute_btn = QPushButton("Execute Update")
        self.execute_btn.setFixedHeight(24)  # Compact height
        self.execute_btn.clicked.connect(self.execute_clicked.emit)

        main_layout.addWidget(self.preview_btn)
        main_layout.addWidget(self.execute_btn)

        # Progress area: label and bar side by side to reduce vertical space.
        progress_widget = QWidget()
        progress_layout = QHBoxLayout(progress_widget)
        progress_layout.setContentsMargins(0, 0, 0, 0)
        progress_layout.setSpacing(4)

        self.progress_label = QLabel("")
        self.progress_label.setObjectName(
            "progressLabel"
        )  # Style handled by QSS
        self.progress_label.setFixedHeight(
            24
        )  # Match button height for alignment

        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")  # Style handled by QSS
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setFixedHeight(24)  # Match button height
        self.progress_bar.setFixedWidth(350)  # Compact width for progress bar

        progress_layout.addWidget(self.progress_label)
        progress_layout.addWidget(self.progress_bar)

        main_layout.addWidget(progress_widget)

    def set_progress(self, step, pct):
        self.progress_label.setText(step)
        self.progress_bar.setValue(pct)

    def set_running(self, running):
        self.preview_btn.setEnabled(not running)
        self.execute_btn.setEnabled(not running)
        if not running:
            self.progress_bar.setValue(0)
            self.progress_label.setText("")
