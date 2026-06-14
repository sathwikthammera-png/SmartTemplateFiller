from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox, QFrame
)
from PySide6.QtCore import Qt, Signal
from docx import Document
from ui.shared_state import set_template, get_template
import os

class DropZone(QFrame):
    file_dropped = Signal(str)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
        self.setStyleSheet("""
            QFrame {
                border: 2px dashed #374151;
                border-radius: 12px;
                background-color: #111827;
            }
            QFrame:hover {
                border-color: #3b82f6;
                background-color: #182235;
            }
        """)
        
        layout = QVBoxLayout(self)
        self.label = QLabel("📂 Drag & Drop Word Template Here\n\n- or -")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("color: #9ca3af; font-size: 16px; font-weight: bold; background-color: transparent;")
        layout.addWidget(self.label)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for f in files:
            if f.endswith(".docx"):
                self.file_dropped.emit(f)
                return

class Screen1(QWidget):
    def __init__(self, stack):
        super().__init__()
        self.stack = stack
        self.next_screen = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(50, 50, 50, 50)
        layout.setSpacing(20)
        self.setStyleSheet("background-color: #090d16;")

        # Title
        title = QLabel("Smart Template Filler")
        title.setObjectName("Title")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: white; font-size: 28px; font-weight: bold; background-color: transparent;")
        layout.addWidget(title)
        
        subtitle = QLabel("Select a Word Template (.docx) to begin auto-filling")
        subtitle.setObjectName("Subtitle")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #9ca3af; font-size: 15px; background-color: transparent;")
        layout.addWidget(subtitle)

        # Drop Zone
        self.drop_zone = DropZone()
        self.drop_zone.setFixedSize(600, 250)
        self.drop_zone.file_dropped.connect(self.load_file)
        layout.addWidget(self.drop_zone, alignment=Qt.AlignCenter)
        
        # Manual Button
        choose_btn = QPushButton("Browse Files...")
        choose_btn.clicked.connect(self.select_template)
        choose_btn.setFixedWidth(200)
        choose_btn.setStyleSheet("""
            QPushButton {
                background-color: #1f2937;
                color: white;
                border: 1px solid #374151;
                border-radius: 6px;
                padding: 10px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #374151;
                border-color: #4b5563;
            }
        """)
        layout.addWidget(choose_btn, alignment=Qt.AlignCenter)

        # Status
        self.selected_path = QLabel("")
        self.selected_path.setAlignment(Qt.AlignCenter)
        self.selected_path.setStyleSheet("color: #10b981; font-weight: bold; font-size: 14px; background-color: transparent;")
        layout.addWidget(self.selected_path)

        # Next Button
        self.next_btn = QPushButton("Next ➡")
        self.next_btn.clicked.connect(self.proceed_to_next)
        self.next_btn.setFixedWidth(150)
        self.next_btn.setEnabled(False) # Disabled until file selected
        self.next_btn.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: white;
                font-weight: bold;
                border-radius: 6px;
                padding: 10px;
            }
            QPushButton:disabled {
                background-color: #1e293b;
                color: #4b5563;
            }
            QPushButton:hover:enabled {
                background-color: #2563eb;
            }
        """)
        layout.addWidget(self.next_btn, alignment=Qt.AlignCenter)

        self.setLayout(layout)

    def set_next_screen(self, screen2):
        self.next_screen = screen2

    def select_template(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Select Template", "", "Word Documents (*.docx)")
        if filepath:
            self.load_file(filepath)

    def load_file(self, filepath):
        try:
            doc = Document(filepath)
            set_template(filepath, doc)
            self.selected_path.setText(f"✅ Selected: {os.path.basename(filepath)}")
            self.drop_zone.setStyleSheet("border: 2px solid #10b981; border-radius: 12px; background-color: #064e3b;")
            self.next_btn.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load document:\n{e}")
            self.selected_path.setText("")
            self.next_btn.setEnabled(False)

    def proceed_to_next(self):
        path, doc = get_template()
        if not doc:
            return
        self.next_screen.build_form_fields()
        self.stack.setCurrentWidget(self.next_screen)
