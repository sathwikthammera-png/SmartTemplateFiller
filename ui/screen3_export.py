import os
import tempfile
import time
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QTextEdit, QPushButton,
    QFileDialog, QMessageBox, QProgressDialog, QHBoxLayout, QScrollArea
)
from PySide6.QtCore import Qt, QSize, QThread, Signal, QObject
from PySide6.QtPrintSupport import QPrinter, QPrintPreviewDialog
from PySide6.QtPdf import QPdfDocument
from PySide6.QtGui import QPainter, QPixmap
from copy import deepcopy
from ui.shared_state import get_template

try:
    import win32com.client  # Requires Microsoft Word installed
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None


class PdfWorker(QObject):
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, docx_path, pdf_path):
        super().__init__()
        self.docx_path = docx_path
        self.pdf_path = pdf_path

    def run(self):
        word = None
        try:
            # Initialize COM for this thread
            if pythoncom:
                pythoncom.CoInitialize()
            
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            word_doc = word.Documents.Open(self.docx_path, False, True)
            word_doc.ExportAsFixedFormat(self.pdf_path, 17) # 17 = wdExportFormatPDF
            word_doc.Close(False)
            
            self.finished.emit(self.pdf_path)

        except Exception as e:
            self.error.emit(str(e))
        finally:
            if word:
                try:
                    word.Quit()
                except:
                    pass
            if pythoncom:
                pythoncom.CoUninitialize()


class Screen3(QWidget):
    def __init__(self, stack):
        super().__init__()
        self.stack = stack
        self.filled_data = {}
        
        # Threading and Data
        self.thread = None
        self.worker = None
        self.current_pdf_path = None
        self.pdf_doc = None # Store for print preview lifetime

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop)
        layout.setContentsMargins(32, 32, 32, 32)
        layout.setSpacing(16)
        self.setStyleSheet("background-color: #090d16;")

        title_layout = QHBoxLayout()
        title = QLabel("Final Document Preview")
        title.setObjectName("Title")
        title.setStyleSheet("color: white; font-size: 24px; font-weight: bold; background-color: transparent;")
        title_layout.addWidget(title)
        
        self.status_label = QLabel("") 
        self.status_label.setStyleSheet("color: #94a3b8; font-weight: 500; background-color: transparent;")
        title_layout.addStretch()
        title_layout.addWidget(self.status_label)
        layout.addLayout(title_layout)

        # High Fidelity Preview Area (Scroll Area for images)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        # Deep slate background workspace for paper layout preview
        self.scroll_area.setStyleSheet("background-color: #0f172a; border: 1px solid #374151; border-radius: 8px;")
        
        self.preview_container = QWidget()
        self.preview_container.setStyleSheet("background-color: transparent;")
        self.preview_layout = QVBoxLayout(self.preview_container)
        self.preview_layout.setAlignment(Qt.AlignCenter)
        self.preview_layout.setSpacing(20)
        self.preview_layout.setContentsMargins(30, 30, 30, 30)
        self.scroll_area.setWidget(self.preview_container)
        
        layout.addWidget(self.scroll_area)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)
        btn_layout.addStretch()

        self.back_btn = QPushButton("⬅ Back to Edit")
        self.back_btn.setFixedSize(160, 42)
        # Inherits premium secondary button theme from styles.py
        self.back_btn.clicked.connect(self.go_back)
        btn_layout.addWidget(self.back_btn)

        self.save_btn = QPushButton("💾 Save as .docx")
        self.save_btn.setFixedSize(160, 42)
        self.save_btn.clicked.connect(self.save_docx)
        btn_layout.addWidget(self.save_btn)

        self.print_btn = QPushButton("🖨 Print / Export")
        self.print_btn.setProperty("class", "primary") # Active bright Indigo primary color
        self.print_btn.setFixedSize(160, 42)
        self.print_btn.clicked.connect(self.print_preview)
        btn_layout.addWidget(self.print_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def go_back(self):
        # Cancel any running threads if user goes back
        if self.thread and self.thread.isRunning():
            self.thread.terminate()
        self.stack.setCurrentIndex(1)

    def clear_preview(self):
        while self.preview_layout.count():
            item = self.preview_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def load_preview(self, filled_data):
        self.filled_data = filled_data
        self.clear_preview()
        
        # Show Loading State
        self.status_label.setText("⌛ Rendering visual preview...")
        loading_lbl = QLabel("\n\nGenerating actual document layout...\nThis requires Microsoft Word.")
        loading_lbl.setAlignment(Qt.AlignCenter)
        loading_lbl.setStyleSheet("color: #94a3b8; font-size: 14px;")
        self.preview_layout.addWidget(loading_lbl)

        # 1. Prepare DOCX
        path, original_doc = get_template()
        if not original_doc: return
        
        doc = deepcopy(original_doc)
        self._batch_replace(doc)
        
        tmp_docx = os.path.join(tempfile.gettempdir(), f"preview_{int(time.time())}.docx")
        doc.save(tmp_docx)
        self.current_pdf_path = tmp_docx.replace(".docx", ".pdf")

        # 2. Trigger Background PDF Generation
        if win32com is None:
            self.status_label.setText("⚠ Print/PDF requires Word")
            return

        self.thread = QThread()
        self.worker = PdfWorker(tmp_docx, self.current_pdf_path)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_preview_pdf_ready)
        self.worker.error.connect(self.on_pdf_error)
        
        # Cleanup chain
        self.worker.finished.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)
        self.thread.start()

    def _batch_replace(self, doc):
        """Simple text replacement without XML processing"""
        # Replace in paragraphs
        for para in doc.paragraphs:
            for key, val in self.filled_data.items():
                if val:
                    for run in para.runs:
                        run.text = run.text.replace(f"{{{{{key}}}}}", val)
                        run.text = run.text.replace(f"<<{key}>>", val)
        
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for key, val in self.filled_data.items():
                            if val:
                                for run in para.runs:
                                    run.text = run.text.replace(f"{{{{{key}}}}}", val)
                                    run.text = run.text.replace(f"<<{key}>>", val)

    def on_preview_pdf_ready(self, pdf_path):
        self.status_label.setText("✅ Preview Ready")
        self.clear_preview()
        
        try:
            pdf_doc = QPdfDocument(self)
            pdf_doc.load(pdf_path)
            
            for i in range(pdf_doc.pageCount()):
                # Render page to a high-res image (A4-ish aspect)
                image = pdf_doc.render(i, QSize(800, 1100)) 
                pixmap = QPixmap.fromImage(image)
                page_label = QLabel()
                page_label.setPixmap(pixmap)
                page_label.setFixedSize(pixmap.size())
                page_label.setStyleSheet("background-color: white; border: 1px solid #94a3b8; margin-bottom: 20px;")
                self.preview_layout.addWidget(page_label, alignment=Qt.AlignCenter)
            
        except Exception as e:
            self.on_pdf_error(str(e))



    def save_docx(self):
        path, original_doc = get_template()
        doc = deepcopy(original_doc)
        try:
            self._batch_replace(doc)
            filepath, _ = QFileDialog.getSaveFileName(self, "Save Document", "", "Word Documents (*.docx)")
            if filepath:
                doc.save(filepath)
                QMessageBox.information(self, "Saved", f"Document saved to:\n{filepath}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save:\n{e}")

    def print_preview(self):
        if not self.current_pdf_path or not os.path.exists(self.current_pdf_path):
            QMessageBox.warning(self, "Not Ready", "Please wait for the preview to finish generating.")
            return

        try:
            self.status_label.setText("Opening Print dialog...")
            self.on_pdf_ready(self.current_pdf_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Print error: {e}")

    def on_pdf_ready(self, pdf_path):
        # This shows the actual QPrintPreviewDialog (OS-level print)
        try:
            if not pdf_path or not os.path.exists(pdf_path):
                 raise FileNotFoundError(f"PDF file not found at {pdf_path}")

            self.pdf_doc = QPdfDocument(self)
            status = self.pdf_doc.load(pdf_path)
            
            if self.pdf_doc.pageCount() == 0:
                QMessageBox.warning(self, "Empty Document", "The generated PDF has no pages.")
                return

            def render_pdf(printer):
                painter = QPainter()
                if not painter.begin(printer):
                    return
                
                # Get the actual page rectangle in device pixels
                page_rect = printer.pageRect(QPrinter.DevicePixel)
                target_size = page_rect.size().toSize()
                
                for page in range(self.pdf_doc.pageCount()):
                    # Render the PDF page to a QImage matching the printer's resolution
                    img = self.pdf_doc.render(page, target_size)
                    if not img.isNull():
                        # Draw the image at the top-left (0,0) of the printer page
                        painter.drawImage(0, 0, img)
                    
                    if page < self.pdf_doc.pageCount() - 1:
                        printer.newPage()
                
                painter.end()

            # HighResolution is important for Word-like quality
            printer = QPrinter(QPrinter.HighResolution)
            preview = QPrintPreviewDialog(printer, self)
            preview.setMinimumSize(1000, 800)
            preview.paintRequested.connect(render_pdf)
            preview.exec()
            
            # Reset after dialog is closed to free resources
            self.pdf_doc = None
            
        except Exception as e:
             QMessageBox.critical(self, "Preview Error", f"Failed to initialize print preview: {e}")

    def on_pdf_error(self, err_msg):
        self.status_label.setText("❌ Error")
        QMessageBox.critical(self, "Error", f"Word export failed:\n{err_msg}")
