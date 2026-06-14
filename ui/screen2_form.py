from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton,
    QMessageBox, QScrollArea, QFormLayout, QHBoxLayout,
    QTextEdit, QSplitter, QSizePolicy, QFrame, QComboBox,
    QListWidget, QListWidgetItem, QTabWidget, QPlainTextEdit,
    QSpinBox, QProgressBar
)
from PySide6.QtCore import Qt, QTimer, QThread, Signal, QObject
from ui.shared_state import get_template
from ui.docx_tools import display_placeholder, extract_placeholders
import re
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Import legal modules - these should always be available
from legal_ai_assistant import LegalAIAssistant
from legal_db_search import IndianLegalDatabase


class AIWorker(QObject):
    """Worker thread for AI operations"""
    finished = Signal(str)
    error = Signal(str)
    
    def __init__(self, task_type: str, **kwargs):
        super().__init__()
        self.task_type = task_type
        self.kwargs = kwargs
        self.assistant = None
    
    def run(self):
        try:
            self.assistant = LegalAIAssistant()
            
            if self.task_type == "generate_clause":
                result = self.assistant.generate_legal_clause(
                    self.kwargs["placeholder"],
                    self.kwargs["context"],
                    self.kwargs["doc_type"]
                )
            elif self.task_type == "suggest_sections":
                result = "\n".join(self.assistant.suggest_legal_sections(
                    self.kwargs["doc_type"],
                    self.kwargs["filled_data"]
                ))
            elif self.task_type == "review_text":
                review = self.assistant.review_legal_text(
                    self.kwargs["text"],
                    self.kwargs["doc_type"]
                )
                result = f"Compliance: {review['compliance']}\nImprovements: {review['improvements']}\nMissing Elements: {review['missing_elements']}"
            else:
                result = "Unknown task"
            
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(f"Error: {str(e)}")


class Screen2(QWidget):
    def __init__(self, stack):
        super().__init__()
        self.stack = stack
        self.entries = {}
        self.next_screen = None

        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(24, 24, 24, 24)
        main_layout.setSpacing(16)
        self.setStyleSheet("background-color: #090d16;")

        # Splitter for three panels
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setHandleWidth(1)

        # Left Panel: Template Input
        left_panel = QWidget()
        left_panel.setStyleSheet("background-color: #111827; border: 1px solid #374151; border-radius: 8px;")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(16, 16, 16, 16)
        left_layout.setSpacing(12)
        
        lbl_input = QLabel("Questionnaire")
        lbl_input.setStyleSheet("color: white; font-weight: bold; font-size: 16px; background-color: transparent;")
        left_layout.addWidget(lbl_input)

        self.form_status = QLabel("Open a reusable template to see its fields.")
        self.form_status.setStyleSheet("color: #94a3b8; font-size: 12px; background-color: transparent;")
        left_layout.addWidget(self.form_status)

        self.form_widget = QWidget()
        self.form_layout = QFormLayout()
        self.form_layout.setLabelAlignment(Qt.AlignLeft)
        self.form_layout.setVerticalSpacing(16)
        self.form_layout.setHorizontalSpacing(10)
        self.form_widget.setLayout(self.form_layout)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.form_widget)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        left_layout.addWidget(self.scroll_area)

        form_actions = QHBoxLayout()
        self.clear_fields_btn = QPushButton("Clear Fields")
        self.clear_fields_btn.clicked.connect(self.clear_fields)
        form_actions.addWidget(self.clear_fields_btn)
        form_actions.addStretch()
        self.completion_label = QLabel("0% complete")
        self.completion_label.setStyleSheet("color: #10b981; font-size: 12px; background-color: transparent;")
        form_actions.addWidget(self.completion_label)
        left_layout.addLayout(form_actions)

        # Middle Panel: Live Preview
        middle_panel = QWidget()
        middle_panel.setStyleSheet("background-color: #0f172a; border: 1px solid #374151; border-radius: 8px;")
        middle_layout = QVBoxLayout(middle_panel)
        middle_layout.setContentsMargins(16, 16, 16, 16)
        middle_layout.setSpacing(12)
        
        lbl_preview = QLabel("Live Preview")
        lbl_preview.setStyleSheet("color: white; font-weight: bold; font-size: 16px; background-color: transparent;")
        middle_layout.addWidget(lbl_preview)

        self.preview_box = QTextEdit()
        self.preview_box.setReadOnly(True)
        # Premium paper-like sheet floating on slate workspace with shadow border
        self.preview_box.setStyleSheet("""
            QTextEdit {
                background-color: white;
                color: #1f2937;
                border: 1px solid #e5e7eb;
                border-radius: 6px;
                padding: 24px;
            }
        """)
        middle_layout.addWidget(self.preview_box)

        # Right Panel: AI Legal Assistance
        right_panel = QWidget()
        right_panel.setStyleSheet("background-color: #111827; border: 1px solid #374151; border-radius: 8px;")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.setSpacing(12)
        
        lbl_ai = QLabel("🤖 AI Legal Assistant")
        lbl_ai.setStyleSheet("color: white; font-weight: bold; font-size: 16px; background-color: transparent;")
        right_layout.addWidget(lbl_ai)

        # Create tabs for AI features - inherits sleek global tab styles
        self.ai_tabs = QTabWidget()
        self.ai_tabs.setStyleSheet("QTabWidget::pane { border: 1px solid #374151; background-color: #090d16; border-radius: 6px; }")
        
        # Tab 1: Legal Database Search
        search_widget = QWidget()
        search_layout = QVBoxLayout(search_widget)
        search_layout.setContentsMargins(12, 12, 12, 12)
        search_layout.setSpacing(10)
        
        search_label = QLabel("Search Legal Sections:")
        search_label.setStyleSheet("color: #9ca3af; font-size: 12px;")
        search_layout.addWidget(search_label)
        
        self.legal_search_input = QLineEdit()
        self.legal_search_input.setPlaceholderText("Search IPC, Contract Act, Property Act...")
        search_layout.addWidget(self.legal_search_input)
        
        self.search_btn = QPushButton("Search")
        self.search_btn.setProperty("class", "primary") # Blue Primary styling
        self.search_btn.clicked.connect(self.search_legal_database)
        search_layout.addWidget(self.search_btn)
        
        self.search_results = QListWidget()
        self.search_results.itemDoubleClicked.connect(self.insert_legal_section)
        search_layout.addWidget(self.search_results)
        
        self.ai_tabs.addTab(search_widget, "Legal Search")
        
        # Tab 2: Generate Clause
        clause_widget = QWidget()
        clause_layout = QVBoxLayout(clause_widget)
        clause_layout.setContentsMargins(12, 12, 12, 12)
        clause_layout.setSpacing(8)
        
        clause_label = QLabel("Generate Legal Clause:")
        clause_label.setStyleSheet("color: #9ca3af; font-size: 12px;")
        clause_layout.addWidget(clause_label)
        
        self.clause_placeholder = QComboBox()
        self.clause_placeholder.setPlaceholderText("Select placeholder...")
        clause_layout.addWidget(self.clause_placeholder)
        
        context_label = QLabel("Context:")
        context_label.setStyleSheet("color: #9ca3af; font-size: 12px;")
        clause_layout.addWidget(context_label)
        
        self.clause_context = QPlainTextEdit()
        self.clause_context.setPlaceholderText("Enter context for clause generation (e.g. rent amount, duration)...")
        self.clause_context.setMaximumHeight(80)
        clause_layout.addWidget(self.clause_context)
        
        self.generate_clause_btn = QPushButton("✨ Generate Clause")
        self.generate_clause_btn.setProperty("class", "primary")
        self.generate_clause_btn.clicked.connect(self.generate_legal_clause)
        clause_layout.addWidget(self.generate_clause_btn)
        
        self.clause_output = QPlainTextEdit()
        self.clause_output.setReadOnly(True)
        # Modern console output styling
        self.clause_output.setStyleSheet("""
            QPlainTextEdit {
                background-color: #030712;
                color: #38bdf8;
                border: 1px solid #1f2937;
                border-radius: 6px;
                font-family: Consolas, monospace;
                font-size: 12px;
                padding: 8px;
            }
        """)
        clause_layout.addWidget(self.clause_output)
        
        self.ai_tabs.addTab(clause_widget, "Generate Clause")
        
        # Tab 3: Suggest Sections
        suggest_widget = QWidget()
        suggest_layout = QVBoxLayout(suggest_widget)
        suggest_layout.setContentsMargins(12, 12, 12, 12)
        suggest_layout.setSpacing(10)
        
        self.suggest_btn = QPushButton("💡 Suggest Legal Sections")
        self.suggest_btn.clicked.connect(self.suggest_legal_sections)
        suggest_layout.addWidget(self.suggest_btn)
        
        self.suggestions_box = QPlainTextEdit()
        self.suggestions_box.setReadOnly(True)
        self.suggestions_box.setStyleSheet("""
            QPlainTextEdit {
                background-color: #030712;
                color: #fbbf24;
                border: 1px solid #1f2937;
                border-radius: 6px;
                font-family: Consolas, monospace;
                font-size: 12px;
                padding: 8px;
            }
        """)
        suggest_layout.addWidget(self.suggestions_box)
        
        self.ai_tabs.addTab(suggest_widget, "AI Suggestions")
        
        right_layout.addWidget(self.ai_tabs)

        # Add all three panels to splitter
        self.splitter.addWidget(left_panel)
        self.splitter.addWidget(middle_panel)
        self.splitter.addWidget(right_panel)
        self.splitter.setSizes([350, 420, 330])
        main_layout.addWidget(self.splitter)

        # Bottom Navigation Buttons
        nav_layout = QHBoxLayout()
        
        self.back_btn = QPushButton("⬅ Back")
        self.back_btn.setFixedSize(120, 40)
        self.back_btn.clicked.connect(self.go_back)
        nav_layout.addWidget(self.back_btn)
        
        nav_layout.addStretch()

        self.next_btn = QPushButton("Generate ➡")
        self.next_btn.setProperty("class", "primary") # Triggers elegant blue accent style
        self.next_btn.setFixedSize(140, 40)
        self.next_btn.clicked.connect(self.proceed)
        nav_layout.addWidget(self.next_btn)

        main_layout.addLayout(nav_layout)

    def set_next_screen(self, screen3):
        self.next_screen = screen3

    def build_form_fields(self):
        self.entries.clear()
        while self.form_layout.count():
            item = self.form_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        path, doc = get_template()
        if not doc:
            self.form_status.setText("No document loaded.")
            return

        ordered_placeholders = extract_placeholders(doc)

        if not ordered_placeholders:
            self.form_status.setText("This document has no reusable fields.")
            QMessageBox.information(self, "No Placeholders", "No reusable fields were found. Add placeholders like {{client_name}} to make this a template.")
            return

        for placeholder in ordered_placeholders:
            entry = self.create_dynamic_field(placeholder)
            self.entries[placeholder] = entry
            
            lbl = QLabel(f"{display_placeholder(placeholder)}:")
            lbl.setToolTip(f"Template key: {{{{{placeholder}}}}}")
            lbl.setStyleSheet("color: #cbd5e1; font-weight: 600; font-size: 12px; background-color: transparent;")
            self.form_layout.addRow(lbl, entry)

        # Also populate the clause placeholder combobox for AI features
        self.clause_placeholder.clear()
        self.clause_placeholder.addItems(ordered_placeholders)
        self.form_status.setText(f"{len(ordered_placeholders)} reusable field(s) found.")
        
        self.update_preview()

    def create_dynamic_field(self, placeholder):
        long_markers = ("clause", "terms", "condition", "description", "details", "address", "facts", "declaration", "content", "body")
        if any(marker in placeholder.lower() for marker in long_markers):
            entry = QPlainTextEdit()
            entry.setPlaceholderText(f"Enter {display_placeholder(placeholder)}...")
            entry.setMaximumHeight(95)
            entry.setStyleSheet("padding: 8px 12px; font-size: 13px;")
            entry.textChanged.connect(self.update_preview)
            return entry

        entry = QLineEdit()
        entry.setPlaceholderText(f"Enter {display_placeholder(placeholder)}...")
        entry.setStyleSheet("padding: 8px 12px; font-size: 13px;")
        entry.textChanged.connect(self.update_preview)
        return entry

    def field_value(self, field):
        if isinstance(field, QPlainTextEdit):
            return field.toPlainText()
        return field.text()

    def set_field_value(self, field, value):
        if isinstance(field, QPlainTextEdit):
            field.setPlainText(value)
        else:
            field.setText(value)

    def clear_fields(self):
        for field in self.entries.values():
            self.set_field_value(field, "")
        self.update_preview()

    def extract_placeholders(self, doc):
        return extract_placeholders(doc)

    def update_preview(self):
        path, doc = get_template()
        filled_data = {key: field.text() for key, field in self.entries.items()}
        preview_html = ""

        for para in doc.paragraphs:
            line = para.text
            for key in self.entries:
                val = filled_data[key]
                if val:
                    if self.entries[key].hasFocus():
                        line = line.replace(f"{{{{{key}}}}}", f"<span style='background-color:#fef08a; padding: 0 4px;'>{val}</span>")
                        line = line.replace(f"<<{key}>>", f"<span style='background-color:#fef08a; padding: 0 4px;'>{val}</span>")
                    else:
                        line = line.replace(f"{{{{{key}}}}}", val)
                        line = line.replace(f"<<{key}>>", val)
                else:
                    line = line.replace(f"{{{{{key}}}}}", f"<span style='color:#9ca3af; text-decoration:underline;'>{{{{{key}}}}}</span>")
                    line = line.replace(f"<<{key}>>", f"<span style='color:#9ca3af; text-decoration:underline;'>&lt;&lt;{key}&gt;&gt;</span>")
            preview_html += f"<p>{line}</p>"
        # render any tables in the document as HTML to support preview
        for table in doc.tables:
            # determine border style from document table style
            border_style = '1px solid black' if table.style and 'Grid' in table.style.name else 'none'
            preview_html += f"<table border='1' style='border-collapse: collapse; margin-bottom:10px; border:{border_style};'>"
            first = True
            for row in table.rows:
                preview_html += "<tr>"
                for cell in row.cells:
                    cell_text = cell.text
                    tag = 'th' if first else 'td'
                    for key in self.entries:
                        val = filled_data[key]
                        if val:
                            if self.entries[key].hasFocus():
                                cell_text = cell_text.replace(f"{{{{{key}}}}}", f"<span style='background-color:#fef08a; padding: 0 4px;'>{val}</span>")
                                cell_text = cell_text.replace(f"<<{key}>>", f"<span style='background-color:#fef08a; padding: 0 4px;'>{val}</span>")
                            else:
                                cell_text = cell_text.replace(f"{{{{{key}}}}}", val)
                                cell_text = cell_text.replace(f"<<{key}>>", val)
                        else:
                            cell_text = cell_text.replace(f"{{{{{key}}}}}", f"<span style='color:#9ca3af; text-decoration:underline;'>{{{{{key}}}}}</span>")
                            cell_text = cell_text.replace(f"<<{key}>>", f"<span style='color:#9ca3af; text-decoration:underline;'>&lt;&lt;{key}&gt;&gt;</span>")
                    preview_html += f"<{tag}>{cell_text}</{tag}>"
                preview_html += "</tr>"
                first = False
            preview_html += "</table>"

        # Using a fixed bright paper-like styling for the preview inner content
        styled_html = f"""
        <html>
        <head>
            <style>
                body {{
                    font-family: 'Times New Roman', serif;
                    font-size: 16px;
                    line-height: 1.6;
                    color: black;
                    background-color: white;
                    padding: 20px;
                }}
                p {{ margin-bottom: 10px; }}
            </style>
        </head>
        <body>
            {preview_html}
        </body>
        </html>
        """

        self.preview_box.setHtml(styled_html)
        # Debounce the refresh call slightly less aggressive
        # QTimer.singleShot(500, self.refresh_preview) 

    # Removed recursive refresh call causing potential lag, update_preview is sufficient on text change

    def go_back(self):
        self.stack.setCurrentIndex(0)

    def proceed(self):
        filled_data = {key: field.text() for key, field in self.entries.items()}
        self.next_screen.load_preview(filled_data)
        self.stack.setCurrentWidget(self.next_screen)

    # ========== AI & Legal Database Methods ==========

    def search_legal_database(self):
        """Search Indian legal database - works without API key"""
        query = self.legal_search_input.text().strip()
        if not query:
            QMessageBox.warning(self, "Empty Query", "Please enter a search term.")
            return
        
        self.search_results.clear()
        results = IndianLegalDatabase.search(query)
        
        if not results:
            self.search_results.addItem("No results found")
            return
        
        for result in results:
            item_text = f"[{result['source']}] {result['reference']}: {result['content'][:50]}..."
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, result)
            self.search_results.addItem(item)

    def insert_legal_section(self, item):
        """Insert selected legal section into current field"""
        result = item.data(Qt.UserRole)
        if not result:
            return
        
        current_placeholder = self.clause_placeholder.currentText()
        if current_placeholder and current_placeholder in self.entries:
            self.entries[current_placeholder].setText(result['content'])
            self.update_preview()
            QMessageBox.information(self, "Inserted", f"Section inserted into {current_placeholder}")

    def generate_legal_clause(self):
        """Generate legal clause using Claude AI"""
        placeholder = self.clause_placeholder.currentText()
        context = self.clause_context.toPlainText().strip()
        
        if not placeholder or not context:
            QMessageBox.warning(self, "Missing Input", "Please select placeholder and enter context.")
            return
        
        # Show progress
        self.clause_output.setText("🔄 Generating clause...\n\nThis may take a few seconds...")
        self.generate_clause_btn.setEnabled(False)
        
        # Run in thread
        self.ai_thread = QThread()
        self.ai_worker = AIWorker(
            "generate_clause",
            placeholder=placeholder,
            context=context,
            doc_type="legal_document"
        )
        self.ai_worker.moveToThread(self.ai_thread)
        self.ai_thread.started.connect(self.ai_worker.run)
        self.ai_worker.finished.connect(self.on_clause_generated)
        self.ai_worker.error.connect(self.on_ai_error)
        self.ai_worker.finished.connect(self.ai_thread.quit)
        self.ai_worker.error.connect(self.ai_thread.quit)
        self.ai_thread.start()

    def on_clause_generated(self, clause_text):
        """Handle generated clause"""
        self.clause_output.setText(clause_text)
        self.generate_clause_btn.setEnabled(True)
        
        # Ask to insert
        reply = QMessageBox.question(
            self, "Insert Clause?",
            "Do you want to insert this clause into the document?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            placeholder = self.clause_placeholder.currentText()
            if placeholder in self.entries:
                self.entries[placeholder].setText(clause_text)
                self.update_preview()

    def suggest_legal_sections(self):
        """Get AI suggestions for legal sections"""
        filled_data = {key: field.text() for key, field in self.entries.items()}
        
        self.suggestions_box.setText("💭 Analyzing document...\n\nGenerating suggestions...")
        self.suggest_btn.setEnabled(False)
        
        # Run in thread
        self.ai_thread = QThread()
        self.ai_worker = AIWorker(
            "suggest_sections",
            doc_type="legal_document",
            filled_data=filled_data
        )
        self.ai_worker.moveToThread(self.ai_thread)
        self.ai_thread.started.connect(self.ai_worker.run)
        self.ai_worker.finished.connect(self.on_suggestions_ready)
        self.ai_worker.error.connect(self.on_ai_error)
        self.ai_worker.finished.connect(self.ai_thread.quit)
        self.ai_worker.error.connect(self.ai_thread.quit)
        self.ai_thread.start()

    def on_suggestions_ready(self, suggestions_text):
        """Handle AI suggestions"""
        self.suggestions_box.setText(suggestions_text)
        self.suggest_btn.setEnabled(True)

    def on_ai_error(self, error_msg):
        """Handle AI errors"""
        self.suggest_btn.setEnabled(True)
        self.generate_clause_btn.setEnabled(True)
        logger.error(f"AI Error: {error_msg}")
        QMessageBox.critical(self, "AI Error", error_msg)
