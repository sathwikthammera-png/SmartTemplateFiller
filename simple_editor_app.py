from PySide6.QtWidgets import QApplication
from ui.simple_editor import SimpleDocumentEditor
from ui.styles import ModernTheme
import sys

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(ModernTheme.STYLESHEET)
    
    editor = SimpleDocumentEditor()
    editor.setWindowTitle("Simple Document Editor")
    editor.resize(1200, 700)
    editor.show()
    
    sys.exit(app.exec())
