from PySide6.QtWidgets import QApplication, QStackedWidget
from ui.screen1_template import Screen1
from ui.screen2_form import Screen2
from ui.screen3_export import Screen3
from ui.startup_screen import StartupScreen
from ui.document_creation import DocumentCreationScreen

from ui.styles import ModernTheme

app = QApplication([])
app.setStyle("Fusion")
app.setStyleSheet(ModernTheme.STYLESHEET)
stack = QStackedWidget()

# Create core screens
screen1 = Screen1(stack)
screen2 = Screen2(stack)
screen3 = Screen3(stack)

# additional screen for creating new documents from scratch
creation_screen = DocumentCreationScreen(stack, screen2)

# Keep existing linkage
screen1.set_next_screen(screen2)
screen2.set_next_screen(screen3)

# Add widgets to stack (startup screen will be the first)
startup = StartupScreen(stack, screen1, screen2, creation_screen=creation_screen)
stack.addWidget(startup)
stack.addWidget(creation_screen)
stack.addWidget(screen1)
stack.addWidget(screen2)
stack.addWidget(screen3)

# Start on the new startup screen (non-destructive)
stack.setCurrentWidget(startup)
stack.setWindowTitle("Smart Template Filler")
stack.resize(1000, 600)
stack.show()

app.exec()