
class ModernTheme:
    PRIMARY_COLOR = "#3b82f6"      # Vibrant Indigo-Blue
    PRIMARY_HOVER = "#2563eb"      # Deep Blue
    SECONDARY_COLOR = "#475569"    # Muted Slate
    BG_DARK = "#090d16"            # Deep Space Dark
    BG_LIGHT = "#111827"           # Card Background (Rich Dark Slate)
    BG_CONTROL = "#1f2937"         # Control Background (Soft Gray-Slate)
    TEXT_MAIN = "#f9fafb"          # White Main Text
    TEXT_SUB = "#9ca3af"           # Slate Gray Secondary Text
    BORDER_COLOR = "#374151"       # Subtle Border Gray

    STYLESHEET = f"""
    /* --- Base Application Styles --- */
    QWidget {{
        background-color: {BG_DARK};
        color: {TEXT_MAIN};
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, sans-serif;
        font-size: 13px;
    }}

    /* --- Windows and Dialogs --- */
    QMainWindow, QDialog, QStackedWidget {{
        background-color: {BG_DARK};
    }}

    /* --- Labels --- */
    QLabel {{
        background-color: transparent;
        color: {TEXT_MAIN};
    }}
    
    QLabel#Title {{
        font-size: 24px;
        font-weight: bold;
        color: {TEXT_MAIN};
        margin-bottom: 5px;
    }}

    QLabel#Subtitle {{
        font-size: 14px;
        color: {TEXT_SUB};
        margin-bottom: 15px;
    }}

    /* --- Buttons --- */
    QPushButton {{
        background-color: {BG_CONTROL};
        border: 1px solid {BORDER_COLOR};
        border-radius: 6px;
        color: {TEXT_MAIN};
        padding: 8px 16px;
        font-weight: 600;
        font-size: 13px;
    }}
    
    QPushButton:hover {{
        background-color: {BORDER_COLOR};
        border-color: {SECONDARY_COLOR};
    }}
    
    QPushButton:pressed {{
        background-color: {PRIMARY_HOVER};
        border-color: {PRIMARY_COLOR};
    }}
    
    QPushButton:disabled {{
        background-color: #1e293b;
        color: #4b5563;
        border-color: #1e293b;
    }}
    
    /* Primary Accent Button Style */
    QPushButton[class="primary"] {{
        background-color: {PRIMARY_COLOR};
        border: none;
        color: white;
    }}
    
    QPushButton[class="primary"]:hover {{
        background-color: {PRIMARY_HOVER};
    }}

    /* --- Line and Text Inputs --- */
    QLineEdit, QTextEdit, QPlainTextEdit {{
        background-color: {BG_LIGHT};
        border: 1px solid {BORDER_COLOR};
        border-radius: 6px;
        color: {TEXT_MAIN};
        padding: 8px;
        selection-background-color: {PRIMARY_COLOR};
        selection-color: white;
    }}
    
    QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {{
        border: 1px solid {PRIMARY_COLOR};
        background-color: {BG_DARK};
    }}

    /* --- Scroll Area --- */
    QScrollArea {{
        border: none;
        background-color: transparent;
    }}
    
    QScrollArea QWidget {{
        background-color: transparent;
    }}

    /* --- Custom ScrollBars (Sleek and Modern) --- */
    QScrollBar:vertical {{
        border: none;
        background: transparent;
        width: 8px;
        margin: 0px;
    }}
    
    QScrollBar::handle:vertical {{
        background: {BORDER_COLOR};
        min-height: 25px;
        border-radius: 4px;
    }}
    
    QScrollBar::handle:vertical:hover {{
        background: {PRIMARY_COLOR};
    }}
    
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        border: none;
        background: none;
        height: 0px;
    }}
    
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
        background: none;
    }}

    QScrollBar:horizontal {{
        border: none;
        background: transparent;
        height: 8px;
        margin: 0px;
    }}
    
    QScrollBar::handle:horizontal {{
        background: {BORDER_COLOR};
        min-width: 25px;
        border-radius: 4px;
    }}
    
    QScrollBar::handle:horizontal:hover {{
        background: {PRIMARY_COLOR};
    }}
    
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        border: none;
        background: none;
        width: 0px;
    }}
    
    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
        background: none;
    }}

    /* --- Tab Widget (IDE / Modern dashboard tabs) --- */
    QTabWidget::pane {{
        border: 1px solid {BORDER_COLOR};
        background-color: {BG_LIGHT};
        border-radius: 8px;
        top: -1px;
    }}
    
    QTabBar::tab {{
        background-color: transparent;
        color: {TEXT_SUB};
        padding: 8px 16px;
        border: 1px solid transparent;
        border-bottom: none;
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
        font-weight: 600;
        margin-right: 4px;
    }}
    
    QTabBar::tab:hover {{
        background-color: {BG_CONTROL};
        color: {TEXT_MAIN};
    }}
    
    QTabBar::tab:selected {{
        background-color: {BG_LIGHT};
        color: {PRIMARY_COLOR};
        border-color: {BORDER_COLOR};
    }}

    /* --- List Widgets --- */
    QListWidget {{
        background-color: {BG_LIGHT};
        border: 1px solid {BORDER_COLOR};
        border-radius: 6px;
        color: {TEXT_MAIN};
        padding: 6px;
    }}
    
    QListWidget::item {{
        padding: 8px 12px;
        border-radius: 4px;
        color: {TEXT_MAIN};
        background-color: transparent;
        margin-bottom: 2px;
    }}
    
    QListWidget::item:hover {{
        background-color: {BG_CONTROL};
    }}
    
    QListWidget::item:selected {{
        background-color: {PRIMARY_COLOR};
        color: white;
    }}

    /* --- Table Widgets --- */
    QTableWidget {{
        background-color: {BG_LIGHT};
        border: 1px solid {BORDER_COLOR};
        border-radius: 6px;
        color: {TEXT_MAIN};
        gridline-color: {BORDER_COLOR};
    }}
    
    QTableWidget::item {{
        padding: 6px;
    }}

    QHeaderView::section {{
        background-color: {BG_CONTROL};
        color: {TEXT_SUB};
        padding: 8px;
        font-weight: bold;
        border: 1px solid {BORDER_COLOR};
    }}

    /* --- Combo Box & Spin Box --- */
    QComboBox, QSpinBox {{
        background-color: {BG_CONTROL};
        border: 1px solid {BORDER_COLOR};
        border-radius: 6px;
        padding: 6px 12px;
        color: {TEXT_MAIN};
    }}
    
    QComboBox:focus, QSpinBox:focus {{
        border: 1px solid {PRIMARY_COLOR};
    }}
    
    QComboBox::drop-down {{
        border: none;
        width: 20px;
    }}
    
    QComboBox::down-arrow {{
        image: none;
        border-left: 5px solid transparent;
        border-right: 5px solid transparent;
        border-top: 5px solid {TEXT_SUB};
        width: 0;
        height: 0;
        margin-right: 8px;
    }}
    
    QComboBox QAbstractItemView {{
        background-color: {BG_LIGHT};
        border: 1px solid {BORDER_COLOR};
        border-radius: 6px;
        color: {TEXT_MAIN};
        selection-background-color: {PRIMARY_COLOR};
    }}

    /* --- Splitter --- */
    QSplitter::handle {{
        background-color: {BORDER_COLOR};
        width: 1px;
    }}

    /* --- Progress Bar --- */
    QProgressBar {{
        border: 1px solid {BORDER_COLOR};
        border-radius: 6px;
        text-align: center;
        background-color: {BG_CONTROL};
        height: 12px;
        font-size: 10px;
        font-weight: bold;
    }}
    
    QProgressBar::chunk {{
        background-color: {PRIMARY_COLOR};
        border-radius: 5px;
    }}
    """

