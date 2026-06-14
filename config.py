"""
Configuration file for API keys and settings
"""

# Claude API Configuration
CLAUDE_API_KEY = ""  # Set via environment variable or edit here
CLAUDE_MODEL = "claude-3-5-sonnet-20241022"  # Latest Claude model
CLAUDE_MAX_TOKENS = 1500

# Legal Database Sources
LEGAL_DATABASES = {
    "ipc": "Indian Penal Code",
    "contract_act": "Indian Contract Act, 1872",
    "property_act": "Transfer of Property Act, 1882",
    "sale_deed": "Sale Deed Templates & Legal Sections",
    "affidavit": "Affidavit Templates & Requirements",
    "succession": "Indian Succession Act, 1925",
}

# Application Settings
DEBUG = True
LOG_LEVEL = "INFO"
