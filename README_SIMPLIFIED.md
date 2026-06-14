# Document Creator - Simplified & Enhanced

## Overview
This application creates documents with UI options for easy editing. Key features:
- **UI to Word**: Fill forms in UI and export to Word
- **Word to UI**: Upload Word documents and edit in UI
- **No XML complexity**: Simple text-based replacement

## Key Improvements

### ✅ Removed Unnecessary XML Processing
**Before (Complex):**
```python
for p in doc.element.body.iterdescendants(qn('w:p')):
    t_nodes = list(p.iterdescendants(qn('w:t')))
    # Complex buffer manipulation...
```

**After (Simple):**
```python
for para in doc.paragraphs:
    for run in para.runs:
        run.text = run.text.replace("{{key}}", value)
```

### ✅ Added Upload Feature
- Upload existing Word documents
- Automatically extract placeholders
- Edit and re-export

## How to Use

### Option 1: Full Application
```bash
python main.py
```

**Features:**
- Startup screen with multiple options
- Create new documents from scratch
- Use templates with placeholders
- AI legal assistant (optional)
- Live preview
- Export to Word/PDF

### Option 2: Simple Editor
```bash
python simple_editor_app.py
```

**Features:**
- Minimal interface
- Upload Word documents
- Load templates
- Edit fields
- Save as Word

## Placeholder Format

Use either format in your templates:
- `{{placeholder_name}}`
- `<<placeholder_name>>`

**Example Template:**
```
AGREEMENT

This agreement is made between {{party1}} and {{party2}}.

Date: {{date}}
Location: <<location>>

Terms:
1. {{term1}}
2. {{term2}}
```

## Workflow

### Creating from Template
1. Click "New Template" or "Upload Word Doc"
2. Select your .docx file
3. Fill in the form fields
4. See live preview
5. Click "Save as Word"

### Uploading Existing Document
1. Click "Upload Word Doc"
2. Select filled document
3. Edit any fields
4. Save changes

## Technical Details

### Simple Text Replacement
```python
def save_word(self):
    for para in self.doc.paragraphs:
        for key, field in self.fields.items():
            val = field.text()
            if val:
                for run in para.runs:
                    run.text = run.text.replace(f"{{{{{key}}}}}", val)
                    run.text = run.text.replace(f"<<{key}>>", val)
```

### Benefits
- **Easier to understand**: No XML knowledge needed
- **Faster**: Direct text operations
- **More reliable**: Less prone to formatting issues
- **Maintainable**: Simple Python string operations

## File Structure

```
SmartTemplateFillerQt_V4 - Copy/
├── main.py                    # Full application
├── simple_editor_app.py       # Standalone simple editor
├── ui/
│   ├── startup_screen.py      # Welcome screen
│   ├── document_creation.py   # Create from scratch
│   ├── screen1_template.py    # Template selection
│   ├── screen2_form.py        # Form filling (simplified)
│   ├── screen3_export.py      # Export (simplified)
│   ├── simple_editor.py       # Simple editor widget
│   ├── shared_state.py        # Document state
│   └── styles.py              # UI styling
├── legal_ai_assistant.py      # AI features (optional)
└── legal_db_search.py         # Legal database (optional)
```

## Requirements

**Core:**
```
PySide6
python-docx
```

**Optional (for AI features):**
```
anthropic
```

**Optional (for PDF/Print):**
```
pywin32
```

## What Was Removed

### Complex XML Processing
- `docx.oxml.ns.qn` imports
- `iterdescendants` traversal
- XML node manipulation
- Buffer-based text assembly

### Why It Was Removed
- **Overcomplicated**: Most use cases don't need XML-level control
- **Error-prone**: XML structure can vary
- **Hard to maintain**: Requires deep docx knowledge
- **Slower**: More processing overhead

## What Was Added

### Upload Functionality
```python
def upload_word(self):
    """Upload existing Word doc and extract data to UI"""
    path, _ = QFileDialog.getOpenFileName(...)
    self.doc = Document(path)
    placeholders = self.extract_placeholders()
    self.build_form(placeholders)
```

### Simple Editor
- Standalone application
- Minimal dependencies
- Focus on core functionality

## Tips

1. **Use consistent placeholder names**: `{{name}}` not `{{Name}}`
2. **Test templates first**: Load and check placeholders
3. **Keep backups**: Save original templates
4. **Use descriptive names**: `{{client_name}}` not `{{cn}}`

## Troubleshooting

**Placeholders not found?**
- Check format: `{{name}}` or `<<name>>`
- Ensure no extra spaces: `{{ name }}` won't work

**Formatting lost?**
- Simple replacement preserves run-level formatting
- For complex formatting, use the full application

**Can't save?**
- Check file permissions
- Ensure .docx extension
- Close file if already open

## Future Enhancements

- [ ] Support for more placeholder formats
- [ ] Batch processing multiple documents
- [ ] Template library
- [ ] Cloud storage integration
- [ ] Collaborative editing

## Support

For issues or questions:
1. Check this documentation
2. Review example templates
3. Test with simple documents first
