# SmartTemplateFillerQt_V4 - AI Enhanced Legal Document Automation

## New Features: AI & Legal Section Integration

Your tool has been enhanced with powerful AI and legal database capabilities for creating professional legal documents like affidavits, sale deeds, and contracts.

### 🤖 New AI Features

#### 1. **Legal Database Search**
- Search through comprehensive Indian legal sections
- Covers: IPC, Contract Act, Transfer of Property Act, Sale Deeds, Affidavits, Succession Act
- Double-click results to insert into your document
- **Tab**: "Legal Search"

#### 2. **Generate Legal Clauses**
- Use Claude AI to auto-generate professionally drafted legal clauses
- Provide context and select placeholder to populate
- AI understands Indian legal language and requirements
- **Tab**: "Generate Clause"

#### 3. **AI Suggestions**
- Analyzes your document and suggests relevant legal sections
- Recommends Indian acts, precedents, and clauses
- Helps ensure compliance and completeness
- **Tab**: "AI Suggestions"

#### 4. **Legal Text Review** (Coming Soon)
- Review drafted legal text for compliance
- Identify missing legal elements
- Get improvement suggestions

---

## 🔧 Setup & Configuration

### Step 1: Get Claude API Key

1. Go to [Anthropic Console](https://console.anthropic.com/)
2. Sign up or log in to your account
3. Navigate to **API Keys**
4. Click **Create Key** and copy it

### Step 2: Set Environment Variable

#### **Windows (PowerShell)**:
```powershell
$env:ANTHROPIC_API_KEY = "your-api-key-here"
```

#### **Windows (Command Prompt)**:
```cmd
set ANTHROPIC_API_KEY=your-api-key-here
```

#### **Windows (Permanent)** - Set System Environment Variable:
1. Press `Win + X` → Select "System"
2. Click "Advanced system settings"
3. Click "Environment Variables"
4. Click "New" under "User variables"
5. Variable name: `ANTHROPIC_API_KEY`
6. Variable value: `your-api-key-here`
7. Click OK and restart VS Code

#### **MacOS/Linux**:
```bash
export ANTHROPIC_API_KEY="your-api-key-here"
```

### Step 3: Test the Connection

The app will automatically test your API key when you first use the AI features. You'll see a status message indicating success or configuration issues.

---

## 📖 How to Use

### **Workflow with AI Assistance**

1. **Load Template** (Screen 1)
   - Select your legal template (affidavit, sale deed, etc.)

2. **Fill Form with AI Help** (Screen 2)
   - **Left Panel**: Fill in placeholders manually or use AI
   - **Middle Panel**: Real-time preview of your document
   - **Right Panel**: AI Assistant with three tabs

### **Using the AI Panel**

#### **Tab 1: Legal Search**
```
1. Enter keyword (e.g., "sale", "cheque", "contract")
2. Click "Search"
3. Review results
4. Double-click to insert into document
```

#### **Tab 2: Generate Clause**
```
1. Select placeholder from dropdown (e.g., "consideration", "notices")
2. Enter context in text box
   - Example: "Sale of residential property in Mumbai"
   - Example: "Party A is individual, Party B is company"
3. Click "✨ Generate Clause"
4. AI generates professional legal text
5. Optionally insert into document
```

#### **Tab 3: AI Suggestions**
```
1. Fill in some document details
2. Click "💡 Suggest Legal Sections"
3. AI recommends relevant Indian legal sections
4. Use suggestions to enhance your document
```

---

## 📋 Supported Legal Areas

### The embedded legal database includes:

- **IPC (Indian Penal Code, 1860)**: Criminal law sections
- **Indian Contract Act, 1872**: Contract formation and law
- **Transfer of Property Act, 1882**: Property transactions
- **Sale Deed Templates**: Specific clauses and elements
- **Affidavit Guidelines**: Format and requirements
- **Succession Act, 1925**: Inheritance and succession
- **General Clauses**: Boilerplate legal language

### Example Document Types:
- Affidavits (sworn statements)
- Sale Deeds (property transfer)
- Contracts
- Agreements
- Legal Notices
- Powers of Attorney

---

## 💰 Pricing (Claude API)

**Free Tier Available**: 
- $5 free trial credit to test
- Enough for 100+ API calls

**Pay-as-You-Go**:
- Standard rates starting at ~$0.08 per 1K input tokens
- Typical legal clause generation: $0.10-0.50 per call

---

## ⚙️ Technical Details

### Files Added:

1. **`config.py`** - Configuration and API settings
   - Store API key reference (uses environment variable)
   - Legal database settings
   - Application configuration

2. **`legal_ai_assistant.py`** - Claude AI integration
   - `LegalAIAssistant` class for all AI operations
   - Methods: generate_legal_clause, suggest_legal_sections, review_legal_text, etc.
   - Handles API calls and responses

3. **`legal_db_search.py`** - Legal database search
   - `IndianLegalDatabase` class with searchable sections
   - Full Indian legal database embedded
   - Fast local search without external API calls

### Modified Files:

4. **`ui/screen2_form.py`** - Enhanced with AI panel
   - Added right panel with tabbed interface
   - `AIWorker` class for threading
   - All AI interaction methods
   - Integrates with existing form filling

---

## 🐛 Troubleshooting

### "Claude API not configured" Error
**Solution**: Set `ANTHROPIC_API_KEY` environment variable as shown in Setup section

### "No results found" in Legal Search
- Try different keywords
- Use broader terms (e.g., "contract" instead of "specific contract type")
- Check placeholder name relevance

### AI response is slow
- First call may be slower (~5-10 seconds)
- Subsequent calls are cached
- This is normal for Claude API
- Wait for "Generating..." message to complete

### Out of free credits
- Upgrade to paid plan on Anthropic Console
- $5 free trial should be sufficient for testing
- Add payment method on Console

---

## 📝 Example: Creating a Sale Deed

```
1. Load sale_deed_template.docx (Screen 1)

2. Screen 2 appears with placeholders:
   - {{seller_name}}
   - {{buyer_name}}
   - {{property_address}}
   - {{consideration}}
   - {{legal_warranty}}

3. Use AI Legal Search:
   - Search: "sale deed clauses"
   - Insert relevant sections

4. Use Generate Clause:
   - Select: {{legal_warranty}}
   - Context: "Residential property in Delhi, clear title"
   - AI generates professional warranty clause

5. Use AI Suggestions:
   - Get recommendations for missing Transfer of Property Act sections
   - Add suggested sections

6. Preview updates in real-time

7. Generate and export document (Screen 3)
```

---

## 🚀 Advanced Tips

### Batch Document Generation
1. Create document template with placeholders
2. Generate one document with AI assistance
3. Use Export to save as PDF or DOCX
4. Repeat with different data

### Custom Legal Database
- Edit `IndianLegalDatabase` in `legal_db_search.py` to add:
  - New legal sections
  - Custom clauses
  - Organization-specific language

### Integrating New Acts
```python
# In legal_db_search.py, add to DATABASE:
"your_act": {
    "name": "Your Act Name",
    "sections": {
        "1": "Section content",
        "2": "More content",
    }
}
```

---

## 📞 Support

If you encounter issues:

1. **Check logs**: Look at terminal output for error details
2. **Verify API key**: Make sure `ANTHROPIC_API_KEY` is set correctly
3. **Test connectivity**: Try using Anthropic Console directly
4. **Check Python version**: Requires Python 3.10+

---

## 🎯 Future Enhancements

Planned features:
- [ ] Document templates library
- [ ] Multi-language legal sections
- [ ] Historical document versions
- [ ] Collaborative editing
- [ ] Integration with e-signature services
- [ ] Document comparison tool

---

## 📜 License & Legal Disclaimer

This tool is for generating document templates. While AI-assisted, always have legal documents reviewed by a qualified lawyer before use in official proceedings.

The embedded legal database contains summaries and references. Always verify against current legislation.

---

**Version**: 4.1  
**Last Updated**: February 24, 2026  
**Status**: AI Features Beta (Testing)
