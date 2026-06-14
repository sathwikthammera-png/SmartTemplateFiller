"""
Legal AI Assistant - Handles Claude API interactions for legal document generation
"""

import os
from typing import Optional
try:
    import anthropic
except Exception:
    anthropic = None
from config import CLAUDE_API_KEY, CLAUDE_MODEL, CLAUDE_MAX_TOKENS

class LegalAIAssistant:
    def __init__(self):
        """Initialize Claude client with API key from environment or config"""
        api_key = os.getenv("ANTHROPIC_API_KEY") or CLAUDE_API_KEY
        
        self.api_key_available = bool(api_key)
        if not self.api_key_available:
            self.client = None
            return
        
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = CLAUDE_MODEL
        self.max_tokens = CLAUDE_MAX_TOKENS

    def generate_legal_clause(self, placeholder: str, context: str, document_type: str) -> str:
        """
        Generate a legal clause based on placeholder and context
        
        Args:
            placeholder: The placeholder name (e.g., "party_details", "consideration")
            context: User context/input
            document_type: Type of document (e.g., "sale_deed", "affidavit")
        
        Returns:
            Generated legal clause text or error message
        """
        if not self.api_key_available or not self.client:
            return "⚠️ API Key Not Configured\n\nTo use AI clause generation, please set ANTHROPIC_API_KEY environment variable or install the `anthropic` package.\n\nAlternatively, use Legal Search tab to find existing clauses from the database."
        
        prompt = f"""You are an expert legal document drafter specializing in Indian law and legal documents.

Generate a professional, legally sound clause for the following:
- Document Type: {document_type}
- Section/Placeholder: {placeholder}
- Context: {context}

Provide only the clause text, without explanations. The clause should:
1. Use proper legal language and formatting
2. Be relevant to Indian law
3. Be concise but comprehensive
4. Be suitable for formal legal documents

Generated Clause:"""

        try:
            message = self.client.messages.create(
                model=self.model,
                max_tokens=self.max_tokens,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            return message.content[0].text.strip()
        except Exception as e:
            return f"❌ Error: {str(e)}"

    def suggest_legal_sections(self, document_type: str, filled_data: dict) -> list:
        """
        Suggest relevant legal sections and acts based on document context
        
        Args:
            document_type: Type of document being created
            filled_data: Dictionary of filled placeholders and values
        
        Returns:
            List of suggested sections with explanations
        """
        if not self.api_key_available or not self.client:
            return ["⚠️ API Key Not Configured",
                "To use AI suggestions, please set ANTHROPIC_API_KEY environment variable or install the `anthropic` package.",
                "Alternatively, use Legal Search tab to find relevant sections manually."]
        
        context_str = "\n".join([f"- {k}: {v}" for k, v in filled_data.items() if v])
        
        prompt = f"""You are an expert in Indian law. Based on the following document details, suggest the most relevant legal sections, acts, and clauses that should be included.

Document Type: {document_type}
Filled Information:
{context_str}

Provide 3-5 specific suggestions in this format:
[Section/Act Name] - [Brief explanation of why it's relevant]

Suggestions:"""

        try:
            message = self.client.messages.create(
                model=self.model,
                max_tokens=self.max_tokens,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            suggestions = message.content[0].text.strip().split("\n")
            return [s.strip() for s in suggestions if s.strip()]
        except Exception as e:
            return [f"❌ Error: {str(e)}"]

    def review_legal_text(self, text: str, document_type: str) -> dict:
        """
        Review drafted legal text for compliance and improvements
        
        Args:
            text: Text to review
            document_type: Type of document
        
        Returns:
            Dictionary with review results and suggestions
        """
        if not self.api_key_available or not self.client:
            return {
                "compliance": "⚠️ API Key Not Configured or `anthropic` not installed",
                "improvements": "Use Legal Search or manual review",
                "missing_elements": "N/A",
                "enhancements": "N/A"
            }

        prompt = f"""You are an expert legal reviewer specializing in Indian law documents.

Review the following legal text from a '{document_type}' document and provide:
1. Compliance issues (if any)
2. Language improvements
3. Missing legal elements
4. Suggestions for enhancement

Text to review:
{text}

Provide response in this format:
COMPLIANCE: [issues or "No issues found"]
IMPROVEMENTS: [suggestions or "Well drafted"]
MISSING_ELEMENTS: [elements or "None"]
ENHANCEMENTS: [suggestions or "N/A"]"""

        try:
            message = self.client.messages.create(
                model=self.model,
                max_tokens=self.max_tokens,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            response = message.content[0].text.strip()
            return self._parse_review_response(response)
        except Exception as e:
            return {
                "compliance": f"❌ Error: {e}",
                "improvements": "",
                "missing_elements": "",
                "enhancements": ""
            }

    def extract_key_information(self, text: str) -> dict:
        """Extract key information from legal text"""
        if not self.api_key_available or not self.client:
            return {"extracted": "⚠️ API Key Not Configured or `anthropic` not installed"}

        prompt = f"""Extract the key information from this legal document text:

{text}

Return as key-value pairs:"""

        try:
            message = self.client.messages.create(
                model=self.model,
                max_tokens=500,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            return {"extracted": message.content[0].text.strip()}
        except Exception as e:
            return {"extracted": f"❌ Error: {e}"}

    def generate_summary(self, text: str) -> str:
        """Generate a concise summary of legal text"""
        if not self.api_key_available or not self.client:
            return "⚠️ API Key Not Configured or `anthropic` not installed"

        prompt = f"""Create a concise, accurate summary of this legal clause or section:

{text}

Summary (max 3 sentences):"""

        try:
            message = self.client.messages.create(
                model=self.model,
                max_tokens=self.max_tokens,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            return message.content[0].text.strip()
        except Exception as e:
            return f"❌ Error: {e}"

    def _parse_review_response(self, response: str) -> dict:
        """Parse the review response into structured format"""
        result = {
            "compliance": "",
            "improvements": "",
            "missing_elements": "",
            "enhancements": ""
        }
        
        for line in response.split("\n"):
            if line.startswith("COMPLIANCE:"):
                result["compliance"] = line.replace("COMPLIANCE:", "").strip()
            elif line.startswith("IMPROVEMENTS:"):
                result["improvements"] = line.replace("IMPROVEMENTS:", "").strip()
            elif line.startswith("MISSING_ELEMENTS:"):
                result["missing_elements"] = line.replace("MISSING_ELEMENTS:", "").strip()
            elif line.startswith("ENHANCEMENTS:"):
                result["enhancements"] = line.replace("ENHANCEMENTS:", "").strip()
        
        return result
