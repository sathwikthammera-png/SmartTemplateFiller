"""
Legal Database Search - Comprehensive Indian Legal Sections Database
"""

class IndianLegalDatabase:
    """Contains searchable Indian legal sections and acts"""
    
    DATABASE = {
        # Indian Penal Code
        "ipc": {
            "name": "Indian Penal Code, 1860",
            "sections": {
                "138": "Cheque Dishonour",
                "415": "Cheating",
                "420": "Cheating and dishonestly inducing delivery of property",
                "498A": "Cruelty by husband or his relatives",
                "304": "Death caused by negligence",
                "354": "Assault or use of criminal force to a woman",
            }
        },
        
        # Indian Contract Act
        "contract_act": {
            "name": "Indian Contract Act, 1872",
            "sections": {
                "2": "Definitions",
                "10": "Who can contract",
                "12": "Capacity of parties",
                "23": "Consideration",
                "24": "Consideration must move at the desire of the promisor",
                "25": "Absence of consideration",
                "29": "Void agreements restricting legal proceedings",
            }
        },
        
        # Transfer of Property Act
        "property_act": {
            "name": "Transfer of Property Act, 1882",
            "sections": {
                "5": "Definition of transfer of property",
                "6": "Property transferable",
                "10": "Conditions of transfer",
                "14": "Right to transfer must be absolute",
                "54": "Sale of immovable property",
                "55": "Conditions prescribed by parties",
            }
        },
        
        # Sale Deed Specific
        "sale_deed": {
            "name": "Sale Deed Guidelines",
            "clauses": {
                "parties": "Identification and capacity of buyer and seller",
                "consideration": "Agreed purchase price and payment terms",
                "property_description": "Legal description of property with survey numbers",
                "title": "Warranty of seller's title and possession",
                "encumbrances": "Declaration of freedom from encumbrances",
                "possession": "Date and mode of possession",
                "defects": "Covenant for further assurance",
                "stamp_duty": "Obligation to pay applicable stamp duty",
            }
        },
        
        # Affidavit
        "affidavit": {
            "name": "Affidavit Guidelines",
            "clauses": {
                "declaration": "I, [Name], son/daughter of [Father's Name], aged..., do hereby solemnly affirm and declare as follows:",
                "facts": "The facts stated herein are true to the best of my knowledge and belief",
                "signature": "Signature of deponent",
                "verification": "Verification: Verified at [Place] on [Date]",
            }
        },
        
        # Succession Act
        "succession": {
            "name": "Indian Succession Act, 1925",
            "sections": {
                "2": "Definitions",
                "3": "Jurisdiction",
                "30": "Letters of administration",
                "52": "Will definition",
                "63": "Requirements of valid will",
                "113": "Succession of widows",
            }
        },
        
        # General Clauses
        "general_clauses": {
            "name": "General Legal Clauses",
            "clauses": {
                "governing_law": "This Agreement shall be governed by and construed in accordance with the laws of India",
                "jurisdiction": "The parties irrevocably agree to the exclusive jurisdiction of courts of [Location]",
                "confidentiality": "The parties agree to maintain confidentiality of this document",
                "amendment": "No amendment of this document shall be valid unless in writing and signed by all parties",
                "severability": "If any provision is found invalid, remaining provisions shall continue in force",
            }
        }
    }

    @staticmethod
    def search(query: str, doc_type: str = None) -> list:
        """
        Search legal database for relevant sections
        
        Args:
            query: Search keyword or phrase
            doc_type: Optional document type filter (e.g., "sale_deed", "affidavit")
        
        Returns:
            List of relevant sections with details
        """
        results = []
        query_lower = query.lower().strip()
        
        # Search in specific document type if provided
        if doc_type and doc_type in IndianLegalDatabase.DATABASE:
            db_section = IndianLegalDatabase.DATABASE[doc_type]
            items = db_section.get("sections", {}).items() or db_section.get("clauses", {}).items()
            
            for key, value in items:
                if query_lower in str(key).lower() or query_lower in str(value).lower():
                    results.append({
                        "source": db_section["name"],
                        "reference": key,
                        "content": value,
                        "type": "section" if "sections" in db_section else "clause"
                    })
        else:
            # Search across all databases
            for db_key, db_content in IndianLegalDatabase.DATABASE.items():
                items = db_content.get("sections", {}).items() or db_content.get("clauses", {}).items()
                
                for key, value in items:
                    if query_lower in str(key).lower() or query_lower in str(value).lower():
                        results.append({
                            "source": db_content["name"],
                            "reference": key,
                            "content": value,
                            "type": "section" if "sections" in db_content else "clause"
                        })
        
        return results

    @staticmethod
    def get_section(act_name: str, section_number: str) -> dict:
        """
        Get specific section by act and number
        
        Args:
            act_name: Name of act (e.g., "ipc", "contract_act")
            section_number: Section/clause identifier
        
        Returns:
            Section details or empty dict if not found
        """
        if act_name not in IndianLegalDatabase.DATABASE:
            return {}
        
        db = IndianLegalDatabase.DATABASE[act_name]
        sections = db.get("sections", {})
        clauses = db.get("clauses", {})
        
        content = sections.get(section_number) or clauses.get(section_number)
        if content:
            return {
                "source": db["name"],
                "reference": section_number,
                "content": content
            }
        
        return {}

    @staticmethod
    def get_all_acts() -> dict:
        """Get all available acts and documents"""
        return {
            key: value["name"] 
            for key, value in IndianLegalDatabase.DATABASE.items()
        }

    @staticmethod
    def get_act_content(act_name: str) -> dict:
        """Get all sections/clauses of a specific act"""
        if act_name not in IndianLegalDatabase.DATABASE:
            return {}
        
        db = IndianLegalDatabase.DATABASE[act_name]
        return {
            "name": db["name"],
            "sections": db.get("sections", {}),
            "clauses": db.get("clauses", {})
        }
