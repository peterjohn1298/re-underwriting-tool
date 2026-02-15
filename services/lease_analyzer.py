"""Lease document NLP analyzer using Google Gemini / Claude API and PyPDF2."""

import logging
import json
import os

logger = logging.getLogger(__name__)


class LeaseAnalyzer:
    """Analyze lease documents using AI (Gemini or Claude) for NLP extraction."""

    def __init__(self, api_key: str = None):
        from config import ANTHROPIC_API_KEY, GOOGLE_API_KEY
        self.anthropic_key = api_key or ANTHROPIC_API_KEY
        self.google_key = GOOGLE_API_KEY

    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """Extract text from a PDF file using PyPDF2."""
        try:
            from PyPDF2 import PdfReader
            reader = PdfReader(pdf_path)
            text_parts = []
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    text_parts.append(text)
            full_text = "\n".join(text_parts)
            logger.info(f"Extracted {len(full_text)} chars from {len(reader.pages)} pages")
            return full_text
        except Exception as e:
            logger.error(f"PDF extraction failed: {e}")
            return ""

    def analyze_lease(self, pdf_path: str) -> dict:
        """Extract and analyze lease terms from a single PDF."""
        if not pdf_path or not os.path.exists(pdf_path):
            return {"error": "PDF file not found"}

        text = self.extract_text_from_pdf(pdf_path)
        if not text or len(text) < 50:
            return {"error": "Could not extract meaningful text from PDF"}

        # Try Gemini first, then Claude, then regex fallback
        if self.google_key:
            try:
                return self._gemini_analysis(text)
            except Exception as e:
                logger.error(f"Gemini analysis failed: {e}")

        if self.anthropic_key and self.anthropic_key != "your_anthropic_api_key_here":
            try:
                return self._claude_analysis(text)
            except Exception as e:
                logger.error(f"Claude analysis failed: {e}")

        return self._fallback_analysis(text)

    def analyze_multiple_leases(self, pdf_paths: list[str]) -> dict:
        """Analyze multiple lease PDFs and return aggregated results.

        Fix #8: Multi-tenant support — analyzes each lease individually,
        then aggregates into a portfolio-level summary.
        """
        if not pdf_paths:
            return {"error": "No lease files provided"}

        individual_leases = []
        all_risk_flags = []
        total_monthly_rent = 0
        total_annual_rent = 0
        lease_count = 0

        for path in pdf_paths:
            result = self.analyze_lease(path)
            if result.get("error"):
                result["file"] = os.path.basename(path)
                individual_leases.append(result)
                continue

            result["file"] = os.path.basename(path)
            individual_leases.append(result)
            lease_count += 1

            if result.get("monthly_rent"):
                total_monthly_rent += result["monthly_rent"]
            if result.get("annual_rent"):
                total_annual_rent += result["annual_rent"]
            for flag in result.get("risk_flags", []):
                all_risk_flags.append(f"[{result.get('tenant_name', os.path.basename(path))}] {flag}")

        # Weighted average escalation
        escalation_rates = [l.get("annual_escalation_pct") for l in individual_leases
                           if l.get("annual_escalation_pct") is not None]
        avg_escalation = sum(escalation_rates) / len(escalation_rates) if escalation_rates else None

        return {
            "lease_count": lease_count,
            "total_files": len(pdf_paths),
            "individual_leases": individual_leases,
            "portfolio_summary": {
                "total_monthly_rent": total_monthly_rent if total_monthly_rent else None,
                "total_annual_rent": total_annual_rent if total_annual_rent else None,
                "avg_escalation_pct": round(avg_escalation, 2) if avg_escalation else None,
                "total_risk_flags": len(all_risk_flags),
            },
            "risk_flags": all_risk_flags,
            "summary": (
                f"Analyzed {lease_count} of {len(pdf_paths)} lease documents. "
                f"Total monthly rent: ${total_monthly_rent:,.0f}. "
                f"{len(all_risk_flags)} risk flag(s) identified across all leases."
                if total_monthly_rent else
                f"Analyzed {lease_count} of {len(pdf_paths)} lease documents."
            ),
            "analysis_method": "multi_lease_aggregated",
        }

    def _get_lease_prompt(self, text: str) -> str:
        """Shared prompt for both Gemini and Claude."""
        max_chars = 80000
        if len(text) > max_chars:
            text = text[:max_chars] + "\n...[TRUNCATED]..."
        return f"""Analyze this commercial real estate lease document and extract key terms.
Return ONLY valid JSON with these fields (use null for any field you cannot determine):

{{
    "tenant_name": "string",
    "landlord_name": "string",
    "lease_type": "string (e.g., NNN, Gross, Modified Gross)",
    "monthly_rent": number or null,
    "annual_rent": number or null,
    "rent_per_sf": number or null,
    "lease_term_months": number or null,
    "lease_start_date": "string or null",
    "lease_end_date": "string or null",
    "escalation_clause": "string description or null",
    "annual_escalation_pct": number or null,
    "renewal_options": "string description or null",
    "security_deposit": number or null,
    "ti_allowance": number or null,
    "ti_per_sf": number or null,
    "cam_charges": "string description or null",
    "cam_annual": number or null,
    "permitted_use": "string or null",
    "key_clauses": ["list of notable clauses"],
    "risk_flags": ["list of risk concerns for landlord/investor"],
    "summary": "2-3 sentence executive summary of the lease"
}}

LEASE DOCUMENT TEXT:
{text}"""

    def _gemini_analysis(self, text: str) -> dict:
        """Use Google Gemini API to extract structured lease data."""
        import google.generativeai as genai

        genai.configure(api_key=self.google_key)
        model = genai.GenerativeModel("gemini-2.0-flash")

        prompt = self._get_lease_prompt(text)
        response = model.generate_content(prompt)
        response_text = response.text

        try:
            if "```json" in response_text:
                json_str = response_text.split("```json")[1].split("```")[0]
            elif "```" in response_text:
                json_str = response_text.split("```")[1].split("```")[0]
            else:
                json_str = response_text

            result = json.loads(json_str.strip())
            result["analysis_method"] = "gemini_nlp"
            result["text_length"] = len(text)
            return result

        except json.JSONDecodeError:
            logger.warning("Failed to parse Gemini JSON response")
            return {
                "summary": response_text[:500],
                "analysis_method": "gemini_nlp_raw",
                "raw_response": response_text,
                "text_length": len(text),
            }

    def _claude_analysis(self, text: str) -> dict:
        """Use Claude API to extract structured lease data (fallback if Gemini unavailable)."""
        import anthropic

        client = anthropic.Anthropic(api_key=self.anthropic_key)
        prompt = self._get_lease_prompt(text)

        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}],
        )

        response_text = message.content[0].text

        # Parse JSON from response
        try:
            # Try to extract JSON from markdown code blocks if present
            if "```json" in response_text:
                json_str = response_text.split("```json")[1].split("```")[0]
            elif "```" in response_text:
                json_str = response_text.split("```")[1].split("```")[0]
            else:
                json_str = response_text

            result = json.loads(json_str.strip())
            result["analysis_method"] = "claude_nlp"
            result["text_length"] = len(text)
            return result

        except json.JSONDecodeError:
            logger.warning("Failed to parse Claude JSON response")
            return {
                "summary": response_text[:500],
                "analysis_method": "claude_nlp_raw",
                "raw_response": response_text,
                "text_length": len(text),
            }

    def _fallback_analysis(self, text: str) -> dict:
        """Basic regex-based fallback when Claude API is unavailable."""
        import re

        result = {
            "analysis_method": "regex_fallback",
            "text_length": len(text),
            "tenant_name": None,
            "landlord_name": None,
            "lease_type": None,
            "monthly_rent": None,
            "annual_rent": None,
            "lease_term_months": None,
            "escalation_clause": None,
            "renewal_options": None,
            "security_deposit": None,
            "ti_allowance": None,
            "cam_charges": None,
            "key_clauses": [],
            "risk_flags": ["Automated analysis only — manual review recommended"],
            "summary": "Lease document uploaded but AI analysis unavailable. "
                       "Configure GOOGLE_API_KEY or ANTHROPIC_API_KEY for full NLP analysis.",
        }

        text_lower = text.lower()

        # Try to detect lease type
        if "triple net" in text_lower or "nnn" in text_lower:
            result["lease_type"] = "NNN (Triple Net)"
        elif "gross lease" in text_lower:
            result["lease_type"] = "Gross"
        elif "modified gross" in text_lower:
            result["lease_type"] = "Modified Gross"

        # Extract dollar amounts
        dollar_amounts = re.findall(r'\$[\d,]+(?:\.\d{2})?', text)
        if dollar_amounts:
            amounts = []
            for d in dollar_amounts:
                try:
                    amounts.append(float(d.replace("$", "").replace(",", "")))
                except ValueError:
                    pass
            if amounts:
                result["detected_amounts"] = sorted(set(amounts), reverse=True)[:10]

        # Extract dates
        dates = re.findall(r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b', text)
        if dates:
            result["detected_dates"] = dates[:5]

        # Extract term mentions
        term_match = re.search(r'(\d+)\s*(?:year|yr)(?:s)?\s*(?:term|lease)', text_lower)
        if term_match:
            years = int(term_match.group(1))
            result["lease_term_months"] = years * 12

        # Detect key clause keywords
        clause_keywords = {
            "Assignment/Subletting": ["assignment", "subletting", "sublet"],
            "Default/Remedies": ["default", "remedies", "cure period"],
            "Insurance Requirements": ["insurance", "liability coverage"],
            "Maintenance/Repairs": ["maintenance", "repairs", "landlord responsible"],
            "Termination": ["early termination", "termination clause"],
            "Escalation": ["escalation", "annual increase", "cpi adjustment"],
            "Renewal Option": ["renewal option", "option to renew", "extend"],
        }
        for clause_name, keywords in clause_keywords.items():
            if any(kw in text_lower for kw in keywords):
                result["key_clauses"].append(clause_name)

        return result
