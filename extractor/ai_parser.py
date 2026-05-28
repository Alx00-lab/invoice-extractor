import json
import logging
from openai import OpenAI

logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """You are an invoice data extraction specialist.

Extract structured data from the invoice text provided and return ONLY a valid JSON object.
No explanations, no markdown, no extra text — just the JSON.

Required JSON structure:
{
  "vendor_name": "string or null",
  "invoice_number": "string or null",
  "invoice_date": "string or null",
  "due_date": "string or null",
  "subtotal": "string or null",
  "tax": "string or null",
  "total": "string or null",
  "payment_terms": "string or null",
  "line_items": [
    {
      "description": "string",
      "quantity": "string or null",
      "unit_price": "string or null",
      "amount": "string or null"
    }
  ],
  "validation_warning": "string or null"
}

Rules:
- If a field is not found, use null.
- For dates, keep the original format found in the document.
- For monetary values, keep the original string including currency symbol.
- In validation_warning: if the sum of line_items amounts does NOT match total, 
  write "Total mismatch: line items sum to X but invoice total shows Y". Otherwise null.
- Never guess or hallucinate values. Only extract what is clearly in the text.
"""


def parse_invoice(text: str, api_key: str) -> dict:
    """
    Sends extracted invoice text to GPT-4o-mini.
    Returns a structured dict or raises RuntimeError.
    """
    if not api_key or not api_key.strip():
        raise ValueError("OpenAI API key is missing.")

    client = OpenAI(api_key=api_key.strip())

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": f"Invoice text:\n\n{text}"}
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )

        raw = response.choices[0].message.content
        data = json.loads(raw)
        logger.info("Invoice parsed successfully.")
        return data

    except json.JSONDecodeError as e:
        logger.error(f"JSON parsing failed: {e}")
        raise RuntimeError("The AI returned an invalid response. Please try again.")
    except Exception as e:
        logger.error(f"OpenAI API error: {e}")
        raise RuntimeError(f"AI processing failed: {e}")
