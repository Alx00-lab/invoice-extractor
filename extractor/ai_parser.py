from functools import lru_cache
import json
import time
import logging

from openai import (
    OpenAI,
    APIError,
    APITimeoutError,
    APIConnectionError,
    RateLimitError,
    InternalServerError,
)

from .validator import validate_invoice

logger = logging.getLogger(__name__)

MODEL = "gpt-4o-mini"
MAX_INPUT_CHARS = 12_000   # ~3-4k tokens; bounds cost on pathologically long PDFs
MAX_RETRIES = 3
BASE_DELAY = 1.0           # seconds; doubles each retry (1s, 2s, 4s)
REQUEST_TIMEOUT = 60.0

# Transient failures worth retrying. A 2-second wait fixes most of these.
_RETRYABLE = (RateLimitError, APITimeoutError, APIConnectionError, InternalServerError)

# One client (and HTTP connection pool) per API key — reused across the whole batch.
_client_cache = {}

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
  ]
}

Rules:
- If a field is not found, use null.
- For dates, keep the original format found in the document.
- For monetary values, keep the original string exactly as written, including the currency symbol.
- Do NOT perform any arithmetic, summation, or validation. Extraction only.
- Never guess or hallucinate values. Only extract what is clearly present in the text.
"""

# Here is where I did the new change to test the new branch and not max out API
@lru_cache(maxsize=32)
def _get_client(api_key: str) -> OpenAI:
    """Cached client per API key — max 32 keys in memory at once."""
    return OpenAI(api_key=api_key.strip(), max_retries=0, timeout=REQUEST_TIMEOUT)


def _truncate(text: str):
    """Bound input size. Returns (text, was_truncated)."""
    if len(text) <= MAX_INPUT_CHARS:
        return text, False
    logger.warning(f"Input text {len(text)} chars exceeds {MAX_INPUT_CHARS}; truncating.")
    cut = text.rfind("\n", 0, MAX_INPUT_CHARS)
    return text[:cut if cut > 0 else MAX_INPUT_CHARS], True


def parse_invoice(text: str, api_key: str) -> dict:
    """
    Send extracted invoice text to GPT-4o-mini and return a structured dict.

    - Reuses a cached client per API key (no new HTTP pool per invoice).
    - Retries transient API errors with exponential backoff.
    - Truncates oversized input to cap cost.
    - Validates totals in pure Python (never via the model).

    Raises ValueError for bad input, RuntimeError for unrecoverable API failures.
    """
    if not api_key or not api_key.strip():
        raise ValueError("OpenAI API key is missing.")
    if not text or not text.strip():
        raise ValueError("No invoice text to parse.")

    client = _get_client(api_key)
    text, truncated = _truncate(text)

    data = None
    last_err = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": f"Invoice text:\n\n{text}"},
                ],
                response_format={"type": "json_object"},
                temperature=0,
            )
            data = json.loads(response.choices[0].message.content)
            break

        except json.JSONDecodeError as e:
            logger.error(f"JSON parsing failed: {e}")
            raise RuntimeError("The AI returned an invalid response. Please try again.")

        except _RETRYABLE as e:
            last_err = e
            if attempt < MAX_RETRIES:
                delay = BASE_DELAY * (2 ** (attempt - 1))
                logger.warning(
                    f"Transient API error (attempt {attempt}/{MAX_RETRIES}): {e}. "
                    f"Retrying in {delay:.0f}s."
                )
                time.sleep(delay)
                continue
            logger.error(f"API failed after {MAX_RETRIES} attempts: {e}")
            raise RuntimeError(
                f"AI service is temporarily unavailable (after {MAX_RETRIES} attempts). "
                "Please try again in a moment."
            )

        except APIError as e:
            logger.error(f"OpenAI API error: {e}")
            raise RuntimeError(f"AI processing failed: {e}")

        except Exception as e:
            logger.error(f"Unexpected error: {e}")
            raise RuntimeError(f"AI processing failed: {e}")

    if data is None:
        raise RuntimeError(f"AI processing failed: {last_err}")

    # Validation happens in Python, not in the model.
    data["validation_warning"] = validate_invoice(data)

    if truncated:
        note = (
            f"Invoice text was long and was truncated to {MAX_INPUT_CHARS:,} characters "
            "for processing — verify totals."
        )
        existing = data.get("validation_warning")
        data["validation_warning"] = f"{note} | {existing}" if existing else note

    logger.info("Invoice parsed successfully.")
    return data
