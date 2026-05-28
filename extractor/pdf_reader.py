import pdfplumber
import logging

logger = logging.getLogger(__name__)


def extract_text_from_pdf(file) -> str:
    """
    Accepts a file path (str) or a file-like object (Streamlit UploadedFile).
    Returns all extracted text or raises ValueError if empty.
    """
    try:
        with pdfplumber.open(file) as pdf:
            pages_text = []
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    pages_text.append(text)
                else:
                    logger.warning(f"Page {i+1} returned no text — may be scanned/image-based.")

        full_text = "\n".join(pages_text).strip()

        if not full_text:
            raise ValueError(
                "No text could be extracted. The PDF may be scanned or image-based. "
                "OCR support is not included in this version."
            )

        logger.info(f"Extracted {len(full_text)} characters from PDF.")
        return full_text

    except ValueError:
        raise
    except Exception as e:
        logger.error(f"Failed to read PDF: {e}")
        raise RuntimeError(f"Could not read the PDF file: {e}")
