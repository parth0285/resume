import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Direct Google Gemini API configuration.
# Confirmed against Google's current API docs (ai.google.dev/api): the real
# REST endpoint is generativelanguage.googleapis.com/v1beta — NOT v1beta2,
# which does not exist for gemini-* models. Method must be :generateContent
# (not :generateText or :generateMessage, which are old PaLM-2 REST methods).
GOOGLE_STUDIO_BASE_URL = os.getenv("GOOGLE_STUDIO_BASE_URL", "https://generativelanguage.googleapis.com/v1beta")
GOOGLE_STUDIO_MODEL = os.getenv("GOOGLE_STUDIO_MODEL", "gemini-3.5-flash")
# NOTE: 19000 bytes was a leftover assumption from the old PaLM-2 text-bison
# API, which really did have a small request-size ceiling. Gemini's actual
# context window is enormous (1M+ tokens for 3.x flash models) — a 19KB cap
# was silently truncating real multi-page resumes before the model even saw
# them. Raised to 2MB, which comfortably fits any realistic resume PDF's
# extracted text while still guarding against a truly pathological input.
GOOGLE_STUDIO_MAX_REQUEST_BYTES = int(os.getenv("GOOGLE_STUDIO_MAX_REQUEST_BYTES", "2000000"))


# Generation settings for structured extraction.
# Temperature is 0 (combined with topK=1 in the API call itself) so extraction
# is as deterministic as the model allows — this minimizes (though cannot
# fully eliminate; some residual variance is inherent to the model/hardware)
# run-to-run differences in which optional list items get extracted.
EXTRACTION_TEMPERATURE = float(os.getenv("EXTRACTION_TEMPERATURE", "0"))

# OCR Fallback settings
# DPI for converting PDF pages to images. 150-200 is usually enough for text.
OCR_DPI = int(os.getenv("OCR_DPI", "200"))
# Maximum number of pages to process with OCR to prevent hanging on huge scanned PDFs.
OCR_MAX_PAGES = int(os.getenv("OCR_MAX_PAGES", "10"))