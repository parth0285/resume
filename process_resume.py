import sys
import os
import json
from analyzer_engine import ResumeAnalyzerEngine

if len(sys.argv) < 2:
    print("Usage: python3 process_resume.py <path_to_pdf> [model_value]")
    sys.exit(1)

pdf_path = sys.argv[1]
model_value = sys.argv[2] if len(sys.argv) > 2 else "google/gemini-3.5-flash-lite"

if not os.path.exists(pdf_path):
    print(f"File not found: {pdf_path}")
    sys.exit(2)

engine = ResumeAnalyzerEngine()
# model_value is expected like 'google/gemini-3.5-flash-lite'
print(f"Using model: {model_value}")

text = engine.extract_text_from_pdf(pdf_path)
print(f"Extracted {len(text)} chars of text from PDF.")
res = engine.extract_faculty_data(text, os.path.basename(pdf_path), model_name=model_value)

# Print errors if any
if isinstance(res, dict) and res.get("error"):
    print("Extraction error:", res.get("error"))
    sys.exit(3)

master = res.get("master_faculty_database", [])
if not master:
    print("No master records returned by extraction.")
    sys.exit(4)

record = master[0]
print("\n=== RAW designation_history ===")
print(json.dumps(record.get("designation_history"), indent=2, ensure_ascii=False))
print("\n=== Computed experience ===")
print(json.dumps(record.get("experience"), indent=2, ensure_ascii=False))
print("\n=== Flags ===")
print(json.dumps(record.get("flags"), indent=2, ensure_ascii=False))

# Show the debug dump file location
dump_path = os.path.join(os.getcwd(), "debug_logs", "designation_history_dump.jsonl")
print(f"\nDebug dump appended to: {dump_path}")
