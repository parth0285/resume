# Faculty CV & Publication Extraction Engine

An application that turns a batch of **Faculty CVs / Academic Resumes / Research Profiles** into two clean, structured datasets — a **Master Faculty Database** and a **Publication Details** list — exported as a single two-sheet Excel workbook. Powered by **Google Studio Gemini** via direct API calls.

---

## 📋 **Project Overview**

Upload any number of PDF resumes and the app will, for each one:
- Extract personal details, contact info, current role, education, experience, and profile links (ORCID, Google Scholar, ResearchGate, Scopus, LinkedIn) into **Sheet 1 — Master Faculty Database** (one row per resume).
- Extract every publication (journals, conferences, books, book chapters, patents) into **Sheet 2 — Publication Details** (one row per publication, linked back to the faculty member).

Missing information is always returned as `null`/blank — the extractor never guesses or fabricates values.

---

## 🔑 **Features**

- **Bulk upload** — analyze multiple resumes in one pass.
- **Two structured datasets** — Master Faculty Database + Publication Details, matching a fixed JSON schema.
- **On-screen review** — both datasets are shown as interactive tables before export.
- **One-click Excel export** — download a single `.xlsx` file with both sheets, auto-fitted columns.
- **OCR fallback** — scanned/image-based PDFs are still readable via `pytesseract`.
- **Per-file error handling** — a bad file doesn't stop the rest of the batch.

---

## 🛠️ **Tech Stack**

| **Component**       | **Technology**                  |  
|----------------------|----------------------------------|  
| **Frontend**         | [Streamlit](https://streamlit.io/) |  
| **Backend**          | Python                          |  
| **AI Model**         | Google Studio Gemini            |  
| **PDF Parsing**      | `pdfplumber`                    |  
| **OCR Fallback**     | `pytesseract`                   |  
| **Excel Export**     | `pandas` + `openpyxl`           |  
| **Environment Config** | `.env` for API key security    |  

---

## 📊 **How It Works**

1. **Resume Parsing**
   - Extracts text from each uploaded PDF using `pdfplumber`, falling back to OCR for scanned/image-based files.

2. **Structured Extraction**
   - Sends each resume's text directly to Google Studio Gemini with a strict extraction prompt and JSON schema.
   - Returns one Master Faculty Database record and zero-or-more Publication Details records per resume.

3. **Review & Export**
   - Results from all uploaded resumes are aggregated and shown as two tables in the app.
   - Download everything as a single Excel workbook with two sheets.

---

## 🔑 **Setup**

1. Get an API key from Google Studio / AI Studio and add it to a `.env` file in the project root:
   ```
   GOOGLE_STUDIO_API_KEY=your_key_here
   GOOGLE_STUDIO_MODEL=gemini-3.5-flash
   ```
   Browse available Gemini model ids in the Google Cloud Console or AI Studio project dashboard.
3. Install dependencies: `pip install -r requirements.txt`
4. Run the app: `streamlit run app.py`

---

![image](https://github.com/user-attachments/assets/418e54ef-82d0-474b-a6bc-9a30d72f27f5)

## 🙌 **Contributing**

Welcome contributions to make this tool better!

1. **Fork** the repository.  
2. **Create a new branch** for your feature or bug fix.  
3. **Submit a pull request** with detailed information about your changes.
