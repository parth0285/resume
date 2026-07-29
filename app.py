import os
import tempfile
import time
import streamlit as st
from dotenv import load_dotenv

import config
from analyzer_engine import ResumeAnalyzerEngine
from export_utils import (
    master_records_to_dataframe,
    publication_records_to_dataframe,
    build_excel_workbook,
    build_single_sheet_workbook,
    build_publication_sheet_grouped,
    build_publication_workbook_per_faculty,
)

# Load environment variables
load_dotenv()

# Bridge Streamlit Cloud secrets into environment variables
if "GOOGLE_STUDIO_API_KEY" in st.secrets:
    os.environ["GOOGLE_STUDIO_API_KEY"] = st.secrets["GOOGLE_STUDIO_API_KEY"]

# Set page config for a clean SaaS look
st.set_page_config(
    page_title="Resume Extraction Engine",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- Minimal theme (Google Fonts + custom CSS) ---
st.markdown("""
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        [data-testid="collapsedControl"] { display: none; }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header[data-testid="stHeader"] { background: transparent; }

        html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

        .block-container {
            padding-top: 2.5rem;
            padding-bottom: 3rem;
            max-width: 1200px;
        }

        :root {
            --accent: #7C5CFC;
            --grad: linear-gradient(90deg, #7C5CFC 0%, #C05CDB 100%);
        }

        /* Header */
        .app-title {
            font-weight: 700;
            font-size: 2rem;
            color: #14121F;
            margin-bottom: 0.15rem;
        }
        .app-title .accent { color: var(--accent); }
        .app-sub {
            color: #6B7280;
            font-size: 1.05rem;
            margin-bottom: 1.8rem;
        }

        /* Step track — minimal text row */
        .step-track {
            display: flex;
            align-items: center;
            gap: 0.4rem;
            font-size: 0.85rem;
            color: #9691A8;
            margin-bottom: 1.8rem;
            flex-wrap: wrap;
        }
        .step-track .step { display: flex; align-items: center; gap: 0.35rem; }
        .step-track .step.active { color: #14121F; font-weight: 600; }
        .step-track .num {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 18px;
            height: 18px;
            border-radius: 50%;
            background: #EFEBFE;
            color: var(--accent);
            font-size: 0.7rem;
            font-weight: 600;
        }
        .step-track .step.active .num { background: var(--accent); color: white; }
        .step-track .sep { color: #D9D5E8; }

        /* Section headers */
        .section-head {
            font-weight: 600;
            font-size: 1rem;
            color: #14121F;
            margin-top: 0.3rem;
            margin-bottom: 0.15rem;
        }
        .section-sub {
            color: #8A8698;
            font-size: 0.85rem;
            margin-bottom: 0.6rem;
        }

        /* Primary buttons -> flat gradient, no heavy shadow */
        .stButton > button[kind="primary"], .stDownloadButton > button {
            background: var(--grad) !important;
            border: none !important;
            color: white !important;
            font-weight: 600 !important;
            border-radius: 8px !important;
            box-shadow: none !important;
        }
        .stButton > button[kind="primary"]:hover, .stDownloadButton > button:hover {
            filter: brightness(1.05);
            color: white !important;
        }

        /* File uploader -> subtle dashed drop zone */
        /* Empty-state placeholder */
        .empty-state {
            border: 1.5px dashed #E5E2F0;
            border-radius: 10px;
            padding: 2.4rem 1.5rem;
            text-align: center;
            color: #9691A8;
            margin-top: 1.6rem;
        }
        .empty-state .icon { font-size: 1.6rem; margin-bottom: 0.4rem; }
        .empty-state .title { font-weight: 600; color: #57536A; font-size: 0.95rem; margin-bottom: 0.2rem; }
        .empty-state .sub { font-size: 0.85rem; }

        [data-testid="stFileUploaderDropzone"] {
            background: #FAFAFC !important;
            border: 1.5px dashed #DCD8EC !important;
            border-radius: 10px !important;
        }

        hr { border-color: #EEECF5 !important; margin: 1.2rem 0 !important; }
    </style>
""", unsafe_allow_html=True)


def step_header(title: str, subtitle: str = ""):
    st.markdown(f'<div class="section-head">{title}</div>', unsafe_allow_html=True)
    if subtitle:
        st.markdown(f'<div class="section-sub">{subtitle}</div>', unsafe_allow_html=True)


# --- Header ---
st.markdown('<div class="app-title">🎓 Resume <span class="accent">Extraction</span> Engine</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="app-sub">Automated, structured data extraction for faculty resumes and academic CVs.</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="step-track">
        <div class="step active"><span class="num">1</span>Select Model</div>
        <span class="sep">→</span>
        <div class="step"><span class="num">2</span>Upload CVs</div>
        <span class="sep">→</span>
        <div class="step"><span class="num">3</span>Review</div>
        <span class="sep">→</span>
        <div class="step"><span class="num">4</span>Export</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Initialize the modular analysis engine
@st.cache_resource
def get_engine():
    return ResumeAnalyzerEngine()

engine = get_engine()

# Initialize session state for persistent results
if "master_records" not in st.session_state:
    st.session_state.master_records = []
if "publication_records" not in st.session_state:
    st.session_state.publication_records = []
if "extraction_errors" not in st.session_state:
    st.session_state.extraction_errors = []
if "processed_files" not in st.session_state:
    st.session_state.processed_files = []
if "extraction_running" not in st.session_state:
    st.session_state.extraction_running = False
if "stop_requested" not in st.session_state:
    st.session_state.stop_requested = False
if "selected_model" not in st.session_state:
    st.session_state.selected_model = "google/gemini-3.5-flash-lite"

# --- STEP 1: Model Selection ---
step_header("Select AI Model")

model_options = [
    "Gemini 3.5 Flash (20 resumes/day)",
    "Gemini 3.6 Flash (20 resumes/day)",
    "Gemini 3.5 Flash Lite (500 resumes/day)",
    "Gemini 3.1 Flash Lite (500 resumes/day)",
]

model_values = {
    "Gemini 3.5 Flash (20 resumes/day)": "google/gemini-3.5-flash",
    "Gemini 3.6 Flash (20 resumes/day)": "google/gemini-3.6-flash",
    "Gemini 3.5 Flash Lite (500 resumes/day)": "google/gemini-3.5-flash-lite",
    "Gemini 3.1 Flash Lite (500 resumes/day)": "google/gemini-3.1-flash-lite",
}

default_model = "google/gemini-3.5-flash-lite"

default_display = next((label for label, model in model_values.items() if model == default_model), model_options[0])

if st.session_state.selected_model not in model_values.values():
    st.session_state.selected_model = default_model

selected_display = st.selectbox(
    "Model",
    options=model_options,
    index=model_options.index(next(label for label, model in model_values.items() if model == st.session_state.selected_model)),
    label_visibility="collapsed"
)
selected_model = model_values[selected_display]
st.session_state.selected_model = selected_model

# Ensure Google Studio key exists
if not os.getenv("GOOGLE_STUDIO_API_KEY"):
    st.error("Missing GOOGLE_STUDIO_API_KEY in .env file.")
    st.stop()


# --- STEP 2: Upload Section ---
step_header("Upload CVs", "Upload PDF or DOCX resumes for structured extraction.")
max_files = 20

uploaded_files = st.file_uploader(
    "Upload CVs",
    type=["pdf", "docx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="Upload up to 20 resumes in PDF or DOCX format."
)

# Enforce limits
over_limit = False
if uploaded_files and len(uploaded_files) > max_files:
    over_limit = True
    st.error(f"Remove {len(uploaded_files) - max_files} files to continue — limit is {max_files} per Google Studio model.")

btn_disabled = not uploaded_files or over_limit

extract_btn = st.button("Extract Data", disabled=btn_disabled, type="primary")

# Triggering extraction
if extract_btn and uploaded_files and not over_limit:
    # Mark extraction as running and clear any previous stop flag
    st.session_state.extraction_running = True
    st.session_state.stop_requested = False
    new_master_records = st.session_state.master_records.copy()
    new_publication_records = st.session_state.publication_records.copy()
    new_errors = []
    failed_files = set()
    new_processed = st.session_state.processed_files.copy()

    progress = st.progress(0.0, text="Starting extraction...")
    total = len(uploaded_files)

    # Stop button — visible only while extraction is running
    stop_col1, stop_col2 = st.columns([1, 3])
    with stop_col1:
        if st.button("⏹️ Stop Extraction"):
            st.session_state.stop_requested = True
    with stop_col2:
        st.write("")

    for idx, uploaded_file in enumerate(uploaded_files, start=1):
        # If user requested stop before starting this iteration, break.
        if st.session_state.stop_requested:
            progress.progress((idx - 1) / total, text=f"Stopping after {idx-1} of {total} files...")
            break
        if uploaded_file.name in new_processed:
            continue

        progress.progress((idx - 1) / total, text=f"Processing {uploaded_file.name} ({idx}/{total})...")
        tmp_path = None

        try:
            file_name = uploaded_file.name
            _, file_ext = os.path.splitext(file_name)
            file_ext = file_ext.lower()

            if file_ext == ".doc":
                new_errors.append(
                    f"'{file_name}': Old .doc format isn't supported — please save this file as .docx or PDF and re-upload."
                )
                continue
            if file_ext not in {".pdf", ".docx"}:
                new_errors.append(
                    f"'{file_name}': Unsupported file type '{file_ext}'. Please upload PDF or DOCX files."
                )
                continue

            suffix = file_ext
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                tmp.write(uploaded_file.getbuffer())
                tmp_path = tmp.name

            def _update_progress(msg):
                progress.progress(
                    (idx - 1) / total,
                    text=f"{uploaded_file.name} ({idx}/{total}): {msg}"
                )

            if file_ext == ".pdf":
                resume_text = engine.extract_text_from_pdf(tmp_path, progress_callback=_update_progress)
            else:
                resume_text = engine.extract_text_from_docx(tmp_path, progress_callback=_update_progress)

            if not resume_text.strip():
                new_errors.append(f"'{uploaded_file.name}': could not extract any readable text.")
                continue

            attempt = 1
            while attempt <= 2:
                progress.progress(
                    (idx - 1) / total,
                    text=f"Processing {uploaded_file.name} ({idx}/{total})... attempt {attempt}"
                )

                result = engine.extract_faculty_data(resume_text, uploaded_file.name, model_name=st.session_state.selected_model)
                if "error" not in result:
                    new_master_records.extend(result.get("master_faculty_database", []))
                    new_publication_records.extend(result.get("publication_details", []))
                    new_processed.append(uploaded_file.name)
                    break

                error_message = result["error"]
                # Avoid appending duplicate errors for the same file
                if uploaded_file.name in failed_files:
                    break
                is_rpm_error = "rate limit" in error_message.lower() or "requests per minute" in error_message.lower()
                if attempt == 1 and is_rpm_error:
                    progress.progress(
                        (idx - 1) / total,
                        text=f"Processing {uploaded_file.name} ({idx}/{total})... rate limit hit, retrying in 15s"
                    )
                    # Sleep conservatively based on the chosen model's RPM
                    wait_s = engine.get_model_wait_seconds(st.session_state.selected_model)
                    time.sleep(wait_s)
                    attempt += 1
                    continue

                # Append the error only once per file
                if uploaded_file.name not in failed_files:
                    new_errors.append(error_message)
                    failed_files.add(uploaded_file.name)
                break

            if idx < total:
                progress.progress(
                    (idx - 1) / total,
                    text=f"Processing {uploaded_file.name} ({idx}/{total})... waiting for rate limit"
                )
                # Inter-file pacing according to selected model's RPM; silent to UI
                time.sleep(engine.get_model_wait_seconds(st.session_state.selected_model))

        except Exception as e:
            # Avoid duplicate error messages for the same file
            msg = f"'{uploaded_file.name}': unexpected error: {str(e)}"
            if uploaded_file.name not in failed_files:
                new_errors.append(msg)
                failed_files.add(uploaded_file.name)

        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass

    progress.progress(1.0, text="Extraction complete.")
    if st.session_state.stop_requested:
        progress.progress(1.0, text=f"Extraction stopped by user. {len(new_processed)} of {total} files processed.")
    st.session_state.extraction_running = False
    progress.empty()

    st.session_state.master_records = new_master_records
    st.session_state.publication_records = new_publication_records
    st.session_state.extraction_errors = new_errors
    st.session_state.processed_files = new_processed
    
    st.rerun()


# --- STEP 3: Extraction Results ---
if st.session_state.master_records or st.session_state.publication_records or st.session_state.extraction_errors:
    st.markdown("---")
    step_header("Extraction Results")

    master_df = master_records_to_dataframe(st.session_state.master_records)
    publication_df = publication_records_to_dataframe(st.session_state.publication_records)

    m_col1, m_col2, m_col3 = st.columns(3)
    with m_col1:
        st.metric("Faculty Records", len(master_df))
    with m_col2:
        st.metric("Publications Extracted", len(publication_df))
    with m_col3:
        st.metric("Files Processed", len(st.session_state.processed_files))
        
    # Show successfully processed files
    if st.session_state.processed_files:
        with st.expander("✅ View Processed Files", expanded=False):
            for f in st.session_state.processed_files:
                st.markdown(f"- {f}")

    # Show errors, if any
    if st.session_state.extraction_errors:
        with st.expander(f"⚠️ {len(st.session_state.extraction_errors)} File(s) Failed", expanded=True):
            for err in st.session_state.extraction_errors:
                st.error(err)

    tab1, tab2 = st.tabs(["Master Faculty Database", "Publication Details"])

    with tab1:
        st.dataframe(master_df, use_container_width=True, hide_index=True)

    with tab2:
        st.dataframe(publication_df, use_container_width=True, hide_index=True)


    # --- STEP 4: Download Options ---
    st.markdown("---")
    step_header("Download & Export")
    
    d_col1, d_col2, d_col3, d_col4 = st.columns([1, 1, 1, 1])
    
    with d_col1:
        master_bytes = build_single_sheet_workbook(master_df, "Master Faculty Database")
        st.download_button(
            label="📥 Master Sheet",
            data=master_bytes,
            file_name="master_faculty_database.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        
    with d_col2:
        pub_bytes = build_publication_workbook_per_faculty(master_df, publication_df)
        st.download_button(
            label="📥 Publication Sheet",
            data=pub_bytes,
            file_name="publication_details.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        
    with d_col3:
        combined_bytes = build_excel_workbook(master_df, publication_df)
        st.download_button(
            label="📥 Combined Workbook",
            data=combined_bytes,
            file_name="faculty_extraction_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        
    with d_col4:
        if st.button("🗑️ Clear Results", use_container_width=True):
            st.session_state.master_records = []
            st.session_state.publication_records = []
            st.session_state.extraction_errors = []
            st.session_state.processed_files = []
            st.rerun()
else:
    st.markdown(
        """
        <div class="empty-state">
            <div class="icon">📄</div>
            <div class="title">No results yet</div>
            <div class="sub">Upload resumes above and click "Extract Data" — your extracted faculty &amp; publication data will appear here.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )