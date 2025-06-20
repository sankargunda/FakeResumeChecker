import os
import re
import pandas as pd
import docx
import PyPDF2
import streamlit as st
import zipfile
import io
import base64
import time
import tempfile
import shutil
import subprocess


# === CONFIGURATION ===
BASE_PATH = os.path.dirname(__file__)
FAKE_COMPANY_LIST_PATH = os.path.join(BASE_PATH, "fake_companies.xlsx")
GENUINE_OUTPUT = os.path.join(BASE_PATH, "Genuine_Results.xlsx")
FAKE_OUTPUT = os.path.join(BASE_PATH, "Fake_Results.xlsx")
TEMP_DIR = os.path.join(BASE_PATH, "temp_files")
os.makedirs(TEMP_DIR, exist_ok=True)

# === TEXT EXTRACTORS ===
def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        st.error(f"Error reading DOCX: {e}")
        return ""

def extract_text_from_pdf(file_path):
    try:
        with open(file_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            return "\n".join([page.extract_text() or "" for page in reader.pages])
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return ""

def extract_text_from_doc(file_path):
    if shutil.which("soffice") is None:
        st.error("LibreOffice (soffice) is not installed or not in PATH. Please install LibreOffice to process .doc files.")
        return ""
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            subprocess.run([
                "soffice", "--headless", "--convert-to", "txt:Text", "--outdir", tmpdir, file_path
            ], check=True)
            base = os.path.splitext(os.path.basename(file_path))[0]
            converted_txt = os.path.join(tmpdir, base + ".txt")
            with open(converted_txt, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
    except Exception as e:
        st.error(f"Error extracting DOC with LibreOffice: {e}")
        return ""

# === LOAD FAKE COMPANIES FROM EXCEL (Only Column A) ===
def load_fake_companies():
    df = pd.read_excel(FAKE_COMPANY_LIST_PATH, usecols=[0])
    print(df.head())
    return df.iloc[:, 0].dropna().astype(str).str.strip().str.lower().tolist()

# === NORMALIZATION FUNCTION TO REMOVE PUNCTUATION & LOWERCASE ===
def normalize(s):
    return re.sub(r"[^\w\s]", "", s).lower().strip()

# === FAKE DETECTION LOGIC ===
def is_fake_resume(text, fake_companies):
    lines = text.splitlines()
    normalized_fakes = [normalize(fake) for fake in fake_companies]

    delimiters = [
        ',', ';', ' at ', ' with ', ' in ', '|', 'joined', 'organization',
        'experience', 'worked', 'working', 'currently', 'employer', 'company',
        'firm', 'served', 'project'
    ]

    def split_entities(line):
        for d in delimiters:
            line = line.replace(d, '|')
        return [e.strip() for e in line.split('|') if e.strip()]

    for line in lines:
        entities = split_entities(line)
        for entity in entities:
            norm_entity = normalize(entity)
            for fake in normalized_fakes:
                if norm_entity == fake or norm_entity.startswith(fake + ' '):
                    return True, fake, line.strip()
    return False, "", ""

# === SAVE RESULTS TO EXCEL ===
def save_result_to_excel(df, output_path):
    if os.path.exists(output_path):
        try:
            existing = pd.read_excel(output_path)
            df = pd.concat([existing, df], ignore_index=True)
        except zipfile.BadZipFile:
            pass
    df.to_excel(output_path, index=False)

# === VISUAL ENHANCEMENTS: CSS STYLES ===
# All custom styles are now included directly in main.py.
st.markdown("""
    <style>
    /* Main Title and Subtitle */
    .title-text {
        text-align: center;
        font-size: 42px;
        font-weight: bold;
        color: #2C5282;
        margin-bottom: 0.2em;
    }
    .subtitle-text {
        text-align: center;
        font-size: 20px;
        color: #718096;
        margin-bottom: 2em;
    }
    /* Upload label */
    label[data-testid="stFileUploaderLabel"] {
        color: #4A5568 !important;
        font-size: 1rem !important;
        font-weight: 500 !important;
    }
    /* Drag-and-drop area */
    section[data-testid="stFileUploadDropzone"] {
        background-color: #F7FAFC !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 0.5rem !important;
    }
    /* Text inside the dropzone */
    section[data-testid="stFileUploadDropzone"] * {
        color: #4A5568 !important;
    }
    /* "Browse files" button */
    button[title="Browse files"] {
        background-color: #FFFFFF !important;
        color: #4A5568 !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 0.5rem !important;
    }
    button[title="Browse files"]:hover {
        background-color: #F7FAFC !important;
        color: #2D3748 !important;
        border-color: #CBD5E0 !important;
    }
    /* Table and download link styles */
    .custom-table {
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: 16px;
        border-collapse: collapse;
        width: 100%;
        table-layout: auto;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    .custom-table th, .custom-table td {
        border: 1px solid #ddd;
        padding: 12px 10px;
        text-align: left;
        vertical-align: top;
        max-width: 320px;
        word-break: break-word;
        white-space: pre-line;
    }
    .custom-table th {
        background-color: transparent !important;
        color: #2C5282 !important;
        font-weight: bold !important;
        white-space: nowrap !important;
    }
    .custom-table td:hover, .custom-table th:hover, td:hover, th:hover {
        background-color: #C6F6D5 !important;
    }
    .custom-table tr:hover {
        background-color: #F7FAFC;
    }
    .tao-logo-absolute {
        position: fixed;
        top: 0;
        left: 0;
        width: 180px;
        z-index: 9999;
    }
    .download-zip-btn {
        background-color: #EBF8FF !important;
        color: #2C5282 !important;
        border: 1px solid #BEE3F8 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 10px 18px !important;
        font-size: 1rem !important;
        margin-top: 10px;
        margin-bottom: 10px;
        box-shadow: 0 2px 8px rgba(44,82,130,0.04);
        transition: background 0.2s, color 0.2s;
    }
    .download-zip-btn:hover {
        background-color: #BEE3F8 !important;
        color: #2C5282 !important;
    }
    /* Target the Streamlit download button by its key */
    button[data-testid="baseButton-download-zip-btn-real"] {
        background-color: #EBF8FF !important;
        color: #2C5282 !important;
        border: 1px solid #BEE3F8 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 10px 18px !important;
        font-size: 1rem !important;
        margin-top: 10px;
        margin-bottom: 10px;
        box-shadow: 0 2px 8px rgba(44,82,130,0.04);
        transition: background 0.2s, color 0.2s;
    }
    button[data-testid="baseButton-download-zip-btn-real"]:hover {
        background-color: #BEE3F8 !important;
        color: #2C5282 !important;
    }
    .genuine-title {
        font-size: 1.5rem;
        font-weight: bold;
        color: #2C5282;
        display: flex;
        align-items: center;
        margin-bottom: 0.5em;
    }
    </style>
    <img src='https://i.postimg.cc/GtzH6R0W/image.jpg' class='tao-logo-absolute' />
""", unsafe_allow_html=True)

# === Streamlit UI ===
st.markdown('<div class="title-text">Resume Validator</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle-text">Fake Company Detection</div>', unsafe_allow_html=True)
uploaded_files = st.file_uploader("Upload Resume(s)", type=["pdf", "docx", "doc"], accept_multiple_files=True)


if uploaded_files:
    fake_companies = load_fake_companies()
    fake_rows, genuine_rows = [], []

    for uploaded_file in uploaded_files:
        # Create a unique temp path using filename (safe)
        safe_filename = uploaded_file.name.replace(" ", "_")
        ext = safe_filename.split(".")[-1].lower()
        temp_file_path = os.path.join(TEMP_DIR, safe_filename)

        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Extract text based on extension
        if ext == "pdf":
            text = extract_text_from_pdf(temp_file_path)
        elif ext == "docx":
            text = extract_text_from_docx(temp_file_path)
        elif ext == "doc":
            text = extract_text_from_doc(temp_file_path)
        else:
            st.error(f"Unsupported file format: {uploaded_file.name}")
            continue

        is_fake, matched_company, matched_line = is_fake_resume(text, fake_companies)

        if is_fake:
            row = {
                "Resume": uploaded_file.name,
                "Matched Fake Company": matched_company,
                "Line": matched_line,
                "Result": "FAKE"
            }
            fake_rows.append(row)
        else:
            row = {
                "Resume": uploaded_file.name,
                "Result": "GENUINE"
            }
            genuine_rows.append(row)

        for _ in range(3):
            try:
                os.remove(temp_file_path)
                break
            except PermissionError:
                time.sleep(0.5)

    # === Display Fake Resumes Table ===
    if fake_rows:
        df_fake = pd.DataFrame(fake_rows)
        df_fake = df_fake[["Resume", "Result", "Matched Fake Company", "Line"]]
        st.markdown("### ❌ Fake Resumes")

        table_html = (
            "<table class='custom-table'>"
            "<tr>"
            "<th>Resume</th>"
            "<th>Result</th>"
            "<th>Matched Fake Company</th>"
            "<th>Line</th>"
            "</tr>"
        )
        for _, row in df_fake.iterrows():
            resume = row['Resume']
            result = row['Result']
            fake_company = row['Matched Fake Company']
            line = row['Line']
            table_html += (
                f"<tr>"
                f"<td title='{resume}'>{resume}</td>"
                f"<td style='color:#2C5282;font-weight:bold;'>{result}</td>"
                f"<td title='{fake_company}'>{fake_company}</td>"
                f"<td title='{line}'>{line}</td>"
                f"</tr>"
            )
        table_html += "</table>"

        st.markdown(table_html, unsafe_allow_html=True)
        save_result_to_excel(df_fake, FAKE_OUTPUT)

    # === Display Genuine Resumes Table ===
    if genuine_rows:
        df_genuine = pd.DataFrame(genuine_rows)
        df_genuine = df_genuine[["Resume", "Result"]]
        df_genuine.insert(0, "S. No", range(1, len(df_genuine) + 1))
        st.markdown('<div class="genuine-title">✅ Genuine Resumes</div>', unsafe_allow_html=True)
        # Custom HTML table for consistent styling
        table_html = (
            "<table class='custom-table'>"
            "<tr>"
            "<th>S. No</th>"
            "<th>Resume</th>"
            "<th>Result</th>"
            "</tr>"
        )
        for i, row in df_genuine.iterrows():
            sno = row['S. No']
            resume = row['Resume']
            result = row['Result']
            table_html += (
                f"<tr>"
                f"<td>{sno}</td>"
                f"<td title='{resume}'>{resume}</td>"
                f"<td style='color:#2C5282;font-weight:bold;'>{result}</td>"
                f"</tr>"
            )
        table_html += "</table>"
        st.markdown(table_html, unsafe_allow_html=True)
        save_result_to_excel(df_genuine, GENUINE_OUTPUT)

        # --- CORRECTED DOWNLOAD LOGIC ---
        # Get a list of the actual uploaded file objects that are genuine
        genuine_names = {row["Resume"] for row in genuine_rows}
        genuine_files = [f for f in uploaded_files if f.name in genuine_names]

        if len(genuine_files) == 1:
            # Only one genuine resume: provide a plain download link (no button, no custom style)
            genuine_file = genuine_files[0]
            resume_name = genuine_file.name
            data = genuine_file.getvalue()
            b64 = base64.b64encode(data).decode()
            href = f'<a href="data:application/octet-stream;base64,{b64}" download="{resume_name}">{resume_name}</a>'
            st.markdown(href, unsafe_allow_html=True)
        elif len(genuine_files) > 1:
            # More than one: provide a ZIP download
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for genuine_file in genuine_files:
                    # Add file to the zip using its name and in-memory content
                    zip_file.writestr(genuine_file.name, genuine_file.getvalue())

            st.download_button(
                label="Download All Genuine Resumes as ZIP",
                data=zip_buffer,
                file_name="genuine_resumes.zip",
                mime="application/zip",
                key="download-zip-btn-real"
            ) 
