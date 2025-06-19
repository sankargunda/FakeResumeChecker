import os
import re
import pandas as pd
import docx
import PyPDF2
import streamlit as st
import platform
import zipfile
import base64
import io
import tempfile


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
    if platform.system() != "Windows":
        st.warning("Skipping .doc file: Not supported on non-Windows.")
        return ""
    try:
        import pythoncom
        import win32com.client
        import os
        with tempfile.NamedTemporaryFile(delete=False, suffix=".doc") as tmp_file:
            tmp_path = tmp_file.name
            tmp_file.write(uploaded_file.read())
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(tmp_path)
        text = doc.Content.Text
        doc.Close(False)
        word.Quit()
        os.remove(tmp_path)
        return text
    except Exception as e:
        st.error(f"Error reading DOC: {e}")
        return ""

# === LOAD FAKE COMPANIES FROM EXCEL (Only Column A) ===
def load_fake_companies():
    df = pd.read_excel(FAKE_COMPANY_LIST_PATH, usecols=[0])
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

# === VISUAL ENHANCEMENTS ===
st.set_page_config(page_title="Resume Validator", layout="centered")

st.markdown("""
    <style>
    /* Removed custom background color, will use Streamlit default */
    /* body, .stApp, .block-container, .main, .css-18e3th9, .css-1d391kg {
        background-color: #F0F9FF !important;
    } */
    .title-text {
        text-align: center;
        font-size: 42px;
        font-weight: bold;
        color: #0F172A;
        margin-bottom: 0.2em;
    }
    .subtitle-text {
        text-align: center;
        font-size: 20px;
        color: #334155;
        margin-bottom: 2em;
    }
    label[for^="file_uploader"] {
        color: #0F172A !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        display: block !important;
        margin-bottom: 1em !important;
    }
    [data-testid="stFileUploadDropzone"] {
        background: linear-gradient(145deg, #EFF6FF, #DBEAFE) !important;
        border: 2px dashed #3B82F6 !important;
        border-radius: 14px !important;
        padding: 25px !important;
    }
    [data-testid="stFileUploadDropzone"] * {
        color: #1E3A8A !important;
        font-size: 1rem !important;
        font-weight: 500 !important;
    }
    .custom-table th {
        background-color: #E0F2FE;
        color: #1E3A8A;
    }
    .custom-table tr:hover {
        background-color: #F0F9FF;
    }
    .tao-logo-absolute {
        position: fixed;
        top: 0;
        left: 0;
        width: 180px;
        z-index: 9999;
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
            text = extract_text_from_doc(uploaded_file)
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
            print(f"❌ {uploaded_file.name} FAKE -> {matched_company}")
        else:
            row = {
                "Resume": uploaded_file.name,
                "Result": "GENUINE"
            }
            genuine_rows.append(row)
            print(f"✅ {uploaded_file.name} GENUINE")

        os.remove(temp_file_path)

    # === Display Fake Resumes Table ===
    if fake_rows:
        df_fake = pd.DataFrame(fake_rows)
        df_fake = df_fake[["Resume", "Result", "Matched Fake Company", "Line"]]
        st.markdown("### ❌ Fake Resumes")

        table_html = (
            "<style>"
            ".custom-table {font-size: 16px; border-collapse: collapse; width: 100%; table-layout: auto; box-shadow: 0 2px 8px rgba(0,0,0,0.04);}"
            ".custom-table th, .custom-table td {border: 1px solid #ddd; padding: 12px 10px; text-align: left; vertical-align: top; max-width: 320px; word-break: break-word; white-space: pre-line;}"
            ".custom-table th {background-color: #f2f2f2; font-weight: bold; color: #22223b;}"
            ".custom-table tr:nth-child(even){background-color: #f9f9f9;}"
            ".custom-table tr:hover {background-color: #e3f2fd;}"
            ".custom-table td.ellipsis {overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 220px; cursor: pointer;}"
            "</style>"
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
                f"<td style='color:red;font-weight:bold;'>{result}</td>"
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
        st.markdown('<div class="section-title">✅ Genuine Resumes</div>', unsafe_allow_html=True)
        st.table(df_genuine)
        save_result_to_excel(df_genuine, GENUINE_OUTPUT)

        genuine_resume_files = [
            (row["Resume"], f) for row, f in zip(genuine_rows, uploaded_files) if row["Resume"] == f.name
        ]

        if len(genuine_resume_files) == 1:
            resume_name, file_obj = genuine_resume_files[0]
            data = file_obj.getbuffer()
            b64 = base64.b64encode(data).decode()
            href = f'''
                <a href="data:application/octet-stream;base64,{b64}" download="{resume_name}" class="simple-download-link">
                    ⬇️ Download {resume_name}
                </a>
            '''
            st.markdown(href, unsafe_allow_html=True)
        elif len(genuine_resume_files) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for resume_name, file_obj in genuine_resume_files:
                    zip_file.writestr(resume_name, file_obj.getbuffer())
            zip_buffer.seek(0)
            st.download_button(
                label="Download All Genuine Resumes as ZIP",
                data=zip_buffer,
                file_name="genuine_resumes.zip",
                mime="application/zip"
            )
