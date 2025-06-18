import os
import re
import pandas as pd
import docx
import PyPDF2
import streamlit as st
import platform

# Windows-specific import for .doc file handling
if platform.system() == "Windows":
    import win32com.client

# === CONFIGURATION ===
BASE_PATH = os.path.dirname(__file__)
FAKE_COMPANY_LIST_PATH = os.path.join(BASE_PATH, "fake_companies.xlsx")
GENUINE_OUTPUT = os.path.join(BASE_PATH, "Genuine_Results.xlsx")
FAKE_OUTPUT = os.path.join(BASE_PATH, "Fake_Results.xlsx")
TEMP_RESUME_PATH = os.path.join(BASE_PATH, "temp_uploaded_resume")

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
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path)
        text = doc.Content.Text
        doc.Close()
        word.Quit()
        return text
    except Exception as e:
        st.error(f"Error reading DOC: {e}")
        return ""

# === LOAD FAKE COMPANIES (only Column A) ===
def load_fake_companies():
    df = pd.read_excel(FAKE_COMPANY_LIST_PATH, usecols=[0])  # Only first column
    return df.iloc[:, 0].dropna().astype(str).str.strip().str.lower().tolist()

# === FAKE CHECK LOGIC ===
def is_fake_resume(text, fake_companies):
    experience_keywords = [
        "currently", "working", "worked", "experience", "organization",
        "employer", "present", "joined", "serving", "professional", "company",
        "designation", "employment", "role", "project", "duration", "job",
        "position", "responsible", "team", "career", "contributed", "served", "assigned"
    ]

    def normalize(s):
        return re.sub(r"[^\w\s]", "", s).lower().strip()

    lines = text.splitlines()
    normalized_fakes = [normalize(fake) for fake in fake_companies]

    # Check lines with experience-related keywords
    keyword_lines = [line for line in lines if any(kw in line.lower() for kw in experience_keywords)]

    if keyword_lines:
        for line in keyword_lines:
            norm_line = normalize(line)
            for fake in normalized_fakes:
                if re.search(r'\b' + re.escape(fake) + r'\b', norm_line):
                    return True, fake, line.strip()
    else:
        # Check all lines word-by-word for exact match
        for line in lines:
            words_in_line = re.findall(r'\b[\w&.\-/]+\b', line.lower())
            for fake in normalized_fakes:
                fake_words = fake.split()
                for i in range(len(words_in_line) - len(fake_words) + 1):
                    if words_in_line[i:i + len(fake_words)] == fake_words:
                        return True, fake, line.strip()

    return False, "", ""

# === EXPORT RESULT TO EXCEL ===
def save_result_to_excel(resume_name, result, matched_company="", line=""):
    if result == "FAKE":
        df = pd.DataFrame([{
            "Resume": resume_name,
            "Matched Fake Company": matched_company,
            "Line": line,
            "Result": result
        }])
        if os.path.exists(FAKE_OUTPUT):
            existing = pd.read_excel(FAKE_OUTPUT)
            df = pd.concat([existing, df], ignore_index=True)
        df.to_excel(FAKE_OUTPUT, index=False)
    else:
        df = pd.DataFrame([{
            "Resume": resume_name,
            "Result": result
        }])
        if os.path.exists(GENUINE_OUTPUT):
            existing = pd.read_excel(GENUINE_OUTPUT)
            df = pd.concat([existing, df], ignore_index=True)
        df.to_excel(GENUINE_OUTPUT, index=False)

# === STREAMLIT UI ===
st.set_page_config(page_title="Resume Validator", layout="centered")
st.markdown("<h3 style='text-align: center;'>📄 Resume Validator – Fake Company Detection</h3>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Resume (.pdf, .docx, .doc)", type=["pdf", "docx", "doc"])

if uploaded_file is not None:
    with open(TEMP_RESUME_PATH, "wb") as f:
        f.write(uploaded_file.getbuffer())

    ext = uploaded_file.name.lower().split(".")[-1]
    if ext == "pdf":
        text = extract_text_from_pdf(TEMP_RESUME_PATH)
    elif ext == "docx":
        text = extract_text_from_docx(TEMP_RESUME_PATH)
    elif ext == "doc":
        text = extract_text_from_doc(TEMP_RESUME_PATH)
    else:
        st.error("Unsupported file format")
        st.stop()

    fake_companies = load_fake_companies()
    is_fake, matched_company, matched_line = is_fake_resume(text, fake_companies)

    if is_fake:
        st.error(f"❌ FAKE: Found '{matched_company}'")
        st.code(matched_line)
        save_result_to_excel(uploaded_file.name, "FAKE", matched_company, matched_line)
    else:
        st.success("✅ GENUINE Resume")
        save_result_to_excel(uploaded_file.name, "GENUINE")

    os.remove(TEMP_RESUME_PATH)
