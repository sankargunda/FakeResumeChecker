import os
import re
import pandas as pd
import docx
import PyPDF2
import streamlit as st
import platform

# 💬 Optional Windows-only support for `.doc` files
if platform.system() == "Windows":
    import win32com.client

# === CONFIGURATION ===
# 💬 Define important paths used for input/output and uploaded file temp location
BASE_PATH = os.path.dirname(__file__)
RESUME_FOLDER = os.path.join(BASE_PATH, "resumes")
FAKE_COMPANY_LIST_PATH = os.path.join(BASE_PATH, "fake_companies.xlsx")
GENUINE_OUTPUT = os.path.join(BASE_PATH, "Genuine_Results.xlsx")
FAKE_OUTPUT = os.path.join(BASE_PATH, "Fake_Results.xlsx")
TEMP_RESUME_PATH = os.path.join(BASE_PATH, "temp_uploaded_resume")

# === HELPER FUNCTIONS ===

# 💬 Extract text from DOCX format
def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        st.error(f"Error reading DOCX: {e}")
        return ""

# 💬 Extract text from PDF format
def extract_text_from_pdf(file_path):
    try:
        with open(file_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            return "\n".join([page.extract_text() or "" for page in reader.pages])
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return ""

# 💬 Extract text from DOC (old MS Word format) – Windows only
def extract_text_from_doc(file_path):
    if platform.system() != "Windows":
        st.warning("Skipping .doc file: Not supported on Streamlit Cloud.")
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

# 💬 Check for full match of any fake company name in the resume text
def is_fake_resume(text, fake_companies):
    lines = text.splitlines()
    for line in lines:
        words_in_line = re.findall(r'\b\w[\w&.\-/]*\b', line.lower())
        for fake in fake_companies:
            if fake.lower() in [" ".join(words_in_line[i:i+len(fake.split())]) for i in range(len(words_in_line))]:
                return True, fake, line.strip()
    return False, "", ""

# 💬 Load fake company names from Excel
def load_fake_companies():
    df = pd.read_excel(FAKE_COMPANY_LIST_PATH)
    return df.iloc[:, 0].dropna().astype(str).str.strip().str.lower().tolist()

# 💬 Save results to appropriate Excel (Fake/Genuine)
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
st.set_page_config(page_title="Resume Screening – Company Legitimacy Check", layout="centered")
st.markdown("<h3 style='text-align: center;'>📄 Resume Validator – Fake Company Detection</h3>", unsafe_allow_html=True)


# 💬 Upload box for the resume file
uploaded_file = st.file_uploader("Upload Resume (.pdf, .docx, .doc)", type=["pdf", "docx", "doc"])

if uploaded_file is not None:
    # 💬 Save the uploaded file temporarily
    with open(TEMP_RESUME_PATH, "wb") as f:
        f.write(uploaded_file.getbuffer())

    ext = uploaded_file.name.lower().split(".")[-1]

    # 💬 Extract text based on file extension
    if ext == "pdf":
        text = extract_text_from_pdf(TEMP_RESUME_PATH)
    elif ext == "docx":
        text = extract_text_from_docx(TEMP_RESUME_PATH)
    elif ext == "doc":
        text = extract_text_from_doc(TEMP_RESUME_PATH)
    else:
        st.error("Unsupported file format")
        st.stop()

    # 💬 Load fake companies and validate resume
    fake_companies = load_fake_companies()
    is_fake, matched_company, line = is_fake_resume(text, fake_companies)

    # 💬 Show results on the UI and save to Excel
    if is_fake:
        st.error(f"❌ FAKE: Found '{matched_company}'")
        st.code(line)
        save_result_to_excel(uploaded_file.name, "FAKE", matched_company, line)
    else:
        st.success("✅ GENUINE Resume")
        save_result_to_excel(uploaded_file.name, "GENUINE")

    # 💬 Delete the temporary resume
    os.remove(TEMP_RESUME_PATH)
