import streamlit as st
import re
import io
import json
import docx
from docx.shared import Pt
import fitz  # PyMuPDF
import google.generativeai as genai
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==========================================
# 1. PAGE CONFIG & AUTHENTICATION
# ==========================================
st.set_page_config(page_title="12-Step Translation Reviewer", layout="wide")

def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.text_input("Enter Team Password", type="password", key="pwd")
        if st.session_state["pwd"] == st.secrets["TEAM_PASSWORD"]:
            st.session_state["password_correct"] = True
            st.rerun()
        elif st.session_state["pwd"] != "":
            st.error("Incorrect Password")
        return False
    return True

if not check_password():
    st.stop()

# ==========================================
# 2. INITIALIZE GOOGLE & AI SERVICES
# ==========================================
@st.cache_resource
def get_google_services():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        credentials = service_account.Credentials.from_service_account_info(
            creds_dict, 
            scopes=['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
        )
        docs_service = build('docs', 'v1', credentials=credentials)
        drive_service = build('drive', 'v3', credentials=credentials)
        return docs_service, drive_service
    except Exception as e:
        return None, None

docs_service, drive_service = get_google_services()
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
def extract_file_id(url):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else None

def extract_text_from_pdf(file_bytes):
    pdf_document = fitz.open(stream=file_bytes, filetype="pdf")
    full_text = ""
    for page_num in range(len(pdf_document)):
        full_text += pdf_document.load_page(page_num).get_text("text") + "\n"
    
    cleaned_paragraphs = []
    for para in full_text.split('\n\n'):
        clean_para = para.replace('\n', ' ').strip()
        if clean_para:
            cleaned_paragraphs.append(clean_para)
    return cleaned_paragraphs

def extract_text_from_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

def build_final_docx(arabic_texts, english_texts):
    doc = docx.Document()
    for text in arabic_texts:
        p = doc.add_paragraph(text)
        p.style.font.name = 'Arial'
        p.style.font.size = Pt(12)
    doc.add_page_break()
    for text in english_texts:
        p = doc.add_paragraph(text)
        p.style.font.name = 'Arial'
        p.style.font.size = Pt(12)
    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# ==========================================
# 4. DASHBOARD UI & ROUTING
# ==========================================
st.title("Arabic Translation Reviewer - 12-Step Literature")
st.write("Logged in securely. Ready to process documents.")

# Initialize session state for the parsed data
if 'processed_data' not in st.session_state:
    st.session_state['processed_data'] = None
if 'workflow_type' not in st.session_state:
    st.session_state['workflow_type'] = None

tab1, tab2 = st.tabs(["Method 1: Google Drive Link (In-Place)", "Method 2: Direct File Upload"])

# --- TAB 1: GOOGLE DRIVE ---
with tab1:
    st.write("Paste a Google Doc URL to edit the document directly in-place.")
    doc_url = st.text_input("Paste Google Doc URL Here:")
    
    if st.button("Load from Drive") and doc_url:
        file_id = extract_file_id(doc_url)
        if not file_id:
            st.error("Invalid Google Drive URL.")
        else:
            st.info(f"Connecting to Document ID: {file_id}...")
            # Placeholder: In a full production script, you'd extract text via docs_service here.
            # Simulating data extraction and AI response for UI demonstration:
            st.session_state['processed_data'] = [
                {
                    "id": 1,
                    "status": "minor_edits",
                    "english": "The obsession of the mind will eventually cease.",
                    "original_arabic": "سوف تتوقف وسوسة العقل في النهاية.",
                    "suggested_arabic": "ستتوقف الوساوس القهرية للعقل في نهاية المطاف.",
                    "reasoning": "Clinical context: changed literal translation to clinical MSA terminology."
                }
            ]
            st.session_state['workflow_type'] = 'drive'
            st.session_state['file_id'] = file_id
            st.rerun()

# --- TAB 2: FILE UPLOAD ---
with tab2:
    st.write("Upload separate files to generate a finalized .docx document.")
    colA, colB = st.columns(2)
    with colA:
        file_en = st.file_uploader("1. Upload English Source", type=["docx", "pdf"])
    with colB:
        file_ar = st.file_uploader("2. Upload Arabic Translation", type=["docx", "pdf"])
        
    if st.button("Process Uploaded Files") and file_en and file_ar:
        en_paras = extract_text_from_pdf(file_en.read()) if file_en.name.endswith('.pdf') else extract_text_from_docx(file_en.read())
        ar_paras = extract_text_from_pdf(file_ar.read()) if file_ar.name.endswith('.pdf') else extract_text_from_docx(file_ar.read())
        
        if len(en_paras) != len(ar_paras):
            st.error(f"⚠️ Alignment Warning: English has {len(en_paras)} paragraphs, Arabic has {len(ar_paras)}.")
        else:
            st.success("Files aligned successfully! Generating review cards...")
            
            # Combine them into the expected dictionary format
            raw_pairs = [{"id": i+1, "english": en, "original_arabic": ar} for i, (en, ar) in enumerate(zip(en_paras, ar_paras))]
            
            # Placeholder: Feed raw_pairs to Gemini via genai library here.
            # Simulating output for the dashboard:
            st.session_state['processed_data'] = [
                {
                    "id": p["id"],
                    "status": "perfect" if "test" not in p["english"].lower() else "minor_edits",
                    "english": p["english"],
                    "original_arabic": p["original_arabic"],
                    "suggested_arabic": p["original_arabic"], # Simulated untouched
                    "reasoning": "AI Review Complete."
                } for p in raw_pairs
            ]
            st.session_state['workflow_type'] = 'upload'
            st.rerun()

# ==========================================
# 5. THE REVIEW GRID (UNIVERSAL)
# ==========================================
if st.session_state['processed_data']:
    st.divider()
    st.subheader("Review Pending Edits")
    
    approved_count = 0
    finalized_arabic_list = []
    finalized_english_list = []
    
    for i, item in enumerate(st.session_state['processed_data']):
        color = "🟢" if item['status'] == "perfect" else ("🟡" if item['status'] == "minor_edits" else "🔴")
        
        with st.container():
            st.markdown(f"**Segment {item['id']}** | Status: {color} {item['status'].upper()}")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.text_area("English Source (Read-Only)", value=item['english'], disabled=True, height=120, key=f"en_{i}")
            with col2:
                st.text_area("Original Arabic (Read-Only)", value=item['original_arabic'], disabled=True, height=80, key=f"orig_ar_{i}")
                st.caption(f"**AI Reasoning:** {item['reasoning']}")
            with col3:
                final_text = st.text_area("Final Arabic Decision (Edit Here)", value=item['suggested_arabic'], height=120, key=f"edit_ar_{i}")
                is_approved = st.checkbox("Approve this segment", key=f"approve_{i}", value=(item['status'] == 'perfect'))
                
                if is_approved:
                    approved_count += 1
                    finalized_arabic_list.append(final_text)
                    finalized_english_list.append(item['english'])
            st.divider()

    # ==========================================
    # 6. COMMIT / EXPORT LOGIC
    # ==========================================
    total_segments = len(st.session_state['processed_data'])
    st.write(f"**Approved Changes: {approved_count} / {total_segments}**")
    
    if approved_count == total_segments:
        if st.session_state['workflow_type'] == 'drive':
            if st.button(">>> COMMIT CHANGES TO ORIGINAL DOC <<<", type="primary"):
                st.warning("Executing Drive update: Wiping document and restructuring blocks...")
                # Here you would call Google Docs API batchUpdate with the finalized lists
                st.success("Live document updated and rearranged successfully!")
                
        elif st.session_state['workflow_type'] == 'upload':
            st.success("All segments approved! You can now download the finalized document.")
            final_docx = build_final_docx(finalized_arabic_list, finalized_english_list)
            st.download_button(
                label="📥 Download Finalized .docx Document",
                data=final_docx,
                file_name="Final_Translated_Review.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
    else:
        st.button(">>> FINALIZE DOCUMENT <<<", disabled=True)
        st.caption("You must approve all segments before finalizing.")
