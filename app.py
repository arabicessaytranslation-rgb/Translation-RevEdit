import streamlit as st
import re
import io
import json
import docx
import fitz  # PyMuPDF
import google.generativeai as genai
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==========================================
# 1. CONFIGURATION & SECRETS
# ==========================================
st.set_page_config(page_title="12-Step Translation Reviewer", layout="wide")

GLOSSARY_SPREADSHEET_ID = "1oc4TCY_iK9R7mBiXgb5rKWssjmrQywYg6UpOBXx8pUQ"
GLOSSARY_RANGE = "'المصطلحات'!C:D" 

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
            scopes=[
                'https://www.googleapis.com/auth/documents', 
                'https://www.googleapis.com/auth/drive',
                'https://www.googleapis.com/auth/spreadsheets.readonly'
            ]
        )
        docs_service = build('docs', 'v1', credentials=credentials)
        drive_service = build('drive', 'v3', credentials=credentials)
        sheets_service = build('sheets', 'v4', credentials=credentials)
        return docs_service, drive_service, sheets_service
    except Exception as e:
        st.error(f"Google Auth Error: {e}")
        return None, None, None

docs_service, drive_service, sheets_service = get_google_services()
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

generation_config = {"response_mime_type": "application/json"}
model = genai.GenerativeModel('gemini-1.5-flash', generation_config=generation_config)

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
@st.cache_data(ttl=3600)
def fetch_glossary():
    try:
        sheet = sheets_service.spreadsheets()
        result = sheet.values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_RANGE).execute()
        values = result.get('values', [])
        
        glossary_string = "12-Step Glossary:\n"
        for row in values:
            if len(row) >= 2:
                glossary_string += f"- {row[0]} -> {row[1]}\n"
        return glossary_string
    except Exception as e:
        st.warning(f"Could not fetch glossary. Please check the ID, Tab Name, and sharing permissions. Error: {e}")
        return ""

def review_with_ai(english, arabic, glossary_text):
    prompt = f"""
    You are an expert translator specializing in 12-step recovery literature. 
    Your tone must be clinical, professional, and non-moralizing.
    
    You MUST adhere to this glossary for specific terms:
    {glossary_text}
    
    Review this translation pair:
    English: "{english}"
    Arabic: "{arabic}"
    
    Respond ONLY with a JSON object using this exact format:
    {{
        "status": "perfect" OR "minor_edits" OR "major_rewrite",
        "suggested_arabic": "The finalized Arabic text (keep original if perfect, or provide the corrected version)",
        "reasoning": "Briefly explain why you made changes, or say 'Matches glossary/tone' if perfect."
    }}
    """
    try:
        response = model.generate_content(prompt)
        return json.loads(response.text)
    except Exception as e:
        return {"status": "major_rewrite", "suggested_arabic": arabic, "reasoning": "AI Error. Please review manually."}

def extract_id(url):
    """Extracts the ID from a Google Docs, Drive File, or Drive Folder URL."""
    match = re.search(r"/(?:d|folders)/([a-zA-Z0-9-_]+)", url)
    if match: return match.group(1)
    match_param = re.search(r"id=([a-zA-Z0-9-_]+)", url)
    if match_param: return match_param.group(1)
    if re.match(r"^[a-zA-Z0-9-_]+$", url.strip()): return url.strip()
    return None

def extract_text_from_pdf(file_bytes):
    pdf_document = fitz.open(stream=file_bytes, filetype="pdf")
    full_text = ""
    for page_num in range(len(pdf_document)):
        full_text += pdf_document.load_page(page_num).get_text("text") + "\n"
    return [p.replace('\n', ' ').strip() for p in full_text.split('\n\n') if p.replace('\n', ' ').strip()]

def extract_text_from_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

def extract_text_from_drive(file_id):
    try:
        file_meta = drive_service.files().get(fileId=file_id, fields="mimeType").execute()
        mime_type = file_meta.get("mimeType")
        
        if mime_type == 'application/vnd.google-apps.document':
            document = docs_service.documents().get(documentId=file_id).execute()
            paragraphs = []
            for element in document.get('body').get('content', []):
                if 'paragraph' in element:
                    para_text = ""
                    for elem in element.get('paragraph').get('elements', []):
                        if 'textRun' in elem:
                            para_text += elem.get('textRun').get('content')
                    clean_text = para_text.strip()
                    if clean_text: paragraphs.append(clean_text)
            return paragraphs
            
        elif mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            request = drive_service.files().get_media(fileId=file_id)
            file_bytes = request.execute()
            return extract_text_from_docx(file_bytes)
        else:
            st.error(f"Unsupported file type. Found: {mime_type}")
            return None
    except Exception as e:
        st.error(f"Could not read document from Drive. Ensure the bot is an Editor. Error: {e}")
        return None

# ==========================================
# 4. DASHBOARD UI & ROUTING
# ==========================================
st.title("Arabic Translation Reviewer - 12-Step Literature")
glossary_data = fetch_glossary()

if "No glossary connected" not in glossary_data and glossary_data != "":
    st.success(f"✅ Glossary connected successfully from tab: {GLOSSARY_RANGE.split('!')[0]}")
else:
    st.warning("⚠️ Glossary not active. Check Spreadsheet ID and Tab Name.")

if 'processed_data' not in st.session_state:
    st.session_state['processed_data'] = None

tab1, tab2 = st.tabs(["Method 1: Google Drive Link (In-Place)", "Method 2: Direct File Upload"])

# --- TAB 1: GOOGLE DRIVE ---
with tab1:
    st.write("Paste a Google Doc or Word URL to review the document.")
    doc_url = st.text_input("Paste Google Drive File URL Here:")
    
    if st.button("Load from Drive") and doc_url:
        file_id = extract_id(doc_url)
        if not file_id:
            st.error("Invalid Google Drive URL.")
        else:
            st.info(f"Connecting to Document ID: {file_id}...")
            paras = extract_text_from_drive(file_id)
            
            if paras:
                en_paras = paras[0::2]
                ar_paras = paras[1::2]
                
                if len(en_paras) != len(ar_paras):
                    st.error(f"⚠️ Alignment Warning: Found {len(en_paras)} English paragraphs and {len(ar_paras)} Arabic paragraphs. The document must alternate exactly.")
                else:
                    st.info("File read successfully! Sending to AI for review. This may take a moment...")
                    progress_bar = st.progress(0)
                    processed_results = []
                    total = len(en_paras)
                    
                    for i, (en, ar) in enumerate(zip(en_paras, ar_paras)):
                        ai_result = review_with_ai(en, ar, glossary_data)
                        processed_results.append({
                            "id": i + 1,
                            "status": ai_result.get("status", "minor_edits"),
                            "english": en,
                            "original_arabic": ar,
                            "suggested_arabic": ai_result.get("suggested_arabic", ar),
                            "reasoning": ai_result.get("reasoning", "Review complete.")
                        })
                        progress_bar.progress((i + 1) / total)
                    
                    st.session_state['processed_data'] = processed_results
                    st.rerun()

# --- TAB 2: FILE UPLOAD ---
with tab2:
    st.write("Upload separate files to review the translation.")
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
            st.info("Files aligned! Sending to AI for review. This may take a moment...")
            progress_bar = st.progress(0)
            processed_results = []
            total = len(en_paras)
            
            for i, (en, ar) in enumerate(zip(en_paras, ar_paras)):
                ai_result = review_with_ai(en, ar, glossary_data)
                processed_results.append({
                    "id": i + 1,
                    "status": ai_result.get("status", "minor_edits"),
                    "english": en,
                    "original_arabic": ar,
                    "suggested_arabic": ai_result.get("suggested_arabic", ar),
                    "reasoning": ai_result.get("reasoning", "Review complete.")
                })
                progress_bar.progress((i + 1) / total)
            
            st.session_state['processed_data'] = processed_results
            st.rerun()

# ==========================================
# 5. THE REVIEW GRID & EXPORT
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

    total_segments = len(st.session_state['processed_data'])
    st.write(f"**Approved Changes: {approved_count} / {total_segments}**")
    
    # NEW LOGIC: Display text blocks for manual copying instead of generating a file
    if approved_count == total_segments:
        st.success("🎉 All segments approved! You can now copy the finalized text below.")
        
        st.subheader("📄 Finalized Translations")
        
        final_arabic_text = "\n\n".join(finalized_arabic_list)
        final_english_text = "\n\n".join(finalized_english_list)
        
        col_final_ar, col_final_en = st.columns(2)
        with col_final_ar:
            st.text_area("Final Arabic Text (Select All and Copy)", value=final_arabic_text, height=400)
        with col_final_en:
            st.text_area("Final English Text (Select All and Copy)", value=final_english_text, height=400)
    else:
        st.caption("You must approve all segments above to reveal the final compiled text.")
