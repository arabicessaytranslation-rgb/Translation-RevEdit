import streamlit as st
import re
import io
import json
import time
import random
import docx
import fitz  # PyMuPDF
import difflib # NEW: Added for visual text comparison
from google.oauth2 import service_account
from googleapiclient.discovery import build
from pydantic import BaseModel, Field

# --- Google GenAI SDK ---
try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

# ==========================================
# 1. CONFIGURATION & SECRETS
# ==========================================
st.set_page_config(page_title="12-Step Translation Reviewer", layout="wide")

GLOSSARY_SPREADSHEET_ID = "1oc4TCY_iK9R7mBiXgb5rKWssjmrQywYg6UpOBXx8pUQ"
GLOSSARY_RANGE = "'المصطلحات'!C:D"

MAX_RETRIES_PER_MODEL = 3
BASE_BACKOFF_SECONDS = 2.0
MAX_BACKOFF_SECONDS = 15.0
RETRYABLE_KEYWORDS = (
    "503", "500", "ServiceUnavailable", "service_unavailable",
    "high demand", "UNAVAILABLE", "429", "ResourceExhausted",
    "DeadlineExceeded", "timeout", "Quota",
)

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

@st.cache_resource
def get_genai_client():
    if not GENAI_AVAILABLE:
        return None
    return genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

client = get_genai_client()

safety_settings = []
if GENAI_AVAILABLE:
    safety_settings = [
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_HARASSMENT,
            threshold=types.HarmBlockThreshold.BLOCK_NONE,
        ),
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_HATE_SPEECH,
            threshold=types.HarmBlockThreshold.BLOCK_NONE,
        ),
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
            threshold=types.HarmBlockThreshold.BLOCK_NONE,
        ),
        types.SafetySetting(
            category=types.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
            threshold=types.HarmBlockThreshold.BLOCK_NONE,
        ),
    ]

class ReviewResult(BaseModel):
    status: str = Field(description="Must be 'perfect', 'minor_edits', or 'major_rewrite'")
    suggested_arabic: str = Field(description="The finalized Arabic text")
    reasoning: str = Field(description="Detailed explanation of what was changed and why")

@st.cache_resource(ttl=3600)
def get_fallback_models():
    secret_model = st.secrets.get("ACTIVE_MODEL", "").strip()
    if secret_model:
        return [secret_model]

    if GENAI_AVAILABLE and client is not None:
        try:
            available_flash_models = []
            for m in client.models.list():
                clean_name = m.name.replace("models/", "")
                if "flash" in clean_name.lower() and not any(tag in clean_name.lower() for tag in ["legacy", "embed", "imagen"]):
                    available_flash_models.append(clean_name)
            
            available_flash_models.sort(reverse=True)
            if available_flash_models:
                return available_flash_models
        except Exception:
            pass

    return [
        "gemini-3.8-flash",
        "gemini-3.6-flash",
        "gemini-3.5-flash",
    ]

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
        st.warning(f"Could not fetch glossary. Error: {e}")
        return ""

# NEW: Advanced Semantic Prompting
def _build_prompt(english: str, arabic: str, glossary_text: str) -> str:
    return f"""
You are an expert bilingual editor specializing in 12-step recovery literature. 
Your goal is to ensure the Arabic translation is accurate, clinical, professional, and flows naturally.

GLOSSARY & CONTEXTUAL REASONING (CRITICAL INSTRUCTIONS):
Do not perform blind word-for-word replacements. You must actively understand semantic meaning.
- If the English uses a pronoun (e.g., "it works") referring to a known concept (e.g., "the program"), ensure the Arabic translation reflects the correct contextual noun/glossary term and gender.
- Adapt to plurals, verb conjugations, and synonyms intelligently. 
- The glossary is a conceptual guide to maintain tone, not just a strict dictionary. Look at the whole sentence before making changes.

GLOSSARY TERMS:
{glossary_text}

Review this specific translation pair:
English Source: "{english}"
Original Arabic Translation: "{arabic}"

Analyze the text. If the original Arabic captures the meaning and tone perfectly, leave it unchanged. If it misses nuances, misapplies glossary concepts, or sounds unnatural, provide the corrected Arabic text.
"""

# NEW: Visual Diff Generator for Arabic text
def generate_html_diff(original: str, suggested: str) -> str:
    if original.strip() == suggested.strip():
        return "<div dir='rtl' style='color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px; text-align: right;'>✨ لا توجد تعديلات (Perfect Translation)</div>"
    
    diff = difflib.ndiff(original.split(), suggested.split())
    html = ["<div dir='rtl' style='line-height: 2; font-size: 18px; text-align: right; background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef;'>"]
    
    for word in diff:
        if word.startswith('- '):
            # Red strikethrough for deleted words
            html.append(f"<span style='background-color: #ffcdd2; color: #b71c1c; text-decoration: line-through; padding: 2px 6px; margin: 0 2px; border-radius: 4px;'>{word[2:]}</span>")
        elif word.startswith('+ '):
            # Green highlight for added words
            html.append(f"<span style='background-color: #c8e6c9; color: #1b5e20; font-weight: bold; padding: 2px 6px; margin: 0 2px; border-radius: 4px;'>{word[2:]}</span>")
        elif word.startswith('  '):
            # Standard text for unchanged words
            html.append(f"<span style='color: #212529; margin: 0 2px;'>{word[2:]}</span>")
            
    html.append("</div>")
    return " ".join(html)

def _is_retryable(err_str: str) -> bool:
    return any(kw.lower() in err_str.lower() for kw in RETRYABLE_KEYWORDS)

def _backoff_sleep(attempt: int):
    delay = min(BASE_BACKOFF_SECONDS * (2 ** attempt), MAX_BACKOFF_SECONDS)
    delay += random.uniform(0, 1.0)
    time.sleep(delay)

def _call_gemini(model_name: str, prompt: str):
    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ReviewResult,
            safety_settings=safety_settings,
            temperature=0.2, 
        ),
    )
    if not response.text:
        feedback = getattr(response, 'prompt_feedback', 'Safety blocked')
        raise ValueError(f"Empty output from model: {feedback}")

    # Switched to single quotes to prevent copy-paste line breaks
    clean_text = response.text.replace('```json', '').replace('```', '').strip()
    match = re.search(r'\{.*\}', clean_text, re.DOTALL)
    if match:
        clean_text = match.group(0)
    return json.loads(clean_text)

def review_with_ai(english: str, arabic: str, glossary_text: str):
    if not GENAI_AVAILABLE or client is None:
        return {"status": "major_rewrite", "suggested_arabic": arabic,
                "reasoning": "google-genai SDK not available. Add 'google-genai' to requirements.txt."}

    if "[MISSING" in english or "[MISSING" in arabic:
        return {"status": "major_rewrite", "suggested_arabic": arabic,
                "reasoning": "Alignment mismatch detected. Manual input required."}

    prompt = _build_prompt(english, arabic, glossary_text)
    active_models = get_fallback_models()
    last_error = None

    for model_name in active_models:
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                parsed = _call_gemini(model_name, prompt)
                return {
                    "status": parsed.get("status", "minor_edits"),
                    "suggested_arabic": parsed.get("suggested_arabic", arabic),
                    "reasoning": parsed.get("reasoning", "Reviewed successfully.")
                }
            except Exception as e:
                err_str = f"{type(e).__name__} - {str(e)}"
                last_error = err_str

                if any(kw in err_str.lower() for kw in ("notfound", "404", "no longer available", "not found", "not supported")):
                    break

                if _is_retryable(err_str) and attempt < MAX_RETRIES_PER_MODEL - 1:
                    _backoff_sleep(attempt)
                    continue

                break

    return {
        "status": "major_rewrite",
        "suggested_arabic": arabic,
        "reasoning": f"⚠️ All model attempts failed. Last error: {last_error}",
    }

def extract_id(url: str):
    match = re.search(r"/(?:d|folders)/([a-zA-Z0-9-_]+)", url)
    if match: return match.group(1)
    match_param = re.search(r"id=([a-zA-Z0-9-_]+)", url)
    if match_param: return match_param.group(1)
    if re.match(r"^[a-zA-Z0-9-_]+$", url.strip()): return url.strip()
    return None

def extract_text_from_pdf(file_bytes: bytes):
    pdf_document = fitz.open(stream=file_bytes, filetype="pdf")
    full_text = ""
    for page_num in range(len(pdf_document)):
        full_text += pdf_document.load_page(page_num).get_text("text") + "\n"
    return [p.replace('\n', ' ').strip() for p in full_text.split('\n\n') if p.replace('\n', ' ').strip()]

def extract_text_from_docx(file_bytes: bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

def extract_text_from_drive(file_id: str):
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
            st.error(f"Unsupported file type: {mime_type}")
            return None
    except Exception as e:
        st.error(f"Could not read document from Drive. Ensure the bot is an Editor. Error: {e}")
        return None

def smart_align_mixed_text(paragraphs: list):
    pairs = []
    current_en = []
    current_ar = []
    current_state = 'en'

    for p in paragraphs:
        has_arabic = bool(re.search(r'[\u0600-\u06FF]', p))
        if not has_arabic:
            if current_state == 'ar':
                pairs.append({"english": "\n".join(current_en), "arabic": "\n".join(current_ar)})
                current_en = []
                current_ar = []
                current_state = 'en'
            current_en.append(p)
        else:
            current_state = 'ar'
            current_ar.append(p)

    if current_en or current_ar:
        pairs.append({"english": "\n".join(current_en), "arabic": "\n".join(current_ar)})
    return pairs

def smart_align_separate_files(en_paras: list, ar_paras: list):
    pairs = []
    max_len = max(len(en_paras), len(ar_paras))
    for i in range(max_len):
        en = en_paras[i] if i < len(en_paras) else "[MISSING ENGLISH SOURCE]"
        ar = ar_paras[i] if i < len(ar_paras) else "[MISSING ARABIC TRANSLATION]"
        pairs.append({"english": en, "arabic": ar})
    return pairs

# ==========================================
# 4. DASHBOARD UI & ROUTING
# ==========================================
st.title("Arabic Translation Reviewer - 12-Step Literature")

if not GENAI_AVAILABLE:
    st.error("🚨 Critical Dependency Missing: The `google-genai` package is not installed. Please add it to `requirements.txt`.")
    st.stop()

glossary_data = fetch_glossary()
active_models_list = get_fallback_models()
st.caption(f"Active AI Routing Models: `{', '.join(active_models_list)}`")

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
                smart_pairs = smart_align_mixed_text(paras)
                st.info(f"File read successfully! Smart detection grouped text into {len(smart_pairs)} translation segments. Processing with AI...")
                progress_bar = st.progress(0)
                processed_results = []
                total = len(smart_pairs)

                for i, pair in enumerate(smart_pairs):
                    en, ar = pair['english'], pair['arabic']
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

        smart_pairs = smart_align_separate_files(en_paras, ar_paras)
        if len(en_paras) != len(ar_paras):
            st.warning(f"⚠️ Alignment Mismatch Detected: {len(en_paras)} English blocks vs {len(ar_paras)} Arabic blocks. Proceeding with missing tags...")
        else:
            st.info("Files aligned! Sending to AI for review...")

        progress_bar = st.progress(0)
        processed_results = []
        total = len(smart_pairs)

        for i, pair in enumerate(smart_pairs):
            en, ar = pair['english'], pair['arabic']
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
# 5. REVIEW INTERFACE & EXPORT BLOCKS (NEW UI)
# ==========================================
if st.session_state['processed_data']:
    st.divider()
    st.subheader("Review Segments & Diff Visualizer")

    approved_count = 0
    finalized_arabic_list = []
    finalized_english_list = []

    for i, item in enumerate(st.session_state['processed_data']):
        color = "🟢" if item['status'] == "perfect" else ("🟡" if item['status'] == "minor_edits" else "🔴")

        with st.container(border=True):
            st.markdown(f"### Segment {item['id']} | Status: {color} {item['status'].upper()}")
            
            # Row 1: English Source vs Original Arabic Side-by-Side
            col_en, col_ar = st.columns(2)
            with col_en:
                st.markdown("**English Source:**")
                st.info(item['english'])
            with col_ar:
                st.markdown("**Original Arabic Translation:**")
                st.info(f"<div dir='rtl' style='text-align: right;'>{item['original_arabic']}</div>", unsafe_allow_html=True)
            
            # Row 2: Visual Diff Highlights
            st.markdown("**Visual Changes (Red = Removed, Green = Added):**")
            diff_html = generate_html_diff(item['original_arabic'], item['suggested_arabic'])
            st.markdown(diff_html, unsafe_allow_html=True)
            
            # Row 3: Reasoning
            with st.expander("💡 View AI Reasoning & Summary", expanded=(item['status'] != 'perfect')):
                st.markdown(item['reasoning'])

            # Row 4: Final Editable Decision
            st.markdown("**Final Decision (Edit if necessary):**")
            final_text = st.text_area(
                "Final Arabic Translation", 
                value=item['suggested_arabic'], 
                height=120, 
                key=f"edit_ar_{i}", 
                label_visibility="collapsed"
            )
            
            is_approved = st.checkbox(f"✅ Approve Segment {item['id']}", key=f"approve_{i}", value=(item['status'] == 'perfect'))

            if is_approved:
                approved_count += 1
                finalized_arabic_list.append(final_text)
                finalized_english_list.append(item['english'])

    total_segments = len(st.session_state['processed_data'])
    st.divider()
    st.write(f"### **Approved Changes: {approved_count} / {total_segments}**")

    if approved_count == total_segments:
        st.success("🎉 All segments approved! Copy the finalized blocks below directly into your master document.")
        st.subheader("📄 Finalized Text Blocks")

        final_arabic_text = "\n\n".join(finalized_arabic_list)
        final_english_text = "\n\n".join(finalized_english_list)

        col_final_ar, col_final_en = st.columns(2)
        with col_final_ar:
            st.text_area("Final Arabic Text (Select All and Copy)", value=final_arabic_text, height=400)
        with col_final_en:
            st.text_area("Final English Text (Select All and Copy)", value=final_english_text, height=400)
    else:
        st.caption("You must check 'Approve this segment' on all segments above to compile the final text.")
