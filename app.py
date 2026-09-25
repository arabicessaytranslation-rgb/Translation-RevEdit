import streamlit as st
import re
import io
import json
import time
import random
import docx
import fitz  # PyMuPDF
import difflib
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
st.set_page_config(page_title="12-Step AI Suite: Translate & Review", layout="wide")

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

# --- STATE MANAGEMENT ---
if "app_mode" not in st.session_state:
    st.session_state["app_mode"] = "Reviewer Mode"
    
if "input_method" not in st.session_state:
    st.session_state["input_method"] = "drive"

if 'processed_data' not in st.session_state:
    st.session_state['processed_data'] = None

if 'current_file_id' not in st.session_state:
    st.session_state['current_file_id'] = None

# ==========================================
# 2. INITIALIZE GOOGLE & AI SERVICES
# ==========================================
@st.cache_resource(ttl=300)
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
        docs_service = build('docs', 'v1', credentials=credentials, cache_discovery=False)
        drive_service = build('drive', 'v3', credentials=credentials, cache_discovery=False)
        sheets_service = build('sheets', 'v4', credentials=credentials, cache_discovery=False)
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
        types.SafetySetting(category=types.HarmCategory.HARM_CATEGORY_HARASSMENT, threshold=types.HarmBlockThreshold.BLOCK_NONE),
        types.SafetySetting(category=types.HarmCategory.HARM_CATEGORY_HATE_SPEECH, threshold=types.HarmBlockThreshold.BLOCK_NONE),
        types.SafetySetting(category=types.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold=types.HarmBlockThreshold.BLOCK_NONE),
        types.SafetySetting(category=types.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold=types.HarmBlockThreshold.BLOCK_NONE),
    ]

class TranslationResult(BaseModel):
    arabic_translation: str = Field(description="The finalized Arabic translation")
    glossary_notes: str = Field(description="Explanation of specific terms used based on context")

class ReviewResult(BaseModel):
    status: str = Field(description="Must be 'perfect', 'minor_edits', or 'major_rewrite'")
    suggested_arabic: str = Field(description="The finalized Arabic text")
    reasoning: str = Field(description="Detailed explanation of changes based on semantics and glossary")

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

    return ["gemini-3.8-flash", "gemini-3.6-flash", "gemini-3.5-flash"]

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

def generate_html_diff(original: str, suggested: str) -> str:
    if not original or original.startswith("[MISSING"):
        return "<div dir='rtl' style='color: #0369a1; background-color: #e0f2fe; padding: 10px; border-radius: 5px; text-align: right;'>✨ ترجمة تم توليدها بالكامل من المسرد (New Translation)</div>"
    if original.strip() == suggested.strip():
        return "<div dir='rtl' style='color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px; text-align: right;'>✨ لا توجد تعديلات (Perfect Match)</div>"
    
    diff = difflib.ndiff(original.split(), suggested.split())
    html = ["<div dir='rtl' style='line-height: 2; font-size: 18px; text-align: right; background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef;'>"]
    
    for word in diff:
        if word.startswith('- '):
            html.append(f"<span style='background-color: #ffcdd2; color: #b71c1c; text-decoration: line-through; padding: 2px 6px; margin: 0 2px; border-radius: 4px;'>{word[2:]}</span>")
        elif word.startswith('+ '):
            html.append(f"<span style='background-color: #c8e6c9; color: #1b5e20; font-weight: bold; padding: 2px 6px; margin: 0 2px; border-radius: 4px;'>{word[2:]}</span>")
        elif word.startswith('  '):
            html.append(f"<span style='color: #212529; margin: 0 2px;'>{word[2:]}</span>")
            
    html.append("</div>")
    return " ".join(html)

def _is_retryable(err_str: str) -> bool:
    return any(kw.lower() in err_str.lower() for kw in RETRYABLE_KEYWORDS)

def _backoff_sleep(attempt: int):
    delay = min(BASE_BACKOFF_SECONDS * (2 ** attempt), MAX_BACKOFF_SECONDS)
    delay += random.uniform(0, 1.0)
    time.sleep(delay)

def _call_gemini(model_name: str, prompt: str, schema_type):
    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=schema_type,
            safety_settings=safety_settings,
            temperature=0.2, 
        ),
    )
    if not response.text:
        feedback = getattr(response, 'prompt_feedback', 'Safety blocked')
        raise ValueError(f"Empty output from model: {feedback}")

    clean_text = response.text.replace("`" * 3 + "json", "").replace("`" * 3, "").strip()
    match = re.search(r'\{.*\}', clean_text, re.DOTALL)
    if match:
        clean_text = match.group(0)
    return json.loads(clean_text)

def translate_with_ai(english: str, glossary_text: str):
    if not GENAI_AVAILABLE or client is None:
        return {"arabic_translation": "", "glossary_notes": "SDK Error."}

    prompt = f"""
You are an expert bilingual translator specializing in 12-step recovery literature. 
Translate the English text into Arabic accurately, ensuring the tone remains clinical, professional, and non-moralizing.

GLOSSARY & CONTEXTUAL REASONING:
Do not perform blind word-for-word replacements. Actively understand semantic meaning.
- When pronouns like "it" appear (e.g., "it works") referring to concepts such as "the program", ensure the Arabic translation reflects the correct contextual noun/glossary term and grammatical gender.
- Apply the glossary terms naturally into the sentence flow.

GLOSSARY TERMS:
{glossary_text}

Translate:
English Source: "{english}"
"""
    active_models = get_fallback_models()
    last_error = None

    for model_name in active_models:
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                parsed = _call_gemini(model_name, prompt, TranslationResult)
                return {
                    "arabic_translation": parsed.get("arabic_translation", ""),
                    "glossary_notes": parsed.get("glossary_notes", f"Translated via {model_name}")
                }
            except Exception as e:
                err_str = f"{type(e).__name__} - {str(e)}"
                last_error = err_str
                if any(kw in err_str.lower() for kw in ("notfound", "404", "no longer available", "not found")):
                    break
                if _is_retryable(err_str) and attempt < MAX_RETRIES_PER_MODEL - 1:
                    _backoff_sleep(attempt)
                    continue
                break

    return {"arabic_translation": "", "glossary_notes": f"⚠️ API Error: {last_error}"}

def review_with_ai(english: str, arabic: str, glossary_text: str):
    if not GENAI_AVAILABLE or client is None:
        return {"status": "major_rewrite", "suggested_arabic": arabic, "reasoning": "SDK Error."}

    prompt = f"""
You are an expert bilingual editor specializing in 12-step recovery literature. 
Ensure the Arabic translation is accurate, clinical, professional, and grammatically sound.

GLOSSARY & CONTEXTUAL REASONING:
Do not perform blind word-for-word replacements.
- If English pronouns refer to specific concepts (e.g. "it works" referring to "the program"), verify the Arabic contextually renders this clearly.
- Respect recovery glossary terms and verify sentence flow.

GLOSSARY TERMS:
{glossary_text}

Review this pair:
English Source: "{english}"
Original Arabic: "{arabic}"

If the translation captures meaning and tone accurately, leave it as is. If it misses glossary nuance or sounds unnatural, provide the polished Arabic translation.
"""
    active_models = get_fallback_models()
    last_error = None

    for model_name in active_models:
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                parsed = _call_gemini(model_name, prompt, ReviewResult)
                return {
                    "status": parsed.get("status", "minor_edits"),
                    "suggested_arabic": parsed.get("suggested_arabic", arabic),
                    "reasoning": parsed.get("reasoning", "")
                }
            except Exception as e:
                err_str = f"{type(e).__name__} - {str(e)}"
                last_error = err_str
                if any(kw in err_str.lower() for kw in ("notfound", "404", "no longer available", "not found")):
                    break
                if _is_retryable(err_str) and attempt < MAX_RETRIES_PER_MODEL - 1:
                    _backoff_sleep(attempt)
                    continue
                break

    return {"status": "major_rewrite", "suggested_arabic": arabic, "reasoning": f"⚠️ API Error: {last_error}"}

# --- DOCUMENT PARSERS & ANOMALY DETECTOR ---
def extract_id(url: str):
    match = re.search(r"/(?:d|folders)/([a-zA-Z0-9-_]+)", url)
    if match: return match.group(1)
    match_param = re.search(r"id=([a-zA-Z0-9-_]+)", url)
    if match_param: return match_param.group(1)
    if re.match(r"^[a-zA-Z0-9-_]+$", url.strip()): return url.strip()
    return None

def extract_text_from_pdf(file_bytes: bytes):
    pdf_document = fitz.open(stream=file_bytes, filetype="pdf")
    blocks = []
    for page_num in range(len(pdf_document)):
        page_text = pdf_document.load_page(page_num).get_text("text")
        for line in page_text.split('\n'):
            line_str = line.strip()
            if bool(re.search(r'[a-zA-Z\u0600-\u06FF]', line_str)):
                blocks.append(line_str)
    return blocks

def extract_text_from_docx(file_bytes: bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    blocks = []
    
    for p in doc.element.body.iter():
        if p.tag.endswith('}p'):
            text = "".join(node.text for node in p.iter() if node.tag.endswith('}t') and node.text)
            clean_text = text.strip()
            if bool(re.search(r'[a-zA-Z\u0600-\u06FF]', clean_text)):
                if not blocks or blocks[-1] != clean_text:
                    blocks.append(clean_text)
                    
    return blocks

def _parse_docs_elements(elements):
    paras = []
    for elem in elements:
        if 'paragraph' in elem:
            para_text = ""
            for run in elem.get('paragraph', {}).get('elements', []):
                if 'textRun' in run:
                    para_text += run.get('textRun', {}).get('content', '')
            clean_text = para_text.strip()
            if bool(re.search(r'[a-zA-Z\u0600-\u06FF]', clean_text)): 
                paras.append(clean_text)
        elif 'table' in elem:
            for row in elem.get('table', {}).get('tableRows', []):
                for cell in row.get('tableCells', []):
                    paras.extend(_parse_docs_elements(cell.get('content', [])))
    return paras

def extract_text_from_drive(file_id: str, is_retry=False):
    try:
        docs_svc, drive_svc, _ = get_google_services()
        file_meta = drive_svc.files().get(fileId=file_id, fields="mimeType").execute()
        mime_type = file_meta.get("mimeType")

        if mime_type == 'application/vnd.google-apps.document':
            try:
                document = docs_svc.documents().get(documentId=file_id, includeTabsContent=True).execute()
            except Exception:
                document = docs_svc.documents().get(documentId=file_id).execute()

            all_paras = []
            
            def sweep_doc_obj(doc_obj):
                temp_paras = []
                temp_paras.extend(_parse_docs_elements(doc_obj.get('body', {}).get('content', [])))
                for footer in doc_obj.get('footers', {}).values():
                    temp_paras.extend(_parse_docs_elements(footer.get('content', [])))
                for header in doc_obj.get('headers', {}).values():
                    temp_paras.extend(_parse_docs_elements(header.get('content', [])))
                for footnote in doc_obj.get('footnotes', {}).values():
                    temp_paras.extend(_parse_docs_elements(footnote.get('content', [])))
                return temp_paras

            tabs = document.get('tabs', [])
            if tabs:
                for tab in tabs:
                    all_paras.extend(sweep_doc_obj(tab.get('documentTab', {})))
            else:
                all_paras.extend(sweep_doc_obj(document))

            return all_paras

        elif mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            request = drive_svc.files().get_media(fileId=file_id)
            file_bytes = request.execute()
            return extract_text_from_docx(file_bytes)
        else:
            st.error(f"Unsupported file type: {mime_type}")
            return None
    except Exception as e:
        err_str = str(e)
        if ("Broken pipe" in err_str or "Errno 32" in err_str) and not is_retry:
            get_google_services.clear()
            return extract_text_from_drive(file_id, is_retry=True)
        st.error(f"Could not read document from Drive. Error: {e}")
        return None

def smart_align_with_anomaly_detection(paragraphs: list):
    en_paras = []
    ar_paras = []
    
    for p in paragraphs:
        if bool(re.search(r'[\u0600-\u06FF]', p)):
            ar_paras.append(p)
        else:
            en_paras.append(p)

    aligned_segments = []
    max_len = max(len(en_paras), len(ar_paras))
    
    anomaly_msg = None
    if len(en_paras) != len(ar_paras):
        anomaly_msg = f"Count Mismatch: Found {len(en_paras)} English blocks vs {len(ar_paras)} Arabic blocks. Alignment may be shifted."

    for i in range(max_len):
        en_text = en_paras[i] if i < len(en_paras) else "[MISSING ENGLISH SOURCE]"
        ar_text = ar_paras[i] if i < len(ar_paras) else "[MISSING ARABIC TRANSLATION]"
        
        if en_text == "[MISSING ENGLISH SOURCE]":
            status = 'misaligned_ar'
        elif ar_text == "[MISSING ARABIC TRANSLATION]":
            status = 'misaligned_en'
        else:
            status = 'normal'

        aligned_segments.append({
            'id': i + 1,
            'status': status,
            'english': en_text,
            'arabic': ar_text,
            'anomaly': anomaly_msg
        })

    return aligned_segments

def prepend_google_doc_with_arabic(file_id: str, arabic_text: str):
    """Prepends Arabic text, applies RTL styling, and pushes original content to a new page."""
    try:
        docs_svc, _, _ = get_google_services()
        
        text_to_insert = arabic_text + "\n"
        len_text = len(text_to_insert)

        requests = [
            {
                'insertText': {
                    'location': {'index': 1},
                    'text': text_to_insert
                }
            },
            {
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': 1,
                        'endIndex': 1 + len_text
                    },
                    'paragraphStyle': {
                        'direction': 'RIGHT_TO_LEFT',
                        'alignment': 'START'
                    },
                    'fields': 'direction,alignment'
                }
            },
            {
                'insertPageBreak': {
                    'location': {'index': 1 + len_text}
                }
            }
        ]

        docs_svc.documents().batchUpdate(
            documentId=file_id,
            body={'requests': requests}
        ).execute()
        return True
    except Exception as e:
        st.error(f"Could not update Google Doc: {e}")
        return False

# ==========================================
# 4. DASHBOARD UI & PUSH BUTTON ROUTING
# ==========================================
st.title("⚙️ 12-Step AI Suite")

if not GENAI_AVAILABLE:
    st.error("🚨 Critical Dependency Missing: The `google-genai` package is not installed.")
    st.stop()

glossary_data = fetch_glossary()
active_models_list = get_fallback_models()

# --- MASTER CONTROL PANEL (PUSH BUTTONS) ---
with st.container(border=True):
    col_main, col_info = st.columns([2, 1])
    
    with col_main:
        st.markdown("### 🎛️ 1. Select Operating Mode")
        col_m1, col_m2 = st.columns(2)
        
        with col_m1:
            if st.button("🌍 Translator Mode", use_container_width=True, type="primary" if st.session_state["app_mode"] == "Translator Mode" else "secondary"):
                if st.session_state["app_mode"] != "Translator Mode":
                    st.session_state["app_mode"] = "Translator Mode"
                    st.session_state['processed_data'] = None
                    st.session_state['current_file_id'] = None
                    st.rerun()
                    
        with col_m2:
            if st.button("📝 Reviewer Mode", use_container_width=True, type="primary" if st.session_state["app_mode"] == "Reviewer Mode" else "secondary"):
                if st.session_state["app_mode"] != "Reviewer Mode":
                    st.session_state["app_mode"] = "Reviewer Mode"
                    st.session_state['processed_data'] = None
                    st.session_state['current_file_id'] = None
                    st.rerun()
        
        st.markdown("### 📥 2. Select Input Method")
        col_meth1, col_meth2 = st.columns(2)
        
        with col_meth1:
            if st.button("☁️ Method 1: Google Drive", use_container_width=True, type="primary" if st.session_state["input_method"] == "drive" else "secondary"):
                if st.session_state["input_method"] != "drive":
                    st.session_state["input_method"] = "drive"
                    st.session_state['processed_data'] = None
                    st.session_state['current_file_id'] = None
                    st.rerun()
                    
        with col_meth2:
            if st.button("📁 Method 2: File Upload", use_container_width=True, type="primary" if st.session_state["input_method"] == "upload" else "secondary"):
                if st.session_state["input_method"] != "upload":
                    st.session_state["input_method"] = "upload"
                    st.session_state['processed_data'] = None
                    st.session_state['current_file_id'] = None
                    st.rerun()

    with col_info:
        st.markdown("### 📊 System Status")
        st.caption(f"🤖 **Models:** `{', '.join(active_models_list)}`")
        
        if "No glossary connected" not in glossary_data and glossary_data != "":
            st.success(f"✅ **Glossary Connected:**\n`{GLOSSARY_RANGE.split('!')[0]}`")
        else:
            st.warning("⚠️ Glossary not active. Check Spreadsheet ID.")
            
        if st.session_state["app_mode"] == "Translator Mode":
            st.info("✨ **Translator Mode Active:** AI generates new translations from an English source document.")
        else:
            st.info("✨ **Reviewer Mode Active:** AI compares and corrects existing Arabic translations.")

st.divider()

# ==========================================
# FILE INGESTION LOGIC
# ==========================================
if st.session_state["input_method"] == "drive":
    if st.session_state["app_mode"] == "Translator Mode":
        st.write("Paste a Google Doc or Word URL containing **English text** to translate.")
    else:
        st.write("Paste a Google Doc or Word URL containing **English & Arabic text** to review.")
        
    doc_url = st.text_input("Paste Google Drive File URL Here:")

    if st.button("Load & Process from Drive") and doc_url:
        file_id = extract_id(doc_url)
        if not file_id:
            st.error("Invalid Google Drive URL.")
        else:
            st.info(f"Connecting to Document ID: {file_id}...")
            st.session_state['current_file_id'] = file_id 
            paras = extract_text_from_drive(file_id)

            if paras:
                if st.session_state["app_mode"] == "Translator Mode":
                    st.info(f"Extracted {len(paras)} segments. Translating via AI...")
                    progress_bar = st.progress(0)
                    processed_results = []
                    for i, en_text in enumerate(paras):
                        ai_result = translate_with_ai(en_text, glossary_data)
                        processed_results.append({
                            "id": i + 1,
                            "english": en_text,
                            "arabic_translation": ai_result.get("arabic_translation", ""),
                            "glossary_notes": ai_result.get("glossary_notes", "")
                        })
                        progress_bar.progress((i + 1) / len(paras))
                else:
                    segments = smart_align_with_anomaly_detection(paras)
                    num_en = sum(1 for p in paras if not bool(re.search(r'[\u0600-\u06FF]', p)))
                    num_ar = sum(1 for p in paras if bool(re.search(r'[\u0600-\u06FF]', p)))
                    st.info(f"Extracted: {num_en} English blocks & {num_ar} Arabic blocks. Processing review...")
                    progress_bar = st.progress(0)
                    processed_results = []

                    for i, item in enumerate(segments):
                        if item['status'] == 'normal':
                            ai_res = review_with_ai(item['english'], item['arabic'], glossary_data)
                            processed_results.append({
                                "id": item['id'],
                                "status": ai_res.get("status", "minor_edits"),
                                "english": item['english'],
                                "original_arabic": item['arabic'],
                                "suggested_arabic": ai_res.get("suggested_arabic", item['arabic']),
                                "reasoning": ai_res.get("reasoning", ""),
                                "anomaly": item['anomaly']
                            })
                        elif item['status'] == 'misaligned_en':
                            trans_res = translate_with_ai(item['english'], glossary_data)
                            processed_results.append({
                                "id": item['id'],
                                "status": "major_rewrite",
                                "english": item['english'],
                                "original_arabic": "[MISSING IN SOURCE DOCUMENT]",
                                "suggested_arabic": trans_res.get("arabic_translation", ""),
                                "reasoning": f"⚠️ Auto-translated from source using Sheet glossary.",
                                "anomaly": item['anomaly']
                            })
                        elif item['status'] == 'misaligned_ar':
                            processed_results.append({
                                "id": item['id'],
                                "status": "major_rewrite",
                                "english": "[MISSING ENGLISH SOURCE]",
                                "original_arabic": item['arabic'],
                                "suggested_arabic": item['arabic'],
                                "reasoning": f"⚠️ No English source detected for comparison.",
                                "anomaly": item['anomaly']
                            })
                        progress_bar.progress((i + 1) / len(segments))

                st.session_state['processed_data'] = processed_results
                st.rerun()

elif st.session_state["input_method"] == "upload":
    if st.session_state["app_mode"] == "Translator Mode":
        st.write("Upload an **English** Word document or PDF to translate.")
        file_en = st.file_uploader("Upload English Source File", type=["docx", "pdf"])
        
        if st.button("Process & Translate File") and file_en:
            paras = extract_text_from_pdf(file_en.read()) if file_en.name.endswith('.pdf') else extract_text_from_docx(file_en.read())
            if paras:
                st.info(f"Extracted {len(paras)} segments. Translating via AI...")
                progress_bar = st.progress(0)
                processed_results = []
                for i, en_text in enumerate(paras):
                    ai_result = translate_with_ai(en_text, glossary_data)
                    processed_results.append({
                        "id": i + 1,
                        "english": en_text,
                        "arabic_translation": ai_result.get("arabic_translation", ""),
                        "glossary_notes": ai_result.get("glossary_notes", "")
                    })
                    progress_bar.progress((i + 1) / len(paras))

                st.session_state['processed_data'] = processed_results
                st.rerun()
                
    else:
        st.write("Choose how you want to upload your document(s):")
        upload_type = st.radio("Upload Format:", ["Option A: Single Bilingual File (Contains both English & Arabic)", "Option B: Two Separate Files"], horizontal=True)

        if upload_type == "Option A: Single Bilingual File (Contains both English & Arabic)":
            file_bilingual = st.file_uploader("Upload Document (.docx or .pdf)", type=["docx", "pdf"], key="single_bilingual")
            if st.button("Process Bilingual File") and file_bilingual:
                paras = extract_text_from_pdf(file_bilingual.read()) if file_bilingual.name.endswith('.pdf') else extract_text_from_docx(file_bilingual.read())
                smart_pairs = smart_align_with_anomaly_detection(paras)
                
                num_en = sum(1 for p in paras if not bool(re.search(r'[\u0600-\u06FF]', p)))
                num_ar = sum(1 for p in paras if bool(re.search(r'[\u0600-\u06FF]', p)))
                st.info(f"Extracted: {num_en} English blocks & {num_ar} Arabic blocks. Generating review...")
                progress_bar = st.progress(0)
                processed_results = []
                
                for i, pair in enumerate(smart_pairs):
                    if pair['status'] == 'normal':
                        ai_res = review_with_ai(pair['english'], pair['arabic'], glossary_data)
                        processed_results.append({
                            "id": pair['id'],
                            "status": ai_res.get("status", "minor_edits"),
                            "english": pair['english'],
                            "original_arabic": pair['arabic'],
                            "suggested_arabic": ai_res.get("suggested_arabic", pair['arabic']),
                            "reasoning": ai_res.get("reasoning", ""),
                            "anomaly": pair['anomaly']
                        })
                    elif pair['status'] == 'misaligned_en':
                        trans_res = translate_with_ai(pair['english'], glossary_data)
                        processed_results.append({
                            "id": pair['id'],
                            "status": "major_rewrite",
                            "english": pair['english'],
                            "original_arabic": "[MISSING IN SOURCE DOCUMENT]",
                            "suggested_arabic": trans_res.get("arabic_translation", ""),
                            "reasoning": f"⚠️ Auto-translated from source using Sheet glossary.",
                            "anomaly": pair['anomaly']
                        })
                    elif pair['status'] == 'misaligned_ar':
                        processed_results.append({
                            "id": pair['id'],
                            "status": "major_rewrite",
                            "english": "[MISSING ENGLISH SOURCE]",
                            "original_arabic": pair['arabic'],
                            "suggested_arabic": pair['arabic'],
                            "reasoning": f"⚠️ No English source detected for comparison.",
                            "anomaly": pair['anomaly']
                        })
                    progress_bar.progress((i + 1) / len(smart_pairs))

                st.session_state['processed_data'] = processed_results
                st.rerun()

        else:
            colA, colB = st.columns(2)
            with colA:
                file_en = st.file_uploader("1. Upload English Source", type=["docx", "pdf"], key="sep_en")
            with colB:
                file_ar = st.file_uploader("2. Upload Arabic Translation", type=["docx", "pdf"], key="sep_ar")

            if st.button("Process Separate Files") and file_en and file_ar:
                en_paras = extract_text_from_pdf(file_en.read()) if file_en.name.endswith('.pdf') else extract_text_from_docx(file_en.read())
                ar_paras = extract_text_from_pdf(file_ar.read()) if file_ar.name.endswith('.pdf') else extract_text_from_docx(file_ar.read())
                
                smart_pairs = smart_align_with_anomaly_detection(en_paras + ar_paras)
                st.info(f"Extracted: {len(en_paras)} English blocks & {len(ar_paras)} Arabic blocks. Generating review...")
                progress_bar = st.progress(0)
                processed_results = []
                
                for i, pair in enumerate(smart_pairs):
                    if pair['status'] == 'normal':
                        ai_res = review_with_ai(pair['english'], pair['arabic'], glossary_data)
                        processed_results.append({
                            "id": pair['id'],
                            "status": ai_res.get("status", "minor_edits"),
                            "english": pair['english'],
                            "original_arabic": pair['arabic'],
                            "suggested_arabic": ai_res.get("suggested_arabic", pair['arabic']),
                            "reasoning": ai_res.get("reasoning", ""),
                            "anomaly": pair['anomaly']
                        })
                    elif pair['status'] == 'misaligned_en':
                        trans_res = translate_with_ai(pair['english'], glossary_data)
                        processed_results.append({
                            "id": pair['id'],
                            "status": "major_rewrite",
                            "english": pair['english'],
                            "original_arabic": "[MISSING IN SOURCE DOCUMENT]",
                            "suggested_arabic": trans_res.get("arabic_translation", ""),
                            "reasoning": f"⚠️ Auto-translated from source using Sheet glossary.",
                            "anomaly": pair['anomaly']
                        })
                    elif pair['status'] == 'misaligned_ar':
                        processed_results.append({
                            "id": pair['id'],
                            "status": "major_rewrite",
                            "english": "[MISSING ENGLISH SOURCE]",
                            "original_arabic": pair['arabic'],
                            "suggested_arabic": pair['arabic'],
                            "reasoning": f"⚠️ No English source detected for comparison.",
                            "anomaly": pair['anomaly']
                        })
                    progress_bar.progress((i + 1) / len(smart_pairs))

                st.session_state['processed_data'] = processed_results
                st.rerun()

# ==========================================
# 5. DYNAMIC OUTPUT GRID
# ==========================================
if st.session_state['processed_data']:
    st.divider()
    
    approved_count = 0
    finalized_arabic_list = []
    finalized_english_list = []

    if st.session_state["app_mode"] == "Translator Mode":
        st.subheader("Translation Editor")
        for i, item in enumerate(st.session_state['processed_data']):
            with st.container(border=True):
                st.markdown(f"### Segment {item['id']}")
                col_en, col_ar = st.columns(2)
                with col_en:
                    st.markdown("**English Source:**")
                    st.info(item['english'])
                    with st.expander("💡 AI Glossary & Context Notes", expanded=False):
                        st.caption(item['glossary_notes'])
                with col_ar:
                    st.markdown("**Final Translation Decision (Edit Here):**")
                    final_text = st.text_area("Final Arabic Translation", value=item['arabic_translation'], height=120, key=f"edit_ar_{i}", label_visibility="collapsed")
                
                is_approved = st.checkbox(f"✅ Approve Segment {item['id']}", key=f"approve_{i}", value=True)
                if is_approved:
                    approved_count += 1
                    finalized_arabic_list.append(final_text)
                    finalized_english_list.append(item['english'])

    else:
        st.subheader("Review Segments & Diff Visualizer")
        for i, item in enumerate(st.session_state['processed_data']):
            color = "🟢" if item['status'] == "perfect" else ("🟡" if item['status'] == "minor_edits" else "🔴")
            with st.container(border=True):
                st.markdown(f"### Segment {item['id']} | Status: {color} {item['status'].upper()}")

                if item.get('anomaly'):
                    st.warning(f"⚠️ **Structural Alert:** {item['anomaly']}")
                
                col_en, col_ar = st.columns(2)
                with col_en:
                    st.markdown("**English Source:**")
                    if item['english'] == "[MISSING ENGLISH SOURCE]":
                        st.error(item['english'])
                    else:
                        st.info(item['english'])

                with col_ar:
                    st.markdown("**Original Arabic Translation:**")
                    if item['original_arabic'] == "[MISSING IN SOURCE DOCUMENT]":
                        st.error(item['original_arabic'])
                    else:
                        st.markdown(f"<div dir='rtl' style='text-align: right; background-color: #e0f2fe; padding: 15px; border-radius: 8px; color: #0369a1;'>{item['original_arabic']}</div>", unsafe_allow_html=True)
                
                st.markdown("**Visual Changes (Red = Removed, Green = Added):**")
                diff_html = generate_html_diff(item['original_arabic'], item['suggested_arabic'])
                st.markdown(diff_html, unsafe_allow_html=True)
                
                with st.expander("💡 View AI Reasoning & Summary", expanded=(item.get('anomaly') is not None or item['status'] != 'perfect')):
                    st.markdown(item['reasoning'])

                st.markdown("**Final Decision (Edit if necessary):**")
                final_text = st.text_area("Final Arabic Translation", value=item['suggested_arabic'], height=120, key=f"edit_ar_{i}", label_visibility="collapsed")
                
                is_approved = st.checkbox(f"✅ Approve Segment {item['id']}", key=f"approve_{i}", value=(item['status'] == 'perfect'))
                if is_approved:
                    approved_count += 1
                    finalized_arabic_list.append(final_text)
                    finalized_english_list.append(item['english'])

    # --- COMPILED EXPORT ---
    total_segments = len(st.session_state['processed_data'])
    st.divider()
    st.write(f"### **Approved Segments: {approved_count} / {total_segments}**")

    if approved_count == total_segments:
        st.success("🎉 All segments approved! Copy the text blocks below, or push them directly to Google Drive.")
        st.subheader("📄 Finalized Text Blocks")

        final_arabic_text = "\n\n".join(finalized_arabic_list)
        final_english_text = "\n\n".join(finalized_english_list)

        col_final_ar, col_final_en = st.columns(2)
        with col_final_ar:
            st.text_area("Final Arabic Text (Select All and Copy)", value=final_arabic_text, height=400)
        with col_final_en:
            st.text_area("Final English Text (Select All and Copy)", value=final_english_text, height=400)

        # --- GOOGLE DRIVE PUSH INTEGRATION ---
        if st.session_state["input_method"] == "drive" and st.session_state.get('current_file_id'):
            st.divider()
            st.markdown("### ☁️ Sync to Google Drive")
            st.info("✨ **Action:** This will insert the finalized Arabic translation at the very top of your Google Doc, set the text direction to Right-To-Left (RTL), and insert a Page Break so your original formatted document is safely preserved on the pages below.")
            
            if st.button("🚀 Push Arabic to Top of Document", type="primary", use_container_width=True):
                with st.spinner("Pushing updates to Google Drive..."):
                    success = prepend_google_doc_with_arabic(
                        st.session_state['current_file_id'], 
                        final_arabic_text
                    )
                    if success:
                        st.success("✅ Document successfully updated in Google Drive!")
                        st.balloons()
    else:
        st.caption("You must check 'Approve this segment' on all segments above to compile the final text.")
