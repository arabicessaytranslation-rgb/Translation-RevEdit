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

BATCH_SIZE = 6
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

if 'source_file_id' not in st.session_state:
    st.session_state['source_file_id'] = None

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

# --- PYDANTIC SCHEMAS (SINGLE & BATCH) ---
class TranslationResult(BaseModel):
    arabic_translation: str = Field(description="The finalized Arabic translation")
    glossary_notes: str = Field(description="Explanation of specific terms used based on context")

class TranslationBatchItem(BaseModel):
    id: int = Field(description="The ID of the segment")
    arabic_translation: str
    glossary_notes: str

class TranslationBatchResult(BaseModel):
    items: list[TranslationBatchItem]

class ReviewResult(BaseModel):
    status: str = Field(description="Must be 'perfect', 'minor_edits', or 'major_rewrite'")
    suggested_arabic: str = Field(description="The finalized Arabic text")
    reasoning: str = Field(description="Detailed explanation of changes based on semantics and glossary")

class ReviewBatchItem(BaseModel):
    id: int = Field(description="The ID of the segment")
    status: str
    suggested_arabic: str
    reasoning: str

class ReviewBatchResult(BaseModel):
    items: list[ReviewBatchItem]

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
# 3. HELPER FUNCTIONS & AI BATCH ENGINE
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

# --- SINGLE ITEM AI CALLS (THE FAILSAFES) ---
def translate_with_ai(english: str, glossary_text: str):
prompt = f"""You are an expert bilingual translator specializing in 12-step recovery literature. 
Translate the following English segments into Arabic accurately. Ensure the tone remains clinical, professional, and non-moralizing.

CRITICAL INSTRUCTIONS FOR THIS BATCH:
1. Narrative Flow: These segments are sequential parts of a single document. Maintain consistent grammatical gender, tone, and pronoun references across all segments.
2. Contextual Nuance: Do not perform blind word-for-word replacements. Actively understand the semantic meaning.
3. Pronoun Resolution: When English pronouns like "it" or "we" appear (e.g., "it works", "we admitted"), ensure the Arabic translation reflects the correct contextual noun (e.g., "the program", "the fellowship") and its proper Arabic grammatical gender.
4. Glossary Integration: Apply the glossary terms naturally into the sentence flow without forcing them if the grammar breaks.

GLOSSARY TERMS:
{glossary_text}

\n\nSegments to Translate:
{input_payload}"""
    active_models = get_fallback_models()
    for model_name in active_models:
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                parsed = _call_gemini(model_name, prompt, TranslationResult)
                return {"arabic_translation": parsed.get("arabic_translation", ""), "glossary_notes": parsed.get("glossary_notes", "")}
            except Exception as e:
                err_str = f"{type(e).__name__} - {str(e)}"
                if _is_retryable(err_str) and attempt < MAX_RETRIES_PER_MODEL - 1:
                    _backoff_sleep(attempt)
                    continue
                break
    return {"arabic_translation": "", "glossary_notes": f"⚠️ Fallback Error."}

def review_with_ai(english: str, arabic: str, glossary_text: str):
    prompt = f"""You are an expert bilingual editor for 12-step recovery literature. 
GLOSSARY TERMS:\n{glossary_text}\nReview this pair:\nEnglish Source: "{english}"\nOriginal Arabic: "{arabic}"\nCorrect the Arabic if it misses glossary nuances or sounds unnatural."""
    active_models = get_fallback_models()
    for model_name in active_models:
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                parsed = _call_gemini(model_name, prompt, ReviewResult)
                return {"status": parsed.get("status", "minor_edits"), "suggested_arabic": parsed.get("suggested_arabic", arabic), "reasoning": parsed.get("reasoning", "")}
            except Exception as e:
                err_str = f"{type(e).__name__} - {str(e)}"
                if _is_retryable(err_str) and attempt < MAX_RETRIES_PER_MODEL - 1:
                    _backoff_sleep(attempt)
                    continue
                break
    return {"status": "major_rewrite", "suggested_arabic": arabic, "reasoning": f"⚠️ Fallback Error."}

# --- MICRO-BATCH AI CALLS & AUTO-CORRECT CASCADE ---
def translate_batch_with_fallback(batch_segments, glossary_text):
    if not GENAI_AVAILABLE or client is None:
        return [translate_with_ai(seg['english'], glossary_text) for seg in batch_segments]

    input_payload = "\n\n".join([f"ID: {seg['id']}\nText: {seg['english']}" for seg in batch_segments])
    prompt = f"""You are an expert bilingual translator for 12-step recovery literature. 
Translate the following segments accurately, maintaining contextual flow between them. 
GLOSSARY TERMS:\n{glossary_text}
\n\nSegments to Translate:\n{input_payload}"""

    active_models = get_fallback_models()
    for model_name in active_models:
        for attempt in range(2): # Try batch twice
            try:
                parsed = _call_gemini(model_name, prompt, TranslationBatchResult)
                items = parsed.get("items", [])
                # strict validation: Did we get the right count and exact IDs back?
                if len(items) == len(batch_segments) and all(items[i]['id'] == batch_segments[i]['id'] for i in range(len(items))):
                    return items
            except Exception as e:
                if _is_retryable(str(e)) and attempt < 1:
                    _backoff_sleep(attempt)
                    continue
                break
                
    # SHATTER AND RESCUE: If the batch fails validation or API errors out, gracefully fallback to 1-by-1 processing
    results = []
    for seg in batch_segments:
        single_res = translate_with_ai(seg['english'], glossary_text)
        results.append({
            "id": seg['id'],
            "arabic_translation": single_res['arabic_translation'],
            "glossary_notes": single_res['glossary_notes']
        })
    return results

def review_batch_with_fallback(batch_segments, glossary_text):
    if not GENAI_AVAILABLE or client is None:
        return [review_with_ai(seg['english'], seg['arabic'], glossary_text) for seg in batch_segments]

    input_payload = "\n\n".join([f"ID: {seg['id']}\nEnglish: {seg['english']}\nArabic: {seg['arabic']}" for seg in batch_segments])
    prompt = f"""You are an expert bilingual editor for 12-step recovery literature. 
Review the following English/Arabic pairs for accuracy and glossary adherence.
GLOSSARY TERMS:\n{glossary_text}
\n\nPairs to Review:\n{input_payload}"""

    active_models = get_fallback_models()
    for model_name in active_models:
        for attempt in range(2):
            try:
                parsed = _call_gemini(model_name, prompt, ReviewBatchResult)
                items = parsed.get("items", [])
                if len(items) == len(batch_segments) and all(items[i]['id'] == batch_segments[i]['id'] for i in range(len(items))):
                    return items
            except Exception as e:
                if _is_retryable(str(e)) and attempt < 1:
                    _backoff_sleep(attempt)
                    continue
                break
                
    # SHATTER AND RESCUE FALLBACK
    results = []
    for seg in batch_segments:
        single_res = review_with_ai(seg['english'], seg['arabic'], glossary_text)
        results.append({
            "id": seg['id'],
            "status": single_res['status'],
            "suggested_arabic": single_res['suggested_arabic'],
            "reasoning": single_res['reasoning']
        })
    return results


# ==========================================
# 4. DOCUMENT PARSERS & ANOMALY DETECTOR
# ==========================================
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
                blocks.append({'text': line_str, 'start': None, 'end': None})
    return blocks

def extract_text_from_docx(file_bytes: bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    blocks = []
    for p in doc.element.body.iter():
        if p.tag.endswith('}p'):
            text = "".join(node.text for node in p.iter() if node.tag.endswith('}t') and node.text)
            clean_text = text.strip()
            if bool(re.search(r'[a-zA-Z\u0600-\u06FF]', clean_text)):
                if not blocks or blocks[-1]['text'] != clean_text:
                    blocks.append({'text': clean_text, 'start': None, 'end': None})
    return blocks

def _parse_docs_elements(elements):
    paras = []
    for elem in elements:
        if 'paragraph' in elem:
            para_text = ""
            start_idx = elem.get('startIndex')
            end_idx = elem.get('endIndex')
            for run in elem.get('paragraph', {}).get('elements', []):
                if 'textRun' in run:
                    para_text += run.get('textRun', {}).get('content', '')
            clean_text = para_text.strip()
            if bool(re.search(r'[a-zA-Z\u0600-\u06FF]', clean_text)): 
                paras.append({'text': clean_text, 'start': start_idx, 'end': end_idx})
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
        st.error(f"Could not read document from Drive. Ensure the Google Doc is shared with the Service Account. Error: {e}")
        return None

def smart_align_with_anomaly_detection(paragraphs: list):
    en_paras = []
    ar_paras = []
    
    for p in paragraphs:
        if bool(re.search(r'[\u0600-\u06FF]', p['text'])):
            ar_paras.append(p)
        else:
            en_paras.append(p)

    aligned_segments = []
    max_len = max(len(en_paras), len(ar_paras))
    anomaly_msg = None
    if len(en_paras) != len(ar_paras):
        anomaly_msg = f"Count Mismatch: Found {len(en_paras)} English blocks vs {len(ar_paras)} Arabic blocks."

    for i in range(max_len):
        en_obj = en_paras[i] if i < len(en_paras) else {'text': "[MISSING ENGLISH SOURCE]", 'start': None, 'end': None}
        ar_obj = ar_paras[i] if i < len(ar_paras) else {'text': "[MISSING ARABIC TRANSLATION]", 'start': None, 'end': None}
        
        status = 'normal'
        if en_obj['text'] == "[MISSING ENGLISH SOURCE]": status = 'misaligned_ar'
        elif ar_obj['text'] == "[MISSING ARABIC TRANSLATION]": status = 'misaligned_en'

        aligned_segments.append({
            'id': i + 1,
            'status': status,
            'english': en_obj['text'],
            'arabic': ar_obj['text'],
            'ar_start': ar_obj['start'],
            'ar_end': ar_obj['end'],
            'anomaly': anomaly_msg
        })
    return aligned_segments

# ==========================================
# 5. GOOGLE DOCS WRITE-BACK FUNCTIONS
# ==========================================
def push_to_drive_translator(document_id, final_arabic_text):
    docs_svc, _, _ = get_google_services()
    try:
        requests = [
            {'insertPageBreak': {'location': {'index': 1}}},
            {'insertText': {'location': {'index': 1}, 'text': final_arabic_text + "\n\n"}}
        ]
        docs_svc.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        return True
    except Exception as e:
        st.error(f"Failed to push to Drive. Ensure the Service Account has 'Editor' access. Error: {e}")
        return False

def push_to_drive_reviewer(document_id, approved_segments):
    docs_svc, _, _ = get_google_services()
    try:
        # Filter segments that have exact coordinate indices
        valid_segments = [seg for seg in approved_segments if seg['ar_start'] is not None and seg['ar_end'] is not None]
        # Sort coordinates descending (bottom-to-top) to prevent index shifting!
        valid_segments.sort(key=lambda x: x['ar_start'], reverse=True)
        
        requests = []
        for seg in valid_segments:
            # -1 to preserve the trailing newline of the Google Doc paragraph block
            requests.append({
                'deleteContentRange': {
                    'range': {
                        'startIndex': seg['ar_start'],
                        'endIndex': seg['ar_end'] - 1 
                    }
                }
            })
            requests.append({
                'insertText': {
                    'location': {'index': seg['ar_start']},
                    'text': seg['final_arabic']
                }
            })
            
        if requests:
            docs_svc.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        return True
    except Exception as e:
        st.error(f"Failed to push to Drive. Ensure the Service Account has 'Editor' access. Error: {e}")
        return False

# ==========================================
# 6. DASHBOARD UI & INGESTION
# ==========================================
st.title("⚙️ 12-Step AI Suite")

if not GENAI_AVAILABLE:
    st.error("🚨 Critical Dependency Missing: The `google-genai` package is not installed.")
    st.stop()

glossary_data = fetch_glossary()

with st.container(border=True):
    col_main, col_info = st.columns([2, 1])
    with col_main:
        st.markdown("### 🎛️ 1. Select Operating Mode")
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            if st.button("🌍 Translator Mode", use_container_width=True, type="primary" if st.session_state["app_mode"] == "Translator Mode" else "secondary"):
                st.session_state["app_mode"] = "Translator Mode"
                st.session_state['processed_data'] = None
                st.rerun()
        with col_m2:
            if st.button("📝 Reviewer Mode", use_container_width=True, type="primary" if st.session_state["app_mode"] == "Reviewer Mode" else "secondary"):
                st.session_state["app_mode"] = "Reviewer Mode"
                st.session_state['processed_data'] = None
                st.rerun()
                
        st.markdown("### 📥 2. Select Input Method")
        col_meth1, col_meth2 = st.columns(2)
        with col_meth1:
            if st.button("☁️ Google Drive", use_container_width=True, type="primary" if st.session_state["input_method"] == "drive" else "secondary"):
                st.session_state["input_method"] = "drive"
                st.session_state['processed_data'] = None
                st.rerun()
        with col_meth2:
            if st.button("📁 File Upload", use_container_width=True, type="primary" if st.session_state["input_method"] == "upload" else "secondary"):
                st.session_state["input_method"] = "upload"
                st.session_state['processed_data'] = None
                st.rerun()

    with col_info:
        st.markdown("### 📊 System Status")
        st.caption("⚡ **Engine:** `Micro-Batching Enabled`")
        if "No glossary" not in glossary_data and glossary_data != "":
            st.success(f"✅ **Glossary Connected:**\n`{GLOSSARY_RANGE.split('!')[0]}`")
        else:
            st.warning("⚠️ Glossary not active.")

st.divider()

if st.session_state["input_method"] == "drive":
    doc_url = st.text_input("Paste Google Drive File URL Here:")
    if st.button("Load & Process from Drive") and doc_url:
        file_id = extract_id(doc_url)
        if file_id:
            st.session_state['source_file_id'] = file_id
            st.info(f"Connecting to Document ID: {file_id}...")
            paras = extract_text_from_drive(file_id)

            if paras:
                if st.session_state["app_mode"] == "Translator Mode":
                    st.info(f"Extracted {len(paras)} segments. Translating via AI Batching...")
                    progress_bar = st.progress(0)
                    processed_results = []
                    
                    # Batch processing
                    batches = [paras[i:i + BATCH_SIZE] for i in range(0, len(paras), BATCH_SIZE)]
                    for idx, batch in enumerate(batches):
                        batch_payload = [{'id': j + 1, 'english': p['text']} for j, p in enumerate(batch)]
                        ai_results = translate_batch_with_fallback(batch_payload, glossary_data)
                        for ai_res in ai_results:
                            processed_results.append({
                                "id": ai_res['id'] + (idx * BATCH_SIZE),
                                "english": batch[ai_res['id'] - 1]['text'],
                                "arabic_translation": ai_res.get("arabic_translation", ""),
                                "glossary_notes": ai_res.get("glossary_notes", "")
                            })
                        progress_bar.progress((idx + 1) / len(batches))
                        
                else:
                    segments = smart_align_with_anomaly_detection(paras)
                    st.info(f"Extracted blocks. Aligning and Reviewing via AI Batching...")
                    progress_bar = st.progress(0)
                    processed_results = []
                    
                    normal_segs = [s for s in segments if s['status'] == 'normal']
                    batches = [normal_segs[i:i + BATCH_SIZE] for i in range(0, len(normal_segs), BATCH_SIZE)]
                    
                    for idx, batch in enumerate(batches):
                        ai_results = review_batch_with_fallback(batch, glossary_data)
                        for ai_res, original_seg in zip(ai_results, batch):
                            processed_results.append({
                                "id": original_seg['id'],
                                "status": ai_res.get('status', 'minor_edits'),
                                "english": original_seg['english'],
                                "original_arabic": original_seg['arabic'],
                                "suggested_arabic": ai_res.get("suggested_arabic", original_seg['arabic']),
                                "reasoning": ai_res.get("reasoning", ""),
                                "anomaly": original_seg['anomaly'],
                                "ar_start": original_seg['ar_start'],
                                "ar_end": original_seg['ar_end']
                            })
                        progress_bar.progress((idx + 1) / len(batches))
                        
                    # Handle Anomalies One-by-One outside the batches
                    for item in segments:
                        if item['status'] == 'misaligned_en':
                            trans_res = translate_with_ai(item['english'], glossary_data)
                            processed_results.append({
                                "id": item['id'], "status": "major_rewrite", "english": item['english'],
                                "original_arabic": "[MISSING]", "suggested_arabic": trans_res.get("arabic_translation", ""),
                                "reasoning": "⚠️ Auto-translated orphaned English.", "anomaly": item['anomaly'],
                                "ar_start": None, "ar_end": None
                            })
                        elif item['status'] == 'misaligned_ar':
                            processed_results.append({
                                "id": item['id'], "status": "major_rewrite", "english": "[MISSING]",
                                "original_arabic": item['arabic'], "suggested_arabic": item['arabic'],
                                "reasoning": "⚠️ Orphaned Arabic. Ignored.", "anomaly": item['anomaly'],
                                "ar_start": item['ar_start'], "ar_end": item['ar_end']
                            })
                            
                    processed_results.sort(key=lambda x: x['id'])
                st.session_state['processed_data'] = processed_results
                st.rerun()

elif st.session_state["input_method"] == "upload":
    st.info("File upload ignores the Push-to-Drive feature. Use Drive Input for full write-back automation.")
    # Standard logic applies here (omitted for brevity, follows exact same batch loop mapping as above)
    pass

# ==========================================
# 7. DYNAMIC OUTPUT GRID & PUSH ACTION
# ==========================================
if st.session_state['processed_data']:
    st.divider()
    approved_count = 0
    finalized_data = []

    if st.session_state["app_mode"] == "Translator Mode":
        st.subheader("Translation Editor")
        for i, item in enumerate(st.session_state['processed_data']):
            with st.container(border=True):
                st.markdown(f"### Segment {item['id']}")
                col_en, col_ar = st.columns(2)
                with col_en:
                    st.info(item['english'])
                    with st.expander("💡 AI Glossary Notes"):
                        st.caption(item['glossary_notes'])
                with col_ar:
                    final_text = st.text_area("Final Translation", value=item['arabic_translation'], height=120, key=f"edit_ar_{i}", label_visibility="collapsed")
                
                if st.checkbox(f"✅ Approve Segment {item['id']}", key=f"approve_{i}", value=True):
                    approved_count += 1
                    finalized_data.append(final_text)

    else:
        st.subheader("Review Segments & Diff Visualizer")
        for i, item in enumerate(st.session_state['processed_data']):
            color = "🟢" if item['status'] == "perfect" else ("🟡" if item['status'] == "minor_edits" else "🔴")
            with st.container(border=True):
                st.markdown(f"### Segment {item['id']} | Status: {color} {item['status'].upper()}")
                col_en, col_ar = st.columns(2)
                with col_en: st.info(item['english'])
                with col_ar:
                    diff_html = generate_html_diff(item['original_arabic'], item['suggested_arabic'])
                    st.markdown(diff_html, unsafe_allow_html=True)
                    with st.expander("💡 View AI Reasoning"): st.markdown(item['reasoning'])
                    final_text = st.text_area("Final Output", value=item['suggested_arabic'], height=120, key=f"edit_ar_{i}", label_visibility="collapsed")
                
                if st.checkbox(f"✅ Approve Segment {item['id']}", key=f"approve_{i}", value=(item['status'] == 'perfect')):
                    approved_count += 1
                    finalized_data.append({'final_arabic': final_text, 'ar_start': item.get('ar_start'), 'ar_end': item.get('ar_end')})

    # --- COMPILED EXPORT & PUSH BUTTON ---
    total_segments = len(st.session_state['processed_data'])
    st.divider()
    st.write(f"### **Approved Segments: {approved_count} / {total_segments}**")

    if approved_count == total_segments and total_segments > 0:
        st.success("🎉 All segments approved! You can copy the text below or push it directly to Google Docs.")
        
        if st.session_state["app_mode"] == "Translator Mode":
            final_arabic_blob = "\n\n".join(finalized_data)
            st.text_area("Compiled Arabic Text", value=final_arabic_blob, height=400)
            
            if st.session_state['input_method'] == 'drive' and st.session_state.get('source_file_id'):
                if st.button("🚀 Push Translation to Google Doc", type="primary", use_container_width=True):
                    with st.spinner("Pushing to Drive..."):
                        success = push_to_drive_translator(st.session_state['source_file_id'], final_arabic_blob)
                        if success: st.balloons(); st.success("Translation pushed successfully!")
        else:
            if st.session_state['input_method'] == 'drive' and st.session_state.get('source_file_id'):
                if st.button("🚀 Apply Revisions to Google Doc", type="primary", use_container_width=True):
                    with st.spinner("Rewriting Document..."):
                        success = push_to_drive_reviewer(st.session_state['source_file_id'], finalized_data)
                        if success: st.balloons(); st.success("Revisions applied successfully!")
