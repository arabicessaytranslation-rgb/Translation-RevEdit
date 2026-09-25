import streamlit as st
import re
import io
import json
import time
import random
import docx
import fitz  # PyMuPDF
import difflib
import concurrent.futures
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

# THROTTLED MULTITHREADING FOR FREE TIER (15 RPM)
MAX_WORKERS = 3 

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
    """Restores dynamic fetching but strictly filters out slow/experimental models to maintain speed."""
    secret_model = st.secrets.get("ACTIVE_MODEL", "").strip()
    if secret_model:
        return [secret_model]

    if GENAI_AVAILABLE and client is not None:
        try:
            available_flash_models = []
            for m in client.models.list():
                clean_name = m.name.replace("models/", "")
                if "flash" in clean_name.lower():
                    bad_tags = ["legacy", "embed", "imagen", "tts", "audio", "preview", "experimental", "vision", "lite"]
                    if not any(tag in clean_name.lower() for tag in bad_tags):
                        available_flash_models.append(clean_name)
            
            available_flash_models.sort(reverse=True)
            if available_flash_models:
                return available_flash_models
        except Exception:
            pass

    return ["gemini-1.5-flash", "gemini-1.5-pro"]

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
    """Safely extracts API output and converts to JSON without regex vulnerabilities."""
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

    clean_text = response.text.strip()
    
    # Strip markdown block formatting safely
    if clean_text.startswith("```json"):
        clean_text = clean_text[7:]
    elif clean_text.startswith("```"):
        clean_text = clean_text[3:]
    if clean_text.endswith("```"):
        clean_text = clean_text[:-3]

    return json.loads(clean_text.strip())

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

    # Replaced blank return with visible error so the text box never mysteriously empties
    return {"arabic_translation": f"⚠️ ERROR: {last_error}", "glossary_notes": "API call failed."}

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

    return {"status": "major_rewrite", "suggested_arabic": f"⚠️ ERROR: {last_error}", "reasoning": "API call failed."}

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

# --- MULTITHREADING RUNNERS ---
def run_parallel_translations(paras, glossary_data, progress_bar):
    processed = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(translate_with_ai, text, glossary_data): (i, text) for i, text in enumerate(paras)}
        for count, future in enumerate(concurrent.futures.as_completed(futures), 1):
            i, text = futures[future]
            try:
                res = future.result()
            except Exception as e:
                res = {"arabic_translation": f"⚠️ ERROR: {e}", "glossary_notes": "Threading Error"}
            
            processed.append({
                "id": i + 1,
                "english": text,
                "arabic_translation": res.get("arabic_translation", ""),
                "glossary_notes": res.get("glossary_notes", "")
            })
            progress_bar.progress(count / len(paras))
            
    processed.sort(key=lambda x: x["id"])
    return processed

def run_parallel_reviews(segments, glossary_data, progress_bar):
    processed = []
    
    def review_router(item):
        if item['status'] == 'normal':
            ai_res = review_with_ai(item['english'], item['arabic'], glossary_data)
            return {
                "id": item['id'],
                "status": ai_res.get("status", "minor_edits"),
                "english": item['english'],
                "original_arabic": item['arabic'],
                "suggested_arabic": ai_res.get("suggested_arabic", item['arabic']),
                "reasoning": ai_res.get("reasoning", ""),
                "anomaly": item['anomaly']
            }
        elif item['status'] == 'misaligned_en':
            trans_res = translate_with_ai(item['english'], glossary_data)
            return {
                "id": item['id'],
                "status": "major_rewrite",
                "english": item['english'],
                "original_arabic": "[MISSING IN SOURCE DOCUMENT]",
                "suggested_arabic": trans_res.get("arabic_translation", ""),
                "reasoning": "⚠️ Auto-translated from source using Sheet glossary.",
                "anomaly": item['anomaly']
            }
        else: # misaligned_ar
            return {
                "id": item['id'],
                "status": "major_rewrite",
                "english": "[MISSING ENGLISH SOURCE]",
                "original_arabic": item['arabic'],
                "suggested_arabic": item['arabic'],
                "reasoning": "⚠️ No English source detected for comparison.",
                "anomaly": item['anomaly']
            }

    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(review_router, item): item for item in segments}
        for count, future in enumerate(concurrent.futures.as_completed(futures), 1):
            try:
                res = future.result()
            except Exception as e:
                item = futures[future]
                res = {
                    "id": item['id'], "status": "major_rewrite", "english": item['english'],
                    "original_arabic": item['arabic'], "suggested_arabic": f"⚠️ ERROR: {e}",
                    "reasoning": "Threading Error", "anomaly": item['anomaly']
                }
