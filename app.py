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
    st.session_state["app_mode"] = "Translator Mode"

if 'processed_data' not in st.session_state:
    st.session_state['processed_data'] = None

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

# --- AI DATA SCHEMAS ---
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

    # String multiplier prevents paste breaks on markdown backticks
    clean_text = response.text.replace("`" * 3 + "json", "").replace("`" * 3, "").strip()
    match = re.search(r'\{.*\}', clean_text, re.DOTALL)
    if match:
        clean_text = match.group(0)
    return json.loads(clean_text)

# --- AI WORKFLOW: TRANSLATION ---
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

# --- AI WORKFLOW: REVIEWER ---
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
    full_text = ""
    for page_num in range(len(pdf_document)):
        full_text += pdf_document.load_page(page_num).get_text("text") + "\n"
    return [p.replace('\n', ' ').strip() for p in full_text.split('\n\n') if p.replace('\n', ' ').strip()]

def extract_text_from_docx(file_bytes: bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

def extract_text_from_drive(file_id: str, is_retry=False):
    try:
        docs_svc, drive_svc, _ = get_google_services()
        file_meta = drive_svc.files().get(fileId=file_id, fields="mimeType").execute()
        mime_type = file_meta.get("mimeType")

        if mime_type == 'application/vnd.google-apps.document':
            document = docs_svc.documents().get(documentId=file_id).execute()
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
    """
    Scans sequential paragraphs, detects language, and flags structural anomalies
    (e.g., consecutive EN-EN or AR-AR) without cascading misalignment down the document.
    """
    tagged = []
    for p in paragraphs:
        is_ar = bool(re.search(r'[\u0600-\u06FF]', p))
        tagged.append({'lang': 'AR' if is_ar else 'EN', 'text': p})

    aligned_segments = []
    i = 0
    seg_id = 1

    while i < len(tagged):
        curr = tagged[i]
        has_next = (i + 1 < len(tagged))
        next_item = tagged[i + 1] if has_next else None

        # Case 1: Standard alternating pair (EN followed immediately by AR)
        if curr['lang'] == 'EN' and has_next and next_item['lang'] == 'AR':
            aligned_segments.append({
                'id': seg_id,
                'status': 'normal',
                'english': curr['text'],
                'arabic': next_item['text'],
                'anomaly': None
            })
            i += 2

        # Case 2: Inverted pair (AR followed immediately by EN)
        elif curr['lang'] == 'AR' and has_next and next_item['lang'] == 'EN':
            aligned_segments.append({
                'id': seg_id,
                'status': 'inverted',
                'english': next_item['text'],
                'arabic': curr['text'],
                'anomaly': 'Inverted Order: Arabic appeared before English in document.'
            })
            i += 2

        # Case 3: Consecutive English (Missing corresponding Arabic block)
        elif curr['lang'] == 'EN':
            aligned_segments.append({
                'id': seg_id,
                'status': 'misaligned_en',
                'english': curr['text'],
                'arabic': '',
                'anomaly': 'Out of Order: Consecutive English paragraphs detected without corresponding Arabic.'
            })
            i += 1

        # Case 4: Consecutive Arabic (Missing English source block)
        elif curr['lang'] == 'AR':
            aligned_segments.append({
                'id': seg_id,
                'status': 'misaligned_ar',
                'english': '',
                'arabic': curr['text'],
                'anomaly': 'Out of Order: Isolated Arabic paragraph found without preceding English.'
            })
            i += 1

        seg_id += 1

    return aligned_segments

def smart_align_separate_files(en_paras: list, ar_paras: list):
    pairs = []
    max_len = max(len(en_paras), len(ar_paras))
    for i in range(max_len):
        en = en_paras[i] if i < len(en_paras) else "[MISSING ENGLISH SOURCE]"
        ar = ar_paras[i] if i < len(ar_paras) else "[MISSING ARABIC TRANSLATION]"
        pairs.append({
            "id": i + 1,
            "status": "normal" if (i < len(en_paras) and i < len(ar_paras)) else "misaligned",
            "english": en,
            "arabic": ar,
            "anomaly": None if (i < len(en_paras) and i < len(ar_paras)) else "File segment count mismatch"
        })
    return pairs

# ==========================================
# 4. DASHBOARD UI & ROUTING
# ==========================================
st.title("12-Step AI Suite")

if not GENAI_AVAILABLE:
    st.error("🚨 Critical Dependency Missing: The `google-genai` package is not installed.")
    st.stop()

col_mode, col_info = st.columns([1, 2])
with col_mode:
    selected_mode = st.radio("Select Tool Mode:", ["Translator Mode", "Reviewer Mode"], horizontal=True)
    if selected_mode != st.session_state["app_mode"]:
        st.session_state["app_mode"] = selected_mode
        st.session_state['processed_data'] = None
        st.rerun()

glossary_data = fetch_glossary()
active_models_list = get_fallback_models()

with col_info:
    st.caption(f"Active AI Routing Models: `{', '.join(active_models_list)}`")
    if "No glossary connected" not in glossary_data and glossary_data != "":
        st.success(f"✅ Glossary connected from: {GLOSSARY_RANGE.split('!')[0]}")
    else:
        st.warning("⚠️ Glossary not active. Check Spreadsheet ID and permissions.")

st.divider()

tab1, tab2 = st.tabs(["Method 1: Google Drive Link", "Method 2: Direct File Upload"])

# --- TAB 1: GOOGLE DRIVE ---
with tab1:
    if st.session_state["app_mode"] == "Translator Mode":
        st.write("Paste a Google Doc or Word URL containing English text to translate.")
    else:
        st.write("Paste a Google Doc or Word URL containing alternating English/Arabic text to review.")
        
    doc_url = st.text_input("Paste Google Drive File URL Here:")

    if st.button("Load & Process from Drive") and doc_url:
        file_id = extract_id(doc_url)
        if not file_id:
            st.error("Invalid Google Drive URL.")
        else:
            st.info(f"Connecting to Document ID: {file_id}...")
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
                    # Reviewer Mode with Smart Anomaly Detection
                    segments = smart_align_with_anomaly_detection(paras)
                    st.info(f"Smart Scanner mapped {len(segments)} segments. Processing reviews & resolving anomalies...")
                    progress_bar = st.progress(0)
                    processed_results = []

                    for i, item in enumerate(segments):
                        if item['status'] in ('normal', 'inverted'):
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
                            # Trigger auto-translation using Google Sheet glossary
                            trans_res = translate_with_ai(item['english'], glossary_data)
                            processed_results.append({
                                "id": item['id'],
                                "status": "major_rewrite",
                                "english": item['english'],
                                "original_arabic": "[MISSING IN SOURCE DOCUMENT]",
                                "suggested_arabic": trans_res.get("arabic_translation", ""),
                                "reasoning": f"⚠️ Anomaly: {item['anomaly']} Auto-translated using Sheet glossary.",
                                "anomaly": item['anomaly']
                            })
                        elif item['status'] == 'misaligned_ar':
                            processed_results.append({
                                "id": item['id'],
                                "status": "major_rewrite",
                                "english": "[MISSING ENGLISH SOURCE]",
                                "original_arabic": item['arabic'],
                                "suggested_arabic": item['arabic'],
                                "reasoning": f"⚠️ Anomaly: {item['anomaly']} No English source detected.",
                                "anomaly": item['anomaly']
                            })
                        progress_bar.progress((i + 1) / len(segments))

                st.session_state['processed_data'] = processed_results
                st.rerun()

# --- TAB 2: FILE UPLOAD ---
with tab2:
    if st.session_state["app_mode"] == "Translator Mode":
        st.write("Upload an English Word document or PDF to translate.")
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
        st.write("Upload separate English and Arabic files to review.")
        colA, colB = st.columns(2)
        with colA:
            file_en = st.file_uploader("1. Upload English Source", type=["docx", "pdf"])
        with colB:
            file_ar = st.file_uploader("2. Upload Arabic Translation", type=["docx", "pdf"])

        if st.button("Process Uploaded Files") and file_en and file_ar:
            en_paras = extract_text_from_pdf(file_en.read()) if file_en.name.endswith('.pdf') else extract_text_from_docx(file_en.read())
            ar_paras = extract_text_from_pdf(file_ar.read()) if file_ar.name.endswith('.pdf') else extract_text_from_docx(file_ar.read())
            smart_pairs = smart_align_separate_files(en_paras, ar_paras)
            
            st.info("Aligning files and generating AI review...")
            progress_bar = st.progress(0)
            processed_results = []
            for i, pair in enumerate(smart_pairs):
                if pair['english'] != "[MISSING ENGLISH SOURCE]" and pair['arabic'] != "[MISSING ARABIC TRANSLATION]":
                    ai_result = review_with_ai(pair['english'], pair['arabic'], glossary_data)
                    suggested = ai_result.get("suggested_arabic", pair['arabic'])
                    status = ai_result.get("status", "minor_edits")
                    reasoning = ai_result.get("reasoning", "")
                elif pair['arabic'] == "[MISSING ARABIC TRANSLATION]":
                    trans_res = translate_with_ai(pair['english'], glossary_data)
                    suggested = trans_res.get("arabic_translation", "")
                    status = "major_rewrite"
                    reasoning = "Missing Arabic translation. Auto-generated from source using Glossary."
                else:
                    suggested = pair['arabic']
                    status = "major_rewrite"
                    reasoning = "Missing English source."

                processed_results.append({
                    "id": pair['id'],
                    "status": status,
                    "english": pair['english'],
                    "original_arabic": pair['arabic'],
                    "suggested_arabic": suggested,
                    "reasoning": reasoning,
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

                # High-visibility warning banner for structural anomalies
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
                
                # Visual Changes Diff
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
        st.success("🎉 All segments approved! Copy the finalized text blocks below directly into your master document.")
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
