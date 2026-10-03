import streamlit as st
import re
import json
import time
import random
import difflib
import smtplib
import pandas as pd
from email.message import EmailMessage
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
# 1. CONFIGURATION, SECRETS & CUSTOM CSS
# ==========================================
st.set_page_config(page_title="12-Step AI Suite: Translate & Review", layout="wide")

def apply_custom_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&family=Inter:wght@400;500;600;700&display=swap');
        
        html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
        
        textarea {
            font-family: 'Cairo', 'Inter', sans-serif !important;
            font-size: 16px !important;
            line-height: 1.6 !important;
            border-radius: 8px !important;
        }
        
        button[kind="primary"] {
            background-color: #10B981 !important;
            border-color: #10B981 !important;
            color: white !important;
            font-size: 16px !important;
            font-weight: 600 !important;
            border-radius: 8px !important;
            padding: 0.5rem 1rem !important;
            box-shadow: 0 4px 6px -1px rgba(16, 185, 129, 0.2) !important;
            transition: all 0.2s ease-in-out !important;
        }
        button[kind="primary"]:hover {
            background-color: #059669 !important;
            border-color: #059669 !important;
            transform: translateY(-2px);
        }

        div[data-testid="stAlert"]:has(p:contains("CRITICAL STEP")) {
            background-color: #FEF2F2 !important;
            border: 1px solid #F87171 !important;
            color: #991B1B !important;
            border-radius: 8px !important;
        }
        
        button:has(p:contains("I have reviewed the final text")) {
            background-color: #EF4444 !important;
            border-color: #EF4444 !important;
            color: white !important;
            font-size: 16px !important;
            font-weight: 600 !important;
            border-radius: 8px !important;
            transition: all 0.2s ease-in-out !important;
        }
        button:has(p:contains("I have reviewed the final text")):hover {
            background-color: #DC2626 !important;
            border-color: #DC2626 !important;
        }
        
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 2rem !important;
        }
    </style>
    """, unsafe_allow_html=True)

apply_custom_css()

GLOSSARY_SPREADSHEET_ID = "1oc4TCY_iK9R7mBiXgb5rKWssjmrQywYg6UpOBXx8pUQ"
GLOSSARY_RANGE = "'المصطلحات'!A:D"
GLOSSARY_DATA_RANGE = "'المصطلحات'!C:D" # Just for AI fetching
SESSIONS_RANGE = "'Sessions'!A:D"
VOLUNTEERS_RANGE = "'Volunteers'!A:D"
LOCK_TIMEOUT_SECONDS = 14400  # 4 hours auto-expiration

BATCH_SIZE = 6
MAX_RETRIES_PER_MODEL = 3
BASE_BACKOFF_SECONDS = 2.0
MAX_BACKOFF_SECONDS = 15.0
RETRYABLE_KEYWORDS = (
    "503", "500", "ServiceUnavailable", "service_unavailable",
    "high demand", "UNAVAILABLE", "429", "ResourceExhausted",
    "DeadlineExceeded", "timeout", "Quota",
)

# ==========================================
# 2. GOOGLE SERVICES & AUTHENTICATION
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
                'https://www.googleapis.com/auth/spreadsheets'
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

def fetch_volunteers():
    """Reads the Volunteers tab to get live user access."""
    try:
        sheet = sheets_service.spreadsheets()
        result = sheet.values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=VOLUNTEERS_RANGE).execute()
        rows = result.get('values', [])
        volunteers = {}
        if len(rows) > 1:
            for r in rows[1:]:
                if len(r) >= 4:
                    email, name, role, status = r[0].strip().lower(), r[1], r[2].lower(), r[3]
                    volunteers[email] = {"name": name, "role": role, "status": status}
        return volunteers
    except Exception as e:
        st.error(f"Error reading Volunteers sheet: {e}")
        return {}

def overwrite_sheet_data(range_name, data_matrix):
    """Safely wipes a sheet range and writes a new dataframe into it."""
    sheet = sheets_service.spreadsheets()
    sheet.values().clear(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name).execute()
    body = {'values': data_matrix}
    sheet.values().update(
        spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name,
        valueInputOption="USER_ENTERED", body=body
    ).execute()

def login_screen():
    if st.session_state.get("authenticated"):
        return True

    st.title("🤝 12-Step AI Suite: Volunteer Portal")
    with st.form("login_form"):
        email = st.text_input("Enter your registered email address").strip().lower()
        submitted = st.form_submit_button("Access Portal", type="primary")

        if submitted:
            volunteers = fetch_volunteers()
            if email in volunteers:
                user_data = volunteers[email]
                if user_data["status"].lower() != "active":
                    st.error("Account suspended. Please contact the coordinator.")
                else:
                    st.session_state["authenticated"] = True
                    st.session_state["user_email"] = email
                    st.session_state["user_role"] = user_data["role"]
                    st.session_state["user_name"] = user_data["name"]
                    
                    if user_data["role"] == "admin":
                        st.session_state["app_mode"] = "God Mode"
                    elif user_data["role"] == "translator":
                        st.session_state["app_mode"] = "Translator Mode"
                    else:
                        st.session_state["app_mode"] = "Reviewer Mode"
                    st.rerun()
            else:
                st.error("Email not recognized. Please verify with the core team.")
    return False

if not login_screen():
    st.stop()

# --- STATE MANAGEMENT ---
if 'processed_data' not in st.session_state:
    st.session_state['processed_data'] = None
if 'source_file_id' not in st.session_state:
    st.session_state['source_file_id'] = None
if 'session_row_index' not in st.session_state:
    st.session_state['session_row_index'] = None

# ==========================================
# 3. GOD MODE (PANDORA BOX DASHBOARD)
# ==========================================
if st.session_state.get("app_mode") == "God Mode":
    st.title("⚡ Central Command: God Mode")
    st.caption("You are logged in as an Administrator. Changes made here instantly affect the live database.")
    
    if st.button("🚪 Logout", type="primary"):
        st.session_state.clear()
        st.rerun()
        
    st.divider()
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "📡 Live Radar", "👥 Volunteers", "📖 Glossary", "📢 Broadcast"
    ])
    
    # --- TAB 1: LIVE RADAR (Sessions) ---
    with tab1:
        st.subheader("Active Document Locks")
        try:
            sheet = sheets_service.spreadsheets()
            res = sheet.values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE).execute()
            session_rows = res.get('values', [])
            
            if len(session_rows) > 1:
                for idx, row in enumerate(session_rows[1:]): # skip header
                    doc_id = row[0] if len(row) > 0 else "Unknown"
                    locked_by = row[1] if len(row) > 1 else "Unknown"
                    ts = float(row[2]) if len(row) > 2 and row[2] else 0
                    time_locked = round((time.time() - ts) / 60, 1) if ts else 0
                    
                    if locked_by and doc_id:
                        with st.container(border=True):
                            col1, col2, col3 = st.columns([2, 1, 1])
                            with col1:
                                st.markdown(f"**Doc ID:** `{doc_id}`")
                                st.markdown(f"🔒 Locked by **{locked_by}** for `{time_locked} minutes`")
                            with col2:
                                if st.button("🧨 Kill Lock", key=f"kill_{idx}", use_container_width=True):
                                    body = {'values': [["", "", "", ""]]}
                                    sheet.values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Sessions'!A{idx+2}:D{idx+2}", valueInputOption="USER_ENTERED", body=body).execute()
                                    st.success("Lock destroyed!")
                                    time.sleep(1)
                                    st.rerun()
                            with col3:
                                vols = fetch_volunteers()
                                active_emails = [e for e, d in vols.items() if d['status'].lower() == 'active']
                                new_owner = st.selectbox("Reassign to:", [""] + active_emails, key=f"reassign_{idx}", label_visibility="collapsed")
                                if new_owner:
                                    if st.button("📥 Transfer", key=f"btn_{idx}", use_container_width=True):
                                        body = {'values': [[new_owner]]}
                                        sheet.values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Sessions'!B{idx+2}", valueInputOption="USER_ENTERED", body=body).execute()
                                        st.success(f"Transferred to {new_owner}!")
                                        time.sleep(1)
                                        st.rerun()
            else:
                st.info("No active sessions right now. All clear!")
        except Exception as e:
            st.error(f"Failed to load sessions: {e}")

    # --- TAB 2: VOLUNTEERS DATABASE ---
    with tab2:
        st.subheader("Manage Volunteer Access")
        st.caption("Edit volunteer info or use the dropdowns to assign roles and update statuses.")
        try:
            res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=VOLUNTEERS_RANGE).execute()
            vol_data = res.get('values', [])
            if not vol_data:
                vol_data = [["Email", "Name", "Role", "Status"]]
            
            df_vol = pd.DataFrame(vol_data[1:], columns=vol_data[0])
            
            edited_vol = st.data_editor(
                df_vol, 
                num_rows="dynamic", 
                use_container_width=True,
                column_config={
                    "Role": st.column_config.SelectboxColumn(
                        "Role",
                        help="Select user permission level",
                        options=["translator", "reviewer", "admin"],
                        required=True
                    ),
                    "Status": st.column_config.SelectboxColumn(
                        "Status",
                        help="Account access state",
                        options=["Active", "Suspended"],
                        required=True
                    )
                }
            )
            
            if st.button("💾 Save Volunteers to Database", type="primary"):
                clean_df = edited_vol.fillna("")
                updated_matrix = [clean_df.columns.tolist()] + clean_df.values.tolist()
                overwrite_sheet_data(VOLUNTEERS_RANGE, updated_matrix)
                st.success("Volunteers database updated securely!")
        except Exception as e:
            st.error(f"Error loading volunteers: {e}")

    # --- TAB 3: GLOSSARY COMMAND CENTER ---
    with tab3:
        st.subheader("Live Terminology Editor")
        st.caption("Changes here will instantly update the AI's translation rules.")
        try:
            res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_RANGE).execute()
            glos_data = res.get('values', [])
            
            headers = ["ID", "Category", "English", "Arabic"]
            
            rows_to_display = []
            if glos_data and len(glos_data) > 0:
                if len(glos_data[0]) == 4 and not glos_data[0][0].isdigit():
                    headers = [str(h).strip() for h in glos_data[0]]
                    rows_to_display = glos_data[1:]
                else:
                    rows_to_display = glos_data
            
            cleaned_rows = []
            for row in rows_to_display:
                padded = list(row) + [""] * (4 - len(row))
                cleaned_rows.append(padded[:4])
                
            if not cleaned_rows:
                cleaned_rows = [["1", "General", "", ""]]
                
            df_glos = pd.DataFrame(cleaned_rows, columns=headers)
            edited_glos = st.data_editor(df_glos, num_rows="dynamic", use_container_width=True)
            
            if st.button("💾 Sync Glossary to AI", type="primary"):
                clean_df = edited_glos.fillna("")
                updated_matrix = [clean_df.columns.tolist()] + clean_df.values.tolist()
                overwrite_sheet_data(GLOSSARY_RANGE, updated_matrix)
                st.success("Glossary synced successfully!")
                st.cache_data.clear()
        except Exception as e:
            st.error(f"Error loading glossary: {e}")

    # --- TAB 4: BROADCAST DESK ---
    with tab4:
        st.subheader("Team Broadcast System")
        st.caption("Send a mass email to all 'Active' volunteers in the database.")
        broadcast_subject = st.text_input("Subject")
        broadcast_message = st.text_area("Message Body", height=150)
        
        if st.button("🚀 Send Broadcast", type="primary"):
            if broadcast_subject and broadcast_message:
                volunteers = fetch_volunteers()
                active_emails = [email for email, data in volunteers.items() if data['status'].lower() == 'active']
                
                if active_emails:
                    with st.spinner("Dispatching emails..."):
                        try:
                            msg = EmailMessage()
                            msg.set_content(f"12-Step Translation Project Update:\n\n{broadcast_message}")
                            msg['Subject'] = f"[12-Step Admin] {broadcast_subject}"
                            msg['From'] = st.secrets["SMTP_EMAIL"]
                            msg['To'] = st.secrets["SMTP_EMAIL"] 
                            msg['Bcc'] = ", ".join(active_emails) 
                            
                            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                                server.login(st.secrets["SMTP_EMAIL"], st.secrets["SMTP_PASSWORD"])
                                server.send_message(msg)
                            st.success(f"Broadcast sent successfully to {len(active_emails)} volunteers!")
                            st.balloons()
                        except Exception as e:
                            st.error(f"Failed to send: {e}")
                else:
                    st.warning("No active volunteers found to email.")
            else:
                st.warning("Please enter a subject and a message.")

    st.stop() # CRITICAL: STOPS GOD MODE FROM RENDERING THE REGULAR APP

# ==========================================
# 4. INITIALIZE AI & HELPER FUNCTIONS
# ==========================================
@st.cache_resource
def get_genai_client():
    if not GENAI_AVAILABLE: return None
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

# --- PYDANTIC SCHEMAS ---
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
    if secret_model: return [secret_model]
    if GENAI_AVAILABLE and client is not None:
        try:
            available_flash_models = []
            for m in client.models.list():
                clean_name = m.name.replace("models/", "")
                if "flash" in clean_name.lower() and not any(tag in clean_name.lower() for tag in ["legacy", "embed", "imagen"]):
                    available_flash_models.append(clean_name)
            available_flash_models.sort(reverse=True)
            if available_flash_models: return available_flash_models
        except Exception:
            pass
    return ["gemini-2.5-flash", "gemini-1.5-flash"]

def manage_document_lock(file_id: str, user_email: str):
    """Inspects the Sessions sheet to block collisions or recover previous work."""
    try:
        sheet = sheets_service.spreadsheets()
        result = sheet.values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE).execute()
        rows = result.get('values', [])
        current_time = time.time()
        target_row_index = max(len(rows) + 1, 2)
        
        for index, row in enumerate(rows):
            if index == 0: continue
            doc_id = row[0] if len(row) > 0 else ""
            if doc_id == file_id:
                target_row_index = index + 1
                locked_by = row[1] if len(row) > 1 else ""
                timestamp = float(row[2]) if (len(row) > 2 and row[2]) else 0
                saved_json = row[3] if len(row) > 3 else ""
                
                if locked_by and locked_by != user_email and (current_time - timestamp) < LOCK_TIMEOUT_SECONDS:
                    return {"status": "blocked", "locked_by": locked_by, "row_index": target_row_index}
                if locked_by == user_email and saved_json:
                    try:
                        recovered_data = json.loads(saved_json)
                        return {"status": "recovered", "data": recovered_data, "row_index": target_row_index}
                    except Exception:
                        pass
                return {"status": "clear", "row_index": target_row_index}
            
            if not doc_id and target_row_index > len(rows):
                target_row_index = index + 1
                
        return {"status": "clear", "row_index": target_row_index}
    except Exception as e:
        return {"status": "clear", "row_index": 2}

def save_draft_to_sheet(file_id: str, user_email: str, row_index: int, processed_data: list):
    """Silently updates the backup row in Google Sheets."""
    if not row_index: return
    try:
        sheet = sheets_service.spreadsheets()
        json_data = json.dumps(processed_data, ensure_ascii=False)
        timestamp = str(time.time())
        body = {'values': [[file_id, user_email, timestamp, json_data]]}
        range_name = f"'Sessions'!A{row_index}:D{row_index}"
        sheet.values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name, valueInputOption="USER_ENTERED", body=body).execute()
    except Exception: pass

def release_document_lock(row_index: int):
    """Clears the lock row once work is successfully pushed."""
    if not row_index: return
    try:
        sheet = sheets_service.spreadsheets()
        body = {'values': [["", "", "", ""]]}
        range_name = f"'Sessions'!A{row_index}:D{row_index}"
        sheet.values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name, valueInputOption="USER_ENTERED", body=body).execute()
    except Exception: pass

def send_email_notification(process_name: str, operator_name: str, operator_email: str):
    try:
        sender_email = st.secrets["SMTP_EMAIL"]
        sender_password = st.secrets["SMTP_PASSWORD"]
        admin_email = "arabicessaytranslation@gmail.com"
        msg = EmailMessage()
        msg.set_content(f"Hello {operator_name},\n\nThank you for your service! The {process_name} process has been successfully completed and pushed to Google Drive.\n\nBest,\n12-Step AI Suite")
        msg['Subject'] = f"Task Completed: {process_name} by {operator_name}"
        msg['From'] = sender_email
        msg['To'] = f"{operator_email}, {admin_email}" 
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.send_message(msg)
    except Exception as e:
        st.error(f"Failed to send email notification. Error: {e}")

@st.cache_data(ttl=3600)
def fetch_glossary():
    try:
        sheet = sheets_service.spreadsheets()
        result = sheet.values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_DATA_RANGE).execute()
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
        return "<div dir='rtl' style='font-family: \"Cairo\", sans-serif; color: #0369a1; background-color: #e0f2fe; padding: 10px; border-radius: 5px; text-align: right;'>✨ ترجمة تم توليدها بالكامل من المسرد (New Translation)</div>"
    if original.strip() == suggested.strip():
        return "<div dir='rtl' style='font-family: \"Cairo\", sans-serif; color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px; text-align: right;'>✨ لا توجد تعديلات (Perfect Match)</div>"
    
    diff = difflib.ndiff(original.split(), suggested.split())
    html = ["<div dir='rtl' style='font-family: \"Cairo\", sans-serif; line-height: 2; font-size: 18px; text-align: right; background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef;'>"]
    for word in diff:
        if word.startswith('- '): html.append(f"<span style='background-color: #ffcdd2; color: #b71c1c; text-decoration: line-through; padding: 2px 6px; margin: 0 2px; border-radius: 4px;'>{word[2:]}</span>")
        elif word.startswith('+ '): html.append(f"<span style='background-color: #c8e6c9; color: #1b5e20; font-weight: bold; padding: 2px 6px; margin: 0 2px; border-radius: 4px;'>{word[2:]}</span>")
        elif word.startswith('  '): html.append(f"<span style='color: #212529; margin: 0 2px;'>{word[2:]}</span>")
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
            response_mime_type="application/json", response_schema=schema_type,
            safety_settings=safety_settings, temperature=0.2, 
        ),
    )
    if not response.text:
        raise ValueError("Empty output from model")
    clean_text = response.text.replace("`" * 3 + "json", "").replace("`" * 3, "").strip()
    match = re.search(r'\{.*\}', clean_text, re.DOTALL)
    if match: clean_text = match.group(0)
    return json.loads(clean_text)

def translate_with_ai(english: str, glossary_text: str):
    prompt = f"""You are an expert bilingual translator specializing in 12-step recovery literature. 
Translate the English text into Arabic accurately, ensuring the tone remains clinical, professional, and non-moralizing.

CRITICAL INSTRUCTIONS:
- Do not perform blind word-for-word replacements. Actively understand semantic meaning.
- When pronouns like "it" appear referring to concepts such as "the program", ensure the Arabic translation reflects the correct contextual noun/glossary term and grammatical gender.
- Apply the glossary terms naturally into the sentence flow.

GLOSSARY TERMS:
{glossary_text}

Translate:
English Source: "{english}"
"""
    for model_name in get_fallback_models():
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                parsed = _call_gemini(model_name, prompt, TranslationResult)
                return {"arabic_translation": parsed.get("arabic_translation", ""), "glossary_notes": parsed.get("glossary_notes", "")}
            except Exception as e:
                if _is_retryable(f"{type(e).__name__} - {str(e)}") and attempt < MAX_RETRIES_PER_MODEL - 1:
                    _backoff_sleep(attempt)
                    continue
                break
    return {"arabic_translation": "", "glossary_notes": "⚠ Fallback Error."}

def review_with_ai(english: str, arabic: str, glossary_text: str):
    prompt = f"""You are an expert bilingual editor specializing in 12-step recovery literature. 
Ensure the Arabic translation is accurate, clinical, professional, and grammatically sound.

CRITICAL INSTRUCTIONS:
- Do not perform blind word-for-word replacements.
- Respect recovery glossary terms and verify sentence flow.

GLOSSARY TERMS:
{glossary_text}

Review this pair:
English Source: "{english}"
Original Arabic: "{arabic}"

If the translation captures meaning and tone accurately, leave it as is. If it misses glossary nuance or sounds unnatural, provide the polished Arabic translation.
"""
    for model_name in get_fallback_models():
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                parsed = _call_gemini(model_name, prompt, ReviewResult)
                return {"status": parsed.get("status", "minor_edits"), "suggested_arabic": parsed.get("suggested_arabic", arabic), "reasoning": parsed.get("reasoning", "")}
            except Exception as e:
                if _is_retryable(f"{type(e).__name__} - {str(e)}") and attempt < MAX_RETRIES_PER_MODEL - 1:
                    _backoff_sleep(attempt)
                    continue
                break
    return {"status": "major_rewrite", "suggested_arabic": arabic, "reasoning": "⚠ Fallback Error."}

def translate_batch_with_fallback(batch_segments, glossary_text):
    if not GENAI_AVAILABLE or client is None:
        return [translate_with_ai(seg['english'], glossary_text) for seg in batch_segments]
    input_payload = "\n\n".join([f"ID: {seg['id']}\nText: {seg['english']}" for seg in batch_segments])
    prompt = f"""You are an expert bilingual translator specializing in 12-step recovery literature. 
Translate the following English segments into Arabic accurately. Ensure the tone remains clinical, professional, and non-moralizing.

CRITICAL INSTRUCTIONS FOR THIS BATCH:
1. Narrative Flow: Maintain consistent grammatical gender, tone, and pronoun references across all segments.
2. Contextual Nuance: Actively understand the semantic meaning.
3. Pronoun Resolution: Reflect the correct contextual noun and proper Arabic grammatical gender.
4. Glossary Integration: Apply glossary terms naturally.

GLOSSARY TERMS:
{glossary_text}

Segments to Translate:
{input_payload}"""

    for model_name in get_fallback_models():
        for attempt in range(2): 
            try:
                parsed = _call_gemini(model_name, prompt, TranslationBatchResult)
                items = parsed.get("items", [])
                if len(items) == len(batch_segments) and all(items[i].get('id') == batch_segments[i].get('id') for i in range(len(items))):
                    return items
            except Exception as e:
                if _is_retryable(str(e)) and attempt < 1:
                    _backoff_sleep(attempt)
                    continue
                break
                
    results = []
    for seg in batch_segments:
        single_res = translate_with_ai(seg['english'], glossary_text)
        results.append({"id": seg['id'], "arabic_translation": single_res['arabic_translation'], "glossary_notes": single_res['glossary_notes']})
    return results

def review_batch_with_fallback(batch_segments, glossary_text):
    if not GENAI_AVAILABLE or client is None:
        return [review_with_ai(seg['english'], seg['arabic'], glossary_text) for seg in batch_segments]
    input_payload = "\n\n".join([f"ID: {seg['id']}\nEnglish: {seg['english']}\nArabic: {seg['arabic']}" for seg in batch_segments])
    prompt = f"""You are an expert bilingual editor specializing in 12-step recovery literature. 
Review the following English/Arabic pairs for accuracy, tone, and glossary adherence. Ensure the tone remains clinical, professional, and non-moralizing.

CRITICAL INSTRUCTIONS FOR THIS BATCH:
1. Narrative Flow: Maintain consistent grammatical gender, tone, and pronoun references across all segments.
2. Glossary Integration: Respect recovery glossary terms. If the translation is accurate, keep it. If not, provide the polished text.

GLOSSARY TERMS:
{glossary_text}

Pairs to Review:
{input_payload}"""

    for model_name in get_fallback_models():
        for attempt in range(2):
            try:
                parsed = _call_gemini(model_name, prompt, ReviewBatchResult)
                items = parsed.get("items", [])
                if len(items) == len(batch_segments) and all(items[i].get('id') == batch_segments[i].get('id') for i in range(len(items))):
                    return items
            except Exception as e:
                if _is_retryable(str(e)) and attempt < 1:
                    _backoff_sleep(attempt)
                    continue
                break
                
    results = []
    for seg in batch_segments:
        single_res = review_with_ai(seg['english'], seg['arabic'], glossary_text)
        results.append({"id": seg['id'], "status": single_res['status'], "suggested_arabic": single_res['suggested_arabic'], "reasoning": single_res['reasoning']})
    return results


# ==========================================
# 5. DOCUMENT PARSERS & WRITE-BACK
# ==========================================
def extract_id(url: str):
    match = re.search(r"/(?:d|folders)/([a-zA-Z0-9-_]+)", url)
    if match: return match.group(1)
    match_param = re.search(r"id=([a-zA-Z0-9-_]+)", url)
    if match_param: return match_param.group(1)
    if re.match(r"^[a-zA-Z0-9-_]+$", url.strip()): return url.strip()
    return None

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
            try: document = docs_svc.documents().get(documentId=file_id, includeTabsContent=True).execute()
            except Exception: document = docs_svc.documents().get(documentId=file_id).execute()

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
                for tab in tabs: all_paras.extend(sweep_doc_obj(tab.get('documentTab', {})))
            else:
                all_paras.extend(sweep_doc_obj(document))
            return all_paras
        else:
            st.error(f"Unsupported file type: {mime_type}. Please use Google Docs.")
            return None
    except Exception as e:
        err_str = str(e)
        if ("Broken pipe" in err_str or "Errno 32" in err_str) and not is_retry:
            get_google_services.clear()
            return extract_text_from_drive(file_id, is_retry=True)
        st.error(f"Could not read document from Drive. Error: {e}")
        return None

def smart_align_with_anomaly_detection(paragraphs: list):
    en_paras, ar_paras = [], []
    for p in paragraphs:
        if bool(re.search(r'[\u0600-\u06FF]', p['text'])): ar_paras.append(p)
        else: en_paras.append(p)

    aligned_segments = []
    max_len = max(len(en_paras), len(ar_paras))
    anomaly_msg = f"Count Mismatch: Found {len(en_paras)} English blocks vs {len(ar_paras)} Arabic blocks." if len(en_paras) != len(ar_paras) else None

    for i in range(max_len):
        en_obj = en_paras[i] if i < len(en_paras) else {'text': "[MISSING ENGLISH SOURCE]", 'start': None, 'end': None}
        ar_obj = ar_paras[i] if i < len(ar_paras) else {'text': "[MISSING ARABIC TRANSLATION]", 'start': None, 'end': None}
        
        status = 'normal'
        if en_obj['text'] == "[MISSING ENGLISH SOURCE]": status = 'misaligned_ar'
        elif ar_obj['text'] == "[MISSING ARABIC TRANSLATION]": status = 'misaligned_en'

        aligned_segments.append({'id': i + 1, 'status': status, 'english': en_obj['text'], 'arabic': ar_obj['text'], 'ar_start': ar_obj.get('start'), 'ar_end': ar_obj.get('end'), 'anomaly': anomaly_msg})
    return aligned_segments

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
        st.error(f"Failed to push to Drive. Error: {e}")
        return False

def push_to_drive_reviewer(document_id, approved_segments):
    docs_svc, _, _ = get_google_services()
    try:
        valid_segments = [seg for seg in approved_segments if seg.get('ar_start') is not None and seg.get('ar_end') is not None]
        valid_segments.sort(key=lambda x: x['ar_start'], reverse=True)
        requests = []
        for seg in valid_segments:
            requests.append({'deleteContentRange': {'range': {'startIndex': seg['ar_start'], 'endIndex': seg['ar_end'] - 1}}})
            requests.append({'insertText': {'location': {'index': seg['ar_start']}, 'text': seg['final_arabic']}})
            
        if requests:
            docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        return True
    except Exception as e:
        st.error(f"Failed to push to Drive. Error: {e}")
        return False


# ==========================================
# 6. DASHBOARD UI & INGESTION
# ==========================================
if not GENAI_AVAILABLE:
    st.error("🚨 Critical Dependency Missing: The `google-genai` package is not installed.")
    st.stop()

glossary_data = fetch_glossary()

col_title, col_logout = st.columns([5, 1])
with col_title: st.title("⚙️ 12-Step AI Suite")
with col_logout:
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.clear()
        st.rerun()

with st.container(border=True):
    col_main, col_info = st.columns([2, 1])
    with col_main:
        st.subheader(f"👋 Welcome, {st.session_state['user_name']}")
        st.markdown(f"**Current Role:** `{st.session_state['user_role'].capitalize()}`")
        st.caption(f"The UI has been customized for the {st.session_state['app_mode']}.")

    with col_info:
        st.markdown("### 📊 System Status")
        st.caption("⚡ **Engine:** `Micro-Batching & Session Guard Active`")
        if "No glossary" not in glossary_data and glossary_data != "":
            st.success(f"✅ **Glossary Connected:**\n`{GLOSSARY_RANGE.split('!')[0]}`")
        else:
            st.warning("⚠️ Glossary not active.")

st.divider()

st.markdown("### ☁️ Process Google Drive Document")
doc_url = st.text_input("Paste Google Docs File URL Here:")

if st.button("Load & Process Document", type="primary") and doc_url:
    file_id = extract_id(doc_url)
    if file_id:
        st.session_state['source_file_id'] = file_id
        lock_status = manage_document_lock(file_id, st.session_state['user_email'])
        
        if lock_status["status"] == "blocked":
            st.error(f"🛑 **Document In Use:** This document is currently being worked on by `{lock_status['locked_by']}`. Please coordinate before opening.")
            st.stop()
            
        elif lock_status["status"] == "recovered":
            st.session_state['session_row_index'] = lock_status["row_index"]
            st.session_state['processed_data'] = lock_status["data"]
            st.success("♻️ **Session Recovered:** Restored your previously saved work and progress.")
            time.sleep(0.5)
            st.rerun()
            
        elif lock_status["status"] == "clear":
            st.session_state['session_row_index'] = lock_status["row_index"]
            st.info(f"Connecting to Document ID: {file_id}...")
            paras = extract_text_from_drive(file_id)

            if paras:
                if st.session_state["app_mode"] == "Translator Mode":
                    st.info(f"Extracted {len(paras)} segments. Translating via AI Batching...")
                    progress_bar = st.progress(0)
                    processed_results = []
                    
                    batches = [paras[i:i + BATCH_SIZE] for i in range(0, len(paras), BATCH_SIZE)]
                    for idx, batch in enumerate(batches):
                        batch_payload = [{'id': j + 1, 'english': p['text']} for j, p in enumerate(batch)]
                        ai_results = translate_batch_with_fallback(batch_payload, glossary_data)
                        for ai_res in ai_results:
                            processed_results.append({
                                "id": ai_res.get('id', 0) + (idx * BATCH_SIZE),
                                "english": batch[ai_res.get('id', 1) - 1]['text'],
                                "arabic_translation": ai_res.get("arabic_translation", ""),
                                "glossary_notes": ai_res.get("glossary_notes", ""),
                                "user_arabic": ai_res.get("arabic_translation", ""),
                                "is_approved": False
                            })
                        progress_bar.progress((idx + 1) / len(batches))
                        
                else:
                    segments = smart_align_with_anomaly_detection(paras)
                    st.info("Extracted blocks. Aligning and Reviewing via AI Batching...")
                    progress_bar = st.progress(0)
                    processed_results = []
                    
                    normal_segs = [s for s in segments if s['status'] == 'normal']
                    batches = [normal_segs[i:i + BATCH_SIZE] for i in range(0, len(normal_segs), BATCH_SIZE)]
                    
                    for idx, batch in enumerate(batches):
                        ai_results = review_batch_with_fallback(batch, glossary_data)
                        for ai_res, original_seg in zip(ai_results, batch):
                            suggested = ai_res.get("suggested_arabic", original_seg.get('arabic', ''))
                            status_val = ai_res.get('status', 'minor_edits')
                            processed_results.append({
                                "id": original_seg.get('id'), "status": status_val,
                                "english": original_seg.get('english', ''), "original_arabic": original_seg.get('arabic', ''),
                                "suggested_arabic": suggested, "reasoning": ai_res.get("reasoning", ""),
                                "anomaly": original_seg.get('anomaly'), "ar_start": original_seg.get('ar_start'), "ar_end": original_seg.get('ar_end'),
                                "user_arabic": suggested, "is_approved": (status_val == 'perfect')
                            })
                        progress_bar.progress((idx + 1) / len(batches))
                        
                    for item in segments:
                        if item.get('status') == 'misaligned_en':
                            trans_res = translate_with_ai(item.get('english', ''), glossary_data)
                            t_arabic = trans_res.get("arabic_translation", "")
                            processed_results.append({
                                "id": item.get('id'), "status": "major_rewrite", "english": item.get('english', ''),
                                "original_arabic": "[MISSING]", "suggested_arabic": t_arabic,
                                "reasoning": "⚠️ Auto-translated orphaned English.", "anomaly": item.get('anomaly'),
                                "ar_start": None, "ar_end": None, "user_arabic": t_arabic, "is_approved": False
                            })
                        elif item.get('status') == 'misaligned_ar':
                            processed_results.append({
                                "id": item.get('id'), "status": "major_rewrite", "english": "[MISSING]",
                                "original_arabic": item.get('arabic', ''), "suggested_arabic": item.get('arabic', ''),
                                "reasoning": "⚠️ Orphaned Arabic. Ignored.", "anomaly": item.get('anomaly'),
                                "ar_start": item.get('ar_start'), "ar_end": item.get('ar_end'), "user_arabic": item.get('arabic', ''), "is_approved": False
                            })
                            
                    processed_results.sort(key=lambda x: x.get('id', 0))
                
                # Silent initial lock registration
                save_draft_to_sheet(
                    file_id, 
                    st.session_state['user_email'], 
                    st.session_state['session_row_index'], 
                    processed_results
                )
                st.session_state['processed_data'] = processed_results
                st.rerun()


# ==========================================
# 7. DYNAMIC OUTPUT GRID & PUSH ACTION
# ==========================================
if st.session_state['processed_data']:
    st.divider()
    approved_count = 0
    finalized_data = []
    state_modified = False

    if st.session_state["app_mode"] == "Translator Mode":
        st.subheader("Translation Editor")
        for i, item in enumerate(st.session_state['processed_data']):
            with st.container(border=True):
                seg_id = item.get('id', i + 1)
                st.markdown(f"### Segment {seg_id}")
                col_en, col_ar = st.columns(2)
                
                with col_en:
                    st.info(item.get('english', ''))
                    if item.get('glossary_notes'):
                        st.markdown("💡 **Glossary Notes:**")
                        st.caption(item.get('glossary_notes'))
                        
                with col_ar:
                    default_text = item.get('user_arabic', item.get('arabic_translation', ''))
                    final_text = st.text_area(
                        "Final Translation", 
                        value=default_text, 
                        height=120, 
                        key=f"edit_ar_{i}", 
                        label_visibility="collapsed"
                    )
                    if final_text != item.get('user_arabic'):
                        item['user_arabic'] = final_text
                        state_modified = True

                default_approval = item.get('is_approved', False)
                is_approved = st.checkbox(f"✅ Approve Segment {seg_id}", key=f"approve_{i}", value=default_approval)
                if is_approved != item.get('is_approved'):
                    item['is_approved'] = is_approved
                    state_modified = True

                if is_approved:
                    approved_count += 1
                    finalized_data.append(final_text)

    else:
        st.subheader("Review Segments & Diff Visualizer")
        for i, item in enumerate(st.session_state['processed_data']):
            # SECURE/DEFENSIVE READS TO PREVENT KEY ERRORS
            status_val = item.get('status', 'minor_edits')
            seg_id = item.get('id', i + 1)
            orig_ar = item.get('original_arabic', '')
            sugg_ar = item.get('suggested_arabic', item.get('arabic_translation', ''))
            reasoning_txt = item.get('reasoning', item.get('glossary_notes', 'No AI reasoning available.'))
            
            color = "🟢" if status_val == "perfect" else ("🟡" if status_val == "minor_edits" else "🔴")
            
            with st.container(border=True):
                st.markdown(f"### Segment {seg_id} | Status: {color} {status_val.upper()}")
                col_en, col_ar = st.columns(2)
                with col_en: 
                    st.info(item.get('english', ''))
                with col_ar:
                    diff_html = generate_html_diff(orig_ar, sugg_ar)
                    st.markdown(diff_html, unsafe_allow_html=True)
                    with st.expander("💡 View AI Reasoning"): 
                        st.markdown(reasoning_txt)
                    
                    default_text = item.get('user_arabic', sugg_ar)
                    final_text = st.text_area(
                        "Final Output", 
                        value=default_text, 
                        height=120, 
                        key=f"edit_ar_{i}", 
                        label_visibility="collapsed"
                    )
                    if final_text != item.get('user_arabic'):
                        item['user_arabic'] = final_text
                        state_modified = True
                
                default_approval = item.get('is_approved', (status_val == 'perfect'))
                is_approved = st.checkbox(f"✅ Approve Segment {seg_id}", key=f"approve_{i}", value=default_approval)
                if is_approved != item.get('is_approved'):
                    item['is_approved'] = is_approved
                    state_modified = True

                if is_approved:
                    approved_count += 1
                    finalized_data.append({'final_arabic': final_text, 'ar_start': item.get('ar_start'), 'ar_end': item.get('ar_end')})

    # Trigger background auto-save whenever edits or approvals change
    if state_modified and st.session_state.get('session_row_index'):
        save_draft_to_sheet(st.session_state['source_file_id'], st.session_state['user_email'], st.session_state['session_row_index'], st.session_state['processed_data'])

    # --- COMPILED EXPORT & PUSH BUTTON ---
    total_segments = len(st.session_state['processed_data'])
    st.divider()
    st.write(f"### **Approved Segments: {approved_count} / {total_segments}**")

    if approved_count == total_segments and total_segments > 0:
        st.success("🎉 All segments approved! Final review before pushing.")

        full_english = "\n\n".join([item.get('english', '') for item in st.session_state['processed_data']])
        if st.session_state["app_mode"] == "Translator Mode":
            full_arabic = "\n\n".join(finalized_data)
        else:
            full_arabic = "\n\n".join([item.get('final_arabic', '') for item in finalized_data])

        st.markdown("### 🔍 Final Full-Text Review")
        col_preview_en, col_preview_ar = st.columns(2)
        
        with col_preview_en:
            st.markdown(f"""<div style="height: 350px; overflow-y: auto; padding: 15px; background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; font-family: 'Inter', sans-serif; font-size: 15px; line-height: 1.6; white-space: pre-wrap; color: #334155;">{full_english}</div>""", unsafe_allow_html=True)
            
        with col_preview_ar:
            st.markdown(f"""<div dir="rtl" style="height: 350px; overflow-y: auto; padding: 15px; background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; font-family: 'Cairo', sans-serif; font-size: 18px; line-height: 1.8; white-space: pre-wrap; color: #0F172A; text-align: right;">{full_arabic}</div>""", unsafe_allow_html=True)

        operator_name = st.session_state["user_name"]
        operator_email = st.session_state["user_email"]

        st.divider()

        if "review_unlocked" not in st.session_state:
            st.session_state["review_unlocked"] = False

        if not st.session_state["review_unlocked"]:
            st.error("🚨 **CRITICAL STEP:** Please give the final text a quick read to ensure the narrative flows naturally and respects the glossary.")
            if st.button("👀 I have reviewed the final text and it looks good", use_container_width=True):
                st.session_state["review_unlocked"] = True
                st.rerun()
        else:
            if st.session_state["app_mode"] == "Translator Mode":
                if st.session_state.get('source_file_id'):
                    if st.button("🚀 Push Translation to Google Doc", type="primary", use_container_width=True):
                        with st.spinner("Pushing to Drive..."):
                            success = push_to_drive_translator(st.session_state['source_file_id'], full_arabic)
                            if success:
                                send_email_notification("Translation", operator_name, operator_email)
                                release_document_lock(st.session_state.get('session_row_index'))
                                st.session_state["review_unlocked"] = False
                                st.session_state["processed_data"] = None
                                st.balloons()
                                st.success("Translation pushed successfully and lock released!")
            else:
                if st.session_state.get('source_file_id'):
                    if st.button("🚀 Apply Revisions to Google Doc", type="primary", use_container_width=True):
                        with st.spinner("Rewriting Document..."):
                            success = push_to_drive_reviewer(st.session_state['source_file_id'], finalized_data)
                            if success:
                                send_email_notification("Review", operator_name, operator_email)
                                release_document_lock(st.session_state.get('session_row_index'))
                                st.session_state["review_unlocked"] = False
                                st.session_state["processed_data"] = None
                                st.balloons()
                                st.success("Revisions applied successfully and lock released!")
