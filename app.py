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
st.set_page_config(page_title="12-Step AI Suite: Workflow Portal", layout="wide", initial_sidebar_state="collapsed")

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
            transition: all 0.2s ease-in-out !important;
        }
        button[kind="primary"]:hover {
            background-color: #059669 !important;
            border-color: #059669 !important;
            transform: translateY(-2px);
        }

        .task-card {
            background-color: #F8FAFC;
            border: 1px solid #E2E8F0;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 15px;
            border-left: 5px solid #3B82F6;
        }
        
        div[data-testid="stAlert"]:has(p:contains("CRITICAL STEP")) {
            background-color: #FEF2F2 !important;
            border: 1px solid #F87171 !important;
            color: #991B1B !important;
            border-radius: 8px !important;
        }
    </style>
    """, unsafe_allow_html=True)

apply_custom_css()

# Google Sheets Configuration
GLOSSARY_SPREADSHEET_ID = "1oc4TCY_iK9R7mBiXgb5rKWssjmrQywYg6UpOBXx8pUQ"
GLOSSARY_RANGE = "'المصطلحات'!A:D"
GLOSSARY_DATA_RANGE = "'المصطلحات'!C:D"
SESSIONS_RANGE = "'Sessions'!A:D"
VOLUNTEERS_RANGE = "'Volunteers'!A:D"
ASSIGNMENTS_RANGE = "'Assignments'!A:E"

LOCK_TIMEOUT_SECONDS = 14400 # 4 hours before a lock naturally expires

# AI Parameters
BATCH_SIZE = 6
MAX_RETRIES_PER_MODEL = 3
BASE_BACKOFF_SECONDS = 2.0
MAX_BACKOFF_SECONDS = 15.0
RETRYABLE_KEYWORDS = ("503", "500", "high demand", "429", "timeout", "Quota")

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
        return (
            build('docs', 'v1', credentials=credentials, cache_discovery=False),
            build('drive', 'v3', credentials=credentials, cache_discovery=False),
            build('sheets', 'v4', credentials=credentials, cache_discovery=False)
        )
    except Exception as e:
        st.error(f"Google Auth Error: {e}")
        return None, None, None

docs_service, drive_service, sheets_service = get_google_services()

def fetch_volunteers():
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=VOLUNTEERS_RANGE).execute()
        rows = res.get('values', [])
        volunteers = {}
        if len(rows) > 1:
            for r in rows[1:]:
                if len(r) >= 4:
                    email, name, role, status = r[0].strip().lower(), r[1], r[2].lower(), r[3]
                    volunteers[email] = {"name": name, "role": role, "status": status}
        return volunteers
    except Exception:
        return {}

def overwrite_sheet_data(range_name, data_matrix):
    sheets_service.spreadsheets().values().clear(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name).execute()
    body = {'values': data_matrix}
    sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name, valueInputOption="USER_ENTERED", body=body).execute()

# --- ASSIGNMENTS DATABASE FUNCTIONS ---
def fetch_assignments():
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=ASSIGNMENTS_RANGE).execute()
        rows = res.get('values', [])
        assignments = []
        if len(rows) > 1:
            for idx, r in enumerate(rows[1:]):
                r.extend([""] * (5 - len(r))) # Pad empty cells for safety
                assignments.append({
                    "row_index": idx + 2,
                    "doc_id": r[0].strip(),
                    "doc_name": r[1],
                    "translator": r[2].lower().strip(),
                    "reviewer": r[3].lower().strip(),
                    "status": r[4]
                })
        return assignments
    except Exception as e:
        return []

def assign_task_to_sheet(doc_id, doc_name, t_email, r_email, status):
    assignments = fetch_assignments()
    row_idx = None
    for row in assignments:
        if row['doc_id'] == doc_id:
            row_idx = row['row_index']
            break

    body = {'values': [[doc_id, doc_name, t_email, r_email, status]]}
    if row_idx:
        range_name = f"'Assignments'!A{row_idx}:E{row_idx}"
        sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name, valueInputOption="USER_ENTERED", body=body).execute()
    else:
        range_name = "'Assignments'!A:E"
        sheets_service.spreadsheets().values().append(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name, valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS", body=body).execute()

def update_assignment_status(doc_id, new_status):
    assignments = fetch_assignments()
    for task in assignments:
        if task['doc_id'] == doc_id:
            range_name = f"'Assignments'!E{task['row_index']}"
            body = {'values': [[new_status]]}
            sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name, valueInputOption="USER_ENTERED", body=body).execute()
            break

# --- LOGIN SCREEN ---
def login_screen():
    if st.session_state.get("authenticated"): return True
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🤝 12-Step AI Suite")
        st.caption("Volunteer Translation & Review Portal")
        with st.container(border=True):
            with st.form("login_form"):
                email = st.text_input("Enter your registered email address:").strip().lower()
                if st.form_submit_button("Access Portal", type="primary", use_container_width=True):
                    volunteers = fetch_volunteers()
                    if email in volunteers:
                        user_data = volunteers[email]
                        if user_data["status"].lower() != "active":
                            st.error("Account suspended. Please contact the coordinator.")
                        else:
                            st.session_state.update({"authenticated": True, "user_email": email, "user_role": user_data["role"], "user_name": user_data["name"]})
                            st.session_state["app_mode"] = "God Mode" if user_data["role"] == "admin" else ("Translator Mode" if user_data["role"] == "translator" else "Reviewer Mode")
                            st.rerun()
                    else:
                        st.error("Email not recognized in the system.")
    return False

if not login_screen(): st.stop()

# Initialize core session variables securely
if 'processed_data' not in st.session_state: st.session_state['processed_data'] = None
if 'source_file_id' not in st.session_state: st.session_state['source_file_id'] = None
if 'session_row_index' not in st.session_state: st.session_state['session_row_index'] = None
if 'active_task' not in st.session_state: st.session_state['active_task'] = None


# ==========================================
# 3. GOD MODE (ADMIN DASHBOARD)
# ==========================================
if st.session_state.get("app_mode") == "God Mode" and not st.session_state.get('source_file_id'):
    col1, col2 = st.columns([5, 1])
    col1.title("⚡ Central Command: God Mode")
    if col2.button("🚪 Logout", type="primary"): 
        st.session_state.clear(); st.rerun()
        
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📡 Live Radar", "👥 Volunteers", "📖 Glossary", "📢 Broadcast", "🗂 Assignment Desk"
    ])
    
    # --- TAB 1: LIVE RADAR ---
    with tab1:
        st.subheader("Active Document Locks")
        try:
            res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE).execute()
            session_rows = res.get('values', [])
            active_locks = False
            if len(session_rows) > 1:
                for idx, row in enumerate(session_rows[1:]): 
                    if len(row) > 1 and row[1]: 
                        active_locks = True
                        doc_id, locked_by = row[0], row[1]
                        ts = float(row[2]) if len(row) > 2 and row[2] else 0
                        time_locked = round((time.time() - ts) / 60, 1) if ts else 0
                        
                        with st.container(border=True):
                            col_l1, col_l2 = st.columns([3, 1])
                            with col_l1:
                                st.markdown(f"🔒 Locked by **{locked_by}** for `{time_locked} minutes`")
                                st.caption(f"Doc ID: {doc_id}")
                            with col_l2:
                                if st.button("🧨 Kill Lock", key=f"kill_{idx}", use_container_width=True):
                                    sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Sessions'!A{idx+2}:D{idx+2}", valueInputOption="USER_ENTERED", body={'values': [["", "", "", ""]]}).execute()
                                    st.rerun()
            if not active_locks:
                st.info("No active sessions right now. All clear!")
        except Exception as e:
            st.error(f"Error fetching radar: {e}")

    # --- TAB 2: VOLUNTEERS DATABASE ---
    with tab2:
        st.subheader("Manage Volunteer Access")
        try:
            res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=VOLUNTEERS_RANGE).execute()
            vol_data = res.get('values', [])
            if not vol_data: vol_data = [["Email", "Name", "Role", "Status"]]
            df_vol = pd.DataFrame(vol_data[1:], columns=vol_data[0])
            edited_vol = st.data_editor(df_vol, num_rows="dynamic", use_container_width=True,
                column_config={
                    "Role": st.column_config.SelectboxColumn("Role", options=["translator", "reviewer", "admin"], required=True),
                    "Status": st.column_config.SelectboxColumn("Status", options=["Active", "Suspended"], required=True)
                })
            if st.button("💾 Save Volunteers", type="primary"):
                overwrite_sheet_data(VOLUNTEERS_RANGE, [edited_vol.columns.tolist()] + edited_vol.fillna("").values.tolist())
                st.success("Volunteers database updated!")
        except Exception as e:
            st.error(f"Error fetching volunteers: {e}")

    # --- TAB 3: GLOSSARY COMMAND CENTER ---
    with tab3:
        st.subheader("Live Terminology Editor")
        try:
            res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_RANGE).execute()
            glos_data = res.get('values', [])
            headers = ["ID", "Category", "English", "Arabic"]
            rows_to_display = glos_data[1:] if (glos_data and len(glos_data[0])==4 and not glos_data[0][0].isdigit()) else glos_data
            cleaned_rows = [list(row) + [""] * (4 - len(row)) for row in rows_to_display]
            if not cleaned_rows: cleaned_rows = [["1", "General", "", ""]]
            df_glos = pd.DataFrame(cleaned_rows, columns=headers)
            edited_glos = st.data_editor(df_glos, num_rows="dynamic", use_container_width=True)
            if st.button("💾 Sync Glossary to AI", type="primary"):
                overwrite_sheet_data(GLOSSARY_RANGE, [edited_glos.columns.tolist()] + edited_glos.fillna("").values.tolist())
                st.success("Glossary synced successfully!")
                st.cache_data.clear()
        except Exception as e:
            st.error(f"Error fetching glossary: {e}")

    # --- TAB 4: BROADCAST DESK ---
    with tab4:
        st.subheader("Team Broadcast System")
        st.caption("Send a mass email to all 'Active' volunteers.")
        broadcast_subject = st.text_input("Subject")
        broadcast_message = st.text_area("Message Body", height=150)
        
        if st.button("🚀 Send Broadcast", type="primary"):
            if broadcast_subject and broadcast_message:
                vols = fetch_volunteers()
                active_emails = [email for email, d in vols.items() if d['status'].lower() == 'active']
                if active_emails:
                    with st.spinner("Dispatching emails via secure SMTP..."):
                        try:
                            msg = EmailMessage()
                            msg.set_content(f"12-Step Translation Project Update:\n\n{broadcast_message}")
                            msg['Subject'] = f"[12-Step Admin] {broadcast_subject}"
                            msg['From'] = st.secrets.get("SMTP_EMAIL", "admin@localhost")
                            msg['To'] = st.secrets.get("SMTP_EMAIL", "admin@localhost")
                            msg['Bcc'] = ", ".join(active_emails) 
                            
                            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                                server.login(st.secrets["SMTP_EMAIL"], st.secrets["SMTP_PASSWORD"])
                                server.send_message(msg)
                            st.success(f"Broadcast sent successfully to {len(active_emails)} volunteers!")
                            st.balloons()
                        except Exception as e:
                            st.error(f"Failed to send email. Check your SMTP secrets. Error: {e}")
                else:
                    st.warning("No active volunteers found to email.")
            else:
                st.warning("Please enter a subject and a message.")

    # --- TAB 5: ASSIGNMENT DESK ---
    with tab5:
        st.subheader("Workflow Assignment Desk")
        st.info("💡 **Note:** Translator assignment is optional. If left unassigned (Bypass), the AI will translate the entire document and route it directly to the Reviewer's inbox.")
        
        try:
            folder_res = drive_service.files().list(q="mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute()
            all_folders = folder_res.get('files', [])
            year_folders = [f for f in all_folders if 'Edition' in f['name'] or '202' in f['name']]
            
            if year_folders:
                year_options = {f['name']: f['id'] for f in year_folders}
                sel_year = st.selectbox("📁 1. Select Year Folder:", list(year_options.keys()))
                year_id = year_options[sel_year]
                
                month_res = drive_service.files().list(q=f"'{year_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute()
                month_folders = month_res.get('files', [])
                
                if month_folders:
                    month_options = {f['name']: f['id'] for f in month_folders}
                    sel_month = st.selectbox("📂 2. Select Month Folder:", list(month_options.keys()))
                    month_id = month_options[sel_month]
                    
                    if st.button("🔄 Fetch Documents", type="primary"):
                        st.session_state['scanned_month_id'] = month_id
                else:
                    st.warning("No subfolders found in this year.")
            else:
                st.warning("No Year folders found in Drive.")
                
            st.divider()

            if 'scanned_month_id' in st.session_state:
                doc_query = f"'{st.session_state['scanned_month_id']}' in parents and mimeType='application/vnd.google-apps.document' and trashed=false"
                doc_res = drive_service.files().list(q=doc_query, fields="files(id, name)").execute()
                docs = doc_res.get('files', [])
                
                if docs:
                    st.markdown(f"### Available Documents ({len(docs)})")
                    vols = fetch_volunteers()
                    t_list = ["[Optional] Bypass - AI Only"] + [e for e, d in vols.items() if d['role'] in ['translator', 'admin'] and d['status'].lower() == 'active']
                    r_list = ["[Mandatory] Select Reviewer..."] + [e for e, d in vols.items() if d['role'] in ['reviewer', 'admin'] and d['status'].lower() == 'active']
                    
                    for doc in docs:
                        with st.container(border=True):
                            c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
                            c1.markdown(f"📄 **{doc['name']}**")
                            t_sel = c2.selectbox("Translator", t_list, key=f"t_{doc['id']}")
                            r_sel = c3.selectbox("Reviewer", r_list, key=f"r_{doc['id']}")
                            
                            if c4.button("🚀 Assign Task", key=f"btn_{doc['id']}", use_container_width=True):
                                if r_sel == r_list[0]:
                                    st.error("❌ You must select a Reviewer!")
                                else:
                                    t_email = "" if t_sel == t_list[0] else t_sel
                                    status = "Pending Review" if t_email == "" else "Pending Translation"
                                    assign_task_to_sheet(doc['id'], doc['name'], t_email, r_sel, status)
                                    msg = f"Task assigned! Skipped human translation." if not t_email else f"Task assigned: {t_email} -> {r_sel}."
                                    st.success(f"✅ {msg}")
                else:
                    st.info("No Google Docs found in this folder.")
        except Exception as e:
            st.error(f"Google Drive Error: {e}")

    st.stop() # Hide user interface from God Mode

# ==========================================
# 4. USER INBOX & TASK DELEGATION
# ==========================================
if not st.session_state.get('source_file_id'):
    col_t, col_l = st.columns([5, 1])
    col_t.title("⚙️ 12-Step AI Suite")
    if col_l.button("🚪 Logout", use_container_width=True): 
        st.session_state.clear(); st.rerun()
    
    st.subheader(f"👋 Welcome, {st.session_state['user_name']} | Role: {st.session_state['user_role'].capitalize()}")
    st.markdown("---")
    
    st.markdown("### 📬 Your Task Queue")
    assignments = fetch_assignments()
    my_tasks = []
    
    for task in assignments:
        if st.session_state["app_mode"] == "Translator Mode" and task.get('translator') == st.session_state['user_email'] and task.get('status') == "Pending Translation":
            my_tasks.append(task)
        elif st.session_state["app_mode"] == "Reviewer Mode" and task.get('reviewer') == st.session_state['user_email'] and task.get('status') == "Pending Review":
            my_tasks.append(task)

    if not my_tasks:
        st.success("🎉 You have no pending tasks in your queue. Great job!")
    else:
        for task in my_tasks:
            with st.container(border=True):
                c1, c2 = st.columns([4, 1])
                c1.markdown(f"### 📄 **{task['doc_name']}**")
                c1.caption(f"Status: `{task['status']}` | ID: {task['doc_id']}")
                if c2.button("🚀 Start Work", key=f"start_{task['doc_id']}", type="primary", use_container_width=True):
                    st.session_state['active_task'] = task
                    st.session_state['source_file_id'] = task['doc_id']
                    st.rerun()
    st.stop() # Wait for user to pick a task


# ==========================================
# 5. AI ENGINE & DOCUMENT PARSING
# ==========================================
client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"]) if GENAI_AVAILABLE else None
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

class TranslationBatchItem(BaseModel):
    id: int
    arabic_translation: str
    glossary_notes: str

class TranslationBatchResult(BaseModel):
    items: list[TranslationBatchItem]

class ReviewResult(BaseModel):
    status: str = Field(description="Must be 'perfect', 'minor_edits', or 'major_rewrite'")
    suggested_arabic: str = Field(description="The finalized Arabic text")
    reasoning: str = Field(description="Detailed explanation of changes based on semantics and glossary")

class ReviewBatchItem(BaseModel):
    id: int
    status: str
    suggested_arabic: str
    reasoning: str

class ReviewBatchResult(BaseModel):
    items: list[ReviewBatchItem]

def get_fallback_models():
    secret_model = st.secrets.get("ACTIVE_MODEL", "").strip()
    if secret_model: return [secret_model]
    return ["gemini-2.5-flash", "gemini-1.5-flash"] 

def manage_document_lock(file_id: str, user_email: str):
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE).execute()
        rows = res.get('values', [])
        current_time = time.time()
        target_row_index = max(len(rows) + 1, 2)
        
        for index, row in enumerate(rows):
            if index == 0: continue
            if (len(row) > 0 and row[0] == file_id):
                locked_by = row[1] if len(row) > 1 else ""
                ts = float(row[2]) if len(row) > 2 and row[2] else 0
                saved_json = row[3] if len(row) > 3 else ""
                
                if locked_by and locked_by != user_email and (current_time - ts) < LOCK_TIMEOUT_SECONDS:
                    return {"status": "blocked", "locked_by": locked_by, "row_index": index + 1}
                if locked_by == user_email and saved_json:
                    try:
                        return {"status": "recovered", "data": json.loads(saved_json), "row_index": index + 1}
                    except Exception: pass
                return {"status": "clear", "row_index": index + 1}
        return {"status": "clear", "row_index": target_row_index}
    except Exception:
        return {"status": "clear", "row_index": 2}

def save_draft_to_sheet(file_id, user_email, row_index, processed_data):
    if not row_index: return
    try:
        json_data = json.dumps(processed_data, ensure_ascii=False)
        body = {'values': [[file_id, user_email, str(time.time()), json_data]]}
        sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Sessions'!A{row_index}:D{row_index}", valueInputOption="USER_ENTERED", body=body).execute()
    except Exception: pass

def release_document_lock(row_index):
    if not row_index: return
    try:
        sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Sessions'!A{row_index}:D{row_index}", valueInputOption="USER_ENTERED", body={'values': [["", "", "", ""]]}).execute()
    except Exception: pass

@st.cache_data(ttl=3600)
def fetch_glossary():
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_DATA_RANGE).execute()
        vals = res.get('values', [])
        glos = "12-Step Glossary:\n"
        for row in vals:
            if len(row) >= 2: glos += f"- {row[0]} -> {row[1]}\n"
        return glos
    except Exception: return ""

def generate_html_diff(original, suggested):
    if not original or original.startswith("[MISSING"): return "<div dir='rtl' style='text-align: right; color: #0369a1; background-color: #e0f2fe; padding: 10px; border-radius: 5px; font-family: \"Cairo\";'>✨ New AI Translation (Bypass Mode)</div>"
    if original.strip() == suggested.strip(): return "<div dir='rtl' style='text-align: right; color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px; font-family: \"Cairo\";'>✨ Perfect Match</div>"
    
    diff = difflib.ndiff(original.split(), suggested.split())
    html = ["<div dir='rtl' style='font-family: \"Cairo\", sans-serif; font-size: 18px; line-height: 2; text-align: right; background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef;'>"]
    for word in diff:
        if word.startswith('- '): html.append(f"<span style='background-color: #ffcdd2; color: #b71c1c; text-decoration: line-through; padding: 2px; border-radius: 4px;'>{word[2:]}</span>")
        elif word.startswith('+ '): html.append(f"<span style='background-color: #c8e6c9; color: #1b5e20; font-weight: bold; padding: 2px; border-radius: 4px;'>{word[2:]}</span>")
        elif word.startswith('  '): html.append(f"<span style='color: #212529;'>{word[2:]}</span>")
    html.append("</div>")
    return " ".join(html)

def _is_retryable(err_str): return any(kw.lower() in err_str.lower() for kw in RETRYABLE_KEYWORDS)
def _backoff_sleep(attempt): time.sleep(min(BASE_BACKOFF_SECONDS * (2 ** attempt), MAX_BACKOFF_SECONDS))

def _call_gemini(model_name, prompt, schema_type):
    response = client.models.generate_content(
        model=model_name, contents=prompt,
        config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=schema_type, safety_settings=safety_settings, temperature=0.2),
    )
    clean_text = response.text.replace("```json", "").replace("```", "").strip()
    match = re.search(r'\{.*\}', clean_text, re.DOTALL)
    if match: clean_text = match.group(0)
    return json.loads(clean_text)

def translate_with_ai(english: str, glossary_text: str):
    prompt = f"Expert translator for 12-step literature.\nGLOSSARY:\n{glossary_text}\nTranslate:\n{english}"
    for model_name in get_fallback_models():
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                parsed = _call_gemini(model_name, prompt, TranslationResult)
                return {"arabic_translation": parsed.get("arabic_translation", ""), "glossary_notes": parsed.get("glossary_notes", "")}
            except Exception as e:
                if _is_retryable(str(e)): _backoff_sleep(attempt)
    return {"arabic_translation": "", "glossary_notes": "⚠ Error"}

def review_with_ai(english: str, arabic: str, glossary_text: str):
    prompt = f"Expert editor for 12-step literature.\nGLOSSARY:\n{glossary_text}\nEnglish: {english}\nArabic: {arabic}"
    for model_name in get_fallback_models():
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                parsed = _call_gemini(model_name, prompt, ReviewResult)
                return {"status": parsed.get("status", "minor_edits"), "suggested_arabic": parsed.get("suggested_arabic", arabic), "reasoning": parsed.get("reasoning", "")}
            except Exception as e:
                if _is_retryable(str(e)): _backoff_sleep(attempt)
    return {"status": "major_rewrite", "suggested_arabic": arabic, "reasoning": "⚠ Error"}

def translate_batch_with_fallback(batch_segments, glossary_text):
    if not GENAI_AVAILABLE or client is None: return [translate_with_ai(s['english'], glossary_text) for s in batch_segments]
    input_payload = "\n\n".join([f"ID: {s.get('id', 0)}\nText: {s.get('english', '')}" for s in batch_segments])
    prompt = f"Expert translator for 12-step literature.\nGLOSSARY:\n{glossary_text}\nSegments:\n{input_payload}"
    for model_name in get_fallback_models():
        for attempt in range(2):
            try:
                parsed = _call_gemini(model_name, prompt, TranslationBatchResult)
                items = parsed.get("items", [])
                if len(items) == len(batch_segments): return items
            except Exception as e:
                if _is_retryable(str(e)): _backoff_sleep(attempt)
    
    results = []
    for s in batch_segments:
        res = translate_with_ai(s.get('english', ''), glossary_text)
        results.append({"id": s.get('id', 0), "arabic_translation": res.get('arabic_translation', ''), "glossary_notes": res.get('glossary_notes', '')})
    return results

def review_batch_with_fallback(batch_segments, glossary_text):
    if not GENAI_AVAILABLE or client is None: return [review_with_ai(s.get('english', ''), s.get('arabic', ''), glossary_text) for s in batch_segments]
    input_payload = "\n\n".join([f"ID: {s.get('id', 0)}\nEnglish: {s.get('english', '')}\nArabic: {s.get('arabic', '')}" for s in batch_segments])
    prompt = f"Expert editor for 12-step literature.\nGLOSSARY:\n{glossary_text}\nPairs:\n{input_payload}"
    for model_name in get_fallback_models():
        for attempt in range(2):
            try:
                parsed = _call_gemini(model_name, prompt, ReviewBatchResult)
                items = parsed.get("items", [])
                if len(items) == len(batch_segments): return items
            except Exception as e:
                if _is_retryable(str(e)): _backoff_sleep(attempt)
                
    results = []
    for s in batch_segments:
        res = review_with_ai(s.get('english', ''), s.get('arabic', ''), glossary_text)
        results.append({"id": s.get('id', 0), "status": res.get('status', 'minor_edits'), "suggested_arabic": res.get('suggested_arabic', ''), "reasoning": res.get('reasoning', '')})
    return results

def _parse_docs_elements(elements):
    paras = []
    for elem in elements:
        if 'paragraph' in elem:
            para_text = "".join([run.get('textRun', {}).get('content', '') for run in elem.get('paragraph', {}).get('elements', []) if 'textRun' in run])
            if re.search(r'[a-zA-Z\u0600-\u06FF]', para_text.strip()): 
                paras.append({'text': para_text.strip(), 'start': elem.get('startIndex'), 'end': elem.get('endIndex')})
        elif 'table' in elem:
            for row in elem.get('table', {}).get('tableRows', []):
                for cell in row.get('tableCells', []):
                    paras.extend(_parse_docs_elements(cell.get('content', [])))
    return paras

def extract_text_from_drive(file_id):
    try:
        document = docs_service.documents().get(documentId=file_id).execute()
        return _parse_docs_elements(document.get('body', {}).get('content', []))
    except Exception as e:
        st.error(f"Error reading doc: {e}")
        return None

def smart_align(paragraphs):
    en_paras, ar_paras = [p for p in paragraphs if not re.search(r'[\u0600-\u06FF]', p['text'])], [p for p in paragraphs if re.search(r'[\u0600-\u06FF]', p['text'])]
    aligned = []
    for i in range(max(len(en_paras), len(ar_paras))):
        en_obj = en_paras[i] if i < len(en_paras) else {'text': "[MISSING ENGLISH SOURCE]"}
        ar_obj = ar_paras[i] if i < len(ar_paras) else {'text': "[MISSING ARABIC TRANSLATION]", 'start': None, 'end': None}
        aligned.append({'id': i + 1, 'english': en_obj.get('text', ''), 'arabic': ar_obj.get('text', ''), 'ar_start': ar_obj.get('start'), 'ar_end': ar_obj.get('end')})
    return aligned

def push_translator(doc_id, text):
    try:
        reqs = [{'insertPageBreak': {'location': {'index': 1}}}, {'insertText': {'location': {'index': 1}, 'text': text + "\n\n"}}]
        docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': reqs}).execute()
        return True
    except Exception as e: 
        st.error(f"Write error: {e}")
        return False

def push_reviewer(doc_id, segments):
    try:
        reqs = []
        for s in sorted([seg for seg in segments if seg.get('ar_start') is not None], key=lambda x: x['ar_start'], reverse=True):
            reqs.append({'deleteContentRange': {'range': {'startIndex': s['ar_start'], 'endIndex': s['ar_end'] - 1}}})
            reqs.append({'insertText': {'location': {'index': s['ar_start']}, 'text': s['final_arabic']}})
        if reqs: docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': reqs}).execute()
        return True
    except Exception as e:
        st.error(f"Failed to apply revisions. Error: {e}")
        return False

# ==========================================
# 6. ACTIVE WORKSPACE (THE EDITOR)
# ==========================================

# 🛡️ DEFENSIVE GUARD: If user hard-refreshes and loses task state, send them back to Inbox safely.
task = st.session_state.get('active_task')
if not task:
    st.session_state['source_file_id'] = None
    st.rerun()

file_id = task.get('doc_id')

col_h1, col_h2 = st.columns([5, 1])
col_h1.markdown(f"## 📝 Workspace: `{task.get('doc_name', 'Document')}`")
if col_h2.button("⬅️ Back to Inbox", use_container_width=True):
    st.session_state['active_task'] = None
    st.session_state['source_file_id'] = None
    st.session_state['processed_data'] = None
    st.rerun()

if not st.session_state.get('processed_data'):
    lock = manage_document_lock(file_id, st.session_state['user_email'])
    if lock["status"] == "blocked":
        st.error(f"🛑 **Document In Use:** This document is currently locked by `{lock.get('locked_by', 'another user')}`.")
        st.stop()
    elif lock["status"] == "recovered":
        st.session_state['session_row_index'] = lock["row_index"]
        st.session_state['processed_data'] = lock["data"]
        st.success("♻️ **Session Recovered.** Restored your previous work.")
    elif lock["status"] == "clear":
        st.session_state['session_row_index'] = lock["row_index"]
        st.info("Extracting document content and analyzing via AI...")
        paras = extract_text_from_drive(file_id)
        
        if paras:
            processed_results = []
            progress_bar = st.progress(0)
            
            if st.session_state["app_mode"] == "Translator Mode":
                batches = [paras[i:i + BATCH_SIZE] for i in range(0, len(paras), BATCH_SIZE)]
                for idx, batch in enumerate(batches):
                    ai_results = translate_batch_with_fallback([{'id': j + 1, 'english': p.get('text', '')} for j, p in enumerate(batch)], glossary_data)
                    for ai_res in ai_results:
                        processed_results.append({
                            "id": ai_res.get('id', 0) + (idx * BATCH_SIZE),
                            "english": batch[ai_res.get('id', 1) - 1].get('text', ''),
                            "arabic_translation": ai_res.get("arabic_translation", ""),
                            "glossary_notes": ai_res.get("glossary_notes", ""),
                            "user_arabic": ai_res.get("arabic_translation", ""),
                            "is_approved": False
                        })
                    progress_bar.progress((idx + 1) / len(batches))
            else:
                segments = smart_align(paras)
                normal_segs = [s for s in segments if s.get('english') != "[MISSING ENGLISH SOURCE]" and s.get('arabic') != "[MISSING ARABIC TRANSLATION]"]
                batches = [normal_segs[i:i + BATCH_SIZE] for i in range(0, len(normal_segs), BATCH_SIZE)]
                
                for idx, batch in enumerate(batches):
                    ai_results = review_batch_with_fallback(batch, glossary_data)
                    for ai_res, o_seg in zip(ai_results, batch):
                        suggested = ai_res.get("suggested_arabic", o_seg.get('arabic', ''))
                        status_val = ai_res.get('status', 'minor_edits')
                        processed_results.append({
                            "id": o_seg.get('id'), "status": status_val,
                            "english": o_seg.get('english', ''), "original_arabic": o_seg.get('arabic', ''),
                            "suggested_arabic": suggested, "reasoning": ai_res.get("reasoning", ""),
                            "ar_start": o_seg.get('ar_start'), "ar_end": o_seg.get('ar_end'),
                            "user_arabic": suggested, "is_approved": (status_val == 'perfect')
                        })
                    progress_bar.progress((idx + 1) / len(batches))
                
                # Handle Anomalies gracefully (Lazy AI Bypass evaluation for Reviewers)
                for item in segments:
                    if item.get('english') == "[MISSING ENGLISH SOURCE]":
                        processed_results.append({
                            "id": item.get('id'), "status": "major_rewrite", "english": "[MISSING]",
                            "original_arabic": item.get('arabic', ''), "suggested_arabic": item.get('arabic', ''),
                            "reasoning": "⚠️️ Orphaned Arabic block.", 
                            "ar_start": item.get('ar_start'), "ar_end": item.get('ar_end'), 
                            "user_arabic": item.get('arabic', ''), "is_approved": False
                        })
                    elif item.get('arabic') == "[MISSING ARABIC TRANSLATION]":
                        trans_res = translate_with_ai(item.get('english', ''), glossary_data)
                        t_arabic = trans_res.get("arabic_translation", "")
                        processed_results.append({
                            "id": item.get('id'), "status": "major_rewrite", "english": item.get('english', ''),
                            "original_arabic": "[MISSING]", "suggested_arabic": t_arabic,
                            "reasoning": "⚠️ Auto-translated orphaned English block.", 
                            "ar_start": None, "ar_end": None, "user_arabic": t_arabic, "is_approved": False
                        })
                        
                processed_results.sort(key=lambda x: x.get('id', 0))
            
            save_draft_to_sheet(file_id, st.session_state['user_email'], st.session_state['session_row_index'], processed_results)
            st.session_state['processed_data'] = processed_results
            st.rerun()

# --- EDITOR UI ---
approved_count, finalized_data, state_modified = 0, [], False
st.divider()

for i, item in enumerate(st.session_state['processed_data']):
    seg_id = item.get('id', i + 1)
    status_val = item.get('status', 'minor_edits')
    eng_txt = item.get('english', '')
    orig_ar = item.get('original_arabic', '')
    sugg_ar = item.get('suggested_arabic', item.get('arabic_translation', ''))
    
    with st.container(border=True):
        if st.session_state["app_mode"] == "Translator Mode":
            st.markdown(f"### Segment {seg_id}")
            col_en, col_ar = st.columns(2)
            with col_en:
                st.info(eng_txt)
                if item.get('glossary_notes'): st.caption(f"💡 **AI Notes:** {item.get('glossary_notes')}")
            with col_ar:
                default_val = item.get('user_arabic', sugg_ar)
                final_text = st.text_area("Final text", value=default_val, height=120, key=f"edit_{i}", label_visibility="collapsed")
                if final_text != item.get('user_arabic'): 
                    item['user_arabic'] = final_text
                    state_modified = True
        else:
            color = "🟢" if status_val == "perfect" else ("🟡" if status_val == "minor_edits" else "🔴")
            st.markdown(f"### Segment {seg_id} | Status: {color} {status_val.upper()}")
            col_en, col_ar = st.columns(2)
            with col_en: 
                st.info(eng_txt)
            with col_ar:
                st.markdown(generate_html_diff(orig_ar, sugg_ar), unsafe_allow_html=True)
                with st.expander("💡 View AI Reasoning"): 
                    st.markdown(item.get('reasoning', item.get('glossary_notes', 'No reasoning provided.')))
                
                default_val = item.get('user_arabic', sugg_ar)
                final_text = st.text_area("Final Output", value=default_val, height=120, key=f"edit_ar_{i}", label_visibility="collapsed")
                if final_text != item.get('user_arabic'): 
                    item['user_arabic'] = final_text
                    state_modified = True
            
        chk = st.checkbox(f"✅ Approve Segment {seg_id}", key=f"chk_{i}", value=item.get('is_approved', False))
        if chk != item.get('is_approved'): 
            item['is_approved'] = chk
            state_modified = True
            
        if chk:
            approved_count += 1
            if st.session_state["app_mode"] == "Translator Mode": 
                finalized_data.append(final_text)
            else: 
                finalized_data.append({'final_arabic': final_text, 'ar_start': item.get('ar_start'), 'ar_end': item.get('ar_end')})

if state_modified: 
    save_draft_to_sheet(file_id, st.session_state['user_email'], st.session_state['session_row_index'], st.session_state['processed_data'])

# --- SUBMISSION LOGIC ---
st.divider()
total_segments = len(st.session_state['processed_data'])
st.write(f"### **Approved Segments: {approved_count} / {total_segments}**")

if approved_count == total_segments and total_segments > 0:
    st.success("🎉 All segments approved! Final review before pushing.")
    
    if "review_unlocked" not in st.session_state: st.session_state["review_unlocked"] = False
    
    if not st.session_state["review_unlocked"]:
        st.error("🚨 **CRITICAL STEP:** Please verify the narrative flow before finalizing.")
        if st.button("👀 I confirm the final text is correct", use_container_width=True):
            st.session_state["review_unlocked"] = True
            st.rerun()
    else:
        if st.button("🚀 Push to Drive & Close Task", type="primary", use_container_width=True):
            with st.spinner("Processing Drive updates and concluding workflow..."):
                if st.session_state["app_mode"] == "Translator Mode":
                    success = push_translator(file_id, "\n\n".join(finalized_data))
                    if success:
                        new_status = "Completed" if task.get('translator') == task.get('reviewer') else "Pending Review"
                        update_assignment_status(file_id, new_status)
                else:
                    success = push_reviewer(file_id, finalized_data)
                    if success: update_assignment_status(file_id, "Completed")
                
                if success:
                    release_document_lock(st.session_state.get('session_row_index'))
                    st.session_state['active_task'] = None
                    st.session_state['source_file_id'] = None
                    st.session_state['processed_data'] = None
                    st.session_state['review_unlocked'] = False
                    st.balloons()
                    st.success("Task completed successfully!")
                    time.sleep(2)
                    st.rerun()
