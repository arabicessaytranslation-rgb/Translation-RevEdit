import base64
import difflib
from email.message import EmailMessage
import io
import json
import re
import smtplib
import time
import html
import requests
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import pandas as pd
from pydantic import BaseModel, Field
import streamlit as st

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
        textarea { font-family: 'Cairo', 'Inter', sans-serif !important; font-size: 16px !important; line-height: 1.6 !important; border-radius: 8px !important; }
        button[kind="primary"] { background-color: #10B981 !important; border-color: #10B981 !important; color: white !important; font-size: 16px !important; font-weight: 600 !important; border-radius: 8px !important; padding: 0.5rem 1rem !important; transition: all 0.2s ease-in-out !important; }
        button[kind="primary"]:hover { background-color: #059669 !important; border-color: #059669 !important; transform: translateY(-2px); }
        .task-card { background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 20px; border-radius: 10px; margin-bottom: 15px; border-left: 5px solid #3B82F6; }
        div[data-testid="stAlert"]:has(p:contains("CRITICAL STEP")) { background-color: #FEF2F2 !important; border: 1px solid #F87171 !important; color: #991B1B !important; border-radius: 8px !important; }
        .reading-mode { font-family: 'Cairo', sans-serif; font-size: 22px; line-height: 2.2; text-align: justify; direction: rtl; background-color: #ffffff !important; color: #0f172a !important; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border: 1px solid #e0e0e0; }
        .badge-green { background-color: #dcfce7; color: #166534; padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;}
        .badge-yellow { background-color: #fef08a; color: #854d0e; padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;}
        .badge-red { background-color: #fee2e2; color: #991b1b; padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;}
    </style>
    """, unsafe_allow_html=True)

apply_custom_css()

GLOSSARY_SPREADSHEET_ID = "1oc4TCY_iK9R7mBiXgb5rKWssjmrQywYg6UpOBXx8pUQ"
GLOSSARY_RANGE = "'المصطلحات'!A:D"
GLOSSARY_DATA_RANGE = "'المصطلحات'!C:D"
SESSIONS_RANGE = "'Sessions'!A:D"
VOLUNTEERS_RANGE = "'Volunteers'!A:D"
ASSIGNMENTS_RANGE = "'Assignments'!A:I" 

LOCK_TIMEOUT_SECONDS = 14400
ADMIN_REFRESH_COOLDOWN_SECONDS = 600

STATUS_TRANS_ASSIGNED, STATUS_TRANS_STARTED, STATUS_TRANS_COMPLETED = "Translation Assigned", "Translation Started", "Translation Completed"
STATUS_REV_ASSIGNED, STATUS_REV_STARTED, STATUS_REV_COMPLETED = "Reviewer Assigned", "Reviewer Started", "Reviewer Completed"
STATUS_REC_PENDING, STATUS_REC_ASSIGNED, STATUS_REC_STARTED, STATUS_REC_COMPLETED = "Pending Audio Recorder", "Recording Assigned", "Recording Started", "Recording Completed"

ALL_STATUSES = [STATUS_TRANS_ASSIGNED, STATUS_TRANS_STARTED, STATUS_TRANS_COMPLETED, STATUS_REV_ASSIGNED, STATUS_REV_STARTED, STATUS_REV_COMPLETED, STATUS_REC_PENDING, STATUS_REC_ASSIGNED, STATUS_REC_STARTED, STATUS_REC_COMPLETED]

BATCH_SIZE = 6
MAX_RETRIES_PER_MODEL = 3
BASE_BACKOFF_SECONDS = 2.0
MAX_BACKOFF_SECONDS = 15.0
RETRYABLE_KEYWORDS = ("503", "500", "high demand", "429", "timeout", "Quota")

# ==========================================
# 2. GOOGLE SERVICES & AUTHENTICATION
# ==========================================
def get_google_services():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        credentials = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=["https://www.googleapis.com/auth/documents", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
        )
        return (
            build("docs", "v1", credentials=credentials, cache_discovery=False),
            build("drive", "v3", credentials=credentials, cache_discovery=False),
            build("sheets", "v4", credentials=credentials, cache_discovery=False),
        )
    except Exception as e:
        st.error(f"Google Auth Error: {e}")
        st.stop()

docs_service, drive_service, sheets_service = get_google_services()

def fetch_volunteers():
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=VOLUNTEERS_RANGE).execute()
        rows = res.get("values", [])
        volunteers = {}
        if len(rows) > 1:
            for r in rows[1:]:
                if len(r) >= 4:
                    email, name, role, status = r[0].strip().lower(), r[1], r[2].lower(), r[3]
                    volunteers[email] = {"name": name, "role": role, "status": status}
        return volunteers
    except Exception as e:
        st.error(f"⚠️ Temporary error connecting to the Volunteers database. Please wait and refresh.\nDetails: {e}")
        st.stop()

def overwrite_sheet_data(range_name, data_matrix):
    sheets_service.spreadsheets().values().clear(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name).execute()
    sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name, valueInputOption="USER_ENTERED", body={"values": data_matrix}).execute()

@st.cache_data(ttl=86400)
def get_quick_word_count(file_id):
    try:
        doc = docs_service.documents().get(documentId=file_id).execute()
        full_text = ""
        for elem in doc.get('body', {}).get('content', []):
            if 'paragraph' in elem:
                for run in elem.get('paragraph', {}).get('elements', []):
                    if 'textRun' in run:
                        full_text += run.get('textRun', {}).get('content', '')
        count = len(full_text.split())
        return count if count > 0 else 0
    except Exception:
        return 0

# --- ASSIGNMENTS & SLA ENGINE ---
def fetch_assignments():
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=ASSIGNMENTS_RANGE).execute()
        rows = res.get("values", [])
        assignments = []
        if len(rows) > 1:
            for idx, r in enumerate(rows[1:]):
                r.extend([""] * (9 - len(r)))
                assignments.append({
                    "row_index": idx + 2, "doc_id": r[0].strip(), "doc_name": r[1],
                    "translator": r[2].lower().strip(), "reviewer": r[3].lower().strip(), "recorder": r[4].lower().strip(),
                    "status": r[5], "sla_track": r[6] if r[6] else "15", "audio_link": r[7], "stage_start_date": r[8] if r[8] else datetime.now().isoformat()
                })
        return assignments
    except Exception as e:
        st.error(f"⚠️ Temporary error connecting to the Assignments database. Please wait and refresh.\nDetails: {e}")
        st.stop()

def assign_task_to_sheet(doc_id, doc_name, t_email, r_email, rec_email, status, sla_track):
    try:
        assignments = fetch_assignments()
        task = next((row for row in assignments if row.get("doc_id") == doc_id), None)
        start_date, audio_link = datetime.now().isoformat(), ""
        
        if task:
            row_idx, audio_link, start_date = task.get("row_index"), task.get("audio_link", ""), task.get("stage_start_date", start_date)
            body = {"values": [[doc_id, doc_name, t_email, r_email, rec_email, status, sla_track, audio_link, start_date]]}
            sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Assignments'!A{row_idx}:I{row_idx}", valueInputOption="USER_ENTERED", body=body).execute()
        else:
            body = {"values": [[doc_id, doc_name, t_email, r_email, rec_email, status, sla_track, audio_link, start_date]]}
            sheets_service.spreadsheets().values().append(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range="'Assignments'!A:I", valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS", body=body).execute()
        return True
    except Exception as e:
        st.error(f"Google Sheets API Error (Assign): {str(e)}")
        return False

def update_assignment_status(doc_id, new_status, reset_timer=False):
    try:
        assignments = fetch_assignments()
        for task in assignments:
            if task.get("doc_id") == doc_id:
                row_idx = task.get('row_index')
                sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Assignments'!F{row_idx}", valueInputOption="USER_ENTERED", body={"values": [[new_status]]}).execute()
                if reset_timer:
                    sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Assignments'!I{row_idx}", valueInputOption="USER_ENTERED", body={"values": [[datetime.now().isoformat()]]}).execute()
                break
    except Exception as e: st.error(f"Status Update Error: {str(e)}")

def update_assignment_audio_link(doc_id, link):
    try:
        assignments = fetch_assignments()
        for task in assignments:
            if task.get("doc_id") == doc_id:
                sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Assignments'!H{task.get('row_index')}", valueInputOption="USER_ENTERED", body={"values": [[link]]}).execute()
                break
    except Exception as e: st.error(f"Audio Link Update Error: {str(e)}")

def update_assignment_team(doc_id, new_t_email=None, new_r_email=None, new_rec_email=None):
    try:
        assignments = fetch_assignments()
        for task in assignments:
            if task.get("doc_id") == doc_id:
                row_idx = task.get('row_index')
                
                current_t = task.get("translator")
                current_r = task.get("reviewer")
                current_rec = task.get("recorder")
                
                final_t = new_t_email if new_t_email is not None else current_t
                final_r = new_r_email if new_r_email is not None else current_r
                final_rec = new_rec_email if new_rec_email is not None else current_rec

                body = {"values": [[final_t, final_r, final_rec]]}
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=GLOSSARY_SPREADSHEET_ID, 
                    range=f"'Assignments'!C{row_idx}:E{row_idx}", 
                    valueInputOption="USER_ENTERED", 
                    body=body
                ).execute()
                return True
        return False
    except Exception as e:
        st.error(f"Team Update Error: {str(e)}")
        return False

def calculate_sla_status(status, sla_track, start_date_str):
    try: start_dt = datetime.fromisoformat(start_date_str) if start_date_str else datetime.now()
    except: start_dt = datetime.now()

    status = status or ""
    if status in [STATUS_TRANS_COMPLETED, STATUS_REV_COMPLETED, STATUS_REC_COMPLETED]:
        return "-", 0, "<span class='badge-green'>✅ Stage Completed</span>"

    max_days = 5
    if sla_track == "20 Days":
        if "Translation" in status: max_days = 7
        elif "Reviewer" in status: max_days = 7
        elif "Recording" in status or "Pending Audio" in status: max_days = 6
    else:
        if "Translation" in status: max_days = 5
        elif "Reviewer" in status: max_days = 5
        elif "Recording" in status or "Pending Audio" in status: max_days = 5

    due_date = start_dt + timedelta(days=max_days)
    remaining = (due_date - datetime.now()).days

    if remaining >= 2: badge = f"<span class='badge-green'>🟢 {remaining} Days Left</span>"
    elif remaining >= 0: badge = f"<span class='badge-yellow'>🟡 Due Soon ({remaining} Days)</span>"
    else: badge = f"<span class='badge-red'>🔴 Overdue ({-remaining} Days)</span>"

    return due_date.strftime("%d %b %Y"), remaining, badge

# --- LOGIN SCREEN ---
def login_screen():
    if st.session_state.get("authenticated"): return True

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🤝 12-Step AI Suite")
        st.caption("Volunteer Translation, Review & Recording Portal")
        with st.container(border=True):
            with st.form("login_form"):
                email = st.text_input("Enter your registered email address:").strip().lower()
                submitted = st.form_submit_button("Access Portal", type="primary", use_container_width=True)
                
                if submitted:
                    if not email:
                        st.warning("Please enter an email address.")
                    else:
                        volunteers = fetch_volunteers()
                        if email in volunteers:
                            user_data = volunteers[email]
                            status = str(user_data.get("status", "")).strip().lower()
                            
                            if status != "active":
                                st.error("⛔ Account is currently suspended. Please contact the coordinator.")
                            else:
                                raw_role = str(user_data.get("role", "")).strip().lower()
                                mode = "God Mode" if raw_role == "admin" else f"{raw_role.capitalize()} Mode"
                                
                                st.session_state.update({
                                    "authenticated": True,
                                    "user_email": email,
                                    "user_role": raw_role,
                                    "user_name": user_data.get("name", "Fellow"),
                                    "app_mode": mode
                                })
                                st.rerun()
                        else:
                            st.error("Email is not listed in the authorized volunteers list.")
    return False

if not login_screen(): st.stop()

if "processed_data" not in st.session_state: st.session_state["processed_data"] = None
if "source_file_id" not in st.session_state: st.session_state["source_file_id"] = None
if "active_task" not in st.session_state: st.session_state["active_task"] = None
if "admin_last_refresh" not in st.session_state: st.session_state["admin_last_refresh"] = 0

# ==========================================
# 3. ⚡ CENTRAL COMMAND HUB
# ==========================================
if st.session_state.get("app_mode") == "God Mode" and not st.session_state.get("source_file_id"):
    col1, col2 = st.columns([5, 1])
    col1.title("⚡ Central Command Hub")
    if col2.button("🚪 Logout", type="primary"):
        st.session_state.clear()
        st.rerun()

    tab_dash, tab_tools, tab_team, tab_bcast = st.tabs([
        "📁 Edition & Workload Hub", "🤖 Cloud Tools & Transfer", "👥 Team & Glossary", "📢 Broadcast & Comm"
    ])

    vols = fetch_volunteers()
    def format_vol_label(email_key):
        if not email_key or email_key.startswith("["): return email_key
        v_name = vols.get(email_key, {}).get("name")
        return f"{v_name} ({email_key})" if v_name else email_key

    # ==========================================
    # TAB 1: EDITION & WORKLOAD HUB
    # ==========================================
    with tab_dash:
        col_head, col_ref = st.columns([4, 1])
        with col_head:
            st.subheader("🎛️ Edition & Workload Dashboard")
            st.caption("Select an edition to pull articles, view word counts, and balance workloads among volunteers.")
        with col_ref:
            curr_time = time.time()
            time_since_refresh = curr_time - st.session_state["admin_last_refresh"]
            if st.button("🔄 Refresh Data", type="primary", use_container_width=True):
                if time_since_refresh < ADMIN_REFRESH_COOLDOWN_SECONDS:
                    mins_left, secs_left = int((ADMIN_REFRESH_COOLDOWN_SECONDS - time_since_refresh) // 60), int((ADMIN_REFRESH_COOLDOWN_SECONDS - time_since_refresh) % 60)
                    st.warning(f"Cooldown active. Wait {mins_left}m {secs_left}s.")
                else:
                    st.session_state["admin_last_refresh"] = curr_time
                    st.cache_data.clear() 
                    st.rerun()

        try:
            folder_res = drive_service.files().list(q="mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute()
            all_folders = folder_res.get("files", [])
            year_folders = [f for f in all_folders if "Edition" in f["name"] or "202" in f["name"]]

            if year_folders:
                col_yr, col_mo, col_scan = st.columns([2, 2, 1])
                year_options = {f["name"]: f["id"] for f in year_folders}
                sel_year = col_yr.selectbox("📁 1. Year Folder:", list(year_options.keys()))
                
                month_res = drive_service.files().list(q=f"'{year_options[sel_year]}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute()
                month_folders = month_res.get("files", [])

                if month_folders:
                    month_options = {f["name"]: f["id"] for f in month_folders}
                    sel_month = col_mo.selectbox("📂 2. Edition Month:", list(month_options.keys()))
                    
                    col_scan.write("")
                    col_scan.write("")
                    if col_scan.button("🔄 Load Edition", type="primary", use_container_width=True):
                        st.session_state["active_edition_id"] = month_options[sel_month]
                        st.session_state["active_edition_name"] = sel_month
                else: st.warning("No month subfolders found inside the selected year.")
            else: st.warning("No Year folders found matching 'Edition' or '202*'.")

            st.divider()

            if "active_edition_id" in st.session_state:
                doc_query = f"'{st.session_state['active_edition_id']}' in parents and mimeType='application/vnd.google-apps.document' and trashed=false"
                doc_res = drive_service.files().list(q=doc_query, fields="files(id, name)").execute()
                docs_in_drive = doc_res.get("files", [])
                assignments = fetch_assignments()
                assgn_map = {a.get("doc_id"): a for a in assignments}
                
                if docs_in_drive:
                    c_total = len(docs_in_drive)
                    c_unassigned = sum(1 for d in docs_in_drive if d.get("id") not in assgn_map)
                    c_translating = sum(1 for d in docs_in_drive if d.get("id") in assgn_map and "Translation" in assgn_map[d.get("id")].get("status", ""))
                    c_reviewing = sum(1 for d in docs_in_drive if d.get("id") in assgn_map and "Reviewer" in assgn_map[d.get("id")].get("status", ""))
                    c_recording = sum(1 for d in docs_in_drive if d.get("id") in assgn_map and "Recording" in assgn_map[d.get("id")].get("status", "") and "Completed" not in assgn_map[d.get("id")].get("status", ""))
                    c_done = sum(1 for d in docs_in_drive if d.get("id") in assgn_map and assgn_map[d.get("id")].get("status") == STATUS_REC_COMPLETED)
                    
                    st.markdown(f"### 📊 Edition Metrics: `{st.session_state['active_edition_name']}`")
                    m1, m2, m3, m4, m5, m6 = st.columns(6)
                    m1.metric("Total Docs", c_total); m2.metric("⚪ Unassigned", c_unassigned); m3.metric("🟠 Translating", c_translating)
                    m4.metric("🔵 Reviewing", c_reviewing); m5.metric("🎙️ Recording", c_recording); m6.metric("🟢 Finalized", c_done)
                    
                    st.divider()

                    # --- ⚖️ AUTO-BALANCER SECTION ---
                    with st.expander("⚖️ Automated Word-Count Workload Balancer", expanded=False):
                        st.caption("Smart tool to distribute articles evenly among available translators and reviewers to ensure a fair workload.")
                        
                        active_translators = [e for e, d in vols.items() if d.get("role") in ["translator", "admin"] and d.get("status", "").lower() == "active"]
                        active_reviewers = [e for e, d in vols.items() if d.get("role") in ["reviewer", "admin"] and d.get("status", "").lower() == "active"]
                        
                        sel_translators = st.multiselect("Select Translators (Optional):", active_translators, format_func=format_vol_label, key="auto_t_list")
                        sel_reviewers = st.multiselect("Select Reviewers (Mandatory):", active_reviewers, format_func=format_vol_label, key="auto_r_list")
                        
                        if st.button("🚀 Generate Proposed Distribution", type="primary"):
                            if not sel_reviewers:
                                st.error("❌ You must select at least one reviewer to proceed.")
                            else:
                                articles_info = []
                                for doc in docs_in_drive:
                                    w_count = get_quick_word_count(doc.get("id"))
                                    articles_info.append({"id": doc.get("id"), "name": doc.get("name"), "words": w_count if isinstance(w_count, int) else 0})
                                
                                articles_info.sort(key=lambda x: x["words"], reverse=True)
                                
                                rev_loads = {r: 0 for r in sel_reviewers}
                                rev_assignments = {r: [] for r in sel_reviewers}
                                for art in articles_info:
                                    lightest_rev = min(rev_loads, key=rev_loads.get)
                                    rev_loads[lightest_rev] += art["words"]
                                    rev_assignments[lightest_rev].append(art)

                                trans_assignments = {}
                                if sel_translators:
                                    trans_loads = {t: 0 for t in sel_translators}
                                    trans_assignments = {t: [] for t in sel_translators}
                                    for art in articles_info:
                                        lightest_t = min(trans_loads, key=trans_loads.get)
                                        trans_loads[lightest_t] += art["words"]
                                        trans_assignments[lightest_t].append(art)

                                planned_rows = []
                                for art in articles_info:
                                    assigned_r = next((r for r, arts in rev_assignments.items() if art in arts), sel_reviewers[0])
                                    assigned_t = next((t for t, arts in trans_assignments.items() if art in arts), "") if sel_translators else ""
                                    planned_rows.append({
                                        "doc_id": art["id"],
                                        "doc_name": art["name"],
                                        "Word Count": art["words"],
                                        "Proposed Translator": assigned_t,
                                        "Proposed Reviewer": assigned_r
                                    })
                                
                                st.session_state["planned_distribution"] = planned_rows
                                st.success("✅ Fair distribution plan generated successfully. Review the table below and click approve to save to Sheets:")

                        if "planned_distribution" in st.session_state and st.session_state["planned_distribution"]:
                            plan_df = pd.DataFrame(st.session_state["planned_distribution"])
                            display_df = plan_df.copy()
                            display_df["Proposed Translator"] = display_df["Proposed Translator"].apply(lambda x: vols.get(x, {}).get("name", x) if x else "AI Bypass")
                            display_df["Proposed Reviewer"] = display_df["Proposed Reviewer"].apply(lambda x: vols.get(x, {}).get("name", x))
                            
                            st.dataframe(display_df[["doc_name", "Word Count", "Proposed Translator", "Proposed Reviewer"]], use_container_width=True)
                            
                            if st.button("💾 Approve & Save Distribution to Sheets", type="primary"):
                                with st.spinner("Saving distribution to Google Sheets..."):
                                    success_count = 0
                                    for row in st.session_state["planned_distribution"]:
                                        t_email = row["Proposed Translator"]
                                        r_email = row["Proposed Reviewer"]
                                        init_status = STATUS_REV_ASSIGNED if not t_email else STATUS_TRANS_ASSIGNED
                                        if assign_task_to_sheet(row["doc_id"], row["doc_name"], t_email, r_email, "", init_status, "15 Days"):
                                            success_count += 1
                                    
                                    if success_count > 0:
                                        st.balloons()
                                        st.success(f"🎉 Successfully approved and dispatched {success_count} articles to Sheets!")
                                        time.sleep(1.5)
                                        del st.session_state["planned_distribution"]
                                        st.rerun()

                    st.divider()
                    st.markdown("### 📋 Task Delegation & Tracking")
                    
                    UNASSIGN_T = "[Clear Translator]"
                    UNASSIGN_R = "[Clear Reviewer]"
                    UNASSIGN_REC = "[Clear Recorder]"

                    t_options = ["[Optional] AI Bypass", UNASSIGN_T] + [e for e, d in vols.items() if d.get("role") in ["translator", "admin"] and d.get("status", "").lower() == "active"]
                    r_options = ["[Mandatory] Reviewer...", UNASSIGN_R] + [e for e, d in vols.items() if d.get("role") in ["reviewer", "admin"] and d.get("status", "").lower() == "active"]
                    rec_options = ["[Optional] Assign Later", UNASSIGN_REC] + [e for e, d in vols.items() if d.get("role") in ["recorder", "admin"] and d.get("status", "").lower() == "active"]

                    for doc in docs_in_drive:
                        doc_id, doc_name = doc.get("id"), doc.get("name")
                        task = assgn_map.get(doc_id)
                        doc_words = get_quick_word_count(doc_id)
                        
                        with st.container(border=True):
                            if task: 
                                cur_st = task.get("status")
                                t_lbl = format_vol_label(task.get("translator")) if task.get("translator") else "🤖 AI Bypass"
                                r_lbl = format_vol_label(task.get("reviewer")) if task.get("reviewer") else "Unassigned"
                                rec_lbl = format_vol_label(task.get("recorder")) if task.get("recorder") else "Unassigned"
                                due_date, _, badge = calculate_sla_status(cur_st, task.get("sla_track"), task.get("stage_start_date"))
                                
                                c_info, c_badge, c_manage = st.columns([3, 1.5, 1])
                                with c_info:
                                    st.markdown(f"📄 **[{doc_name}](https://docs.google.com/document/d/{doc_id}/edit)** &nbsp; 📝 `{doc_words} Words`")
                                    st.caption(f"**T:** {t_lbl} | **R:** {r_lbl} | **Rec:** {rec_lbl}")
                                    if task.get("audio_link"): st.markdown(f"🎧 [Listen in Telegram / Drive]({task.get('audio_link')})")
                                with c_badge:
                                    st.markdown(f"`{cur_st}`"); st.markdown(badge, unsafe_allow_html=True); st.caption(f"Due: {due_date}")
                                with c_manage:
                                    with st.popover("⚙️ Update", use_container_width=True):
                                        new_s = st.selectbox("Stage:", ALL_STATUSES, index=ALL_STATUSES.index(cur_st) if cur_st in ALL_STATUSES else 0, key=f"s_{doc_id}")
                                        
                                        cur_t = task.get("translator")
                                        t_idx = t_options.index(cur_t) if cur_t in t_options else 0
                                        
                                        cur_r = task.get("reviewer")
                                        r_idx = r_options.index(cur_r) if cur_r in r_options else 0

                                        cur_rec = task.get("recorder")
                                        rec_idx = rec_options.index(cur_rec) if cur_rec in rec_options else 0
                                        
                                        new_t = st.selectbox("Translator:", t_options, index=t_idx, format_func=format_vol_label, key=f"t_up_{doc_id}")
                                        new_r = st.selectbox("Reviewer:", r_options, index=r_idx, format_func=format_vol_label, key=f"r_up_{doc_id}")
                                        new_rec = st.selectbox("Recorder:", rec_options, index=rec_idx, format_func=format_vol_label, key=f"rec_up_{doc_id}")
                                        
                                        if st.button("Save", key=f"b_up_{doc_id}", type="primary", use_container_width=True):
                                            final_t = "" if new_t in ["[Optional] AI Bypass", UNASSIGN_T] else new_t
                                            final_r = "" if new_r in ["[Mandatory] Reviewer...", UNASSIGN_R] else new_r
                                            final_rec = "" if new_rec in ["[Optional] Assign Later", UNASSIGN_REC] else new_rec
                                            
                                            if cur_st == STATUS_REC_PENDING and final_rec != "" and new_s == cur_st:
                                                new_s = STATUS_REC_ASSIGNED

                                            reset_clock = (new_s != cur_st and "Completed" in cur_st) or (new_s != cur_st and "Started" in new_s)
                                            
                                            with st.spinner("Updating assignment..."):
                                                if new_s != cur_st: 
                                                    update_assignment_status(doc_id, new_s, reset_timer=reset_clock)
                                                if final_t != cur_t or final_r != cur_r or final_rec != cur_rec:
                                                    update_assignment_team(doc_id, new_t_email=final_t, new_r_email=final_r, new_rec_email=final_rec)
                                                st.rerun()
                            else: 
                                cn, ct, cr, crec, csla, cb = st.columns([2.5, 1.5, 1.5, 1.5, 1, 1])
                                with cn: st.markdown(f"📄 **[{doc_name}](https://docs.google.com/document/d/{doc_id}/edit)**"); st.caption(f"⚪ *Unassigned* | 📝 {doc_words} Words")
                                
                                initial_t_options = [opt for opt in t_options if opt != UNASSIGN_T]
                                initial_r_options = [opt for opt in r_options if opt != UNASSIGN_R]
                                initial_rec_options = [opt for opt in rec_options if opt != UNASSIGN_REC]

                                with ct: t_sel = st.selectbox("Translator", initial_t_options, format_func=format_vol_label, key=f"t_{doc_id}", label_visibility="collapsed")
                                with cr: r_sel = st.selectbox("Reviewer", initial_r_options, format_func=format_vol_label, key=f"r_{doc_id}", label_visibility="collapsed")
                                with crec: rec_sel = st.selectbox("Recorder", initial_rec_options, format_func=format_vol_label, key=f"rec_{doc_id}", label_visibility="collapsed")
                                with csla: sla_sel = st.selectbox("SLA", ["15 Days", "20 Days"], key=f"sla_{doc_id}", label_visibility="collapsed")
                                with cb:
                                    if st.button("🚀 Assign", key=f"btn_{doc_id}", type="primary", use_container_width=True):
                                        if r_sel == initial_r_options[0]: st.error("Reviewer mandatory!")
                                        else:
                                            final_t = "" if t_sel == initial_t_options[0] else t_sel
                                            final_rec = "" if rec_sel == initial_rec_options[0] else rec_sel
                                            sla_val = "15 Days" if sla_sel == "15 Days" else "20 Days"
                                            init_status = STATUS_REV_ASSIGNED if not final_t else STATUS_TRANS_ASSIGNED
                                            with st.spinner("Assigning task..."):
                                                if assign_task_to_sheet(doc_id, doc_name, final_t, r_sel, final_rec, init_status, sla_val):
                                                    st.success("Assigned Successfully!"); time.sleep(1); st.rerun()
                else: st.info("No documents found in this folder.")
        except Exception as e: st.error(f"Error loading edition workspace: {e}")

        st.divider()
        st.subheader("📩 Bulk Assignment Dispatcher")
        st.caption("Compile current edition assignments and dispatch official emails to all volunteers at once.")
        if "active_edition_id" in st.session_state:
            assignments = fetch_assignments()
            try:
                doc_query = f"'{st.session_state['active_edition_id']}' in parents and mimeType='application/vnd.google-apps.document' and trashed=false"
                doc_res = drive_service.files().list(q=doc_query, fields="files(id)").execute()
                edition_doc_ids = [d["id"] for d in doc_res.get("files", [])]
                edition_tasks = [t for t in assignments if t["doc_id"] in edition_doc_ids]
                
                if edition_tasks:
                    if st.button("🚀 Dispatch Official Assignment Emails", type="primary"):
                        with st.spinner("Preparing and sending emails..."):
                            user_tasks = {}
                            for task in edition_tasks:
                                primary_user = task.get("translator") if task.get("translator") else task.get("reviewer")
                                if not primary_user: continue
                                if primary_user not in user_tasks: user_tasks[primary_user] = []
                                user_tasks[primary_user].append(task)
                            admin_email = st.secrets.get("SMTP_EMAIL", "arabicessaytranslation@gmail.com")
                            for assignee_email, tasks in user_tasks.items():
                                assignee_name = vols.get(assignee_email, {}).get("name", "Dear Fellow")
                                cc_list = {admin_email}
                                rows_html = ""
                                for t in tasks:
                                    if t.get("reviewer") and t.get("reviewer") != assignee_email: cc_list.add(t["reviewer"])
                                    if t.get("recorder") and t.get("recorder") != assignee_email: cc_list.add(t["recorder"])
                                    due_date, _, _ = calculate_sla_status(t.get("status"), t.get("sla_track"), t.get("stage_start_date"))
                                    rows_html += f"<tr style='border-bottom: 1px solid #ddd;'><td style='padding: 8px;'><b>{t.get('doc_name')}</b></td><td style='padding: 8px; color: #b91c1c;'>{due_date}</td><td style='padding: 8px;'>{vols.get(t.get('reviewer'), {}).get('name', t.get('reviewer'))}</td><td style='padding: 8px;'>{vols.get(t.get('recorder'), {}).get('name', 'Unassigned')}</td></tr>"
                                cc_list.discard(assignee_email)
                                
                                html_body = f"""<html dir="ltr"><body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;"><h2>Hello {assignee_name},</h2><p>You have been assigned new tasks for the current edition.</p><h3>📍 To start:</h3><ol><li>Log in to the portal: <a href="https://translation-revedit-2026.streamlit.app/">Translation Portal</a></li><li>Use your registered email to access your queue.</li></ol><h3>📋 Your Tasks:</h3><table style="width: 100%; border-collapse: collapse; text-align: left;"><tr style="background-color: #f3f4f6;"><th style="padding: 8px;">Document</th><th style="padding: 8px;">Deadline</th><th style="padding: 8px;">Reviewer</th><th style="padding: 8px;">Recorder</th></tr>{rows_html}</table><p>Note: The portal is equipped with an automated auditor linked to the <a href="https://docs.google.com/spreadsheets/d/{GLOSSARY_SPREADSHEET_ID}/edit">Official Glossary</a>.</p></body></html>"""
                                try:
                                    msg = EmailMessage(); msg.set_content("Please enable HTML."); msg.add_alternative(html_body, subtype='html')
                                    msg["Subject"] = f"🔔 New Tasks Await You (Edition {st.session_state['active_edition_name']})"; msg["From"] = admin_email; msg["To"] = assignee_email; msg["Cc"] = ", ".join(cc_list)
                                    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                                        server.login(admin_email, st.secrets["SMTP_PASSWORD"]); server.send_message(msg)
                                except Exception as e: st.error(f"Failed to send email to {assignee_email}: {e}")
                            st.balloons(); st.success("Assignment notifications sent successfully to the entire team!")
            except Exception as e: st.error(f"Error dispatching emails: {e}")

    # ==========================================
    # TAB 2: CLOUD TOOLS & TRANSFER
    # ==========================================
    with tab_tools:
        st.subheader("🤖 Automated Edition Prep-Bot")
        with st.container(border=True):
            c_link, c_year, c_month = st.columns([3, 1, 1])
            raw_link = c_link.text_input("🔗 Raw Folder Link (Google Drive):", placeholder="https://drive.google.com/drive/folders/...")
            current_year = datetime.now().year
            sel_year = c_year.selectbox("Year:", [str(y) for y in range(current_year - 1, current_year + 3)], index=1)
            sel_month = c_month.selectbox("Month:", ["1 - January", "2 - February", "3 - March", "4 - April", "5 - May", "6 - June", "7 - July", "8 - August", "9 - September", "10 - October", "11 - November", "12 - December"])

            if st.button("🚀 Trigger Preparation Bot", type="primary", use_container_width=True):
                if not raw_link: st.error("❌ Please enter the folder link first.")
                else:
                    with st.spinner("⏳ Sending command to Backend Bot..."):
                        try:
                            gas_webhook_url = st.secrets.get("GAS_WEBAPP_URL", "")
                            gas_token = st.secrets.get("GAS_SECRET_TOKEN", "")
                            if gas_webhook_url and gas_token:
                                payload = {"url": raw_link, "year": sel_year, "monthNum": int(sel_month.split(" - ")[0]), "chatId": st.secrets.get("TELEGRAM_ADMIN_CHAT_ID", ""), "token": gas_token}
                                response = requests.post(gas_webhook_url, json=payload, timeout=15)
                                if response.status_code == 200:
                                    res_data = response.json()
                                    if res_data.get("status") == "success": st.success(f"✅ {res_data.get('message')}"); st.balloons()
                                    else: st.error(f"⚠️ Bot Error: {res_data.get('message')}")
                                else: st.error("Failed to connect to the cloud server.")
                            else: st.error("⚠️ Please ensure 'GAS_WEBAPP_URL' and 'GAS_SECRET_TOKEN' are set in secrets.")
                        except Exception as e: st.error(f"Error communicating with the bot: {e}")

        st.divider()
        st.subheader("📤 Export Edition Contents")
        st.caption("Copy or move all current edition files to an external folder where you have editor permissions.")
        
        if "active_edition_id" not in st.session_state:
            st.warning("⚠ Please load an edition folder first from the 'Edition & Workload Hub' tab.")
        else:
            st.success(f"📂 Active Folder: {st.session_state['active_edition_name']}")
            with st.container(border=True):
                target_url = st.text_input("🔗 Target Folder URL:", placeholder="https://drive.google.com/drive/folders/...")
                operation_type = st.radio("⚙️ Operation Type:", ["Copy Files (Safe - Keeps Originals)", "Move Files (Pulls from current folder)"])
                is_move = (operation_type == "Move Files (Pulls from current folder)")
                
                if st.button("🚀 Start Cloud Export", type="primary", use_container_width=True):
                    if not target_url: st.error("❌ Please enter the target folder link.")
                    else:
                        folder_match = re.search(r'[-\w]{25,}', target_url)
                        if not folder_match: st.error("❌ Invalid target folder link.")
                        else:
                            target_id = folder_match.group(0)
                            with st.spinner("⏳ Communicating with the server for export..."):
                                try:
                                    gas_webhook_url = st.secrets.get("GAS_WEBAPP_URL", "")
                                    gas_token = st.secrets.get("GAS_SECRET_TOKEN", "")
                                    if gas_webhook_url and gas_token:
                                        payload = {"action": "transfer_contents", "source_folder_id": st.session_state["active_edition_id"], "target_folder_id": target_id, "move_files": is_move, "chatId": st.secrets.get("TELEGRAM_ADMIN_CHAT_ID", ""), "token": gas_token}
                                        response = requests.post(gas_webhook_url, json=payload, timeout=20)
                                        if response.status_code == 200:
                                            res_data = response.json()
                                            if res_data.get("status") == "success": st.balloons(); st.success(f"✅ Operation successful! Exported {res_data.get('count')} files.")
                                            else: st.error(f"⚠️ Server error: {res_data.get('message')}")
                                        else: st.error("Failed to connect to the cloud server.")
                                except Exception as e: st.error(f"Export Error: {e}")

    # ==========================================
    # TAB 3: TEAM & GLOSSARY
    # ==========================================
    with tab_team:
        st.subheader("👥 Manage Volunteer Access")
        try:
            res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=VOLUNTEERS_RANGE).execute()
            vol_data = res.get("values", [])
            if not vol_data: vol_data = [["Email", "Name", "Role", "Status"]]
            df_vol = pd.DataFrame(vol_data[1:], columns=vol_data[0])
            edited_vol = st.data_editor(df_vol, num_rows="dynamic", use_container_width=True, column_config={"Role": st.column_config.SelectboxColumn("Role", options=["translator", "reviewer", "recorder", "admin"], required=True), "Status": st.column_config.SelectboxColumn("Status", options=["Active", "Suspended"], required=True)})
            if st.button("💾 Save Volunteers", type="primary"):
                overwrite_sheet_data(VOLUNTEERS_RANGE, [edited_vol.columns.tolist()] + edited_vol.fillna("").values.tolist())
                st.success("Volunteers database updated!")
        except Exception as e: st.error(f"Error fetching volunteers: {e}")

        st.divider()
        st.subheader("📖 Live Terminology Editor")
        try:
            res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_RANGE).execute()
            glos_data = res.get("values", [])
            headers = ["ID", "Category", "English", "Arabic"]
            rows_to_display = glos_data[1:] if (glos_data and len(glos_data[0]) == 4 and not glos_data[0][0].isdigit()) else glos_data
            cleaned_rows = [list(row) + [""] * (4 - len(row)) for row in rows_to_display]
            if not cleaned_rows: cleaned_rows = [["1", "General", "", ""]]
            edited_glos = st.data_editor(pd.DataFrame(cleaned_rows, columns=headers), num_rows="dynamic", use_container_width=True)
            if st.button("💾 Sync Glossary to AI", type="primary"):
                overwrite_sheet_data(GLOSSARY_RANGE, [edited_glos.columns.tolist()] + edited_glos.fillna("").values.tolist())
                st.success("Glossary synced successfully!"); st.cache_data.clear()
        except Exception as e: st.error(f"Error fetching glossary: {e}")

        st.divider()
        st.subheader("🔥 Glossary Heatmap")
        if st.button("🔄 Generate Heatmap", type="primary"):
            with st.spinner("Scanning sessions and analyzing feedback..."):
                try:
                    res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE).execute()
                    rows, term_counts = res.get("values", []), {}
                    for r in rows[1:]:
                        if len(r) > 3 and r[3]:
                            try:
                                for item in json.loads(r[3]):
                                    matches = re.findall(r'([A-Za-z\s]+)\s*(?:->|:|-)\s*([\u0600-\u06FF\s]+)', item.get("reasoning", "") + " " + item.get("glossary_notes", ""))
                                    for eng, ar in matches:
                                        eng_clean = eng.strip().lower()
                                        if len(eng_clean) > 2 and eng_clean != "none": term_counts[eng_clean] = term_counts.get(eng_clean, 0) + 1
                            except Exception: pass
                    if not term_counts: st.info("Not enough data for analysis currently.")
                    else:
                        df_heat = pd.DataFrame(sorted(term_counts.items(), key=lambda x: x[1], reverse=True)[:15], columns=["English Term", "Auto-Correction Frequency"])
                        st.dataframe(df_heat, use_container_width=True)
                except Exception as e: st.error(f"Heatmap Error: {e}")

    # ==========================================
    # TAB 4: BROADCAST & COMM
    # ==========================================
    with tab_bcast:
        st.subheader("📢 Team Broadcast System")
        broadcast_subject = st.text_input("Subject")
        broadcast_message = st.text_area("Message Body", height=150)
        if st.button("🚀 Send Broadcast", type="primary"):
            if broadcast_subject and broadcast_message:
                active_emails = [email for email, d in vols.items() if d.get("status", "").lower() == "active"]
                if active_emails:
                    with st.spinner("Dispatching emails via secure SMTP..."):
                        try:
                            msg = EmailMessage(); msg.set_content(f"12-Step Translation Project Update:\n\n{broadcast_message}")
                            msg["Subject"] = f"[12-Step Admin] {broadcast_subject}"; msg["From"] = st.secrets.get("SMTP_EMAIL", "admin@localhost"); msg["To"] = st.secrets.get("SMTP_EMAIL", "admin@localhost"); msg["Bcc"] = ", ".join(active_emails)
                            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                                server.login(st.secrets["SMTP_EMAIL"], st.secrets["SMTP_PASSWORD"]); server.send_message(msg)
                            st.success(f"Broadcast sent successfully to {len(active_emails)} volunteers!"); st.balloons()
                        except Exception as e: st.error(f"Failed to send email. Check SMTP secrets. Error: {e}")
                else: st.warning("No active volunteers found to email.")
            else: st.warning("Please enter a subject and a message.")

    st.stop()


# ==========================================
# 4. USER INBOX & TASK DELEGATION
# ==========================================
if not st.session_state.get("source_file_id"):
    col_t, col_l = st.columns([5, 1])
    col_t.title("⚙ 12-Step AI Suite")
    if col_l.button("🚪 Logout", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    st.subheader(f"👋 Welcome, {st.session_state.get('user_name')} | Role: {st.session_state.get('user_role', '').capitalize()}")
    st.markdown("---")

    st.markdown("### 📬 Your Task Queue")
    assignments = fetch_assignments()
    my_tasks = []

    for task in assignments:
        t_status = task.get("status")
        app_mode = st.session_state.get("app_mode")
        usr_email = st.session_state.get("user_email")
        
        if app_mode == "Translator Mode" and task.get("translator") == usr_email and t_status in [STATUS_TRANS_ASSIGNED, STATUS_TRANS_STARTED]:
            my_tasks.append(task)
        elif app_mode == "Reviewer Mode" and task.get("reviewer") == usr_email and t_status in [STATUS_REV_ASSIGNED, STATUS_REV_STARTED]:
            my_tasks.append(task)
        elif app_mode == "Recorder Mode" and task.get("recorder") == usr_email and t_status in [STATUS_REC_ASSIGNED, STATUS_REC_STARTED]:
            my_tasks.append(task)

    if not my_tasks:
        st.success("🎉 You have no pending tasks in your queue. Great job!")
    else:
        for task in my_tasks:
            with st.container(border=True):
                c1, c2 = st.columns([4, 1.2])
                cur_s = task.get("status")
                due_date, _, badge = calculate_sla_status(cur_s, task.get("sla_track"), task.get("stage_start_date"))
                doc_words = get_quick_word_count(task.get('doc_id'))
                
                with c1:
                    st.markdown(f"### 📄 **{task.get('doc_name')}** &nbsp; 📝 `{doc_words} Words`")
                    st.markdown(f"Stage: `{cur_s}` &nbsp;|&nbsp; Deadline: **{due_date}** &nbsp; {badge}", unsafe_allow_html=True)
                with c2:
                    st.write("")
                    btn_text = "🚀 Start Work" if "Assigned" in cur_s else "🔄 Continue Working"
                    if st.button(btn_text, key=f"start_{task.get('doc_id')}", type="primary", use_container_width=True):
                        reset_clock = "Assigned" in cur_s
                        if cur_s == STATUS_TRANS_ASSIGNED:
                            update_assignment_status(task.get("doc_id"), STATUS_TRANS_STARTED, reset_timer=reset_clock)
                            task["status"] = STATUS_TRANS_STARTED
                        elif cur_s == STATUS_REV_ASSIGNED:
                            update_assignment_status(task.get("doc_id"), STATUS_REV_STARTED, reset_timer=reset_clock)
                            task["status"] = STATUS_REV_STARTED
                        elif cur_s == STATUS_REC_ASSIGNED:
                            update_assignment_status(task.get("doc_id"), STATUS_REC_STARTED, reset_timer=reset_clock)
                            task["status"] = STATUS_REC_STARTED

                        st.session_state["active_task"] = task
                        st.session_state["source_file_id"] = task.get("doc_id")
                        st.rerun()
    st.stop()


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
    glossary_notes: str = Field(description="List of exact glossary terms applied with explanations, or 'None'")

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
        except Exception: pass
    return ["gemini-2.5-flash", "gemini-1.5-flash"]

def _is_empty_row(row):
    return len(row) == 0 or all(not str(c).strip() for c in row)

def _get_live_session_row(file_id):
    """Live lookup for the session row to avoid Row-Index Fragility."""
    try:
        res = sheets_service.spreadsheets().values().get(
            spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE
        ).execute()
        for idx, row in enumerate(res.get("values", [])):
            if len(row) > 0 and row[0] == file_id:
                return idx + 1
    except Exception: pass
    return None

def manage_document_lock(file_id: str, user_email: str):
    try:
        res = sheets_service.spreadsheets().values().get(
            spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE
        ).execute()
        rows = res.get("values", [])
        current_time = time.time()
        first_recyclable = None 

        for index, row in enumerate(rows):
            if index == 0: 
                continue

            if len(row) > 0 and row[0] == file_id:
                locked_by = row[1] if len(row) > 1 else ""
                ts = float(row[2]) if len(row) > 2 and row[2] else 0
                has_draft = row[3] if len(row) > 3 else ""

                if locked_by and locked_by != user_email and (current_time - ts) < LOCK_TIMEOUT_SECONDS:
                    return {"status": "blocked", "locked_by": locked_by, "row_index": index + 1}

                if locked_by == user_email and has_draft == "DRIVE_DRAFT":
                    try:
                        draft_name = f".draft_{file_id}_{user_email}.json"
                        res_drive = drive_service.files().list(
                            q=f"name='{draft_name}' and trashed=false", fields="files(id)"
                        ).execute()
                        files = res_drive.get('files', [])
                        if files:
                            request = drive_service.files().get_media(fileId=files[0]['id'])
                            file_bytes = request.execute()
                            return {"status": "recovered",
                                    "data": json.loads(file_bytes.decode('utf-8')),
                                    "row_index": index + 1}
                    except Exception:
                        pass
                return {"status": "clear", "row_index": index + 1}

            if first_recyclable is None and _is_empty_row(row):
                first_recyclable = index + 1

        if first_recyclable is not None:
            return {"status": "clear", "row_index": first_recyclable}
        return {"status": "clear", "row_index": max(len(rows) + 1, 2)}

    except Exception:
        return {"status": "clear", "row_index": 2}

def acquire_document_lock(file_id: str, user_email: str, row_index: int):
    try:
        body = {"values": [[file_id, user_email, str(time.time()), ""]]}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=GLOSSARY_SPREADSHEET_ID, 
            range=f"'Sessions'!A{row_index}:D{row_index}", 
            valueInputOption="USER_ENTERED", 
            body=body
        ).execute()
        
        verify = sheets_service.spreadsheets().values().get(
            spreadsheetId=GLOSSARY_SPREADSHEET_ID,
            range=f"'Sessions'!A{row_index}:B{row_index}"
        ).execute().get("values", [[]])[0]
        
        if not verify or verify[0] != file_id:
            pass 
    except Exception: pass

def save_draft_to_drive(file_id, user_email, processed_data):
    try:
        draft_name = f".draft_{file_id}_{user_email}.json"
        json_data = json.dumps(processed_data, ensure_ascii=False).encode('utf-8')
        media = MediaIoBaseUpload(io.BytesIO(json_data), mimetype='application/json', resumable=True)
        
        res_drive = drive_service.files().list(q=f"name='{draft_name}' and trashed=false", fields="files(id)").execute()
        files = res_drive.get('files', [])
        if files:
            drive_service.files().update(fileId=files[0]['id'], media_body=media).execute()
        else:
            parent_id = st.session_state.get("active_task_parent_folder", "")
            file_metadata = {'name': draft_name, 'parents': [parent_id] if parent_id else []}
            drive_service.files().create(body=file_metadata, media_body=media).execute()

        real_row = _get_live_session_row(file_id)
        if real_row:
            body = {"values": [[file_id, user_email, str(time.time()), "DRIVE_DRAFT"]]}
            sheets_service.spreadsheets().values().update(
                spreadsheetId=GLOSSARY_SPREADSHEET_ID, 
                range=f"'Sessions'!A{real_row}:D{real_row}", 
                valueInputOption="USER_ENTERED", 
                body=body
            ).execute()
        return True
    except Exception: return False

def delete_draft_from_drive(file_id, user_email):
    try:
        draft_name = f".draft_{file_id}_{user_email}.json"
        res_drive = drive_service.files().list(q=f"name='{draft_name}' and trashed=false", fields="files(id)").execute()
        files = res_drive.get('files', [])
        if files:
            drive_service.files().delete(fileId=files[0]['id']).execute()
    except Exception: 
        pass

def release_document_lock(file_id, user_email):
    real_row = _get_live_session_row(file_id)
    if real_row:
        try:
            sheets_service.spreadsheets().values().update(
                spreadsheetId=GLOSSARY_SPREADSHEET_ID, 
                range=f"'Sessions'!A{real_row}:D{real_row}", 
                valueInputOption="USER_ENTERED", 
                body={"values": [["", "", "", ""]]}
            ).execute()
        except Exception: pass
        
    if file_id and user_email:
        delete_draft_from_drive(file_id, user_email)

@st.cache_data(ttl=3600)
def fetch_glossary():
    try:
        res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_DATA_RANGE).execute()
        vals = res.get("values", [])
        glos_lines = []
        glos_dict = {}
        count = 0
        for row in vals:
            if len(row) >= 2 and row[0].strip() and row[1].strip():
                if row[0].strip().lower() == "english": continue
                en = row[0].strip()
                ar = row[1].strip()
                glos_lines.append(f"- {en} -> {ar}")
                glos_dict[en.lower()] = ar
                count += 1
        return "\n".join(glos_lines), count, glos_dict
    except Exception:
        return "", 0, {}

def check_glossary_violations(eng_text, ar_text, glossary_dict):
    violations = []
    if not eng_text or not ar_text or not glossary_dict: return violations
    eng_lower = eng_text.lower()
    for en_term, ar_term in glossary_dict.items():
        pattern = r'\b' + re.escape(en_term) + r'\b'
        if re.search(pattern, eng_lower):
            if ar_term not in ar_text:
                violations.append((en_term, ar_term))
    return violations

def generate_html_diff(original, suggested):
    if not original or original.startswith("[MISSING"): return ("<div dir='rtl' style='text-align: right; color: #0369a1; background-color: #e0f2fe; padding: 10px; border-radius: 5px; font-family: \"Cairo\";'>✨ Initial AI Draft</div>")
    if original.strip() == suggested.strip(): return ("<div dir='rtl' style='text-align: right; color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px; font-family: \"Cairo\";'>✨ Perfect Match — No Edits Needed</div>")
    diff = difflib.ndiff(original.split(), suggested.split())
    html_out = ["<div dir='rtl' style='font-family: \"Cairo\", sans-serif; font-size: 18px; line-height: 2; text-align: right; background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef;'>"]
    for word in diff:
        clean_word = html.escape(word[2:]) 
        if word.startswith("- "): html_out.append(f"<span style='background-color: #ffcdd2; color: #b71c1c; text-decoration: line-through; padding: 2px; border-radius: 4px;'>{clean_word}</span>")
        elif word.startswith("+ "): html_out.append(f"<span style='background-color: #c8e6c9; color: #1b5e20; font-weight: bold; padding: 2px; border-radius: 4px;'>{clean_word}</span>")
        elif word.startswith("  "): html_out.append(f"<span style='color: #212529;'>{clean_word}</span>")
    html_out.append("</div>")
    return " ".join(html_out)

def _is_retryable(err_str): return any(kw.lower() in err_str.lower() for kw in RETRYABLE_KEYWORDS)
def _backoff_sleep(attempt): time.sleep(min(BASE_BACKOFF_SECONDS * (2**attempt), MAX_BACKOFF_SECONDS))

def _call_gemini(model_name, prompt, schema_type):
    response = client.models.generate_content(
        model=model_name, contents=prompt,
        config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=schema_type, safety_settings=safety_settings, temperature=0.2),
    )
    clean_text = response.text.replace("```json", "").replace("```", "").strip()
    match = re.search(r"\{.*\}", clean_text, re.DOTALL)
    if match: clean_text = match.group(0)
    return json.loads(clean_text)

def translate_with_ai(english: str, glossary_text: str):
    prompt = f"""You are a master literary and technical translator between Arabic and English, endowed with deep bilingual erudition and extensive literary appreciation in both tongues. Above all, you possess profound, specialized expertise in the culture, ethos, and literature of 12-Step recovery fellowships.

YOUR IDENTITY & PHILOSOPHY:
- Tone of Fellowship: You convey the core spirit of recovery: humble, compassionate, clinically sound, non-judgmental, and non-moralizing. You write as an experienced fellow speaking to another.
- Literary Eloquence without Affectation: Your Arabic is fluid, resonant, and natural (فصيحة، رصينة، منسابة بلا تقعر). You respect Arabic rhetoric, avoiding robotic literalism.
- Pronoun & Context Awareness: Actively resolve pronouns (e.g., "it", "the program") to reflect the correct grammatical gender and cultural noun in Arabic context.

MANDATORY GLOSSARY ENFORCEMENT:
- You treat the provided recovery glossary as inviolable dogma.
- Whenever an English recovery term matches the glossary, you MUST use the exact Arabic term provided. Never substitute it with synonyms.
- Document every matched term transparently in 'glossary_notes'.

OFFICIAL GLOSSARY:
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
                if _is_retryable(str(e)): _backoff_sleep(attempt)
    return {"arabic_translation": "", "glossary_notes": "⚠ Error"}

def review_with_ai(english: str, arabic: str, glossary_text: str):
    prompt = f"""You are a senior bilingual literary editor and a renowned 12-Step recovery literature specialist. You master the rhetoric of both Arabic and English, with acute sensitivity to fellowship semantics.

YOUR EDITORIAL CRITERIA:
1. Spiritual & Emotional Integrity: Ensure the Arabic translation captures the exact nuance of the original English—neither diluting its psychological gravity nor turning it into moralistic preaching.
2. Narrative Flow & Arabic Idiom: Ensure sentences flow naturally with proper Arabic syntax, punctuation, and rhythm. Eliminate clunky translative traces.
3. Strict Terminology Audit: Verify that recovery concepts strictly adhere to the official glossary. If a term misses the specific fellowship consensus, correct it to the exact glossary equivalent and explain the recovery rationale in 'reasoning'.

OFFICIAL GLOSSARY:
{glossary_text}

Review this pair:
English Source: "{english}"
Original Arabic: "{arabic}"
"""
    for model_name in get_fallback_models():
        for attempt in range(MAX_RETRIES_PER_MODEL):
            try:
                parsed = _call_gemini(model_name, prompt, ReviewResult)
                return {"status": parsed.get("status", "minor_edits"), "suggested_arabic": parsed.get("suggested_arabic", arabic), "reasoning": parsed.get("reasoning", "")}
            except Exception as e:
                if _is_retryable(str(e)): _backoff_sleep(attempt)
    return {"status": "major_rewrite", "suggested_arabic": arabic, "reasoning": "⚠ Error"}

def translate_batch_with_fallback(batch_segments, glossary_text):
    if not GENAI_AVAILABLE or client is None:
        return [translate_with_ai(s.get("english", ""), glossary_text) for s in batch_segments]
    input_payload = "\n\n".join([f"ID: {s.get('id', 0)}\nText: {s.get('english', '')}" for s in batch_segments])

    prompt = f"""You are a master literary and technical translator between Arabic and English, endowed with deep bilingual erudition. Above all, you possess profound, specialized expertise in the culture, ethos, and literature of 12-Step recovery fellowships.

YOUR IDENTITY & PHILOSOPHY:
- Tone of Fellowship: You convey the core spirit of recovery: humble, compassionate, clinically sound, non-judgmental, and non-moralizing. You write as an experienced fellow speaking to another.
- Literary Eloquence without Affectation: Your Arabic is fluid, resonant, and natural (فصيحة، رصينة، منسابة بلا تقعر). You respect Arabic rhetoric, avoiding robotic literalism.
- Pronoun & Context Awareness: Actively resolve pronouns (e.g., "it", "the program") to reflect the correct grammatical gender and cultural noun in Arabic context.

MANDATORY GLOSSARY ENFORCEMENT:
- You treat the provided recovery glossary as inviolable dogma.
- Whenever an English recovery term matches the glossary, you MUST use the exact Arabic term provided. Never substitute it with synonyms.
- In 'glossary_notes', explicitly list every glossary term matched and applied (e.g., "recovery -> تعافي"). If none, write "None".

OFFICIAL GLOSSARY:
{glossary_text}

Segments to Translate:
{input_payload}"""

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
        res = translate_with_ai(s.get("english", ""), glossary_text)
        results.append({"id": s.get("id", 0), "arabic_translation": res.get("arabic_translation", ""), "glossary_notes": res.get("glossary_notes", "")})
    return results

def review_batch_with_fallback(batch_segments, glossary_text):
    if not GENAI_AVAILABLE or client is None:
        return [review_with_ai(s.get("english", ""), s.get("arabic", ""), glossary_text) for s in batch_segments]
    input_payload = "\n\n".join([f"ID: {s.get('id', 0)}\nEnglish: {s.get('english', '')}\nArabic: {s.get('arabic', '')}" for s in batch_segments])

    prompt = f"""You are a senior bilingual literary editor and a renowned 12-Step recovery literature specialist. You master the rhetoric of both Arabic and English, with acute sensitivity to fellowship semantics.

YOUR EDITORIAL CRITERIA FOR THIS BATCH:
1. Spiritual & Emotional Integrity: Ensure the Arabic translation captures the exact nuance of the original English—neither diluting its psychological gravity nor turning it into moralistic preaching.
2. Narrative Flow & Arabic Idiom: Ensure sentences flow naturally with proper Arabic syntax, punctuation, and rhythm. Eliminate clunky translative traces.
3. Strict Terminology Audit: Verify that recovery concepts strictly adhere to the official glossary. If a term misses the specific fellowship consensus, correct it to the exact glossary equivalent and explain the recovery rationale in 'reasoning'.
4. Status Assignment: 'perfect' (no edits needed), 'minor_edits' (small grammar/term corrections), 'major_rewrite' (missed core meaning or severely awkward).

OFFICIAL GLOSSARY:
{glossary_text}

Pairs to Review:
{input_payload}"""

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
        res = review_with_ai(s.get("english", ""), s.get("arabic", ""), glossary_text)
        results.append({"id": s.get("id", 0), "status": res.get("status", "minor_edits"), "suggested_arabic": res.get("suggested_arabic", ""), "reasoning": res.get("reasoning", "")})
    return results

# --- DOCUMENT & DRIVE HANDLERS ---
def _parse_docs_elements(elements):
    paras = []
    for elem in elements:
        if "paragraph" in elem:
            para_text = ""
            start_idx = elem.get("startIndex")
            end_idx = elem.get("endIndex")
            for run in elem.get("paragraph", {}).get("elements", []):
                if "textRun" in run: para_text += run.get("textRun", {}).get("content", "")
            clean_text = para_text.strip()
            if bool(re.search(r"[a-zA-Z\u0600-\u06FF]", clean_text)):
                paras.append({"text": clean_text, "start": start_idx, "end": end_idx})
        elif "table" in elem:
            for row in elem.get("table", {}).get("tableRows", []):
                for cell in row.get("tableCells", []):
                    paras.extend(_parse_docs_elements(cell.get("content", [])))
    return paras

def extract_text_from_drive(file_id, is_retry=False):
    try:
        docs_svc, drive_svc, _ = get_google_services()
        file_meta = drive_svc.files().get(fileId=file_id, fields="mimeType, parents").execute()
        mime_type = file_meta.get("mimeType")
        st.session_state["active_task_parent_folder"] = file_meta.get("parents", [""])[0]

        if mime_type == "application/vnd.google-apps.document":
            try: document = docs_svc.documents().get(documentId=file_id, includeTabsContent=True).execute()
            except Exception: document = docs_svc.documents().get(documentId=file_id).execute()

            st.session_state["doc_revision_id"] = document.get("revisionId")

            all_paras = []
            def sweep_doc_obj(doc_obj):
                temp_paras = []
                temp_paras.extend(_parse_docs_elements(doc_obj.get("body", {}).get("content", [])))
                for footer in doc_obj.get("footers", {}).values(): temp_paras.extend(_parse_docs_elements(footer.get("content", [])))
                for header in doc_obj.get("headers", {}).values(): temp_paras.extend(_parse_docs_elements(header.get("content", [])))
                return temp_paras

            tabs = document.get("tabs", [])
            if tabs:
                for tab in tabs: all_paras.extend(sweep_doc_obj(tab.get("documentTab", {})))
            else:
                all_paras.extend(sweep_doc_obj(document))
            return all_paras
        else:
            st.error(f"Unsupported file type: {mime_type}. Please use Google Docs.")
            return None
    except Exception as e:
        err_str = str(e)
        if ("Broken pipe" in err_str or "Errno 32" in err_str) and not is_retry:
            return extract_text_from_drive(file_id, is_retry=True)
        st.error(f"Could not read document from Drive. Error: {e}")
        return None

def upload_audio_to_drive(uploaded_file, doc_name, parent_folder_id):
    try:
        gas_webhook_url = st.secrets.get("GAS_WEBAPP_URL")
        gas_token = st.secrets.get("GAS_SECRET_TOKEN") 
        if not gas_webhook_url or not gas_token:
            st.error("⚠ Please ensure 'GAS_WEBAPP_URL' and 'GAS_SECRET_TOKEN' are set.")
            return None

        clean_title = re.sub(r'[\\/*?:"<>|]', '', doc_name).strip()
        file_ext = uploaded_file.name.split('.')[-1] if hasattr(uploaded_file, 'name') and uploaded_file.name else 'wav'
        mime_type = uploaded_file.type if hasattr(uploaded_file, 'type') and uploaded_file.type else 'audio/wav'
        
        payload = {
            "action": "upload_audio", "parent_folder_id": parent_folder_id,
            "file_name": f"{clean_title}.{file_ext}", "mime_type": mime_type,
            "file_base64": base64.b64encode(uploaded_file.getvalue()).decode('utf-8'),
            "token": gas_token
        }

        res_data = requests.post(gas_webhook_url, json=payload, timeout=50).json()
        if res_data.get("status") == "success": return res_data.get("webViewLink")
        else: st.error(f"GAS Upload Error: {res_data.get('message')}"); return None
    except requests.exceptions.Timeout:
        st.error("⏳ Server timeout. Please try again.")
        return None
    except Exception as e:
        st.error(f'Upload Error: {e}')
        return None

def smart_align(paragraphs):
    en_paras = [p for p in paragraphs if not re.search(r"[\u0600-\u06FF]", p.get("text", ""))]
    ar_paras = [p for p in paragraphs if re.search(r"[\u0600-\u06FF]", p.get("text", ""))]
    aligned = []
    for i in range(max(len(en_paras), len(ar_paras))):
        en_obj = en_paras[i] if i < len(en_paras) else {"text": "[MISSING ENGLISH SOURCE]"}
        ar_obj = ar_paras[i] if i < len(ar_paras) else {"text": "[MISSING ARABIC TRANSLATION]", "start": None, "end": None}
        aligned.append({"id": i + 1, "english": en_obj.get("text", ""), "arabic": ar_obj.get("text", ""), "ar_start": ar_obj.get("start"), "ar_end": ar_obj.get("end")})
    return aligned

def push_to_drive_translator(document_id, final_arabic_text):
    docs_svc, _, _ = get_google_services()
    try:
        requests = [{'insertPageBreak': {'location': {'index': 1}}}, {'insertText': {'location': {'index': 1}, 'text': final_arabic_text + "\n\n"}}]
        docs_svc.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()
        return True
    except Exception as e:
        st.error(f"Failed to push to Drive. Error: {e}")
        return False

def push_to_drive_reviewer(document_id, approved_segments, revision_id):
    docs_svc, _, _ = get_google_services()
    try:
        requests = []
        positional_segments = [seg for seg in approved_segments if seg.get("ar_start") is not None and seg.get("ar_end") is not None]
        positional_segments.sort(key=lambda x: x["ar_start"], reverse=True)
        
        for seg in positional_segments:
            requests.append({"deleteContentRange": {"range": {"startIndex": seg["ar_start"], "endIndex": seg["ar_end"]}}})
            requests.append({"insertText": {"location": {"index": seg["ar_start"]}, "text": seg.get("final_arabic", "") + "\n"}})
        
        orphaned_segments = [seg for seg in approved_segments if seg.get("ar_start") is None]
        if orphaned_segments:
            orphaned_text = "\n\n--- Added Translations ---\n\n" + "\n\n".join([seg.get("final_arabic", "") for seg in orphaned_segments]) + "\n"
            doc = docs_svc.documents().get(documentId=document_id).execute()
            doc_body_elements = doc.get('body', {}).get('content', [])
            if doc_body_elements:
                last_index = doc_body_elements[-1].get('endIndex', 2) - 1
                requests.append({"insertText": {"location": {"index": last_index}, "text": orphaned_text}})

        body = {"requests": requests}
        if revision_id: body["writeControl"] = {"requiredRevisionId": revision_id}
        
        if requests: docs_svc.documents().batchUpdate(documentId=document_id, body=body).execute()
        return True
    except Exception as e:
        if "requiredRevisionId" in str(e) or "400" in str(e):
            st.error("❌ Failed to save: The original document was modified by another user. Please refresh and try again.")
        else:
            st.error(f"Failed to push to Drive. Error: {e}")
        return False

# ==========================================
# 6. ACTIVE WORKSPACE
# ==========================================
task = st.session_state.get("active_task")
if not task:
    st.session_state["source_file_id"] = None
    st.rerun()

file_id = task.get("doc_id")
app_mode = st.session_state.get("app_mode")

# ---------------------------------------------------------
# RECORDER WORKSPACE
# ---------------------------------------------------------
if app_mode == "Recorder Mode":
    col_h1, col_h2 = st.columns([5, 1])
    col_h1.markdown(f"## 🎙️ Recording Studio: `{task.get('doc_name')}`")
    if col_h2.button("⬅️ Back to Inbox", use_container_width=True):
        st.session_state["active_task"] = None; st.session_state["source_file_id"] = None; st.rerun()

    st.markdown(f"[🔗 Open Original Document in Google Docs](https://docs.google.com/document/d/{file_id}/edit)")
    
    paras = extract_text_from_drive(file_id)
    if paras:
        ar_paras = [p.get("text", "") for p in paras if re.search(r"[\u0600-\u06FF]", p.get("text", ""))]
        st.markdown("### Reading Material (Final Arabic)")
        safe_ar_paras = [html.escape(p) for p in ar_paras]
        st.markdown(f"<div class='reading-mode'>{'<br><br>'.join(safe_ar_paras)}</div>", unsafe_allow_html=True)
        
        st.divider()
        st.subheader("📤 Submit Audio Recording")
        col_rec, col_up = st.columns(2)
        with col_rec: recorded_audio = st.audio_input("🎙️ Record directly from your mic:")
        with col_up: uploaded_audio = st.file_uploader("📂 Or upload a ready audio file (MP3/WAV):", type=["mp3", "wav", "m4a"])
        
        final_audio_file = recorded_audio if recorded_audio else uploaded_audio
        if final_audio_file is not None:
            st.success("✅ Audio Ready for submission!"); st.audio(final_audio_file)
            if st.button("🚀 Upload & Complete Task", type="primary", use_container_width=True):
                with st.spinner("Uploading to Google Drive..."):
                    parent_folder = st.session_state.get("active_task_parent_folder", "")
                    file_link = upload_audio_to_drive(final_audio_file, task.get("doc_name"), parent_folder)
                    if file_link:
                        update_assignment_audio_link(file_id, file_link)
                        update_assignment_status(file_id, STATUS_REC_COMPLETED, reset_timer=False)
                        st.balloons(); st.success("Audio uploaded successfully! Task closed.")
                        time.sleep(2); st.session_state["active_task"] = None; st.session_state["source_file_id"] = None; st.rerun()
    else: st.error("Could not fetch document content.")
    st.stop()

# ---------------------------------------------------------
# TRANSLATOR / REVIEWER WORKSPACE
# ---------------------------------------------------------
glossary_data, glossary_term_count, glossary_dict = fetch_glossary()

col_h1, col_h2 = st.columns([4, 2])
col_h1.markdown(f"## 📝 Workspace: `{task.get('doc_name', 'Document')}`")

with col_h2:
    c_btn, c_glos = st.columns([1, 1])
    if c_btn.button("⬅️ Back to Inbox", use_container_width=True):
        st.session_state["active_task"] = None; st.session_state["source_file_id"] = None; st.session_state["processed_data"] = None; st.rerun()
    with c_glos:
        if glossary_term_count > 0: st.success(f"📖 Glossary: {glossary_term_count} terms")
        else: st.warning("⚠️ Glossary: Not Loaded")

st.markdown(f"[🔗 Open Document in Google Docs](https://docs.google.com/document/d/{file_id}/edit)")
is_bypass_task = task.get("translator") == ""

if st.session_state.get("processed_data") is None:
    lock = manage_document_lock(file_id, st.session_state.get("user_email"))
    if lock["status"] == "blocked":
        st.error(f"🛑 **Document In Use:** This document is currently locked by `{lock.get('locked_by', 'another user')}`."); st.stop()
    elif lock["status"] == "recovered":
        st.session_state["processed_data"] = lock["data"]
        st.success("♻️ **Session Recovered.** Restored your previous work.")
    elif lock["status"] == "clear":
        acquire_document_lock(file_id, st.session_state.get("user_email"), lock["row_index"])
        paras = extract_text_from_drive(file_id)

        if paras:
            processed_results = []
            progress_bar = st.progress(0)

            if app_mode == "Translator Mode":
                st.info("Extracting document content and translating via AI Batching...")
                batches = [paras[i : i + BATCH_SIZE] for i in range(0, len(paras), BATCH_SIZE)]
                for idx, batch in enumerate(batches):
                    batch_payload = [{"id": j + 1, "english": p.get("text", "")} for j, p in enumerate(batch)]
                    ai_results = translate_batch_with_fallback(batch_payload, glossary_data)

                    for ai_res in ai_results:
                        ai_id = int(ai_res.get("id", 1)); array_index = ai_id - 1
                        eng_text = batch[array_index].get("text", "") if 0 <= array_index < len(batch) else batch[0].get("text", "")
                        processed_results.append({
                            "id": ai_id + (idx * BATCH_SIZE), "english": eng_text, "arabic_translation": ai_res.get("arabic_translation", ""),
                            "glossary_notes": ai_res.get("glossary_notes", ""), "user_arabic": ai_res.get("arabic_translation", ""), "is_approved": False,
                        })
                    progress_bar.progress((idx + 1) / len(batches))
            else:
                has_arabic = any(re.search(r"[\u0600-\u06FF]", p.get("text", "")) for p in paras)

                if not has_arabic or is_bypass_task:
                    st.info("🤖 **AI Bypass Mode:** Generating initial draft and analyzing narrative flow for your review...")
                    batches = [paras[i : i + BATCH_SIZE] for i in range(0, len(paras), BATCH_SIZE)]
                    draft_segments = []
                    for idx, batch in enumerate(batches):
                        batch_payload = [{"id": j + 1, "english": p.get("text", "")} for j, p in enumerate(batch)]
                        ai_trans = translate_batch_with_fallback(batch_payload, glossary_data)
                        for ai_t in ai_trans:
                            ai_id = int(ai_t.get("id", 1)); array_index = ai_id - 1
                            eng_text = batch[array_index].get("text", "") if 0 <= array_index < len(batch) else batch[0].get("text", "")
                            draft_segments.append({"id": ai_id + (idx * BATCH_SIZE), "english": eng_text, "arabic": ai_t.get("arabic_translation", ""), "glossary_notes": ai_t.get("glossary_notes", "")})

                    rev_batches = [draft_segments[i : i + BATCH_SIZE] for i in range(0, len(draft_segments), BATCH_SIZE)]
                    for idx, r_batch in enumerate(rev_batches):
                        ai_revs = review_batch_with_fallback(r_batch, glossary_data)
                        for ai_r, d_seg in zip(ai_revs, r_batch):
                            suggested, status_val = ai_r.get("suggested_arabic", d_seg.get("arabic", "")), ai_r.get("status", "minor_edits")
                            processed_results.append({
                                "id": d_seg.get("id"), "status": status_val, "english": d_seg.get("english", ""), "original_arabic": d_seg.get("arabic", ""),
                                "suggested_arabic": suggested, "reasoning": (ai_r.get("reasoning", "") + (f" | Glossary: {d_seg.get('glossary_notes')}" if d_seg.get("glossary_notes") else "")),
                                "ar_start": None, "ar_end": None, "user_arabic": suggested, "is_approved": (status_val == "perfect"),
                            })
                        progress_bar.progress((idx + 1) / len(rev_batches))
                else:
                    st.info("Extracting segments and comparing human translation against AI audit...")
                    segments = smart_align(paras)
                    normal_segs = [s for s in segments if s.get("english") != "[MISSING ENGLISH SOURCE]" and s.get("arabic") != "[MISSING ARABIC TRANSLATION]"]
                    batches = [normal_segs[i : i + BATCH_SIZE] for i in range(0, len(normal_segs), BATCH_SIZE)]

                    for idx, batch in enumerate(batches):
                        ai_results = review_batch_with_fallback(batch, glossary_data)
                        for ai_res, o_seg in zip(ai_results, batch):
                            suggested, status_val = ai_res.get("suggested_arabic", o_seg.get("arabic", "")), ai_res.get("status", "minor_edits")
                            processed_results.append({
                                "id": o_seg.get("id"), "status": status_val, "english": o_seg.get("english", ""), "original_arabic": o_seg.get("arabic", ""),
                                "suggested_arabic": suggested, "reasoning": ai_res.get("reasoning", ""), "ar_start": o_seg.get("ar_start"), "ar_end": o_seg.get("ar_end"),
                                "user_arabic": suggested, "is_approved": (status_val == "perfect"),
                            })
                        progress_bar.progress((idx + 1) / len(batches))

                    for item in segments:
                        if item.get("english") == "[MISSING ENGLISH SOURCE]":
                            processed_results.append({"id": item.get("id"), "status": "major_rewrite", "english": "[MISSING]", "original_arabic": item.get("arabic", ""), "suggested_arabic": item.get("arabic", ""), "reasoning": "⚠️ Orphaned Arabic block.", "ar_start": item.get("ar_start"), "ar_end": item.get("ar_end"), "user_arabic": item.get("arabic"), "is_approved": False })
                        elif item.get("arabic") == "[MISSING ARABIC TRANSLATION]":
                            trans_res = translate_with_ai(item.get("english", ""), glossary_data)
                            t_arabic = trans_res.get("arabic_translation", "")
                            processed_results.append({"id": item.get("id"), "status": "major_rewrite", "english": item.get("english", ""), "original_arabic": "[MISSING]", "suggested_arabic": t_arabic, "reasoning": "⚠️ Auto-translated orphaned English block.", "ar_start": None, "ar_end": None, "user_arabic": t_arabic, "is_approved": False })
                    processed_results.sort(key=lambda x: x.get("id", 0))

            if not processed_results:
                processed_results = []
                
            if not save_draft_to_drive(file_id, st.session_state.get("user_email"), processed_results):
                st.error("⚠️ Autosave failed. Your work is only saved in this screen — please keep the page open and notify the coordinator to prevent data loss.")
            else:
                st.session_state["processed_data"] = processed_results
                st.rerun()

# --- EDITOR UI ---
approved_count, finalized_data, pending_autosave = 0, [], False
st.divider()

if st.button("💾 Save Draft Progress", type="secondary", use_container_width=True):
    if save_draft_to_drive(file_id, st.session_state.get("user_email"), st.session_state.get("processed_data")):
        st.toast("✅ Progress saved to Drive successfully!", icon="💾")
    else:
        st.error("⚠️ Failed to save progress to Drive.")

for i, item in enumerate(st.session_state.get("processed_data", [])):
    seg_id, status_val = item.get("id", i + 1), item.get("status", "minor_edits")
    eng_txt, orig_ar, sugg_ar = item.get("english", ""), item.get("original_arabic", ""), item.get("suggested_arabic", item.get("arabic_translation", ""))

    with st.container(border=True):
        if app_mode == "Translator Mode":
            st.markdown(f"### Segment {seg_id}")
            col_en, col_ar = st.columns(2)
            with col_en:
                st.info(html.escape(eng_txt))
                if item.get("glossary_notes"): st.caption(f"💡 **Glossary Matched:** {html.escape(item.get('glossary_notes'))}")
            with col_ar:
                default_val = item.get("user_arabic", item.get("arabic_translation", ""))
                final_text = st.text_area("Final text", value=default_val, height=120, key=f"edit_{i}", label_visibility="collapsed")
                item["user_arabic"] = final_text
                
                violations = check_glossary_violations(eng_txt, final_text, glossary_dict)
                if violations:
                    st.error("⚠️ **Glossary Violation:** " + " | ".join([f"`{html.escape(en)}` ⟵ `{html.escape(ar)}`" for en, ar in violations]))
        else:
            color = "🟢" if status_val == "perfect" else ("🟡" if status_val == "minor_edits" else "🔴")
            st.markdown(f"### Segment {seg_id} | Status: {color} {status_val.upper()}")
            col_en, col_ar = st.columns(2)
            with col_en: st.info(html.escape(eng_txt))
            with col_ar:
                st.markdown(generate_html_diff(orig_ar, sugg_ar), unsafe_allow_html=True)
                with st.expander("💡 View AI Reasoning & Glossary Audit"): st.markdown(html.escape(item.get("reasoning", item.get("glossary_notes", "No reasoning provided."))))
                default_val = item.get("user_arabic", sugg_ar)
                
                final_text = st.text_area("Final Output", value=default_val, height=120, key=f"edit_ar_{i}", label_visibility="collapsed")
                item["user_arabic"] = final_text 

                violations = check_glossary_violations(eng_txt, final_text, glossary_dict)
                if violations:
                    st.error("⚠️ **Glossary Violation:** " + " | ".join([f"`{html.escape(en)}` ⟵ `{html.escape(ar)}`" for en, ar in violations]))

        chk = st.checkbox(f"✅ Approve Segment {seg_id}", key=f"chk_{i}", value=item.get("is_approved", False))
        if chk != item.get("is_approved"): 
            item["is_approved"] = chk
            pending_autosave = True 
            
        if chk:
            approved_count += 1
            if app_mode == "Translator Mode": finalized_data.append(final_text)
            else: finalized_data.append({"final_arabic": final_text, "ar_start": item.get("ar_start"), "ar_end": item.get("ar_end")})

if pending_autosave:
    if save_draft_to_drive(file_id, st.session_state.get("user_email"), st.session_state.get("processed_data")):
        st.toast("✅ Edits auto-saved", icon="💾")

# --- SUBMISSION LOGIC ---
st.divider()
total_segments = len(st.session_state.get("processed_data", []))
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
        if st.button("🚀 Push to Drive & Conclude Stage", type="primary", use_container_width=True):
            with st.spinner("Processing Drive updates and concluding workflow..."):
                if app_mode == "Translator Mode":
                    ar_compiled = "\n\n".join([str(item) for item in finalized_data])
                    success = push_to_drive_translator(file_id, ar_compiled)
                    if success:
                        r_email = task.get("reviewer", "")
                        if task.get("translator") == r_email and r_email != "": new_status = STATUS_REV_COMPLETED
                        elif r_email != "": new_status = STATUS_REV_ASSIGNED
                        else: new_status = STATUS_TRANS_COMPLETED
                        update_assignment_status(file_id, new_status, reset_timer=True)
                else:
                    if is_bypass_task:
                        ar_compiled = "\n\n".join([item.get("final_arabic", "") if isinstance(item, dict) else str(item) for item in finalized_data])
                        success = push_to_drive_translator(file_id, ar_compiled)
                    else:
                        success = push_to_drive_reviewer(file_id, finalized_data, st.session_state.get("doc_revision_id"))

                    if success:
                        rec_email = task.get("recorder", "")
                        new_status = STATUS_REC_ASSIGNED if rec_email else STATUS_REC_PENDING
                        update_assignment_status(file_id, new_status, reset_timer=True)

                if success:
                    release_document_lock(file_id, st.session_state.get("user_email"))
                    st.session_state["active_task"] = None; st.session_state["source_file_id"] = None
                    st.session_state["processed_data"] = None; st.session_state["review_unlocked"] = False
                    st.balloons(); st.success("Task stage completed successfully!")
                    time.sleep(2); st.rerun()
