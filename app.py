import base64
import difflib
from email.message import EmailMessage
import io
import json
import re
import smtplib
import time
import html
import random
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
RETRYABLE_KEYWORDS = ("503", "500", "high demand", "429", "timeout", "quota", "exhausted")

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

    st.subheader(f"👋 Welcome, {st.session_state.get('user_name')} | Role
