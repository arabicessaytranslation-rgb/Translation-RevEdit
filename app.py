import difflib
from email.message import EmailMessage
import io
import json
import random
import re
import smtplib
import time
import requests
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd
from pydantic import BaseModel, Field
import streamlit as st
import base64
import gc

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
st.set_page_config(
    page_title="Arabic Essay Service Coordinator Hub",
    layout="wide",
    initial_sidebar_state="collapsed",
)

def apply_custom_css():
  st.markdown(
      """
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
        
        .reading-mode {
            font-family: 'Cairo', sans-serif;
            font-size: 22px;
            line-height: 2.2;
            text-align: justify;
            direction: rtl;
            background-color: #ffffff !important;
            color: #0f172a !important;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            border: 1px solid #e0e0e0;
        }
    </style>
    """,
      unsafe_allow_html=True,
  )

apply_custom_css()

GLOSSARY_SPREADSHEET_ID = "1oc4TCY_iK9R7mBiXgb5rKWssjmrQywYg6UpOBXx8pUQ"
GLOSSARY_RANGE = "'المصطلحات'!A:D"
GLOSSARY_DATA_RANGE = "'المصطلحات'!C:D"
SESSIONS_RANGE = "'Sessions'!A:D"
VOLUNTEERS_RANGE = "'Volunteers'!A:D"
ASSIGNMENTS_RANGE = "'Assignments'!A:K" 

LOCK_TIMEOUT_SECONDS = 14400
ADMIN_REFRESH_COOLDOWN_SECONDS = 600

STATUS_TRANS_ASSIGNED = "Translation Assigned"
STATUS_TRANS_STARTED = "Translation Started"
STATUS_TRANS_COMPLETED = "Translation Completed"
STATUS_REV_ASSIGNED = "Reviewer Assigned"
STATUS_REV_STARTED = "Reviewer Started"
STATUS_REV_COMPLETED = "Reviewer Completed"
STATUS_REC_PENDING = "Pending Audio Recorder"
STATUS_REC_ASSIGNED = "Recording Assigned"
STATUS_REC_STARTED = "Recording Started"
STATUS_REC_COMPLETED = "Recording Completed"
STATUS_SUBMITTED = "Submitted to Client"

ALL_STATUSES = [
    STATUS_TRANS_ASSIGNED, STATUS_TRANS_STARTED, STATUS_TRANS_COMPLETED,
    STATUS_REV_ASSIGNED, STATUS_REV_STARTED, STATUS_REV_COMPLETED,
    STATUS_REC_PENDING, STATUS_REC_ASSIGNED, STATUS_REC_STARTED, 
    STATUS_REC_COMPLETED, STATUS_SUBMITTED
]

BATCH_SIZE = 6
MAX_RETRIES_PER_MODEL = 3
BASE_BACKOFF_SECONDS = 2.0
MAX_BACKOFF_SECONDS = 15.0
RETRYABLE_KEYWORDS = ("503", "500", "high demand", "429", "timeout", "Quota")

# ==========================================
# 2. GOOGLE SERVICES & SYSTEM EMAILS
# ==========================================
def get_google_services():
  try:
    creds_dict = dict(st.secrets["gcp_service_account"])
    credentials = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=[
            "https://www.googleapis.com/auth/documents",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets",
        ],
    )
    return (
        build("docs", "v1", credentials=credentials, cache_discovery=False),
        build("drive", "v3", credentials=credentials, cache_discovery=False),
        build("sheets", "v4", credentials=credentials, cache_discovery=False),
    )
  except Exception as e:
    st.error(f"Google Auth Error: {e}")
    return None, None, None

docs_service, drive_service, sheets_service = get_google_services()

def send_system_email(to_email, subject, body):
    if not to_email: return
    try:
        msg = EmailMessage()
        msg.set_content(body)
        msg["Subject"] = subject
        msg["From"] = st.secrets.get("SMTP_EMAIL", "admin@localhost")
        msg["To"] = to_email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(st.secrets.get("SMTP_EMAIL"), st.secrets.get("SMTP_PASSWORD"))
            server.send_message(msg)
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

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
  except Exception:
    return {}

def overwrite_sheet_data(range_name, data_matrix):
  sheets_service.spreadsheets().values().clear(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name).execute()
  sheets_service.spreadsheets().values().update(
      spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name,
      valueInputOption="USER_ENTERED", body={"values": data_matrix}
  ).execute()

def fetch_assignments():
  try:
    res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=ASSIGNMENTS_RANGE).execute()
    rows = res.get("values", [])
    assignments = []
    if len(rows) > 1:
      for idx, r in enumerate(rows[1:]):
        r.extend([""] * (11 - len(r)))
        assignments.append({
            "row_index": idx + 2,
            "doc_id": r[0].strip(),
            "doc_name": r[1],
            "translator": r[2].lower().strip(),
            "reviewer": r[3].lower().strip(),
            "recorder": r[4].lower().strip(),
            "status": r[5],
            "final_due": r[6],
            "audio_link": r[7],
            "t_due": r[8],
            "r_due": r[9],
            "rec_due": r[10]
        })
    return assignments
  except Exception:
    return []

def assign_task_to_sheet(doc_id, doc_name, t_email, r_email, rec_email, status, final_due, t_due="", r_due="", rec_due=""):
  assignments = fetch_assignments()
  row_idx = None
  for row in assignments:
    if row.get("doc_id") == doc_id:
      row_idx = row.get("row_index")
      break

  body = {"values": [[doc_id, doc_name, t_email, r_email, rec_email, status, final_due, "", t_due, r_due, rec_due]]}
  if row_idx:
    range_name = f"'Assignments'!A{row_idx}:K{row_idx}"
    sheets_service.spreadsheets().values().update(
        spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name,
        valueInputOption="USER_ENTERED", body=body,
    ).execute()
  else:
    range_name = "'Assignments'!A:K"
    sheets_service.spreadsheets().values().append(
        spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name,
        valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS", body=body,
    ).execute()

def update_assignment_status(doc_id, new_status):
  assignments = fetch_assignments()
  for task in assignments:
    if task.get("doc_id") == doc_id:
      range_name = f"'Assignments'!F{task.get('row_index')}"
      sheets_service.spreadsheets().values().update(
          spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name,
          valueInputOption="USER_ENTERED", body={"values": [[new_status]]},
      ).execute()
      break

def update_assignment_roles(doc_id, t_email, r_email, rec_email):
  assignments = fetch_assignments()
  for task in assignments:
    if task.get("doc_id") == doc_id:
      range_name = f"'Assignments'!C{task.get('row_index')}:E{task.get('row_index')}"
      sheets_service.spreadsheets().values().update(
          spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name,
          valueInputOption="USER_ENTERED", body={"values": [[t_email, r_email, rec_email]]},
      ).execute()
      break

def update_assignment_audio_link(doc_id, link):
  assignments = fetch_assignments()
  for task in assignments:
    if task.get("doc_id") == doc_id:
      range_name = f"'Assignments'!H{task.get('row_index')}"
      sheets_service.spreadsheets().values().update(
          spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name,
          valueInputOption="USER_ENTERED", body={"values": [[link]]},
      ).execute()
      break

def extract_id_from_url(url):
    match_folder = re.search(r'folders/([a-zA-Z0-9-_]+)', url)
    if match_folder: return match_folder.group(1)
    match_file = re.search(r'd/([a-zA-Z0-9-_]+)', url)
    if match_file: return match_file.group(1)
    match_id = re.search(r'id=([a-zA-Z0-9-_]+)', url)
    if match_id: return match_id.group(1)
    return None

# --- LOGIN SCREEN ---
def login_screen():
  if st.session_state.get("authenticated"):
    return True
  col1, col2, col3 = st.columns([1, 2, 1])
  with col2:
    st.title("🤝 12-Step AI Suite")
    st.caption("Volunteer Translation, Review & Recording Portal")
    with st.container(border=True):
      with st.form("login_form"):
        email = st.text_input("Enter your registered email address:").strip().lower()
        password = st.text_input("Admin Password (Leave blank if you are a volunteer):", type="password")
        
        if st.form_submit_button("Access Portal", type="primary", width="stretch"):
          admin_email = st.secrets.get("ADMIN_EMAIL", "").strip().lower()
          admin_password = st.secrets.get("ADMIN_PASSWORD", "")
          
          if admin_email and email == admin_email:
              if password == admin_password:
                  st.session_state.update({
                      "authenticated": True, "user_email": email, "user_role": "admin",
                      "user_name": "System Admin", "app_mode": "Coordinator Hub"
                  })
                  st.rerun()
              else:
                  st.error("Incorrect Admin Password.")
          else:
              volunteers = fetch_volunteers()
              if email in volunteers:
                user_data = volunteers[email]
                if user_data.get("status", "").lower() != "active":
                  st.error("Account suspended. Please contact the coordinator.")
                else:
                  st.session_state.update({
                      "authenticated": True, "user_email": email, "user_role": user_data.get("role"),
                      "user_name": user_data.get("name"), "app_mode": "Volunteer Dashboard"
                  })
                  st.rerun()
              else:
                st.error("Email not recognized in the system.")
  return False

if not login_screen():
  st.stop()

if "processed_data" not in st.session_state: st.session_state["processed_data"] = None
if "source_file_id" not in st.session_state: st.session_state["source_file_id"] = None
if "session_row_index" not in st.session_state: st.session_state["session_row_index"] = None
if "active_task" not in st.session_state: st.session_state["active_task"] = None
if "admin_last_refresh" not in st.session_state: st.session_state["admin_last_refresh"] = 0

# ==========================================
# 3. COORDINATOR HUB (ADMIN DASHBOARD)
# ==========================================
if st.session_state.get("app_mode") == "Coordinator Hub" and not st.session_state.get("source_file_id"):
  col1, col2 = st.columns([5, 1])
  col1.title("⚡ Arabic Essay Service Coordinator Hub")
  if col2.button("🚪 Logout", type="primary"):
    st.session_state.clear()
    st.rerun()

  tab_pipe, tab_team, tab_sys = st.tabs(["🚀 The Pipeline", "👥 Team Operations", "⚙️ System Tools"])
  vols = fetch_volunteers()

  def format_vol_label(email_key):
    if not email_key or email_key.startswith("["): return email_key
    v_info = vols.get(email_key, {})
    return f"{v_info.get('name', email_key)} ({email_key})"

  def get_status_badge(st_str):
    badges = {
        STATUS_TRANS_ASSIGNED: "🟠", STATUS_TRANS_STARTED: "🟡", STATUS_TRANS_COMPLETED: "📑",
        STATUS_REV_ASSIGNED: "🔵", STATUS_REV_STARTED: "🟣", STATUS_REV_COMPLETED: "📑",
        STATUS_REC_PENDING: "⚠️", STATUS_REC_ASSIGNED: "🎧", STATUS_REC_STARTED: "🎙️",
        STATUS_REC_COMPLETED: "🟢", STATUS_SUBMITTED: "✅"
    }
    return f"{badges.get(st_str, '⚪')} {st_str}"

  # ---------------------------------------------------------
  # TAB 1: THE PIPELINE
  # ---------------------------------------------------------
  with tab_pipe:
    with st.expander("➕ Expand to Auto-Dispatch New Edition"):
      t_active = [e for e, d in vols.items() if "translator" in d.get("role") and d.get("status", "").lower() == "active"]
      r_active = [e for e, d in vols.items() if "reviewer" in d.get("role") and d.get("status", "").lower() == "active"]
      rec_active = [e for e, d in vols.items() if "recorder" in d.get("role") and d.get("status", "").lower() == "active"]

      try:
        folder_res = drive_service.files().list(q="mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute()
        year_folders = [f for f in folder_res.get("files", []) if "Edition" in f["name"] or "202" in f["name"]]

        if year_folders:
          c1, c2 = st.columns(2)
          year_options = {f["name"]: f["id"] for f in year_folders}
          sel_year = c1.selectbox("📁 1. Target Year:", list(year_options.keys()))
          year_id = year_options[sel_year]

          month_res = drive_service.files().list(q=f"'{year_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute()
          month_folders = month_res.get("files", [])

          if month_folders:
            month_options = {f["name"]: f["id"] for f in month_folders}
            sel_month = c2.selectbox("📂 2. Target Month:", list(month_options.keys()))
            month_id = month_options[sel_month]

            st.divider()
            c3, c4 = st.columns(2)
            sub_date = c3.date_input("📅 3. Select Final Submission Date:", value=datetime.today() + timedelta(days=20))
            
            with c4:
                st.write("**4. Select Volunteer Pool**")
                pool_t = st.multiselect("Translators:", t_active, format_func=format_vol_label)
                pool_r = st.multiselect("Reviewers (Mandatory):", r_active, format_func=format_vol_label)
                pool_rec = st.multiselect("Recorders:", rec_active, format_func=format_vol_label)

            if st.button("🔄 Analyze Word Counts & Auto-Assign", type="primary", width="stretch"):
              if not pool_r:
                st.error("❌ At least one Reviewer must be selected.")
              else:
                with st.spinner("Scanning for new documents and balancing workloads..."):
                  doc_res = drive_service.files().list(q=f"'{month_id}' in parents and mimeType='application/vnd.google-apps.document' and trashed=false", fields="files(id, name)").execute()
                  docs = doc_res.get("files", [])
                  
                  existing_tasks = fetch_assignments()
                  existing_ids = {t.get("doc_id") for t in existing_tasks}
                  
                  new_docs = [d for d in docs if d['id'] not in existing_ids]
                  
                  if not docs:
                      st.warning("No Google Docs found in this folder.")
                  elif not new_docs:
                      st.success("✅ All documents in this folder are already assigned in the pipeline. No new tasks to add.")
                  else:
                      doc_stats = []
                      for d in new_docs:
                          try:
                              document = docs_service.documents().get(documentId=d['id']).execute()
                              content = document.get('body', {}).get('content', [])
                              text_str = ""
                              for elem in content:
                                  if 'paragraph' in elem:
                                      for run in elem['paragraph'].get('elements', []):
                                          if 'textRun' in run: text_str += run['textRun'].get('content', '')
                              word_count = len(text_str.split())
                              doc_stats.append({'id': d['id'], 'name': d['name'], 'wc': word_count})
                              
                              del document
                              gc.collect()
                          except:
                              doc_stats.append({'id': d['id'], 'name': d['name'], 'wc': 500})
                      
                      doc_stats.sort(key=lambda x: x['wc'], reverse=True)

                      today = datetime.today().date()
                      total_days = (sub_date - today).days
                      if total_days < 4: total_days = 4
                      
                      t_days = max(1, int(total_days * 0.35))
                      r_days = max(1, int(total_days * 0.35))
                      rec_days = max(1, int(total_days * 0.20))
                      
                      t_due = today + timedelta(days=t_days)
                      r_due = t_due + timedelta(days=r_days)
                      rec_due = r_due + timedelta(days=rec_days)

                      s_t_due = t_due.strftime("%b %d")
                      s_r_due = r_due.strftime("%b %d")
                      s_rec_due = rec_due.strftime("%b %d")
                      s_final = sub_date.strftime("%b %d, %Y")

                      loads_t = {email: 0 for email in pool_t}
                      loads_r = {email: 0 for email in pool_r}
                      loads_rec = {email: 0 for email in pool_rec}

                      for ds in doc_stats:
                          a_t = min(loads_t, key=loads_t.get) if loads_t else ""
                          if a_t: loads_t[a_t] += ds['wc']
                          
                          a_r = min(loads_r, key=loads_r.get)
                          loads_r[a_r] += ds['wc']
                          
                          a_rec = min(loads_rec, key=loads_rec.get) if loads_rec else ""
                          if a_rec: loads_rec[a_rec] += ds['wc']

                          init_status = STATUS_TRANS_ASSIGNED if a_t else STATUS_REV_ASSIGNED
                          assign_task_to_sheet(ds['id'], ds['name'], a_t, a_r, a_rec, init_status, s_final, s_t_due, s_r_due, s_rec_due)
                      
                      st.success(f"Successfully integrated and assigned {len(doc_stats)} NEW articles to the active pipeline!")
                      time.sleep(1.5)
                      st.rerun()
          else:
            st.warning("No month subfolders found inside the selected year.")
        else:
          st.warning("No Year folders found.")
      except Exception as e:
        st.error(f"Google Drive Error: {e}")

    st.subheader("📊 Active Task Tracker")
    curr_time = time.time()
    time_since_refresh = curr_time - st.session_state["admin_last_refresh"]
    if st.button("🔄 Refresh Pipeline Data", type="primary"):
        if time_since_refresh < ADMIN_REFRESH_COOLDOWN_SECONDS:
            mins_left = int((ADMIN_REFRESH_COOLDOWN_SECONDS - time_since_refresh) // 60)
            secs_left = int((ADMIN_REFRESH_COOLDOWN_SECONDS - time_since_refresh) % 60)
            st.warning(f"Cooldown active. Please wait {mins_left}m {secs_left}s.")
        else:
            st.session_state["admin_last_refresh"] = curr_time
            st.cache_data.clear()
            st.rerun()

    live_tasks = fetch_assignments()
    active_tasks = [t for t in live_tasks if t.get("status") != STATUS_SUBMITTED]
    ready_tasks = [t for t in live_tasks if t.get("status") == STATUS_REC_COMPLETED]

    if not active_tasks:
      st.info("No active assignments in the pipeline.")
    else:
      m1, m2, m3, m4, m5 = st.columns(5)
      m1.metric("Total Active", len(active_tasks))
      m2.metric("🟠 Translating", sum(1 for t in active_tasks if "Translation" in t.get("status", "")))
      m3.metric("🔵 Reviewing", sum(1 for t in active_tasks if "Reviewer" in t.get("status", "")))
      m4.metric("🎙️ Recording", sum(1 for t in active_tasks if "Recording" in t.get("status", "") and "Completed" not in t.get("status", "")))
      m5.metric("🟢 Ready for Deploy", len(ready_tasks))

      cs1, cs2 = st.columns([3, 1])
      srch = cs1.text_input("🔍 Search active pipeline:", "").strip().lower()
      filt = cs2.selectbox("Filter Stage:", ["All Active"] + ALL_STATUSES[:-1])

      for t in active_tasks:
        if (filt != "All Active" and t.get("status") != filt) or (srch and srch not in str(t).lower()):
            continue
        
        d_id = t.get("doc_id")
        t_disp = format_vol_label(t.get("translator")) if t.get("translator") else "🤖 AI Bypass"
        r_disp = format_vol_label(t.get("reviewer"))
        rec_disp = format_vol_label(t.get("recorder")) if t.get("recorder") else "Unassigned"
        
        dl_t = f" (Due {t.get('t_due')})" if t.get('t_due') else ""
        dl_r = f" (Due {t.get('r_due')})" if t.get('r_due') else ""
        dl_rec = f" (Due {t.get('rec_due')})" if t.get('rec_due') else ""

        with st.container(border=True):
          ci, ca = st.columns([4, 1.3])
          with ci:
            st.markdown(f"**📄 [{t.get('doc_name')}](https://docs.google.com/document/d/{d_id}/edit)** &nbsp;&nbsp; `{get_status_badge(t.get('status'))}`")
            st.caption(f"**T:** {t_disp}{dl_t} &nbsp;|&nbsp; **R:** {r_disp}{dl_r} &nbsp;|&nbsp; **REC:** {rec_disp}{dl_rec}")
          with ca:
            with st.popover("⚙️ Override & Reassign", width="stretch"):
              new_s = st.selectbox("Force Stage:", ALL_STATUSES, index=ALL_STATUSES.index(t.get("status")), key=f"s_{d_id}")
              
              t_opts = [""] + [e for e, d in vols.items() if "translator" in d.get("role") and d.get("status", "").lower() == "active"]
              r_opts = [""] + [e for e, d in vols.items() if "reviewer" in d.get("role") and d.get("status", "").lower() == "active"]
              rec_opts = [""] + [e for e, d in vols.items() if "recorder" in d.get("role") and d.get("status", "").lower() == "active"]
              
              new_t = st.selectbox("Translator:", t_opts, index=t_opts.index(t.get("translator")) if t.get("translator") in t_opts else 0, format_func=lambda x: format_vol_label(x) if x else "Unassigned", key=f"t_{d_id}")
              new_r = st.selectbox("Reviewer:", r_opts, index=r_opts.index(t.get("reviewer")) if t.get("reviewer") in r_opts else 0, format_func=lambda x: format_vol_label(x) if x else "Unassigned", key=f"r_{d_id}")
              new_rec = st.selectbox("Recorder:", rec_opts, index=rec_opts.index(t.get("recorder")) if t.get("recorder") in rec_opts else 0, format_func=lambda x: format_vol_label(x) if x else "Unassigned", key=f"rec_{d_id}")

              c_save, c_unlock = st.columns(2)
              if c_save.button("💾 Save", key=f"b_up_{d_id}", type="primary", width="stretch"):
                update_assignment_status(d_id, new_s)
                update_assignment_roles(d_id, new_t, new_r, new_rec)
                st.rerun()
              
              def admin_force_unlock(doc_id):
                  try:
                      res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE).execute()
                      for index, row in enumerate(res.get("values", [])):
                          if len(row) > 0 and row[0] == doc_id:
                              sheets_service.spreadsheets().values().update(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Sessions'!A{index + 1}:D{index + 1}", valueInputOption="USER_ENTERED", body={"values": [["", "", "", ""]]}).execute()
                  except Exception: pass

              if c_unlock.button("🔓 Unlock", key=f"unlock_{d_id}", width="stretch"):
                admin_force_unlock(d_id)
                st.success("Session cleared!")

    st.divider()
    st.divider()
    st.subheader("📤 Final Deployment Desk")
    sub_link = st.text_input("🔗 Paste Edition Submission Folder Link (Google Drive):", placeholder="https://drive.google.com/drive/folders/...")
    
    if ready_tasks:
        st.markdown(f"**{len(ready_tasks)} Assets Ready for Final Transfer:**")
        client_email = st.text_input("📧 CC Client Email (Optional - Leave blank to only notify Admin):")
        
        # --- NEW: BATCH DEPLOYMENT BUTTON ---
        if len(ready_tasks) > 1:
            if st.button(f"🚀 Deploy ALL {len(ready_tasks)} Ready Assets", type="primary", use_container_width=True):
                sub_id = extract_id_from_url(sub_link)
                if not sub_id:
                    st.error("Please paste a valid Google Drive Folder URL above.")
                else:
                    with st.spinner(f"Batch copying {len(ready_tasks)} assets..."):
                        success_count = 0
                        deployed_names = []
                        for rt in ready_tasks:
                            try:
                                # Copy Document
                                drive_service.files().copy(fileId=rt.get("doc_id"), body={'name': f"[Final Arabic] {rt.get('doc_name')}", 'parents': [sub_id]}).execute()
                                # Copy Audio
                                if rt.get("audio_link"):
                                    aud_id = extract_id_from_url(rt.get("audio_link"))
                                    if aud_id: drive_service.files().copy(fileId=aud_id, body={'parents': [sub_id]}).execute()
                                
                                update_assignment_status(rt.get("doc_id"), STATUS_SUBMITTED)
                                success_count += 1
                                deployed_names.append(rt.get("doc_name"))
                            except Exception as e:
                                st.error(f"Failed to transfer '{rt.get('doc_name')}': {e}")
                        
                        # Send 1 Consolidated Email Receipt
                        if success_count > 0:
                            admin_mail = st.secrets.get("ADMIN_EMAIL", "")
                            doc_list = "\n".join([f"- {name}" for name in deployed_names])
                            msg_body = f"Successfully deployed {success_count} assets to the client folder:\n\n{doc_list}"
                            
                            send_system_email(admin_mail, f"✅ Batch Deployed: {success_count} Assets", msg_body)
                            if client_email: 
                                send_system_email(client_email, f"New 12-Step Literature Available: {success_count} Assets", msg_body)
                            
                            st.success(f"Batch transfer complete! {success_count} assets deployed.")
                            time.sleep(2)
                            st.rerun()
            st.markdown("---")
        
        # --- INDIVIDUAL ASSET CARDS (Kept for granular control & QA Links) ---
        for rt in ready_tasks:
            with st.container(border=True):
                cd1, cd2 = st.columns([4, 1])
                cd1.markdown(f"📄 **{rt.get('doc_name')}**")
                
                qa_links = f"[📝 Review Text](https://docs.google.com/document/d/{rt.get('doc_id')}/edit)"
                if rt.get("audio_link"): qa_links += f" &nbsp;|&nbsp; [🎧 Listen to Audio]({rt.get('audio_link')})"
                cd1.caption(f"**QA Check:** {qa_links}")
                
                if cd2.button("Push to Client", key=f"push_{rt.get('doc_id')}"):
                    sub_id = extract_id_from_url(sub_link)
                    if not sub_id:
                        st.error("Please paste a valid Google Drive Folder URL above.")
                    else:
                        with st.spinner("Copying assets..."):
                            try:
                                drive_service.files().copy(fileId=rt.get("doc_id"), body={'name': f"[Final Arabic] {rt.get('doc_name')}", 'parents': [sub_id]}).execute()
                                if rt.get("audio_link"):
                                    aud_id = extract_id_from_url(rt.get("audio_link"))
                                    if aud_id: drive_service.files().copy(fileId=aud_id, body={'parents': [sub_id]}).execute()
                                
                                update_assignment_status(rt.get("doc_id"), STATUS_SUBMITTED)
                                
                                admin_mail = st.secrets.get("ADMIN_EMAIL", "")
                                msg_body = f"The final translated document and audio for '{rt.get('doc_name')}' have been successfully deployed to the client folder."
                                send_system_email(admin_mail, f"✅ Deployed: {rt.get('doc_name')}", msg_body)
                                if client_email: send_system_email(client_email, f"New 12-Step Literature Available: {rt.get('doc_name')}", msg_body)
                                
                                st.success("Transfer complete & notifications sent!")
                                time.sleep(1.5)
                                st.rerun()
                            except Exception as e:
                                st.error(f"Transfer failed: {e}")
    else:
        st.info("No tasks are currently at 'Recording Completed' ready for deployment.")

  # ---------------------------------------------------------
  # TAB 2: TEAM OPERATIONS
  # ---------------------------------------------------------
  with tab_team:
    st.subheader("Manage Volunteer Access")
    try:
      res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=VOLUNTEERS_RANGE).execute()
      vol_data = res.get("values", [])
      if not vol_data:
        vol_data = [["Email", "Name", "Role", "Status"]]
      df_vol = pd.DataFrame(vol_data[1:], columns=vol_data[0])
      edited_vol = st.data_editor(
          df_vol, num_rows="dynamic", use_container_width=True,
          column_config={
              "Role": st.column_config.SelectboxColumn("Role", options=[
                  "translator", "reviewer", "recorder", 
                  "translator, reviewer", "reviewer, recorder", 
                  "translator, recorder", "translator, reviewer, recorder"
              ], required=True),
              "Status": st.column_config.SelectboxColumn("Status", options=["Active", "Suspended"], required=True),
          },
      )
      if st.button("💾 Save Volunteers", type="primary"):
        overwrite_sheet_data(VOLUNTEERS_RANGE, [edited_vol.columns.tolist()] + edited_vol.fillna("").values.tolist())
        st.success("Volunteers database updated!")
    except Exception as e:
      st.error(f"Error fetching volunteers: {e}")

    st.divider()
    st.subheader("Team Broadcast System")
    broadcast_subject = st.text_input("Subject")
    broadcast_message = st.text_area("Message Body", height=150)
    if st.button("🚀 Send Broadcast", type="primary"):
      if broadcast_subject and broadcast_message:
        active_emails = [email for email, d in vols.items() if d.get("status", "").lower() == "active"]
        if active_emails:
          with st.spinner("Dispatching emails via secure SMTP..."):
            try:
              msg = EmailMessage()
              msg.set_content(f"12-Step Translation Project Update:\n\n{broadcast_message}")
              msg["Subject"] = f"[Coordinator Hub] {broadcast_subject}"
              msg["From"] = st.secrets.get("SMTP_EMAIL", "admin@localhost")
              msg["To"] = st.secrets.get("SMTP_EMAIL", "admin@localhost")
              msg["Bcc"] = ", ".join(active_emails)

              with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(st.secrets["SMTP_EMAIL"], st.secrets["SMTP_PASSWORD"])
                server.send_message(msg)
              st.success(f"Broadcast sent to {len(active_emails)} volunteers!")
            except Exception as e:
              st.error(f"Failed to send email. Error: {e}")
        else:
          st.warning("No active volunteers found.")
          
    st.divider()
    with st.expander("🚨 Automated Overdue Nudges"):
        st.write("Click below to scan the active pipeline and email anyone who has missed their phase deadline.")
        if st.button("🔔 Send Overdue Reminders", type="primary"):
            with st.spinner("Scanning deadlines..."):
                active = [t for t in fetch_assignments() if t.get("status") not in [STATUS_SUBMITTED, STATUS_REC_COMPLETED]]
                reminders_sent = 0
                today = datetime.today()
                
                for t in active:
                    status = t.get("status", "")
                    due_str, target_email, phase = "", "", ""
                    
                    if "Translation" in status: due_str, target_email, phase = t.get("t_due"), t.get("translator"), "Translation"
                    elif "Reviewer" in status: due_str, target_email, phase = t.get("r_due"), t.get("reviewer"), "Review"
                    elif "Recording" in status: due_str, target_email, phase = t.get("rec_due"), t.get("recorder"), "Recording"
                    
                    if due_str and target_email:
                        try:
                            due_date = datetime.strptime(f"{due_str} {today.year}", "%b %d %Y")
                            if due_date < today:
                                send_system_email(target_email, f"Urgent: '{t.get('doc_name')}' is Overdue", f"Hello,\n\nOur records indicate that the {phase} phase for '{t.get('doc_name')}' was due on {due_str}.\n\nPlease log in to the portal to complete this task as soon as possible so the next volunteer can begin.")
                                reminders_sent += 1
                        except Exception: pass
                        
                st.success(f"Scanned pipeline. Sent {reminders_sent} overdue reminders!")

  # ---------------------------------------------------------
  # TAB 3: SYSTEM TOOLS
  # ---------------------------------------------------------
  with tab_sys:
    st.subheader("🤖 Automated Edition Prep-Bot")
    with st.container(border=True):
      c_link, c_year, c_month = st.columns([3, 1, 1])
      raw_link = c_link.text_input("🔗 Raw Folder Link (Google Drive):")
      years = [str(y) for y in range(datetime.now().year - 1, datetime.now().year + 3)]
      sel_year = c_year.selectbox("Year:", years, index=1)
      months = ["1 - January", "2 - February", "3 - March", "4 - April", "5 - May", "6 - June", 
                "7 - July", "8 - August", "9 - September", "10 - October", "11 - November", "12 - December"]
      sel_month = c_month.selectbox("Month:", months)

      if st.button("🚀 Trigger Preparation Bot", type="primary", width="stretch"):
        if not raw_link: 
            st.error("❌ Please enter folder link.")
        else:
          gas_webhook_url = st.secrets.get("GAS_WEBAPP_URL", "")
          gas_token = st.secrets.get("GAS_ACCESS_TOKEN")
          
          if not gas_webhook_url or not gas_token:
              st.error("❌ Missing required GAS Webhook URL or Access Token in Streamlit Secrets.")
          else:
              with st.spinner("⏳ Sending command to Backend Bot..."):
                try:
                  payload = {
                      "url": raw_link, 
                      "year": sel_year, 
                      "monthNum": int(sel_month.split(" - ")[0]), 
                      "chatId": st.secrets.get("TELEGRAM_ADMIN_CHAT_ID", ""),
                      "token": gas_token
                  }
                  response = requests.post(gas_webhook_url, json=payload, timeout=15)
                  if response.status_code == 200 and response.json().get("status") == "success":
                      st.success(f"✅ {response.json().get('message')}")
                  else:
                      st.error(f"⚠️ Error: {response.json().get('message', 'Unknown API Error')}")
                except Exception as e:
                  st.error(f"Connection Error: {e}")

    st.divider()
    st.subheader("Live Terminology Editor")
    try:
      res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_RANGE).execute()
      glos_data = res.get("values", [])
      headers = ["ID", "Category", "English", "Arabic"]
      rows_to_display = glos_data[1:] if (glos_data and len(glos_data[0]) == 4 and not glos_data[0][0].isdigit()) else glos_data
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

    st.divider()
    st.subheader("🧹 Emergency Session Reset")
    if st.button("⚠️ Force Clear All Active Sessions", type="primary"):
      try:
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=GLOSSARY_SPREADSHEET_ID, range="'Sessions'!A2:D"
        ).execute()
        st.session_state.clear()
        st.success("All locks and sessions have been purged!")
        time.sleep(1)
        st.rerun()
      except Exception as e:
        st.error(f"Failed to reset sessions: {e}")

  st.stop()

# ==========================================
# 4. USER INBOX & TASK DELEGATION
# ==========================================
if not st.session_state.get("source_file_id") and st.session_state.get("app_mode") != "Coordinator Hub":
  col_t, col_l = st.columns([5, 1])
  col_t.title("⚙ Arabic Essay Service Portal")
  if col_l.button("🚪 Logout", width="stretch"):
    st.session_state.clear()
    st.rerun()

  st.subheader(f"👋 Welcome, {st.session_state.get('user_name')}")
  st.markdown("---")
  st.markdown("### 📬 Your Task Queue")
  
  assignments = fetch_assignments()
  my_tasks = []
  
  # Dynamic Workspace Routing (Ignores global role, focuses on active task phase)
  usr_email = st.session_state.get("user_email")
  for task in assignments:
    t_status = task.get("status")
    
    if task.get("translator") == usr_email and t_status in [STATUS_TRANS_ASSIGNED, STATUS_TRANS_STARTED]:
        my_tasks.append({**task, "task_mode": "Translator Mode"})
    elif task.get("reviewer") == usr_email and t_status in [STATUS_TRANS_COMPLETED, STATUS_REV_ASSIGNED, STATUS_REV_STARTED]:
        my_tasks.append({**task, "task_mode": "Reviewer Mode"})
    elif task.get("recorder") == usr_email and t_status in [STATUS_REC_PENDING, STATUS_REC_ASSIGNED, STATUS_REC_STARTED]:
        my_tasks.append({**task, "task_mode": "Recorder Mode"})

  if not my_tasks:
    st.success("🎉 You have no pending tasks in your queue. Great job!")
  else:
    for task in my_tasks:
      with st.container(border=True):
        c1, c2 = st.columns([4, 1.2])
        cur_s = task.get("status")
        task_mode = task.get("task_mode")
        
        if task_mode == "Translator Mode": due_str = task.get('t_due')
        elif task_mode == "Reviewer Mode": due_str = task.get('r_due')
        else: due_str = task.get('rec_due')
        
        badge = "🟠" if "Assigned" in cur_s else ("🟡" if "Started" in cur_s else "⏳")
        with c1:
          st.markdown(f"### 📄 **[{task.get('doc_name')}](https://docs.google.com/document/d/{task.get('doc_id')}/edit)**")
          st.caption(f"Current Stage: `{badge} {cur_s}` | **Phase Deadline:** `{due_str if due_str else 'N/A'}`")
        with c2:
          st.write("")
          btn_text = "🔄 Continue Working" if "Started" in cur_s else "🚀 Start Work"
          if st.button(btn_text, key=f"start_{task.get('doc_id')}", type="primary", width="stretch"):
            if cur_s in [STATUS_TRANS_ASSIGNED]: update_assignment_status(task.get("doc_id"), STATUS_TRANS_STARTED)
            elif cur_s in [STATUS_TRANS_COMPLETED, STATUS_REV_ASSIGNED]: update_assignment_status(task.get("doc_id"), STATUS_REV_STARTED)
            elif cur_s in [STATUS_REC_PENDING, STATUS_REC_ASSIGNED]: update_assignment_status(task.get("doc_id"), STATUS_REC_STARTED)
            
            # Switch the app into the correct mode for this specific task
            st.session_state["app_mode"] = task_mode
            st.session_state["active_task"] = task
            st.session_state["source_file_id"] = task.get("doc_id")
            st.rerun()
  st.stop()


# ==========================================
# 5. AI ENGINE & DOCUMENT PARSING
# ==========================================
safety_settings = []
if GENAI_AVAILABLE:
  safety_settings = [
      types.SafetySetting(category=types.HarmCategory.HARM_CATEGORY_HARASSMENT, threshold=types.HarmBlockThreshold.BLOCK_NONE),
      types.SafetySetting(category=types.HarmCategory.HARM_CATEGORY_HATE_SPEECH, threshold=types.HarmBlockThreshold.BLOCK_NONE),
      types.SafetySetting(category=types.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold=types.HarmBlockThreshold.BLOCK_NONE),
      types.SafetySetting(category=types.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold=types.HarmBlockThreshold.BLOCK_NONE),
  ]

def get_api_key_list():
  keys = st.secrets.get("GEMINI_API_KEYS")
  if isinstance(keys, str): return [k.strip() for k in keys.split(",") if k.strip()]
  elif isinstance(keys, list): return [str(k).strip() for k in keys]
  elif "GEMINI_API_KEY" in st.secrets: return [st.secrets["GEMINI_API_KEY"]]
  return []

def get_gemini_client():
  if not GENAI_AVAILABLE: return None
  keys = get_api_key_list()
  if not keys: return None
  return genai.Client(api_key=random.choice(keys))

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
  client = get_gemini_client()
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

def manage_document_lock(file_id: str, user_email: str):
  try:
    res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE).execute()
    rows = res.get("values", [])
    current_time = time.time()
    target_row_index = max(len(rows) + 1, 2)
    for index, row in enumerate(rows):
      if index == 0: continue
      if len(row) > 0 and row[0] == file_id:
        locked_by = row[1] if len(row) > 1 else ""
        ts = float(row[2]) if len(row) > 2 and row[2] else 0
        saved_json = row[3] if len(row) > 3 else ""
        if locked_by and locked_by != user_email and (current_time - ts) < LOCK_TIMEOUT_SECONDS:
          return {"status": "blocked", "locked_by": locked_by, "row_index": index + 1}
        if locked_by == user_email and saved_json:
          try: return {"status": "recovered", "data": json.loads(saved_json), "row_index": index + 1}
          except Exception: pass
        return {"status": "clear", "row_index": index + 1}
    return {"status": "clear", "row_index": target_row_index}
  except Exception:
    return {"status": "clear", "row_index": 2}

def acquire_document_lock(file_id: str, user_email: str, row_index: int):
  try:
    sheets_service.spreadsheets().values().update(
        spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Sessions'!A{row_index}:D{row_index}",
        valueInputOption="USER_ENTERED", body={"values": [[file_id, user_email, str(time.time()), ""]]},
    ).execute()
  except Exception: pass

def save_draft_to_sheet(file_id, user_email, row_index, processed_data):
  if not row_index: return
  try:
    json_data = json.dumps(processed_data, ensure_ascii=False)
    sheets_service.spreadsheets().values().update(
        spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Sessions'!A{row_index}:D{row_index}",
        valueInputOption="USER_ENTERED", body={"values": [[file_id, user_email, str(time.time()), json_data]]},
    ).execute()
  except Exception: pass

def release_document_lock(row_index):
  if not row_index: return
  try:
    sheets_service.spreadsheets().values().update(
        spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=f"'Sessions'!A{row_index}:D{row_index}",
        valueInputOption="USER_ENTERED", body={"values": [["", "", "", ""]]},
    ).execute()
  except Exception: pass

@st.cache_data(ttl=3600)
def fetch_glossary():
  try:
    res = sheets_service.spreadsheets().values().get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_DATA_RANGE).execute()
    vals = res.get("values", [])
    glos_lines = []
    count = 0
    for row in vals:
      if len(row) >= 2 and row[0].strip() and row[1].strip():
        if row[0].strip().lower() == "english": continue
        glos_lines.append(f"- {row[0].strip()} -> {row[1].strip()}")
        count += 1
    return "\n".join(glos_lines), count
  except Exception:
    return "", 0

def generate_html_diff(original, suggested):
  if not original or original.startswith("[MISSING"):
    return ("<div dir='rtl' style='text-align: right; color: #0369a1; background-color: #e0f2fe; padding: 10px; border-radius: 5px; font-family: \"Cairo\";'>✨ Initial AI Draft</div>")
  if original.strip() == suggested.strip():
    return ("<div dir='rtl' style='text-align: right; color: #155724; background-color: #d4edda; padding: 10px; border-radius: 5px; font-family: \"Cairo\";'>✨ Perfect Match — No Edits Needed</div>")
  diff = difflib.ndiff(original.split(), suggested.split())
  html = ["<div dir='rtl' style='font-family: \"Cairo\", sans-serif; font-size: 18px; line-height: 2; text-align: right; background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef;'>"]
  for word in diff:
    if word.startswith("- "): html.append(f"<span style='background-color: #ffcdd2; color: #b71c1c; text-decoration: line-through; padding: 2px; border-radius: 4px;'>{word[2:]}</span>")
    elif word.startswith("+ "): html.append(f"<span style='background-color: #c8e6c9; color: #1b5e20; font-weight: bold; padding: 2px; border-radius: 4px;'>{word[2:]}</span>")
    elif word.startswith("  "): html.append(f"<span style='color: #212529;'>{word[2:]}</span>")
  html.append("</div>")
  return " ".join(html)

def _is_retryable(err_str): return any(kw.lower() in err_str.lower() for kw in RETRYABLE_KEYWORDS)
def _backoff_sleep(attempt): time.sleep(min(BASE_BACKOFF_SECONDS * (2**attempt), MAX_BACKOFF_SECONDS))

def _call_gemini(model_name, prompt, schema_type):
  client = get_gemini_client()
  if not client: raise ValueError("No Gemini API keys configured.")
  response = client.models.generate_content(
      model=model_name, contents=prompt,
      config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=schema_type, safety_settings=safety_settings, temperature=0.2),
  )
  clean_text = response.text.replace("```json", "").replace("```", "").strip()
  match = re.search(r"\{.*\}", clean_text, re.DOTALL)
  if match: clean_text = match.group(0)
  return json.loads(clean_text)

# --- RECOVERY FELLOWSHIP LITERARY PROMPTS ---
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
        if _is_retryable(str(e)):
          _backoff_sleep(attempt)
        else:
          break
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
        if _is_retryable(str(e)):
          _backoff_sleep(attempt)
        else:
          break
  return {"status": "major_rewrite", "suggested_arabic": arabic, "reasoning": "⚠ Error"}

def translate_batch_with_fallback(batch_segments, glossary_text):
  if not GENAI_AVAILABLE or not get_api_key_list():
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
        if _is_retryable(str(e)):
          _backoff_sleep(attempt)
        else:
          break
  results = []
  for s in batch_segments:
    res = translate_with_ai(s.get("english", ""), glossary_text)
    results.append({"id": s.get("id", 0), "arabic_translation": res.get("arabic_translation", ""), "glossary_notes": res.get("glossary_notes", "")})
  return results

def review_batch_with_fallback(batch_segments, glossary_text):
  if not GENAI_AVAILABLE or not get_api_key_list():
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
        if _is_retryable(str(e)):
          _backoff_sleep(attempt)
        else:
          break
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
      for run in elem.get("paragraph", {}).get("elements", []):
        if "textRun" in run: para_text += run.get("textRun", {}).get("content", "")
      clean_text = para_text.strip()
      if bool(re.search(r"[a-zA-Z\u0600-\u06FF]", clean_text)):
        paras.append({"text": clean_text, "start": elem.get("startIndex"), "end": elem.get("endIndex")})
    elif "table" in elem:
      for row in elem.get("table", {}).get("tableRows", []):
        for cell in row.get("tableCells", []): paras.extend(_parse_docs_elements(cell.get("content", [])))
  return paras

def extract_text_from_drive(file_id: str, is_retry=False):
  try:
    docs_svc, drive_svc, _ = get_google_services()
    file_meta = drive_svc.files().get(fileId=file_id, fields="mimeType, parents").execute()
    st.session_state["active_task_parent_folder"] = file_meta.get("parents", [""])[0]
    if file_meta.get("mimeType") == "application/vnd.google-apps.document":
      try: document = docs_svc.documents().get(documentId=file_id, includeTabsContent=True).execute()
      except: document = docs_svc.documents().get(documentId=file_id).execute()
      all_paras = []
      def sweep_doc_obj(doc_obj):
        tp = []
        tp.extend(_parse_docs_elements(doc_obj.get("body", {}).get("content", [])))
        for footer in doc_obj.get("footers", {}).values(): tp.extend(_parse_docs_elements(footer.get("content", [])))
        for header in doc_obj.get("headers", {}).values(): tp.extend(_parse_docs_elements(header.get("content", [])))
        return tp
      tabs = document.get("tabs", [])
      if tabs:
        for tab in tabs: all_paras.extend(sweep_doc_obj(tab.get("documentTab", {})))
      else: all_paras.extend(sweep_doc_obj(document))
      return all_paras
    return None
  except Exception as e:
    if ("Broken pipe" in str(e) or "Errno 32" in str(e)) and not is_retry:
      return extract_text_from_drive(file_id, True)
    return None

def upload_audio_to_drive(uploaded_file, doc_name, parent_folder_id):
    try:
        gas_url = st.secrets.get("GAS_WEBAPP_URL")
        gas_token = st.secrets.get("GAS_ACCESS_TOKEN")
        cc_api_key = st.secrets.get("CLOUDCONVERT_API_KEY") 
        
        if not gas_url or not gas_token: 
            return None
            
        file_name_clean = re.sub(r'[\\/*?:<>]', '', doc_name).strip()
        file_ext = uploaded_file.name.split('.')[-1].lower() if hasattr(uploaded_file, 'name') and '.' in uploaded_file.name else 'wav'
        file_bytes = uploaded_file.getvalue()
        mime_type = "audio/wav" if file_ext == "wav" else "audio/mpeg"

        # --- CloudConvert API Engine (Transforms Non-MP3 to MP3 on the fly) ---
        if cc_api_key and file_ext != "mp3":
            headers = {"Authorization": f"Bearer {cc_api_key}"}
            job_payload = {
                "tasks": {
                    "import-task": {"operation": "import/upload"},
                    "convert-task": {"operation": "convert", "input": "import-task", "output_format": "mp3"},
                    "export-task": {"operation": "export/url", "input": "convert-task"}
                }
            }
            res = requests.post("https://api.cloudconvert.com/v2/jobs", json=job_payload, headers=headers).json()
            job_data = res.get("data", {})
            
            import_task = next((t for t in job_data.get("tasks", []) if t["name"] == "import-task"), None)
            if import_task and "result" in import_task:
                upload_url = import_task["result"]["form"]["url"]
                upload_params = import_task["result"]["form"]["parameters"]
                
                requests.post(upload_url, data=upload_params, files={'file': (f"audio.{file_ext}", file_bytes)})
                wait_res = requests.get(f"https://sync.api.cloudconvert.com/v2/jobs/{job_data['id']}", headers=headers).json()
                
                export_task = next((t for t in wait_res.get("data", {}).get("tasks", []) if t["name"] == "export-task"), None)
                if export_task and "result" in export_task and export_task["result"].get("files"):
                    mp3_url = export_task["result"]["files"][0]["url"]
                    file_bytes = requests.get(mp3_url).content
                    file_ext = "mp3"
                    mime_type = "audio/mpeg"
            
        # --- Final Push to Google Apps Script Webhook ---
        payload = {
            "action": "upload_audio", "parent_folder_id": parent_folder_id,
            "file_name": f"{file_name_clean}.{file_ext}", 
            "mime_type": mime_type,
            "file_base64": base64.b64encode(file_bytes).decode('utf-8'),
            "token": gas_token
        }
        response = requests.post(gas_url, json=payload).json()
        return response.get("webViewLink") if response.get("status") == "success" else None
    except Exception as e: 
        print(f"Audio upload error: {e}")
        return None

def smart_align(paras):
    arabic_paras, english_paras = [], []
    for p in paras:
        text = p.get("text", "").strip()
        if not text or text.startswith("--- Original"): continue
        if bool(re.search(r"[\u0600-\u06FF]", text)): arabic_paras.append(p)
        else: english_paras.append(p)
    segments = []
    for i in range(max(len(arabic_paras), len(english_paras))):
        ar_p = arabic_paras[i] if i < len(arabic_paras) else None
        en_p = english_paras[i] if i < len(english_paras) else None
        segments.append({
            "id": i + 1, "english": en_p.get("text") if en_p else "[MISSING ENGLISH SOURCE]",
            "en_start": en_p.get("start") if en_p else None,
            "arabic": ar_p.get("text") if ar_p else "[MISSING ARABIC TRANSLATION]",
            "ar_start": ar_p.get("start") if ar_p else None, "ar_end": ar_p.get("end") if ar_p else None
        })
    return segments

def push_to_drive_translator(file_id, finalized_data):
    try:
        docs_svc, _, _ = get_google_services()
        arabic_text = "\n\n".join([item.get('final_arabic', '') if isinstance(item, dict) else str(item) for item in finalized_data])
        insert_text = arabic_text + "\n\n\n--- Original English Document ---\n\n"
        docs_svc.documents().batchUpdate(documentId=file_id, body={'requests': [{'insertText': {'location': {'index': 1},'text': insert_text}}]}).execute()
        return True
    except: return False

def push_to_drive_reviewer(file_id, finalized_data):
    try:
        docs_svc, _, _ = get_google_services()
        valid_edits = sorted([i for i in finalized_data if i.get('ar_start') and i.get('ar_end')], key=lambda x: x['ar_start'], reverse=True)
        requests_arr = []
        for item in valid_edits:
            requests_arr.append({'deleteContentRange': {'range': {'startIndex': item['ar_start'], 'endIndex': item['ar_end'] - 1 }}})
            requests_arr.append({'insertText': {'location': {'index': item['ar_start']}, 'text': item['final_arabic']}})
        if requests_arr: docs_svc.documents().batchUpdate(documentId=file_id, body={'requests': requests_arr}).execute()
        return True
    except: return False


# ==========================================
# 6. ACTIVE WORKSPACE
# ==========================================
app_mode = st.session_state.get("app_mode")
task = st.session_state.get("active_task")

if not task:
  st.session_state["source_file_id"] = None
  st.rerun()

file_id = task.get("doc_id")

# ---------------------------------------------------------
# RECORDER WORKSPACE
# ---------------------------------------------------------
if app_mode == "Recorder Mode":
    c1, c2 = st.columns([5, 1])
    c1.markdown(f"## 🎙️ Recording Studio: `{task.get('doc_name')}`")
    if c2.button("⬅️ Back to Inbox", width="stretch"):
        st.session_state.update({"active_task": None, "source_file_id": None, "app_mode": "Volunteer Dashboard"}); st.rerun()
    st.markdown(f"[🔗 Open Original Document](https://docs.google.com/document/d/{file_id}/edit)")
    
    paras = extract_text_from_drive(file_id)
    if paras:
        ar_paras = [p.get("text", "") for p in paras if re.search(r"[\u0600-\u06FF]", p.get("text", ""))]
        st.markdown(f"<div class='reading-mode'>{'<br><br>'.join(ar_paras)}</div>", unsafe_allow_html=True)
        st.divider()
        c_rec, c_up = st.columns(2)
        recorded_audio = c_rec.audio_input("🎙️ Record from mic:")
        uploaded_audio = c_up.file_uploader("📂 Upload audio (MP3/WAV):", type=["mp3", "wav", "m4a"])
        final_audio = recorded_audio or uploaded_audio
        if final_audio:
            st.audio(final_audio)
            if st.button("🚀 Upload & Complete Task", type="primary", width="stretch"):
                with st.spinner("Processing & Uploading... (this may take a moment if converting formats)"):
                    file_link = upload_audio_to_drive(final_audio, task.get("doc_name"), st.session_state.get("active_task_parent_folder", ""))
                    if file_link:
                        update_assignment_audio_link(file_id, file_link)
                        update_assignment_status(file_id, STATUS_REC_COMPLETED)
                        
                        admin_mail = st.secrets.get("ADMIN_EMAIL", "")
                        send_system_email(admin_mail, f"🎙️ Audio Ready: {task.get('doc_name')}", f"Hello,\n\n{st.session_state.get('user_name')} has successfully uploaded the audio for '{task.get('doc_name')}'.\n\nIt is now ready for deployment in the Coordinator Hub.")
                        
                        st.session_state.update({"active_task": None, "source_file_id": None, "app_mode": "Volunteer Dashboard"})
                        st.rerun()
    st.stop()

# ---------------------------------------------------------
# TRANSLATOR / REVIEWER WORKSPACE
# ---------------------------------------------------------
glossary_data, glossary_term_count = fetch_glossary()
c1, c2 = st.columns([4, 2])
c1.markdown(f"## 📝 Workspace: `{task.get('doc_name', 'Document')}`")
with c2:
  b_col, g_col = st.columns([1, 1])
  if b_col.button("⬅️ Back", width="stretch"):
    st.session_state.update({"active_task": None, "source_file_id": None, "processed_data": None, "app_mode": "Volunteer Dashboard"}); st.rerun()
  if glossary_term_count > 0: g_col.success(f"📖 Glossary: {glossary_term_count} terms")
  else: g_col.warning("⚠️ Glossary: Not Loaded")

is_bypass_task = task.get("translator") == ""

if not st.session_state.get("processed_data"):
  lock = manage_document_lock(file_id, st.session_state.get("user_email"))
  if lock["status"] == "blocked": st.error(f"🛑 **Locked:** by `{lock.get('locked_by')}`."); st.stop()
  elif lock["status"] == "recovered":
    st.session_state.update({"session_row_index": lock["row_index"], "processed_data": lock["data"]})
    st.success("♻️ Session Recovered.")
  elif lock["status"] == "clear":
    st.session_state["session_row_index"] = lock["row_index"]
    acquire_document_lock(file_id, st.session_state.get("user_email"), lock["row_index"])
    paras = extract_text_from_drive(file_id)

    if paras:
      processed_results = []
      progress_bar = st.progress(0)

      if app_mode == "Translator Mode":
        batches = [paras[i : i + BATCH_SIZE] for i in range(0, len(paras), BATCH_SIZE)]
        for idx, batch in enumerate(batches):
          batch_payload = [{"id": j + 1, "english": p.get("text", "")} for j, p in enumerate(batch)]
          ai_results = translate_batch_with_fallback(batch_payload, glossary_data)
          for ai_res in ai_results:
            ai_id = int(ai_res.get("id", 1))
            a_idx = ai_id - 1
            eng_text = batch[a_idx].get("text", "") if 0 <= a_idx < len(batch) else batch[0].get("text", "")
            eng_start = batch[a_idx].get("start") if 0 <= a_idx < len(batch) else batch[0].get("start")
            processed_results.append({
                "id": ai_id + (idx * BATCH_SIZE), "english": eng_text, "en_start": eng_start,
                "arabic_translation": ai_res.get("arabic_translation", ""),
                "glossary_notes": ai_res.get("glossary_notes", ""),
                "user_arabic": ai_res.get("arabic_translation", ""), "is_approved": False,
            })
          progress_bar.progress((idx + 1) / len(batches))

      else:
        has_arabic = any(re.search(r"[\u0600-\u06FF]", p.get("text", "")) for p in paras)
        if not has_arabic or is_bypass_task:
          batches = [paras[i : i + BATCH_SIZE] for i in range(0, len(paras), BATCH_SIZE)]
          draft_segments = []
          for idx, batch in enumerate(batches):
            batch_payload = [{"id": j + 1, "english": p.get("text", "")} for j, p in enumerate(batch)]
            ai_trans = translate_batch_with_fallback(batch_payload, glossary_data)
            for ai_t in ai_trans:
              a_idx = int(ai_t.get("id", 1)) - 1
              eng_text = batch[a_idx].get("text", "") if 0 <= a_idx < len(batch) else batch[0].get("text", "")
              eng_start = batch[a_idx].get("start") if 0 <= a_idx < len(batch) else batch[0].get("start")
              draft_segments.append({"id": ai_t.get("id", 1) + (idx * BATCH_SIZE), "english": eng_text, "en_start": eng_start, "arabic": ai_t.get("arabic_translation", ""), "glossary_notes": ai_t.get("glossary_notes", "")})
          rev_batches = [draft_segments[i : i + BATCH_SIZE] for i in range(0, len(draft_segments), BATCH_SIZE)]
          for idx, r_batch in enumerate(rev_batches):
            ai_revs = review_batch_with_fallback(r_batch, glossary_data)
            for ai_r, d_seg in zip(ai_revs, r_batch):
              suggested = ai_r.get("suggested_arabic", d_seg.get("arabic", ""))
              processed_results.append({"id": d_seg.get("id"), "status": ai_r.get("status", "minor_edits"), "english": d_seg.get("english", ""), "en_start": d_seg.get("en_start"), "original_arabic": d_seg.get("arabic", ""), "suggested_arabic": suggested, "reasoning": ai_r.get("reasoning", "") + f" | {d_seg.get('glossary_notes')}", "ar_start": None, "ar_end": None, "user_arabic": suggested, "is_approved": ai_r.get("status") == "perfect"})
            progress_bar.progress((idx + 1) / len(rev_batches))
        else:
          segments = smart_align(paras)
          normal_segs = [s for s in segments if s.get("english") != "[MISSING ENGLISH SOURCE]" and s.get("arabic") != "[MISSING ARABIC TRANSLATION]"]
          batches = [normal_segs[i : i + BATCH_SIZE] for i in range(0, len(normal_segs), BATCH_SIZE)]
          for idx, batch in enumerate(batches):
            ai_results = review_batch_with_fallback(batch, glossary_data)
            for ai_res, o_seg in zip(ai_results, batch):
              suggested = ai_res.get("suggested_arabic", o_seg.get("arabic", ""))
              processed_results.append({"id": o_seg.get("id"), "status": ai_res.get("status", "minor_edits"), "english": o_seg.get("english", ""), "en_start": None, "original_arabic": o_seg.get("arabic", ""), "suggested_arabic": suggested, "reasoning": ai_res.get("reasoning", ""), "ar_start": o_seg.get("ar_start"), "ar_end": o_seg.get("ar_end"), "user_arabic": suggested, "is_approved": ai_res.get("status") == "perfect"})
            progress_bar.progress((idx + 1) / len(batches))

      save_draft_to_sheet(file_id, st.session_state.get("user_email"), st.session_state.get("session_row_index"), processed_results)
      st.session_state["processed_data"] = processed_results
      st.rerun()

approved_count, finalized_data, state_modified = 0, [], False
st.divider()

for i, item in enumerate(st.session_state.get("processed_data", [])):
  seg_id = item.get("id", i + 1)
  with st.container(border=True):
    if app_mode == "Translator Mode":
      st.markdown(f"### Segment {seg_id}")
      c_en, c_ar = st.columns(2)
      c_en.info(item.get("english", ""))
      if item.get("glossary_notes"): c_en.caption(f"💡 {item.get('glossary_notes')}")
      final_text = c_ar.text_area("Final text", value=item.get("user_arabic", ""), height=120, key=f"edit_{i}", label_visibility="collapsed")
      if final_text != item.get("user_arabic"): item["user_arabic"] = final_text; state_modified = True
    else:
      color = "🟢" if item.get("status") == "perfect" else ("🟡" if item.get("status") == "minor_edits" else "🔴")
      st.markdown(f"### Segment {seg_id} | Status: {color} {item.get('status').upper()}")
      c_en, c_ar = st.columns(2)
      c_en.info(item.get("english", ""))
      c_ar.markdown(generate_html_diff(item.get("original_arabic", ""), item.get("suggested_arabic", "")), unsafe_allow_html=True)
      c_ar.expander("💡 Reasoning").write(item.get("reasoning", ""))
      final_text = c_ar.text_area("Final Output", value=item.get("user_arabic", ""), height=120, key=f"edit_ar_{i}", label_visibility="collapsed")
      if final_text != item.get("user_arabic"): item["user_arabic"] = final_text; state_modified = True

    chk = st.checkbox(f"✅ Approve Segment {seg_id}", key=f"chk_{i}", value=item.get("is_approved", False))
    if chk != item.get("is_approved"): item["is_approved"] = chk; state_modified = True
    if chk:
      approved_count += 1
      finalized_data.append({"final_arabic": final_text, "ar_start": item.get("ar_start"), "ar_end": item.get("ar_end"), "en_start": item.get("en_start")})

if state_modified: save_draft_to_sheet(file_id, st.session_state.get("user_email"), st.session_state.get("session_row_index"), st.session_state.get("processed_data"))

st.divider()
total_segments = len(st.session_state.get("processed_data", []))
st.write(f"### **Approved Segments: {approved_count} / {total_segments}**")

if approved_count == total_segments and total_segments > 0:
  st.success("🎉 All segments approved!")
  if "review_unlocked" not in st.session_state: st.session_state["review_unlocked"] = False
  if not st.session_state["review_unlocked"]:
    st.error("🚨 **CRITICAL STEP:** Verify narrative flow.")
    if st.button("👀 Confirm final text"): st.session_state["review_unlocked"] = True; st.rerun()
  else:
    if st.button("🚀 Push to Drive & Conclude", type="primary", width="stretch"):
      with st.spinner("Processing Drive updates..."):
        success = push_to_drive_translator(file_id, finalized_data) if (app_mode == "Translator Mode" or is_bypass_task) else push_to_drive_reviewer(file_id, finalized_data)
        if success:
          # --- BUG FIX: Fetch fresh data to prevent race conditions ---
          fresh_assignments = fetch_assignments()
          fresh_task = next((t for t in fresh_assignments if t.get("doc_id") == file_id), task)
          
          t_email = fresh_task.get("translator", "")
          r_email = fresh_task.get("reviewer", "")
          rec_email = fresh_task.get("recorder", "")
          
          # --- Smart Baton Pass (Handles Super Volunteers) ---
          if app_mode == "Translator Mode": 
              if r_email == t_email and r_email != "":
                  # Multi-Role Bypass: The Translator is the Reviewer. Skip straight to Recording.
                  new_status = STATUS_REC_ASSIGNED if rec_email else STATUS_REC_PENDING
                  if new_status == STATUS_REC_ASSIGNED:
                      send_system_email(rec_email, f"🟢 New Recording Task: {fresh_task.get('doc_name')}", f"Hello,\n\nThe text for '{fresh_task.get('doc_name')}' is finalized! It is now ready for Audio Recording.\n\nPlease log in to the portal to begin.")
              else:
                  new_status = STATUS_REV_ASSIGNED if r_email else STATUS_TRANS_COMPLETED
                  if new_status == STATUS_REV_ASSIGNED:
                      send_system_email(r_email, f"🟢 New Review Task: {fresh_task.get('doc_name')}", f"Hello,\n\nThe translation for '{fresh_task.get('doc_name')}' is complete! It is now in your queue for Review.\n\nPlease log in to the portal to begin.")
          else: 
              new_status = STATUS_REC_ASSIGNED if rec_email else STATUS_REC_PENDING
              if new_status == STATUS_REC_ASSIGNED:
                  send_system_email(rec_email, f"🟢 New Recording Task: {fresh_task.get('doc_name')}", f"Hello,\n\nThe review for '{fresh_task.get('doc_name')}' is complete! It is now ready for Audio Recording.\n\nPlease log in to the portal to begin.")
          
          update_assignment_status(file_id, new_status)
          release_document_lock(st.session_state.get("session_row_index"))
          st.session_state.update({"active_task": None, "source_file_id": None, "processed_data": None, "review_unlocked": False, "app_mode": "Volunteer Dashboard"})
          st.rerun()
