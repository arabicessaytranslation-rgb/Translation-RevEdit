import difflib
from email.message import EmailMessage
import json
import random
import re
import smtplib
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
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
st.set_page_config(
    page_title="12-Step AI Suite: Workflow Portal",
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
    </style>
    """,
      unsafe_allow_html=True,
  )


apply_custom_css()

# Google Sheets Configuration
GLOSSARY_SPREADSHEET_ID = "1oc4TCY_iK9R7mBiXgb5rKWssjmrQywYg6UpOBXx8pUQ"
GLOSSARY_RANGE = "'المصطلحات'!A:D"
GLOSSARY_DATA_RANGE = "'المصطلحات'!C:D"
SESSIONS_RANGE = "'Sessions'!A:D"
VOLUNTEERS_RANGE = "'Volunteers'!A:D"
ASSIGNMENTS_RANGE = "'Assignments'!A:E"

LOCK_TIMEOUT_SECONDS = 14400

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


def fetch_volunteers():
  try:
    res = (
        sheets_service.spreadsheets()
        .values()
        .get(
            spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=VOLUNTEERS_RANGE
        )
        .execute()
    )
    rows = res.get("values", [])
    volunteers = {}
    if len(rows) > 1:
      for r in rows[1:]:
        if len(r) >= 4:
          email, name, role, status = (
              r[0].strip().lower(),
              r[1],
              r[2].lower(),
              r[3],
          )
          volunteers[email] = {"name": name, "role": role, "status": status}
    return volunteers
  except Exception:
    return {}


def overwrite_sheet_data(range_name, data_matrix):
  sheets_service.spreadsheets().values().clear(
      spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=range_name
  ).execute()
  body = {"values": data_matrix}
  sheets_service.spreadsheets().values().update(
      spreadsheetId=GLOSSARY_SPREADSHEET_ID,
      range=range_name,
      valueInputOption="USER_ENTERED",
      body=body,
  ).execute()


# --- ASSIGNMENTS DATABASE FUNCTIONS ---
def fetch_assignments():
  try:
    res = (
        sheets_service.spreadsheets()
        .values()
        .get(
            spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=ASSIGNMENTS_RANGE
        )
        .execute()
    )
    rows = res.get("values", [])
    assignments = []
    if len(rows) > 1:
      for idx, r in enumerate(rows[1:]):
        r.extend([""] * (5 - len(r)))
        assignments.append({
            "row_index": idx + 2,
            "doc_id": r[0].strip(),
            "doc_name": r[1],
            "translator": r[2].lower().strip(),
            "reviewer": r[3].lower().strip(),
            "status": r[4],
        })
    return assignments
  except Exception:
    return []


def assign_task_to_sheet(doc_id, doc_name, t_email, r_email, status):
  assignments = fetch_assignments()
  row_idx = None
  for row in assignments:
    if row.get("doc_id") == doc_id:
      row_idx = row.get("row_index")
      break

  body = {"values": [[doc_id, doc_name, t_email, r_email, status]]}
  if row_idx:
    range_name = f"'Assignments'!A{row_idx}:E{row_idx}"
    sheets_service.spreadsheets().values().update(
        spreadsheetId=GLOSSARY_SPREADSHEET_ID,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()
  else:
    range_name = "'Assignments'!A:E"
    sheets_service.spreadsheets().values().append(
        spreadsheetId=GLOSSARY_SPREADSHEET_ID,
        range=range_name,
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body=body,
    ).execute()


def update_assignment_status(doc_id, new_status):
  assignments = fetch_assignments()
  for task in assignments:
    if task.get("doc_id") == doc_id:
      range_name = f"'Assignments'!E{task.get('row_index')}"
      body = {"values": [[new_status]]}
      sheets_service.spreadsheets().values().update(
          spreadsheetId=GLOSSARY_SPREADSHEET_ID,
          range=range_name,
          valueInputOption="USER_ENTERED",
          body=body,
      ).execute()
      break


# --- LOGIN SCREEN ---
def login_screen():
  if st.session_state.get("authenticated"):
    return True

  col1, col2, col3 = st.columns([1, 2, 1])
  with col2:
    st.title("🤝 12-Step AI Suite")
    st.caption("Volunteer Translation & Review Portal")
    with st.container(border=True):
      with st.form("login_form"):
        email = (
            st.text_input("Enter your registered email address:")
            .strip()
            .lower()
        )
        if st.form_submit_button(
            "Access Portal", type="primary", use_container_width=True
        ):
          volunteers = fetch_volunteers()
          if email in volunteers:
            user_data = volunteers[email]
            if user_data.get("status", "").lower() != "active":
              st.error("Account suspended. Please contact the coordinator.")
            else:
              st.session_state.update({
                  "authenticated": True,
                  "user_email": email,
                  "user_role": user_data.get("role"),
                  "user_name": user_data.get("name"),
              })
              st.session_state["app_mode"] = (
                  "God Mode"
                  if user_data.get("role") == "admin"
                  else (
                      "Translator Mode"
                      if user_data.get("role") == "translator"
                      else "Reviewer Mode"
                  )
              )
              st.rerun()
          else:
            st.error("Email not recognized in the system.")
  return False


if not login_screen():
  st.stop()

# Initialize core session variables
if "processed_data" not in st.session_state:
  st.session_state["processed_data"] = None
if "source_file_id" not in st.session_state:
  st.session_state["source_file_id"] = None
if "session_row_index" not in st.session_state:
  st.session_state["session_row_index"] = None
if "active_task" not in st.session_state:
  st.session_state["active_task"] = None


# ==========================================
# 3. GOD MODE (ADMIN DASHBOARD)
# ==========================================
if st.session_state.get("app_mode") == "God Mode" and not st.session_state.get(
    "source_file_id"
):
  col1, col2 = st.columns([5, 1])
  col1.title("⚡ Central Command: God Mode")
  if col2.button("🚪 Logout", type="primary"):
    st.session_state.clear()
    st.rerun()

  # 6 Full Tabs for the Admin
  tab_track, tab_assign, tab_vols, tab_glos, tab_radar, tab_bcast = st.tabs([
      "📊 Task Tracker",
      "🗂️ Assignment Desk",
      "👥 Volunteers",
      "📖 Glossary",
      "📡 Live Radar",
      "📢 Broadcast",
  ])

  vols = fetch_volunteers()
  all_assignments = fetch_assignments()

  # Helper to display clean names instead of bare emails
  def format_vol_label(email_key):
    if not email_key or email_key.startswith("["):
      return email_key
    v_info = vols.get(email_key, {})
    v_name = v_info.get("name")
    return f"{v_name} ({email_key})" if v_name else email_key

  # =========================================================
  # TAB 1: TASK TRACKER (THE COMPLETE PIPELINE OVERVIEW)
  # =========================================================
  with tab_track:
    st.subheader("📊 Live Assignment Tracker")

    if not all_assignments:
      st.info(
          "No tasks assigned yet. Go to '🗂️ Assignment Desk' to assign your"
          " first document."
      )
    else:
      # Top KPI Metrics Cards
      total_count = len(all_assignments)
      p_trans_count = sum(
          1 for a in all_assignments if a.get("status") == "Pending Translation"
      )
      p_rev_count = sum(
          1 for a in all_assignments if a.get("status") == "Pending Review"
      )
      comp_count = sum(
          1 for a in all_assignments if a.get("status") == "Completed"
      )

      k1, k2, k3, k4 = st.columns(4)
      k1.metric("Total Assigned", total_count)
      k2.metric("Pending Translation", p_trans_count)
      k3.metric("Pending Review", p_rev_count)
      k4.metric("Completed", comp_count)

      st.divider()

      # Filter & Search Controls
      col_srch, col_st = st.columns([3, 1])
      srch = (
          col_srch.text_input(
              "🔍 Search tasks by document title, translator, or reviewer:", ""
          )
          .strip()
          .lower()
      )
      st_filter = col_st.selectbox(
          "Filter by Status:",
          ["All", "Pending Translation", "Pending Review", "Completed"],
      )

      filtered_tasks = []
      for a in all_assignments:
        match_st = (st_filter == "All") or (a.get("status") == st_filter)
        match_txt = (
            srch in a.get("doc_name", "").lower()
            or srch in a.get("translator", "").lower()
            or srch in a.get("reviewer", "").lower()
            or srch in vols.get(a.get("translator"), {}).get("name", "").lower()
            or srch in vols.get(a.get("reviewer"), {}).get("name", "").lower()
        )
        if match_st and match_txt:
          filtered_tasks.append(a)

      if not filtered_tasks:
        st.warning("No assignments match your search filter.")
      else:
        for t in filtered_tasks:
          d_id = t.get("doc_id")
          d_name = t.get("doc_name")
          t_em = t.get("translator")
          r_em = t.get("reviewer")
          status = t.get("status")

          t_disp = (
              f"{vols.get(t_em, {}).get('name', t_em)} ({t_em})"
              if t_em
              else "🤖 AI Bypass (No Translator)"
          )
          r_disp = (
              f"{vols.get(r_em, {}).get('name', r_em)} ({r_em})"
              if r_em
              else "Unassigned"
          )

          badge = (
              "🟡"
              if status == "Pending Translation"
              else ("🔵" if status == "Pending Review" else "🟢")
          )

          with st.container(border=True):
            ci, ca = st.columns([4, 1.2])
            with ci:
              st.markdown(f"**📄 {d_name}** &nbsp;&nbsp; `{badge} {status}`")
              st.caption(
                  f"**Translator:** {t_disp} &nbsp;|&nbsp; **Reviewer:**"
                  f" {r_disp} &nbsp;|&nbsp; **ID:** `{d_id}`"
              )
            with ca:
              with st.popover("⚙️ Update Status", use_container_width=True):
                st.write(f"Task: `{d_name[:25]}...`")
                stat_options = [
                    "Pending Translation",
                    "Pending Review",
                    "Completed",
                ]
                cur_idx = (
                    stat_options.index(status) if status in stat_options else 0
                )
                new_s = st.selectbox(
                    "Set New Status:", stat_options, index=cur_idx, key=f"s_{d_id}"
                )
                if st.button(
                    "Save Status",
                    key=f"b_up_{d_id}",
                    type="primary",
                    use_container_width=True,
                ):
                  update_assignment_status(d_id, new_s)
                  st.success("Status updated!")
                  time.sleep(0.5)
                  st.rerun()

  # =========================================================
  # TAB 2: ASSIGNMENT DESK (FOLDER SCANNER & DISPATCHER)
  # =========================================================
  with tab_assign:
    st.subheader("🗂️ Drive Document Dispatcher")
    st.info(
        "💡 **Assignment Rule:** Assigning a Translator is optional. Leaving"
        " Translator on **Bypass** routes the document through the AI direct"
        " draft engine and lands it straight into the Reviewer's inbox."
    )

    t_options = ["[Optional] Bypass - AI Only"] + [
        e
        for e, d in vols.items()
        if d.get("role") in ["translator", "admin"]
        and d.get("status", "").lower() == "active"
    ]
    r_options = ["[Mandatory] Select Reviewer..."] + [
        e
        for e, d in vols.items()
        if d.get("role") in ["reviewer", "admin"]
        and d.get("status", "").lower() == "active"
    ]

    assigned_map = {a.get("doc_id"): a for a in all_assignments}

    try:
      folder_res = (
          drive_service.files()
          .list(
              q="mimeType='application/vnd.google-apps.folder' and trashed=false",
              fields="files(id, name)",
          )
          .execute()
      )
      all_folders = folder_res.get("files", [])
      year_folders = [
          f
          for f in all_folders
          if "Edition" in f["name"] or "202" in f["name"]
      ]

      if year_folders:
        col_yr, col_mo, col_scan = st.columns([2, 2, 1])
        year_options = {f["name"]: f["id"] for f in year_folders}
        sel_year = col_yr.selectbox(
            "📁 1. Year Folder:", list(year_options.keys())
        )
        year_id = year_options[sel_year]

        month_res = (
            drive_service.files()
            .list(
                q=(
                    f"'{year_id}' in parents and"
                    " mimeType='application/vnd.google-apps.folder' and"
                    " trashed=false"
                ),
                fields="files(id, name)",
            )
            .execute()
        )
        month_folders = month_res.get("files", [])

        if month_folders:
          month_options = {f["name"]: f["id"] for f in month_folders}
          sel_month = col_mo.selectbox(
              "📂 2. Month Folder:", list(month_options.keys())
          )
          month_id = month_options[sel_month]

          col_scan.write("")
          col_scan.write("")
          if col_scan.button(
              "🔄 Scan Folder", type="primary", use_container_width=True
          ):
            st.session_state["scanned_month_id"] = month_id
        else:
          st.warning("No month subfolders found inside the selected year.")
      else:
        st.warning("No Year folders found matching 'Edition' or '202*'.")

      st.divider()

      if "scanned_month_id" in st.session_state:
        doc_query = (
            f"'{st.session_state['scanned_month_id']}' in parents and"
            " mimeType='application/vnd.google-apps.document' and trashed=false"
        )
        doc_res = (
            drive_service.files()
            .list(q=doc_query, fields="files(id, name)")
            .execute()
        )
        docs = doc_res.get("files", [])

        if docs:
          st.markdown(f"### 📋 Files in Folder (`{len(docs)}` found)")

          for doc in docs:
            doc_id = doc.get("id")
            doc_name = doc.get("name")
            prior_task = assigned_map.get(doc_id)

            with st.container(border=True):
              if prior_task:
                # State: Already Assigned
                cur_st = prior_task.get("status")
                badge_icon = (
                    "🟡"
                    if cur_st == "Pending Translation"
                    else ("🔵" if cur_st == "Pending Review" else "🟢")
                )
                t_lbl = (
                    format_vol_label(prior_task.get("translator"))
                    if prior_task.get("translator")
                    else "🤖 AI Bypass"
                )
                r_lbl = format_vol_label(prior_task.get("reviewer"))

                c_info, c_badge, c_re = st.columns([3, 1.2, 1])
                with c_info:
                  st.markdown(f"📄 **{doc_name}**")
                  st.caption(f"Assigned: **{t_lbl}** ➔ **{r_lbl}**")
                with c_badge:
                  st.markdown(f"Status: `{badge_icon} {cur_st}`")
                with c_re:
                  with st.popover("🔄 Reassign", use_container_width=True):
                    re_t = st.selectbox(
                        "New Translator:",
                        t_options,
                        format_func=format_vol_label,
                        key=f"ret_{doc_id}",
                    )
                    re_r = st.selectbox(
                        "New Reviewer:",
                        r_options,
                        format_func=format_vol_label,
                        key=f"rer_{doc_id}",
                    )
                    if st.button(
                        "Confirm",
                        key=f"btn_re_{doc_id}",
                        type="primary",
                        use_container_width=True,
                    ):
                      if re_r == r_options[0]:
                        st.error("Select a reviewer.")
                      else:
                        f_t = "" if re_t == t_options[0] else re_t
                        init_s = (
                            "Pending Review"
                            if not f_t
                            else "Pending Translation"
                        )
                        assign_task_to_sheet(doc_id, doc_name, f_t, re_r, init_s)
                        st.success("Reassigned!")
                        time.sleep(0.5)
                        st.rerun()
              else:
                # State: Unassigned
                cn, ct, cr, cb = st.columns([3, 2, 2, 1.2])
                with cn:
                  st.markdown(f"📄 **{doc_name}**")
                  st.caption("⚪ *Unassigned*")
                with ct:
                  t_sel = st.selectbox(
                      "Translator:",
                      t_options,
                      format_func=format_vol_label,
                      key=f"t_{doc_id}",
                      label_visibility="collapsed",
                  )
                with cr:
                  r_sel = st.selectbox(
                      "Reviewer:",
                      r_options,
                      format_func=format_vol_label,
                      key=f"r_{doc_id}",
                      label_visibility="collapsed",
                  )
                with cb:
                  if st.button(
                      "🚀 Assign",
                      key=f"btn_{doc_id}",
                      type="primary",
                      use_container_width=True,
                  ):
                    if r_sel == r_options[0]:
                      st.error("Select a reviewer!")
                    else:
                      final_t = "" if t_sel == t_options[0] else t_sel
                      init_status = (
                          "Pending Review"
                          if not final_t
                          else "Pending Translation"
                      )
                      assign_task_to_sheet(
                          doc_id, doc_name, final_t, r_sel, init_status
                      )
                      st.success("Assigned!")
                      time.sleep(0.5)
                      st.rerun()
        else:
          st.info("No Google Docs found in this month folder.")
    except Exception as e:
      st.error(f"Google Drive Error: {e}")

  # =========================================================
  # TAB 3: VOLUNTEERS DATABASE
  # =========================================================
  with tab_vols:
    st.subheader("Manage Volunteer Access")
    try:
      res = (
          sheets_service.spreadsheets()
          .values()
          .get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=VOLUNTEERS_RANGE)
          .execute()
      )
      vol_data = res.get("values", [])
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
                  options=["translator", "reviewer", "admin"],
                  required=True,
              ),
              "Status": st.column_config.SelectboxColumn(
                  "Status", options=["Active", "Suspended"], required=True
              ),
          },
      )
      if st.button("💾 Save Volunteers", type="primary"):
        overwrite_sheet_data(
            VOLUNTEERS_RANGE,
            [edited_vol.columns.tolist()] + edited_vol.fillna("").values.tolist(),
        )
        st.success("Volunteers database updated!")
    except Exception as e:
      st.error(f"Error fetching volunteers: {e}")

  # =========================================================
  # TAB 4: GLOSSARY COMMAND CENTER
  # =========================================================
  with tab_glos:
    st.subheader("Live Terminology Editor")
    try:
      res = (
          sheets_service.spreadsheets()
          .values()
          .get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_RANGE)
          .execute()
      )
      glos_data = res.get("values", [])
      headers = ["ID", "Category", "English", "Arabic"]
      rows_to_display = (
          glos_data[1:]
          if (glos_data and len(glos_data[0]) == 4 and not glos_data[0][0].isdigit())
          else glos_data
      )
      cleaned_rows = [
          list(row) + [""] * (4 - len(row)) for row in rows_to_display
      ]
      if not cleaned_rows:
        cleaned_rows = [["1", "General", "", ""]]
      df_glos = pd.DataFrame(cleaned_rows, columns=headers)
      edited_glos = st.data_editor(
          df_glos, num_rows="dynamic", use_container_width=True
      )
      if st.button("💾 Sync Glossary to AI", type="primary"):
        overwrite_sheet_data(
            GLOSSARY_RANGE,
            [edited_glos.columns.tolist()]
            + edited_glos.fillna("").values.tolist(),
        )
        st.success("Glossary synced successfully!")
        st.cache_data.clear()
    except Exception as e:
      st.error(f"Error fetching glossary: {e}")

  # =========================================================
  # TAB 5: LIVE RADAR (CONCURRENCY LOCKS)
  # =========================================================
  with tab_radar:
    st.subheader("Active Document Locks")
    try:
      res = (
          sheets_service.spreadsheets()
          .values()
          .get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE)
          .execute()
      )
      session_rows = res.get("values", [])
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
                st.markdown(
                    f"🔒 Locked by **{locked_by}** for `{time_locked} minutes`"
                )
                st.caption(f"Doc ID: {doc_id}")
              with col_l2:
                if st.button(
                    "🧨 Kill Lock", key=f"kill_{idx}", use_container_width=True
                ):
                  sheets_service.spreadsheets().values().update(
                      spreadsheetId=GLOSSARY_SPREADSHEET_ID,
                      range=f"'Sessions'!A{idx+2}:D{idx+2}",
                      valueInputOption="USER_ENTERED",
                      body={"values": [["", "", "", ""]]},
                  ).execute()
                  st.rerun()
      if not active_locks:
        st.info("No active sessions right now. All clear!")
    except Exception as e:
      st.error(f"Error fetching radar: {e}")

  # =========================================================
  # TAB 6: BROADCAST SYSTEM
  # =========================================================
  with tab_bcast:
    st.subheader("Team Broadcast System")
    st.caption("Send a mass email to all 'Active' volunteers.")
    broadcast_subject = st.text_input("Subject")
    broadcast_message = st.text_area("Message Body", height=150)

    if st.button("🚀 Send Broadcast", type="primary"):
      if broadcast_subject and broadcast_message:
        active_emails = [
            email
            for email, d in vols.items()
            if d.get("status", "").lower() == "active"
        ]
        if active_emails:
          with st.spinner("Dispatching emails via secure SMTP..."):
            try:
              msg = EmailMessage()
              msg.set_content(
                  f"12-Step Translation Project Update:\n\n{broadcast_message}"
              )
              msg["Subject"] = f"[12-Step Admin] {broadcast_subject}"
              msg["From"] = st.secrets.get("SMTP_EMAIL", "admin@localhost")
              msg["To"] = st.secrets.get("SMTP_EMAIL", "admin@localhost")
              msg["Bcc"] = ", ".join(active_emails)

              with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(
                    st.secrets["SMTP_EMAIL"], st.secrets["SMTP_PASSWORD"]
                )
                server.send_message(msg)
              st.success(
                  f"Broadcast sent successfully to {len(active_emails)}"
                  " volunteers!"
              )
              st.balloons()
            except Exception as e:
              st.error(
                  "Failed to send email. Check your SMTP secrets. Error:"
                  f" {e}"
              )
        else:
          st.warning("No active volunteers found to email.")
      else:
        st.warning("Please enter a subject and a message.")

  st.stop()


# ==========================================
# 4. USER INBOX & TASK DELEGATION
# ==========================================
if not st.session_state.get("source_file_id"):
  col_t, col_l = st.columns([5, 1])
  col_t.title("⚙️ 12-Step AI Suite")
  if col_l.button("🚪 Logout", use_container_width=True):
    st.session_state.clear()
    st.rerun()

  st.subheader(
      f"👋 Welcome, {st.session_state.get('user_name')} | Role:"
      f" {st.session_state.get('user_role', '').capitalize()}"
  )
  st.markdown("---")

  st.markdown("### 📬 Your Task Queue")
  assignments = fetch_assignments()
  my_tasks = []

  for task in assignments:
    if (
        st.session_state.get("app_mode") == "Translator Mode"
        and task.get("translator") == st.session_state.get("user_email")
        and task.get("status") == "Pending Translation"
    ):
      my_tasks.append(task)
    elif (
        st.session_state.get("app_mode") == "Reviewer Mode"
        and task.get("reviewer") == st.session_state.get("user_email")
        and task.get("status") == "Pending Review"
    ):
      my_tasks.append(task)

  if not my_tasks:
    st.success("🎉 You have no pending tasks in your queue. Great job!")
  else:
    for task in my_tasks:
      with st.container(border=True):
        c1, c2 = st.columns([4, 1])
        c1.markdown(f"### 📄 **{task.get('doc_name')}**")
        c1.caption(
            f"Status: `{task.get('status')}` | ID: {task.get('doc_id')}"
        )
        if c2.button(
            "🚀 Start Work",
            key=f"start_{task.get('doc_id')}",
            type="primary",
            use_container_width=True,
        ):
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


class TranslationResult(BaseModel):
  arabic_translation: str = Field(
      description="The finalized Arabic translation"
  )
  glossary_notes: str = Field(
      description=(
          "List of exact glossary terms applied with explanations, or 'None'"
      )
  )


class TranslationBatchItem(BaseModel):
  id: int
  arabic_translation: str
  glossary_notes: str


class TranslationBatchResult(BaseModel):
  items: list[TranslationBatchItem]


class ReviewResult(BaseModel):
  status: str = Field(
      description="Must be 'perfect', 'minor_edits', or 'major_rewrite'"
  )
  suggested_arabic: str = Field(description="The finalized Arabic text")
  reasoning: str = Field(
      description=(
          "Detailed explanation of changes based on semantics and glossary"
      )
  )


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
  if secret_model:
    return [secret_model]
  if GENAI_AVAILABLE and client is not None:
    try:
      available_flash_models = []
      for m in client.models.list():
        clean_name = m.name.replace("models/", "")
        if "flash" in clean_name.lower() and not any(
            tag in clean_name.lower() for tag in ["legacy", "embed", "imagen"]
        ):
          available_flash_models.append(clean_name)
      available_flash_models.sort(reverse=True)
      if available_flash_models:
        return available_flash_models
    except Exception:
      pass
  return ["gemini-2.5-flash", "gemini-1.5-flash"]


def manage_document_lock(file_id: str, user_email: str):
  try:
    res = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=SESSIONS_RANGE)
        .execute()
    )
    rows = res.get("values", [])
    current_time = time.time()
    target_row_index = max(len(rows) + 1, 2)

    for index, row in enumerate(rows):
      if index == 0:
        continue
      if len(row) > 0 and row[0] == file_id:
        locked_by = row[1] if len(row) > 1 else ""
        ts = float(row[2]) if len(row) > 2 and row[2] else 0
        saved_json = row[3] if len(row) > 3 else ""

        if (
            locked_by
            and locked_by != user_email
            and (current_time - ts) < LOCK_TIMEOUT_SECONDS
        ):
          return {
              "status": "blocked",
              "locked_by": locked_by,
              "row_index": index + 1,
          }
        if locked_by == user_email and saved_json:
          try:
            return {
                "status": "recovered",
                "data": json.loads(saved_json),
                "row_index": index + 1,
            }
          except Exception:
            pass
        return {"status": "clear", "row_index": index + 1}
    return {"status": "clear", "row_index": target_row_index}
  except Exception:
    return {"status": "clear", "row_index": 2}


def save_draft_to_sheet(file_id, user_email, row_index, processed_data):
  if not row_index:
    return
  try:
    json_data = json.dumps(processed_data, ensure_ascii=False)
    body = {"values": [[file_id, user_email, str(time.time()), json_data]]}
    sheets_service.spreadsheets().values().update(
        spreadsheetId=GLOSSARY_SPREADSHEET_ID,
        range=f"'Sessions'!A{row_index}:D{row_index}",
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()
  except Exception:
    pass


def release_document_lock(row_index):
  if not row_index:
    return
  try:
    sheets_service.spreadsheets().values().update(
        spreadsheetId=GLOSSARY_SPREADSHEET_ID,
        range=f"'Sessions'!A{row_index}:D{row_index}",
        valueInputOption="USER_ENTERED",
        body={"values": [["", "", "", ""]]},
    ).execute()
  except Exception:
    pass


@st.cache_data(ttl=3600)
def fetch_glossary():
  try:
    res = (
        sheets_service.spreadsheets()
        .values()
        .get(
            spreadsheetId=GLOSSARY_SPREADSHEET_ID, range=GLOSSARY_DATA_RANGE
        )
        .execute()
    )
    vals = res.get("values", [])
    glos_lines = []
    count = 0
    for row in vals:
      if len(row) >= 2 and row[0].strip() and row[1].strip():
        if row[0].strip().lower() == "english":
          continue
        glos_lines.append(f"- {row[0].strip()} -> {row[1].strip()}")
        count += 1
    return "\n".join(glos_lines), count
  except Exception:
    return "", 0


def generate_html_diff(original, suggested):
  if not original or original.startswith("[MISSING"):
    return (
        "<div dir='rtl' style='text-align: right; color: #0369a1;"
        " background-color: #e0f2fe; padding: 10px; border-radius: 5px;"
        " font-family: \"Cairo\";'>✨ Initial AI Draft</div>"
    )
  if original.strip() == suggested.strip():
    return (
        "<div dir='rtl' style='text-align: right; color: #155724;"
        " background-color: #d4edda; padding: 10px; border-radius: 5px;"
        " font-family: \"Cairo\";'>✨ Perfect Match — No Edits Needed</div>"
    )

  diff = difflib.ndiff(original.split(), suggested.split())
  html = [
      "<div dir='rtl' style='font-family: \"Cairo\", sans-serif; font-size:"
      " 18px; line-height: 2; text-align: right; background-color: #f8f9fa;"
      " padding: 15px; border-radius: 8px; border: 1px solid #e9ecef;'>"
  ]
  for word in diff:
    if word.startswith("- "):
      html.append(
          "<span style='background-color: #ffcdd2; color: #b71c1c;"
          " text-decoration: line-through; padding: 2px; border-radius:"
          f" 4px;'>{word[2:]}</span>"
      )
    elif word.startswith("+ "):
      html.append(
          "<span style='background-color: #c8e6c9; color: #1b5e20; font-weight:"
          f" bold; padding: 2px; border-radius: 4px;'>{word[2:]}</span>"
      )
    elif word.startswith("  "):
      html.append(f"<span style='color: #212529;'>{word[2:]}</span>")
  html.append("</div>")
  return " ".join(html)


def _is_retryable(err_str):
  return any(kw.lower() in err_str.lower() for kw in RETRYABLE_KEYWORDS)


def _backoff_sleep(attempt):
  time.sleep(min(BASE_BACKOFF_SECONDS * (2**attempt), MAX_BACKOFF_SECONDS))


def _call_gemini(model_name, prompt, schema_type):
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
  clean_text = response.text.replace("```json", "").replace("```", "").strip()
  match = re.search(r"\{.*\}", clean_text, re.DOTALL)
  if match:
    clean_text = match.group(0)
  return json.loads(clean_text)


def translate_with_ai(english: str, glossary_text: str):
  prompt = f"""You are an expert bilingual translator specializing in 12-step recovery literature. 
Translate the English text into Arabic accurately, clinical, professional, and non-moralizing.

MANDATORY GLOSSARY INSTRUCTIONS:
- Below is the official terminology glossary.
- Whenever an English term appears in the text that exists in the glossary, you MUST translate it using the exact Arabic equivalent specified. NEVER substitute it with creative synonyms.
- In 'glossary_notes', explicitly list every glossary term matched and applied (e.g., "Matched: recovery -> تعافي"). If none, write "None".
- When pronouns like "it" refer to recovery concepts ("the program"), render the Arabic noun and grammatical gender accurately.

OFFICIAL GLOSSARY:
{glossary_text}

Translate:
English Source: "{english}"
"""
  for model_name in get_fallback_models():
    for attempt in range(MAX_RETRIES_PER_MODEL):
      try:
        parsed = _call_gemini(model_name, prompt, TranslationResult)
        return {
            "arabic_translation": parsed.get("arabic_translation", ""),
            "glossary_notes": parsed.get("glossary_notes", ""),
        }
      except Exception as e:
        if _is_retryable(str(e)):
          _backoff_sleep(attempt)
  return {"arabic_translation": "", "glossary_notes": "⚠ Error"}


def review_with_ai(english: str, arabic: str, glossary_text: str):
  prompt = f"""You are an expert bilingual editor specializing in 12-step recovery literature. 
Review and audit the Arabic translation against the English source. Ensure clinical, non-moralizing tone.

MANDATORY GLOSSARY AUDIT:
- Verify that recovery terminology strictly conforms to the official glossary below.
- If the original Arabic used a wrong term, correct it to the glossary term and explain the correction in 'reasoning'.
- If the translation is accurate and natural, keep it as is and mark status as 'perfect'.

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
        return {
            "status": parsed.get("status", "minor_edits"),
            "suggested_arabic": parsed.get("suggested_arabic", arabic),
            "reasoning": parsed.get("reasoning", ""),
        }
      except Exception as e:
        if _is_retryable(str(e)):
          _backoff_sleep(attempt)
  return {
      "status": "major_rewrite",
      "suggested_arabic": arabic,
      "reasoning": "⚠ Error",
  }


def translate_batch_with_fallback(batch_segments, glossary_text):
  if not GENAI_AVAILABLE or client is None:
    return [
        translate_with_ai(s.get("english", ""), glossary_text)
        for s in batch_segments
    ]
  input_payload = "\n\n".join(
      [f"ID: {s.get('id', 0)}\nText: {s.get('english', '')}" for s in batch_segments]
  )

  prompt = f"""You are an expert bilingual translator specializing in 12-step recovery literature. 
Translate the following English segments into Arabic accurately. The tone must remain clinical, professional, and non-moralizing.

MANDATORY GLOSSARY & BATCH INSTRUCTIONS:
1. STRICT GLOSSARY ADHERENCE: You MUST use the exact Arabic terms provided in the official glossary below whenever an English term appears. Never substitute them with synonyms.
2. GLOSSARY NOTES: In the 'glossary_notes' field for each segment, state explicitly which glossary terms you matched and used (e.g., "recovery -> تعافي, fellowship -> زمالة"). If no glossary terms were present, state "None".
3. Narrative Flow: Maintain consistent grammatical gender, tone, and pronoun references across all segments.
4. Pronoun Resolution: Reflect the correct contextual noun and proper Arabic grammatical gender.

OFFICIAL GLOSSARY:
{glossary_text}

Segments to Translate:
{input_payload}"""

  for model_name in get_fallback_models():
    for attempt in range(2):
      try:
        parsed = _call_gemini(model_name, prompt, TranslationBatchResult)
        items = parsed.get("items", [])
        if len(items) == len(batch_segments):
          return items
      except Exception as e:
        if _is_retryable(str(e)):
          _backoff_sleep(attempt)

  results = []
  for s in batch_segments:
    res = translate_with_ai(s.get("english", ""), glossary_text)
    results.append({
        "id": s.get("id", 0),
        "arabic_translation": res.get("arabic_translation", ""),
        "glossary_notes": res.get("glossary_notes", ""),
    })
  return results


def review_batch_with_fallback(batch_segments, glossary_text):
  if not GENAI_AVAILABLE or client is None:
    return [
        review_with_ai(
            s.get("english", ""), s.get("arabic", ""), glossary_text
        )
        for s in batch_segments
    ]
  input_payload = "\n\n".join([
      f"ID: {s.get('id', 0)}\nEnglish: {s.get('english', '')}\nArabic:"
      f" {s.get('arabic', '')}"
      for s in batch_segments
  ])

  prompt = f"""You are an expert bilingual editor specializing in 12-step recovery literature. 
Review the following English/Arabic pairs for accuracy, tone, and strict glossary adherence.

MANDATORY EDITING RULES:
1. Glossary Enforcement: Ensure terms strictly match the official glossary below. If the Arabic translation deviated from the glossary, correct it to the exact glossary term.
2. Status Assignment:
   - 'perfect': The translation is completely accurate, adheres to glossary, and flows naturally.
   - 'minor_edits': Small improvements to grammar, syntax, or term corrections.
   - 'major_rewrite': Mistranslated, missed core meaning, or severely awkward phrasing.
3. Detailed Reasoning: Explain clearly in 'reasoning' why any corrections were made.

OFFICIAL GLOSSARY:
{glossary_text}

Pairs to Review:
{input_payload}"""

  for model_name in get_fallback_models():
    for attempt in range(2):
      try:
        parsed = _call_gemini(model_name, prompt, ReviewBatchResult)
        items = parsed.get("items", [])
        if len(items) == len(batch_segments):
          return items
      except Exception as e:
        if _is_retryable(str(e)):
          _backoff_sleep(attempt)

  results = []
  for s in batch_segments:
    res = review_with_ai(
        s.get("english", ""), s.get("arabic", ""), glossary_text
    )
    results.append({
        "id": s.get("id", 0),
        "status": res.get("status", "minor_edits"),
        "suggested_arabic": res.get("suggested_arabic", ""),
        "reasoning": res.get("reasoning", ""),
    })
  return results


def _parse_docs_elements(elements):
  paras = []
  for elem in elements:
    if "paragraph" in elem:
      para_text = ""
      start_idx = elem.get("startIndex")
      end_idx = elem.get("endIndex")
      for run in elem.get("paragraph", {}).get("elements", []):
        if "textRun" in run:
          para_text += run.get("textRun", {}).get("content", "")
      clean_text = para_text.strip()
      if bool(re.search(r"[a-zA-Z\u0600-\u06FF]", clean_text)):
        paras.append({"text": clean_text, "start": start_idx, "end": end_idx})
    elif "table" in elem:
      for row in elem.get("table", {}).get("tableRows", []):
        for cell in row.get("tableCells", []):
          paras.extend(_parse_docs_elements(cell.get("content", [])))
  return paras


def extract_text_from_drive(file_id: str, is_retry=False):
  try:
    docs_svc, drive_svc, _ = get_google_services()
    file_meta = drive_svc.files().get(fileId=file_id, fields="mimeType").execute()
    mime_type = file_meta.get("mimeType")

    if mime_type == "application/vnd.google-apps.document":
      try:
        document = (
            docs_svc.documents()
            .get(documentId=file_id, includeTabsContent=True)
            .execute()
        )
      except Exception:
        document = docs_svc.documents().get(documentId=file_id).execute()

      all_paras = []

      def sweep_doc_obj(doc_obj):
        temp_paras = []
        temp_paras.extend(
            _parse_docs_elements(doc_obj.get("body", {}).get("content", []))
        )
        for footer in doc_obj.get("footers", {}).values():
          temp_paras.extend(_parse_docs_elements(footer.get("content", [])))
        for header in doc_obj.get("headers", {}).values():
          temp_paras.extend(_parse_docs_elements(header.get("content", [])))
        return temp_paras

      tabs = document.get("tabs", [])
      if tabs:
        for tab in tabs:
          all_paras.extend(sweep_doc_obj(tab.get("documentTab", {})))
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


def smart_align(paragraphs):
  en_paras = [
      p
      for p in paragraphs
      if not re.search(r"[\u0600-\u06FF]", p.get("text", ""))
  ]
  ar_paras = [
      p for p in paragraphs if re.search(r"[\u0600-\u06FF]", p.get("text", ""))
  ]
  aligned = []
  for i in range(max(len(en_paras), len(ar_paras))):
    en_obj = (
        en_paras[i] if i < len(en_paras) else {"text": "[MISSING ENGLISH SOURCE]"}
    )
    ar_obj = (
        ar_paras[i]
        if i < len(ar_paras)
        else {"text": "[MISSING ARABIC TRANSLATION]", "start": None, "end": None}
    )
    aligned.append({
        "id": i + 1,
        "english": en_obj.get("text", ""),
        "arabic": ar_obj.get("text", ""),
        "ar_start": ar_obj.get("start"),
        "ar_end": ar_obj.get("end"),
    })
  return aligned


def push_to_drive_translator(document_id, final_arabic_text):
  docs_svc, _, _ = get_google_services()
  try:
    requests = [
        {"insertPageBreak": {"location": {"index": 1}}},
        {
            "insertText": {
                "location": {"index": 1},
                "text": final_arabic_text + "\n\n",
            }
        },
    ]
    docs_svc.documents().batchUpdate(
        documentId=document_id, body={"requests": requests}
    ).execute()
    return True
  except Exception as e:
    st.error(f"Failed to push to Drive. Error: {e}")
    return False


def push_to_drive_reviewer(document_id, approved_segments):
  docs_svc, _, _ = get_google_services()
  try:
    valid_segments = [
        seg
        for seg in approved_segments
        if seg.get("ar_start") is not None and seg.get("ar_end") is not None
    ]
    valid_segments.sort(key=lambda x: x["ar_start"], reverse=True)
    requests = []
    for seg in valid_segments:
      requests.append({
          "deleteContentRange": {
              "range": {
                  "startIndex": seg["ar_start"],
                  "endIndex": seg["ar_end"] - 1,
              }
          }
      })
      requests.append({
          "insertText": {
              "location": {"index": seg["ar_start"]},
              "text": seg.get("final_arabic", ""),
          }
      })

    if requests:
      docs_svc.documents().batchUpdate(
          documentId=document_id, body={"requests": requests}
      ).execute()
    return True
  except Exception as e:
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
glossary_data, glossary_term_count = fetch_glossary()

col_h1, col_h2 = st.columns([4, 2])
col_h1.markdown(f"## 📝 Workspace: `{task.get('doc_name', 'Document')}`")

with col_h2:
  c_btn, c_glos = st.columns([1, 1])
  if c_btn.button("⬅️ Back to Inbox", use_container_width=True):
    st.session_state["active_task"] = None
    st.session_state["source_file_id"] = None
    st.session_state["processed_data"] = None
    st.rerun()
  with c_glos:
    if glossary_term_count > 0:
      st.success(f"📖 Glossary: {glossary_term_count} terms")
    else:
      st.warning("⚠️ Glossary: Not Loaded")

is_bypass_task = task.get("translator") == ""

if not st.session_state.get("processed_data"):
  lock = manage_document_lock(file_id, st.session_state.get("user_email"))
  if lock["status"] == "blocked":
    st.error(
        "🛑 **Document In Use:** This document is currently locked by"
        f" `{lock.get('locked_by', 'another user')}`."
    )
    st.stop()
  elif lock["status"] == "recovered":
    st.session_state["session_row_index"] = lock["row_index"]
    st.session_state["processed_data"] = lock["data"]
    st.success("♻️ **Session Recovered.** Restored your previous work.")
  elif lock["status"] == "clear":
    st.session_state["session_row_index"] = lock["row_index"]

    paras = extract_text_from_drive(file_id)

    if paras:
      processed_results = []
      progress_bar = st.progress(0)

      # --- CASE A: TRANSLATOR MODE ---
      if st.session_state.get("app_mode") == "Translator Mode":
        st.info("Extracting document content and translating via AI Batching...")
        batches = [
            paras[i : i + BATCH_SIZE] for i in range(0, len(paras), BATCH_SIZE)
        ]
        for idx, batch in enumerate(batches):
          batch_payload = [
              {"id": j + 1, "english": p.get("text", "")}
              for j, p in enumerate(batch)
          ]
          ai_results = translate_batch_with_fallback(
              batch_payload, glossary_data
          )

          for ai_res in ai_results:
            ai_id = int(ai_res.get("id", 1))
            array_index = ai_id - 1
            eng_text = (
                batch[array_index].get("text", "")
                if 0 <= array_index < len(batch)
                else batch[0].get("text", "")
            )

            processed_results.append({
                "id": ai_id + (idx * BATCH_SIZE),
                "english": eng_text,
                "arabic_translation": ai_res.get("arabic_translation", ""),
                "glossary_notes": ai_res.get("glossary_notes", ""),
                "user_arabic": ai_res.get("arabic_translation", ""),
                "is_approved": False,
            })
          progress_bar.progress((idx + 1) / len(batches))

      # --- CASE B: REVIEWER MODE ---
      else:
        has_arabic = any(
            re.search(r"[\u0600-\u06FF]", p.get("text", "")) for p in paras
        )

        if not has_arabic or is_bypass_task:
          st.info(
              "🤖 **AI Bypass Mode:** Generating initial draft and analyzing"
              " narrative flow for your review..."
          )
          batches = [
              paras[i : i + BATCH_SIZE]
              for i in range(0, len(paras), BATCH_SIZE)
          ]

          # Step 1: Translate draft
          draft_segments = []
          for idx, batch in enumerate(batches):
            batch_payload = [
                {"id": j + 1, "english": p.get("text", "")}
                for j, p in enumerate(batch)
            ]
            ai_trans = translate_batch_with_fallback(
                batch_payload, glossary_data
            )
            for ai_t in ai_trans:
              ai_id = int(ai_t.get("id", 1))
              array_index = ai_id - 1
              eng_text = (
                  batch[array_index].get("text", "")
                  if 0 <= array_index < len(batch)
                  else batch[0].get("text", "")
              )
              draft_segments.append({
                  "id": ai_id + (idx * BATCH_SIZE),
                  "english": eng_text,
                  "arabic": ai_t.get("arabic_translation", ""),
                  "glossary_notes": ai_t.get("glossary_notes", ""),
              })

          # Step 2: Review draft to generate diff highlights
          rev_batches = [
              draft_segments[i : i + BATCH_SIZE]
              for i in range(0, len(draft_segments), BATCH_SIZE)
          ]
          for idx, r_batch in enumerate(rev_batches):
            ai_revs = review_batch_with_fallback(r_batch, glossary_data)
            for ai_r, d_seg in zip(ai_revs, r_batch):
              suggested = ai_r.get("suggested_arabic", d_seg.get("arabic", ""))
              status_val = ai_r.get("status", "minor_edits")
              processed_results.append({
                  "id": d_seg.get("id"),
                  "status": status_val,
                  "english": d_seg.get("english", ""),
                  "original_arabic": d_seg.get("arabic", ""),
                  "suggested_arabic": suggested,
                  "reasoning": (
                      ai_r.get("reasoning", "")
                      + (
                          f" | Glossary: {d_seg.get('glossary_notes')}"
                          if d_seg.get("glossary_notes")
                          else ""
                      )
                  ),
                  "ar_start": None,
                  "ar_end": None,
                  "user_arabic": suggested,
                  "is_approved": (status_val == "perfect"),
              })
            progress_bar.progress((idx + 1) / len(rev_batches))

        else:
          st.info(
              "Extracting segments and comparing human translation against AI"
              " audit..."
          )
          segments = smart_align(paras)
          normal_segs = [
              s
              for s in segments
              if s.get("english") != "[MISSING ENGLISH SOURCE]"
              and s.get("arabic") != "[MISSING ARABIC TRANSLATION]"
          ]
          batches = [
              normal_segs[i : i + BATCH_SIZE]
              for i in range(0, len(normal_segs), BATCH_SIZE)
          ]

          for idx, batch in enumerate(batches):
            ai_results = review_batch_with_fallback(batch, glossary_data)
            for ai_res, o_seg in zip(ai_results, batch):
              suggested = ai_res.get("suggested_arabic", o_seg.get("arabic", ""))
              status_val = ai_res.get("status", "minor_edits")
              processed_results.append({
                  "id": o_seg.get("id"),
                  "status": status_val,
                  "english": o_seg.get("english", ""),
                  "original_arabic": o_seg.get("arabic", ""),
                  "suggested_arabic": suggested,
                  "reasoning": ai_res.get("reasoning", ""),
                  "ar_start": o_seg.get("ar_start"),
                  "ar_end": o_seg.get("ar_end"),
                  "user_arabic": suggested,
                  "is_approved": (status_val == "perfect"),
              })
            progress_bar.progress((idx + 1) / len(batches))

          for item in segments:
            if item.get("english") == "[MISSING ENGLISH SOURCE]":
              processed_results.append({
                  "id": item.get("id"),
                  "status": "major_rewrite",
                  "english": "[MISSING]",
                  "original_arabic": item.get("arabic", ""),
                  "suggested_arabic": item.get("arabic", ""),
                  "reasoning": "⚠️ Orphaned Arabic block.",
                  "ar_start": item.get("ar_start"),
                  "ar_end": item.get("ar_end"),
                  "user_arabic": item.get("arabic"),
                  "is_approved": False,
              })
            elif item.get("arabic") == "[MISSING ARABIC TRANSLATION]":
              trans_res = translate_with_ai(
                  item.get("english", ""), glossary_data
              )
              t_arabic = trans_res.get("arabic_translation", "")
              processed_results.append({
                  "id": item.get("id"),
                  "status": "major_rewrite",
                  "english": item.get("english", ""),
                  "original_arabic": "[MISSING]",
                  "suggested_arabic": t_arabic,
                  "reasoning": "⚠️ Auto-translated orphaned English block.",
                  "ar_start": None,
                  "ar_end": None,
                  "user_arabic": t_arabic,
                  "is_approved": False,
              })

          processed_results.sort(key=lambda x: x.get("id", 0))

      save_draft_to_sheet(
          file_id,
          st.session_state.get("user_email"),
          st.session_state.get("session_row_index"),
          processed_results,
      )
      st.session_state["processed_data"] = processed_results
      st.rerun()

# --- EDITOR UI ---
approved_count, finalized_data, state_modified = 0, [], False
st.divider()

for i, item in enumerate(st.session_state.get("processed_data", [])):
  seg_id = item.get("id", i + 1)
  status_val = item.get("status", "minor_edits")
  eng_txt = item.get("english", "")
  orig_ar = item.get("original_arabic", "")
  sugg_ar = item.get(
      "suggested_arabic", item.get("arabic_translation", "")
  )

  with st.container(border=True):
    if st.session_state.get("app_mode") == "Translator Mode":
      st.markdown(f"### Segment {seg_id}")
      col_en, col_ar = st.columns(2)
      with col_en:
        st.info(eng_txt)
        if item.get("glossary_notes"):
          st.caption(f"💡 **Glossary Matched:** {item.get('glossary_notes')}")
      with col_ar:
        sugg_trans = item.get("arabic_translation", "")
        default_val = item.get("user_arabic", sugg_trans)
        final_text = st.text_area(
            "Final text",
            value=default_val,
            height=120,
            key=f"edit_{i}",
            label_visibility="collapsed",
        )
        if final_text != item.get("user_arabic"):
          item["user_arabic"] = final_text
          state_modified = True
    else:
      color = (
          "🟢"
          if status_val == "perfect"
          else ("🟡" if status_val == "minor_edits" else "🔴")
      )
      st.markdown(
          f"### Segment {seg_id} | Status: {color} {status_val.upper()}"
      )
      col_en, col_ar = st.columns(2)
      with col_en:
        st.info(eng_txt)
      with col_ar:
        st.markdown(
            generate_html_diff(orig_ar, sugg_ar), unsafe_allow_html=True
        )

        with st.expander("💡 View AI Reasoning & Glossary Audit"):
          st.markdown(
              item.get(
                  "reasoning",
                  item.get("glossary_notes", "No reasoning provided."),
              )
          )

        default_val = item.get("user_arabic", sugg_ar)
        final_text = st.text_area(
            "Final Output",
            value=default_val,
            height=120,
            key=f"edit_ar_{i}",
            label_visibility="collapsed",
        )
        if final_text != item.get("user_arabic"):
          item["user_arabic"] = final_text
          state_modified = True

    chk = st.checkbox(
        f"✅ Approve Segment {seg_id}",
        key=f"chk_{i}",
        value=item.get("is_approved", False),
    )
    if chk != item.get("is_approved"):
      item["is_approved"] = chk
      state_modified = True

    if chk:
      approved_count += 1
      if st.session_state.get("app_mode") == "Translator Mode":
        finalized_data.append(final_text)
      else:
        finalized_data.append({
            "final_arabic": final_text,
            "ar_start": item.get("ar_start"),
            "ar_end": item.get("ar_end"),
        })

if state_modified:
  save_draft_to_sheet(
      file_id,
      st.session_state.get("user_email"),
      st.session_state.get("session_row_index"),
      st.session_state.get("processed_data"),
  )

# --- SUBMISSION LOGIC ---
st.divider()
total_segments = len(st.session_state.get("processed_data", []))
st.write(f"### **Approved Segments: {approved_count} / {total_segments}**")

if approved_count == total_segments and total_segments > 0:
  st.success("🎉 All segments approved! Final review before pushing.")

  if "review_unlocked" not in st.session_state:
    st.session_state["review_unlocked"] = False

  if not st.session_state["review_unlocked"]:
    st.error(
        "🚨 **CRITICAL STEP:** Please verify the narrative flow before"
        " finalizing."
    )
    if st.button(
        "👀 I confirm the final text is correct", use_container_width=True
    ):
      st.session_state["review_unlocked"] = True
      st.rerun()
  else:
    if st.button(
        "🚀 Push to Drive & Close Task",
        type="primary",
        use_container_width=True,
    ):
      with st.spinner("Processing Drive updates and concluding workflow..."):
        if (
            st.session_state.get("app_mode") == "Translator Mode"
            or is_bypass_task
        ):
          if is_bypass_task:
            ar_compiled = "\n\n".join([
                item.get("final_arabic", "")
                if isinstance(item, dict)
                else str(item)
                for item in finalized_data
            ])
          else:
            ar_compiled = "\n\n".join([str(item) for item in finalized_data])

          success = push_to_drive_translator(file_id, ar_compiled)
          if success:
            new_status = (
                "Completed"
                if is_bypass_task
                or (task.get("translator") == task.get("reviewer"))
                else "Pending Review"
            )
            update_assignment_status(file_id, new_status)
        else:
          success = push_to_drive_reviewer(file_id, finalized_data)
          if success:
            update_assignment_status(file_id, "Completed")

        if success:
          release_document_lock(st.session_state.get("session_row_index"))
          st.session_state["active_task"] = None
          st.session_state["source_file_id"] = None
          st.session_state["processed_data"] = None
          st.session_state["review_unlocked"] = False
          st.balloons()
          st.success("Task completed successfully!")
          time.sleep(2)
          st.rerun()
