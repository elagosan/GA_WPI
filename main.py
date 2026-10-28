import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
import os
import io
import base64
import requests
from datetime import datetime

# ==============================================================================
# Page Configuration & Metadata
# ==============================================================================
st.set_page_config(
    page_title="WPI Criteria & Station Assessment Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# Custom CSS Styling: High-Contrast, Perfectly Matched Tabs & Modern UI
# ==============================================================================
st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
      html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      }
      .stApp {
        background-color: #f8fafc;
      }
      [data-testid="stSidebar"] {
        background: #0f172a !important;
        border-right: 1px solid #1e293b;
      }
      [data-testid="stSidebar"] * {
        color: #f1f5f9 !important;
      }
      [data-testid="stSidebar"] hr {
        border-color: #334155 !important;
      }
      .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2.5rem;
        max-width: 1480px;
      }

      /* -------------------------------------------------------------
         HIGH-CONTRAST TAB NAVIGATION (Crystal-clear visibility)
         ------------------------------------------------------------- */
      div[data-testid="stTabs"] {
        background: transparent;
        margin-top: 4px;
      }
      div[data-testid="stTabs"] [data-baseweb="tab-list"] {
        background-color: #ffffff !important;
        border-radius: 14px !important;
        padding: 8px 10px !important;
        border: 1.5px solid #cbd5e1 !important;
        box-shadow: 0 4px 16px rgba(15, 23, 42, 0.06) !important;
        gap: 10px !important;
        margin-bottom: 20px !important;
      }
      div[data-testid="stTabs"] button[role="tab"] {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
        font-weight: 750 !important;
        font-size: 14px !important;
        border-radius: 10px !important;
        padding: 10px 20px !important;
        border: 1.5px solid #e2e8f0 !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        cursor: pointer !important;
      }
      div[data-testid="stTabs"] button[role="tab"] p {
        color: #1e293b !important;
        font-weight: 750 !important;
        font-size: 14px !important;
        margin: 0 !important;
      }
      div[data-testid="stTabs"] button[role="tab"]:hover {
        background-color: #e2e8f0 !important;
        border-color: #94a3b8 !important;
        transform: translateY(-1px);
      }
      div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%) !important;
        color: #ffffff !important;
        border: 1.5px solid #1d4ed8 !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35) !important;
      }
      div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] p {
        color: #ffffff !important;
        font-weight: 800 !important;
      }

      /* -------------------------------------------------------------
         HERO BANNER
         ------------------------------------------------------------- */
      .hero-banner {
        padding: 22px 26px;
        border-radius: 16px;
        color: white;
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 45%, #2563eb 85%, #38bdf8 130%);
        box-shadow: 0 10px 25px -5px rgba(30, 64, 175, 0.22);
        margin-bottom: 18px;
        position: relative;
        overflow: hidden;
      }
      .hero-banner h1 {
        font-size: 26px;
        line-height: 1.25;
        margin: 0 0 4px 0;
        font-weight: 800;
        letter-spacing: -0.5px;
      }
      .hero-banner p {
        color: #e0f2fe;
        margin: 0;
        font-size: 13.5px;
        font-weight: 400;
      }
      .hero-tag {
        display: inline-block;
        text-transform: uppercase;
        letter-spacing: 1.6px;
        font-size: 10px;
        font-weight: 700;
        color: #93c5fd;
        background: rgba(147, 197, 253, 0.15);
        padding: 3px 9px;
        border-radius: 20px;
        margin-bottom: 8px;
        border: 1px solid rgba(147, 197, 253, 0.3);
      }

      /* -------------------------------------------------------------
         METRIC & PANEL CARDS
         ------------------------------------------------------------- */
      .metric-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 15px 18px;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        margin-bottom: 12px;
      }
      .metric-label {
        font-size: 11px;
        font-weight: 700;
        color: #64748b;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin-bottom: 5px;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      .metric-value {
        font-size: 27px;
        font-weight: 800;
        letter-spacing: -0.8px;
        line-height: 1.1;
      }
      .metric-sub {
        font-size: 11px;
        color: #64748b;
        margin-top: 5px;
      }

      /* Home column card containers */
      .home-panel {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 20px 22px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
        margin-bottom: 16px;
      }
      .home-panel-title {
        font-size: 17px;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 3px;
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .home-panel-sub {
        font-size: 12px;
        color: #64748b;
        margin-bottom: 14px;
      }

      /* Category Badges */
      .badge-closed {
        background: #dcfce7;
        color: #15803d;
        padding: 3px 9px;
        border-radius: 9999px;
        font-size: 11.5px;
        font-weight: 700;
      }
      .badge-prog {
        background: #fef3c7;
        color: #b45309;
        padding: 3px 9px;
        border-radius: 9999px;
        font-size: 11.5px;
        font-weight: 700;
      }
      .badge-affected {
        background: #dbeafe;
        color: #1e40af;
        padding: 3px 9px;
        border-radius: 9999px;
        font-size: 11.5px;
        font-weight: 700;
      }

      .footer-bar {
        text-align: center;
        color: #94a3b8;
        font-size: 11.5px;
        padding: 24px 0 10px;
        border-top: 1px solid #e2e8f0;
        margin-top: 24px;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# Helper Functions: Data Normalization & Resilient File Reading
# ==============================================================================
DEFAULT_EXCEL_PATH = "WPI Implementation at TVR - Time Line.xlsx"
GITHUB_REPO_OWNER = "elagosan"
GITHUB_REPO_NAME = "GA_WPI"
GITHUB_TARGET_FILE = "WPI Implementation at TVR - Time Line.xlsx"

def sync_excel_to_github(file_bytes, target_filename, token):
    """
    Auto-uploads newly uploaded Excel file to GitHub repository via GitHub REST API.
    """
    if not token:
        return False, "GitHub Token is not set."
    
    url = f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/contents/{target_filename}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    sha = None
    try:
        r_get = requests.get(url, headers=headers, timeout=10)
        if r_get.status_code == 200:
            sha = r_get.json().get("sha")
    except Exception as e:
        return False, f"GitHub check error: {e}"

    b64_content = base64.b64encode(file_bytes).decode("utf-8")
    payload = {
        "message": f"Auto-update Excel dataset: {target_filename}",
        "content": b64_content,
    }
    if sha:
        payload["sha"] = sha

    try:
        r_put = requests.put(url, headers=headers, json=payload, timeout=20)
        if r_put.status_code in [200, 201]:
            return True, "Successfully synced to GitHub!"
        else:
            err_msg = r_put.json().get("message", r_put.text)
            return False, f"GitHub API ({r_put.status_code}): {err_msg}"
    except Exception as e:
        return False, f"GitHub connection error: {e}"

STANDARD_CATEGORIES = [
    "Health and Safety",
    "Customer Experience",
    "Logistics",
    "Workplace Organization",
]

CATEGORY_PALETTE = {
    "Health and Safety": "#ef4444",
    "Customer Experience": "#6366f1",
    "Logistics": "#0284c7",
    "Workplace Organization": "#059669",
    "All Categories": "#2563eb",
}

STATUS_PALETTE = {
    "Closed": "#10b981",       # Emerald Green
    "In Progress": "#f59e0b",  # Amber
    "Affected": "#2563eb",     # Royal Blue
    "Overall": "#2563eb",
    "Completed": "#10b981",
    "Other": "#94a3b8",
}

def clean_column_name(col):
    if not isinstance(col, str):
        col = str(col)
    return col.strip().replace("\n", " ")

def normalize_category_value(val):
    if pd.isna(val) or val is None:
        return "Uncategorized"
    s = str(val).strip().upper()
    if "HEALTH" in s or "SAFETY" in s:
        return "Health and Safety"
    elif "CUSTOMER" in s or "EXPERI" in s:
        return "Customer Experience"
    elif "LOGISTIC" in s:
        return "Logistics"
    elif "WORK" in s or "ORGAN" in s or "WORKPLACE" in s:
        return "Workplace Organization"
    return str(val).strip()

def normalize_status_value(val):
    if pd.isna(val) or val is None:
        return "Unknown"
    s = str(val).strip().upper().replace(" ", "").replace("-", "").replace("_", "")
    if "CLOSE" in s or "COMPLET" in s or "DONE" in s:
        return "Closed"
    elif "PROG" in s or "OPEN" in s or "PEND" in s:
        return "In Progress"
    return str(val).strip().title()

def read_file_bytes_safe(file_path):
    """
    Safely reads file bytes even if Microsoft Excel or OneDrive has the file open.
    Uses Windows win32file sharing flags as fallback to bypass exclusive locking.
    """
    try:
        with open(file_path, "rb") as f:
            return f.read(), None
    except PermissionError:
        try:
            import win32file
            handle = win32file.CreateFile(
                file_path,
                win32file.GENERIC_READ,
                win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE,
                None,
                win32file.OPEN_EXISTING,
                win32file.FILE_ATTRIBUTE_NORMAL,
                None,
            )
            file_size = win32file.GetFileSize(handle)
            hr, data = win32file.ReadFile(handle, file_size)
            win32file.CloseHandle(handle)
            return bytes(data), None
        except Exception as e:
            return None, f"Excel file locked by another program: {e}"
    except Exception as e:
        return None, str(e)

def get_sheet_names_from_bytes(file_bytes):
    try:
        bio = io.BytesIO(file_bytes)
        xl = pd.ExcelFile(bio)
        return xl.sheet_names, None
    except Exception as e:
        return [], str(e)

@st.cache_data(show_spinner=False)
def load_dataframe_from_bytes(file_bytes, sheet_name):
    """
    Loads sheet data into a pandas DataFrame from raw bytes.
    DataFrame is 100% pickle-serializable.
    """
    try:
        bio = io.BytesIO(file_bytes)
        df = pd.read_excel(bio, sheet_name=sheet_name)
        # Convert object columns to string to guarantee Arrow serialization compatibility
        for col in df.columns:
            if df[col].dtype == object or str(df[col].dtype).startswith("datetime"):
                df[col] = df[col].astype(str)
        return df, None
    except Exception as e:
        return None, str(e)

def parse_observation_tracker(raw_df):
    """
    Parses and cleans the Observation Tracker sheet.
    Detects criteria, status, and action category columns.
    """
    if raw_df is None or raw_df.empty:
        return None, "Empty dataset provided."

    df = raw_df.loc[:, ~raw_df.columns.astype(str).str.contains("^Unnamed")].copy()
    df.columns = [clean_column_name(c) for c in df.columns]

    # Detect Criteria Column
    criteria_col = None
    for col in df.columns:
        c_upper = col.upper()
        if "STANDARD" in c_upper or "CRETERIA" in c_upper or "CRITERIA" in c_upper or "CATEGORY" in c_upper:
            criteria_col = col
            break
    if not criteria_col:
        for col in df.columns:
            sample_vals = " ".join(df[col].dropna().astype(str).head(10).tolist()).upper()
            if "HEALTH" in sample_vals or "CUSTOMER" in sample_vals or "LOGISTIC" in sample_vals:
                criteria_col = col
                break

    # Detect Status Column
    status_col = None
    for col in df.columns:
        if "STATUS" in col.upper():
            status_col = col
            break

    # Detect Action Category Column (ACTION CATEGREY)
    action_cat_col = None
    for col in df.columns:
        c_u = col.upper()
        if "ACTION" in c_u and ("CATEG" in c_u or "TYPE" in c_u or "CLASS" in c_u):
            action_cat_col = col
            break

    # Detect Station Column
    station_col = None
    for col in df.columns:
        if "STATION" in col.upper():
            station_col = col
            break

    # Detect Lead Column
    lead_col = None
    for col in df.columns:
        if "RESP" in col.upper() or "LEAD" in col.upper() or "PILOT" in col.upper():
            lead_col = col
            break

    # Detect Observation text column
    obs_col = None
    for col in df.columns:
        if "OBSERVATION" in col.upper() and "DATE" not in col.upper():
            obs_col = col
            break

    if not criteria_col or not status_col:
        return None, "Could not detect Standard Criteria or Status columns in selected sheet."

    # Normalized fields
    df["Category"] = df[criteria_col].apply(normalize_category_value)
    df["Clean_Status"] = df[status_col].apply(normalize_status_value)

    if action_cat_col:
        df["Action_Category"] = (
            df[action_cat_col]
            .fillna("Unassigned")
            .astype(str)
            .str.strip()
            .str.title()
            .replace({"Standrad": "Standard", "Nan": "Unassigned"})
        )
    else:
        df["Action_Category"] = "General"

    if station_col:
        df["Station"] = df[station_col].fillna("Unknown").astype(str).str.strip()
    else:
        df["Station"] = "General"

    if lead_col:
        df["Lead"] = df[lead_col].fillna("Unassigned").astype(str).str.strip()
    else:
        df["Lead"] = "Unassigned"

    if obs_col:
        df["Observation_Text"] = df[obs_col].fillna("").astype(str).str.strip()
    else:
        df["Observation_Text"] = ""

    return df, None

# ==============================================================================
# Sidebar: File Management & Interactive Filters
# ==============================================================================
with st.sidebar:
    st.markdown("## 📊 Data Source & Setup")
    st.caption("LOAD EXCEL FILE & RUN DASHBOARD")

    uploaded_file = st.file_uploader(
        "Upload Excel file (.xlsx, .xls)",
        type=["xlsx", "xls"],
        help="Upload an updated Excel sheet to dynamically generate charts and tables.",
    )

    data_source_info = ""
    file_bytes = None
    sheet_options = []
    selected_sheet = None
    raw_df = None
    load_err = None
    local_mtime = ""

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        data_source_info = f"📤 Uploaded: {uploaded_file.name}"
        sheet_options, s_err = get_sheet_names_from_bytes(file_bytes)
        if s_err:
            st.error(f"Error inspecting uploaded file: {s_err}")
        else:
            # Check for GitHub token for automated sync
            gh_token = None
            try:
                if "GITHUB_TOKEN" in st.secrets:
                    gh_token = st.secrets["GITHUB_TOKEN"]
            except Exception:
                pass
            if not gh_token:
                gh_token = os.environ.get("GITHUB_TOKEN")

            if gh_token:
                sync_key = f"gh_synced_{uploaded_file.name}_{len(file_bytes)}"
                if sync_key not in st.session_state:
                    with st.spinner("🚀 Syncing updated Excel to GitHub..."):
                        synced, s_msg = sync_excel_to_github(file_bytes, GITHUB_TARGET_FILE, gh_token)
                        if synced:
                            st.session_state[sync_key] = True
                            st.toast("✅ Auto-synced new Excel file to GitHub repository!", icon="🚀")
                        else:
                            st.warning(f"GitHub Sync Note: {s_msg}")
    else:
        if os.path.exists(DEFAULT_EXCEL_PATH):
            mtime = os.path.getmtime(DEFAULT_EXCEL_PATH)
            local_mtime = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            data_source_info = f"📁 Local: {DEFAULT_EXCEL_PATH}"
            file_bytes, b_err = read_file_bytes_safe(DEFAULT_EXCEL_PATH)
            if b_err:
                st.error(b_err)
            elif file_bytes:
                sheet_options, s_err = get_sheet_names_from_bytes(file_bytes)
                if s_err:
                    st.error(f"Error reading sheets: {s_err}")
        else:
            st.error(f"Default file '{DEFAULT_EXCEL_PATH}' not found. Please upload an Excel file.")

    c_reload, c_clear = st.columns([1, 1])
    with c_reload:
        if st.button("🔄 Refresh Data", use_container_width=True, help="Reread local file or uploaded Excel"):
            st.cache_data.clear()
            st.rerun()

    if data_source_info:
        st.markdown(
            f"""
            <div style="background:#1e293b;border-radius:8px;padding:8px 12px;margin:8px 0;font-size:11.5px;color:#cbd5e1">
              <b>Source:</b> {data_source_info}<br/>
              {f'<span style="color:#94a3b8">Updated: {local_mtime}</span>' if local_mtime else ''}
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.divider()

    if sheet_options and file_bytes:
        default_idx = 0
        for idx, sname in enumerate(sheet_options):
            if "OBSERVATION" in sname.upper() or "TRACKER" in sname.upper():
                default_idx = idx
                break
        selected_sheet = st.selectbox("Excel Sheet to Analyze", sheet_options, index=default_idx)

        # Load Sheet Data safely from bytes
        raw_loaded_df, load_err = load_dataframe_from_bytes(file_bytes, selected_sheet)
        if load_err:
            st.error(f"Error loading sheet: {load_err}")
        else:
            raw_df, parse_err = parse_observation_tracker(raw_loaded_df)
            if parse_err:
                st.warning(f"Note: {parse_err}")

    st.markdown("### 🔍 Filters")
    st.caption("Apply filters across category views")

    filter_category = "All Categories"
    selected_statuses = ["Closed", "In Progress"]
    selected_actions = []
    search_keyword = ""

    if raw_df is not None and not raw_df.empty:
        available_categories = ["All Categories"] + sorted([c for c in raw_df["Category"].dropna().unique().tolist() if c])
        filter_category = st.selectbox("Standard Criteria / Category", available_categories, index=0)

        available_statuses = sorted(raw_df["Clean_Status"].dropna().unique().tolist())
        selected_statuses = st.multiselect(
            "Status Filter",
            options=available_statuses,
            default=available_statuses,
        )

        actions_list = sorted([a for a in raw_df["Action_Category"].dropna().unique().tolist() if a != "Unassigned"])
        selected_actions = st.multiselect("Action Category Filter", options=actions_list, default=[])

        search_keyword = st.text_input("Search Observation / Action", placeholder="Type keywords...")

    st.divider()
    st.markdown("### ⚙️ Chart Display Options")
    chart_orientation = st.radio("2D Bar Orientation", ["Vertical", "Horizontal"], index=0)
    show_data_labels = st.toggle("Show Values on Bars", value=True)

# ==============================================================================
# Apply Filters
# ==============================================================================
filtered_df = pd.DataFrame()
if raw_df is not None and not raw_df.empty:
    df_temp = raw_df.copy()

    if filter_category != "All Categories":
        df_temp = df_temp[df_temp["Category"] == filter_category]

    if selected_statuses:
        df_temp = df_temp[df_temp["Clean_Status"].isin(selected_statuses)]

    if selected_actions:
        df_temp = df_temp[df_temp["Action_Category"].isin(selected_actions)]

    if search_keyword.strip():
        k = search_keyword.strip().lower()
        obs_match = df_temp["Observation_Text"].str.lower().str.contains(k, na=False)
        action_match = df_temp.get("ACTION DETAILS", pd.Series("", index=df_temp.index)).astype(str).str.lower().str.contains(k, na=False)
        df_temp = df_temp[obs_match | action_match]

    filtered_df = df_temp

total_obs = len(raw_df) if raw_df is not None else 0
filtered_total = len(filtered_df)
closed_total = len(filtered_df[filtered_df["Clean_Status"] == "Closed"]) if not filtered_df.empty else 0
inprogress_total = len(filtered_df[filtered_df["Clean_Status"] == "In Progress"]) if not filtered_df.empty else 0
closure_rate = (closed_total / filtered_total * 100) if filtered_total > 0 else 0

# ==============================================================================
# Header Banner
# ==============================================================================
st.markdown(
    f"""
    <div class="hero-banner">
      <div class="hero-tag">WPI OPERATIONS • EXCEL INTEGRATED DASHBOARD</div>
      <h1>Standard Criteria & Station Assessment Dashboard</h1>
      <p>Automated extraction from local Excel tables · Side-by-side Station Assessment & 4 Criteria Analytics · 2D Clustered Charts</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Top KPI Metric Cards
k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(
        f"""
        <div class="metric-card">
          <div class="metric-label"><span>AFFECTED OBSERVATIONS</span><span>📋</span></div>
          <div class="metric-value" style="color:#0f172a">{filtered_total:,}</div>
          <div class="metric-sub">{f"Filtered from {total_obs:,} total" if filtered_total != total_obs else "Total observation scope"}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k2:
    st.markdown(
        f"""
        <div class="metric-card">
          <div class="metric-label"><span>CLOSED ITEMS</span><span>✅</span></div>
          <div class="metric-value" style="color:#10b981">{closed_total:,}</div>
          <div class="metric-sub"><span style="color:#10b981;font-weight:700">{closure_rate:.1f}%</span> completed of filtered</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k3:
    prog_pct = (inprogress_total / filtered_total * 100) if filtered_total > 0 else 0
    st.markdown(
        f"""
        <div class="metric-card">
          <div class="metric-label"><span>IN PROGRESS ITEMS</span><span>⏳</span></div>
          <div class="metric-value" style="color:#f59e0b">{inprogress_total:,}</div>
          <div class="metric-sub"><span style="color:#f59e0b;font-weight:700">{prog_pct:.1f}%</span> currently active</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k4:
    active_cats = filtered_df["Category"].nunique() if not filtered_df.empty else 0
    st.markdown(
        f"""
        <div class="metric-card">
          <div class="metric-label"><span>CRITERIA CATEGORIES</span><span>🏷️</span></div>
          <div class="metric-value" style="color:#2563eb">{active_cats}</div>
          <div class="metric-sub">Active categories analyzed</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.write("")

# ==============================================================================
# Helper Function: Render Category Action Charts Pair
# Left (Chart 1): Status Overview Column Chart (Closed, In Progress, Affected)
# Right (Chart 2): Action Category Horizontal Bar Chart (Total Items per Action)
# Compact small-size charts with fully visible text labels
# ==============================================================================
def render_category_charts_pair(cat_name, cat_icon, cat_color, base_df):
    cat_df = base_df[base_df["Category"] == cat_name].copy() if not base_df.empty else pd.DataFrame()
    c_total = len(cat_df)
    c_closed = len(cat_df[cat_df["Clean_Status"] == "Closed"]) if c_total > 0 else 0
    c_inprog = len(cat_df[cat_df["Clean_Status"] == "In Progress"]) if c_total > 0 else 0

    st.markdown(
        f"""
        <div style="background:white;border-left:5px solid {cat_color};border-radius:10px;padding:8px 14px;margin-top:12px;margin-bottom:8px;box-shadow:0 2px 6px rgba(15,23,42,0.04);display:flex;justify-content:space-between;align-items:center;">
          <div style="font-size:15px;font-weight:800;color:#0f172a;">{cat_icon} {cat_name}</div>
          <div style="display:flex;gap:6px;">
            <span class="badge-closed">{c_closed} Closed</span>
            <span class="badge-prog">{c_inprog} In Progress</span>
            <span class="badge-affected">{c_total} Affected</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if c_total == 0:
        st.info(f"No records found for '{cat_name}' with the current filters.")
        return

    # Action categories counts
    act_counts = (
        cat_df["Action_Category"]
        .value_counts()
        .reset_index()
    )
    act_counts.columns = ["Action_Category", "Total"]
    act_counts = act_counts.sort_values("Total", ascending=False)
    act_names = act_counts["Action_Category"].tolist()
    act_totals = [int(v) for v in act_counts["Total"]]

    c_left, c_right = st.columns([1, 1], gap="medium")

    # ------------------ LEFT (CHART 1): STATUS OVERVIEW (CLOSED, IN PROGRESS, AFFECTED) ------------------
    with c_left:
        stat_cats = ["Closed", "In Progress", "Affected"]
        stat_vals = [c_closed, c_inprog, c_total]
        stat_colors = ["#10b981", "#f59e0b", "#2563eb"]  # Green, Amber, Blue

        fig_status = go.Figure()
        fig_status.add_trace(
            go.Bar(
                x=stat_cats,
                y=stat_vals,
                marker=dict(color=stat_colors, line=dict(width=0)),
                text=[f"{v:,}" for v in stat_vals] if show_data_labels else None,
                textposition="outside" if show_data_labels else "none",
                cliponaxis=False,
                textfont=dict(size=10.5, color="#0f172a", weight=700),
                hovertemplate="<b>%{x}</b><br>Count: <b>%{y:,}</b><extra></extra>",
            )
        )
        fig_status.update_layout(
            title=dict(text=f"<b>{cat_name}: Status (Closed, In Progress, Affected)</b>", font=dict(size=12, color="#0f172a"), x=0.01, y=0.96),
            height=240,
            margin=dict(l=25, r=20, t=32, b=20),
            paper_bgcolor="white",
            plot_bgcolor="white",
            font=dict(family="Inter, sans-serif"),
            showlegend=False,
            xaxis=dict(showgrid=False, zeroline=False, tickfont=dict(size=10.5, color="#1e293b", weight=700)),
            yaxis=dict(showgrid=True, gridcolor="#edf2f7", zeroline=False, title=dict(text="Count", font=dict(size=10.5, color="#64748b"))),
        )
        st.plotly_chart(fig_status, use_container_width=True, theme=None, config={"displayModeBar": False})

    # ------------------ RIGHT (CHART 2): ACTION CATEGORY TOTAL ITEMS (HORIZONTAL) ------------------
    with c_right:
        ACTION_PALETTE = {
            "Facility": "#3b82f6",     # Blue
            "Process": "#06b6d4",      # Cyan
            "Standard": "#8b5cf6",     # Purple
            "Training": "#f59e0b",     # Amber
            "Part": "#10b981",         # Emerald
            "General": "#64748b",      # Slate
            "Unassigned": "#94a3b8",   # Gray
        }
        bar_colors = [ACTION_PALETTE.get(a, cat_color) for a in act_names]

        fig_horiz = go.Figure()
        fig_horiz.add_trace(
            go.Bar(
                name="Total Items",
                y=act_names,
                x=act_totals,
                orientation="h",
                marker=dict(color=bar_colors, line=dict(width=0)),
                text=[f"{v:,}" for v in act_totals] if show_data_labels else None,
                textposition="outside" if show_data_labels else "none",
                cliponaxis=False,
                textfont=dict(size=10.5, color="#0f172a", weight=700),
                hovertemplate="<b>Action Category:</b> %{y}<br>Total Items: <b>%{x:,}</b><extra></extra>",
            )
        )
        fig_horiz.update_layout(
            title=dict(text=f"<b>{cat_name}: Action Category Total Items</b>", font=dict(size=12, color="#0f172a"), x=0.01, y=0.96),
            height=240,
            margin=dict(l=95, r=40, t=32, b=20),
            paper_bgcolor="white",
            plot_bgcolor="white",
            font=dict(family="Inter, sans-serif"),
            showlegend=False,
            yaxis=dict(autorange="reversed", showgrid=False, tickfont=dict(size=10.5, color="#1e293b", weight=600)),
            xaxis=dict(showgrid=True, gridcolor="#edf2f7", zeroline=False, title=dict(text="Total Items", font=dict(size=10.5, color="#64748b"))),
        )
        st.plotly_chart(fig_horiz, use_container_width=True, theme=None, config={"displayModeBar": False})

# ==============================================================================
# Dashboard Tabs Setup
# Tab 1: 📊 All Charts Dashboard (All Category Charts in Home Page) [DEFAULT]
# Tab 2: 📁 Raw Data Explorer & Export
# ==============================================================================
tab_home, tab_explorer = st.tabs([
    "📊 All Charts Dashboard",
    "📁 Raw Data Explorer & Export",
])

# ------------------------------------------------------------------------------
# TAB 1: ALL CHARTS DASHBOARD (HOME PAGE)
# ALL CATEGORY CHARTS DISPLAYED DIRECTLY ON HOME PAGE (NO TABLES, PURE CHARTS)
# ------------------------------------------------------------------------------
with tab_home:

    # TWO COLUMNS ALONGSIDE EACH OTHER:
    # Left: Station Assessment (form, metrics, progress, 2D bar chart)
    # Right: Overall 4 Categories Dashboard (2D clustered chart, summary table, status breakdown)
    c_left_station, c_right_criteria = st.columns([1, 1], gap="large")

    # ========================== LEFT: STATION ASSESSMENT ==========================
    with c_left_station:
        st.markdown(
            """
            <div class="home-panel-title">🏢 Station Assessment</div>
            <div class="home-panel-sub">LHD affected station assessment status · Interactive live inputs & 2D bar view</div>
            """,
            unsafe_allow_html=True,
        )

        # Form Inputs for Station Assessment
        with st.expander("📝 Edit Station Assessment Values", expanded=True):
            st.caption("Update numbers below. Chart and completion rates recalculate automatically.")
            l1_overall = st.number_input("Overall assessments", min_value=0, max_value=1000000, value=55, step=1, key="l1_ov")
            l1_in_progress = st.number_input("In progress", min_value=0, max_value=1000000, value=45, step=1, key="l1_ip")
            l1_completed = st.number_input("Completed", min_value=0, max_value=1000000, value=10, step=1, key="l1_cp")

            if l1_in_progress + l1_completed > l1_overall:
                st.warning("⚠️ In progress + Completed exceeds Overall. Please review.")

        l1_comp_rate = (l1_completed / l1_overall * 100) if l1_overall else 0
        l1_ip_rate = (l1_in_progress / l1_overall * 100) if l1_overall else 0

        # Station Assessment KPI cards
        sk1, sk2, sk3 = st.columns(3)
        with sk1:
            st.markdown(
                f'<div class="metric-card"><div class="metric-label">OVERALL</div><div class="metric-value" style="color:#2563eb">{l1_overall:,}</div><div class="metric-sub">Total Scope</div></div>',
                unsafe_allow_html=True,
            )
        with sk2:
            st.markdown(
                f'<div class="metric-card"><div class="metric-label">IN PROGRESS</div><div class="metric-value" style="color:#f59e0b">{l1_in_progress:,}</div><div class="metric-sub">{l1_ip_rate:.1f}% active</div></div>',
                unsafe_allow_html=True,
            )
        with sk3:
            st.markdown(
                f'<div class="metric-card"><div class="metric-label">COMPLETED</div><div class="metric-value" style="color:#10b981">{l1_completed:,}</div><div class="metric-sub">{l1_comp_rate:.1f}% done</div></div>',
                unsafe_allow_html=True,
            )

        # Progress bar
        st.markdown(
            f"""
            <div style="background:#f8fafc;padding:10px 14px;border-radius:10px;border:1px solid #e2e8f0;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center;">
              <span style="font-size:12px;font-weight:700;color:#334155;">Assessment Completion</span>
              <span style="font-size:15px;font-weight:800;color:#059669;">{l1_comp_rate:.1f}%</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.progress(min(l1_comp_rate / 100, 1.0))

        # Station Assessment 2D Bar Chart
        l1_cats = ["Overall", "In Progress", "Completed"]
        l1_vals = [l1_overall, l1_in_progress, l1_completed]
        l1_colors = ["#2563eb", "#f59e0b", "#10b981"]

        fig_l1 = go.Figure()
        if chart_orientation == "Horizontal":
            fig_l1.add_trace(
                go.Bar(
                    y=l1_cats,
                    x=l1_vals,
                    orientation="h",
                    marker=dict(color=l1_colors, line=dict(width=0)),
                    text=[f"{v:,}" for v in l1_vals] if show_data_labels else None,
                    textposition="outside" if show_data_labels else "none",
                    cliponaxis=False,
                    hovertemplate="<b>%{y}</b><br>Count: %{x:,}<extra></extra>",
                )
            )
            fig_l1.update_layout(yaxis=dict(autorange="reversed"))
        else:
            fig_l1.add_trace(
                go.Bar(
                    x=l1_cats,
                    y=l1_vals,
                    marker=dict(color=l1_colors, line=dict(width=0)),
                    text=[f"{v:,}" for v in l1_vals] if show_data_labels else None,
                    textposition="outside" if show_data_labels else "none",
                    cliponaxis=False,
                    hovertemplate="<b>%{x}</b><br>Count: %{y:,}<extra></extra>",
                )
            )

        fig_l1.update_layout(
            title=dict(text="<b>Station Assessment Status</b>", font=dict(size=12, color="#0f172a"), x=0.01, y=0.96),
            height=240,
            margin=dict(l=25, r=20, t=32, b=20),
            paper_bgcolor="white",
            plot_bgcolor="white",
            showlegend=False,
            xaxis=dict(showgrid=False, zeroline=False, tickfont=dict(size=10.5, color="#1e293b", weight=700)),
            yaxis=dict(showgrid=True, gridcolor="#edf1f7", zeroline=False, title=dict(text="Count", font=dict(size=10.5, color="#64748b"))),
        )
        st.plotly_chart(fig_l1, use_container_width=True, theme=None, config={"displayModeBar": False})

    # ========================== RIGHT: OVERALL 4 CATEGORIES (TOTAL AFFECTED) ==========================
    with c_right_criteria:
        st.markdown(
            """
            <div class="home-panel-title">📊 Overall Criteria (All 4 Categories)</div>
            """,
            unsafe_allow_html=True,
        )

        if filtered_df.empty:
            st.warning("No observation data loaded. Check Excel file source.")
        else:
            # Pivot table for the 4 categories
            cat_pivot = (
                filtered_df.groupby(["Category", "Clean_Status"]).size().unstack(fill_value=0).reset_index()
            )
            if "Closed" not in cat_pivot.columns:
                cat_pivot["Closed"] = 0
            if "In Progress" not in cat_pivot.columns:
                cat_pivot["In Progress"] = 0
            cat_pivot["Total Affected"] = cat_pivot["Closed"] + cat_pivot["In Progress"]
            cat_pivot["Completion_Rate"] = (
                (cat_pivot["Closed"] / cat_pivot["Total Affected"] * 100).fillna(0).round(1)
            )

            # Ensure all 4 standard categories appear
            for cat_name in STANDARD_CATEGORIES:
                if cat_name not in cat_pivot["Category"].values:
                    cat_pivot = pd.concat([
                        cat_pivot,
                        pd.DataFrame([{"Category": cat_name, "Closed": 0, "In Progress": 0, "Total Affected": 0, "Completion_Rate": 0.0}])
                    ], ignore_index=True)

            cat_names = cat_pivot["Category"].tolist()
            affected_vals = [int(v) for v in cat_pivot["Total Affected"]]
            cat_bar_colors = [CATEGORY_PALETTE.get(c, "#2563eb") for c in cat_names]

            # SINGLE COLUMN BAR CHART: Only Total Affected across all 4 categories
            fig_overall = go.Figure()
            if chart_orientation == "Vertical":
                fig_overall.add_trace(
                    go.Bar(
                        name="Total Affected",
                        x=cat_names,
                        y=affected_vals,
                        marker=dict(color=cat_bar_colors, line=dict(width=0)),
                        text=[f"{v:,}" for v in affected_vals] if show_data_labels else None,
                        textposition="outside" if show_data_labels else "none",
                        cliponaxis=False,
                        hovertemplate="<b>Standard Criteria:</b> %{x}<br><b>Total Affected:</b> %{y:,}<extra></extra>",
                    )
                )
                fig_overall.update_layout(
                    xaxis=dict(showgrid=False, zeroline=False, tickfont=dict(size=11.5, color="#1e293b", weight=700)),
                    yaxis=dict(showgrid=True, gridcolor="#edf2f7", zeroline=False, title=dict(text="Affected Observations", font=dict(size=12, color="#64748b"))),
                )
            else:
                fig_overall.add_trace(
                    go.Bar(
                        name="Total Affected",
                        y=cat_names,
                        x=affected_vals,
                        orientation="h",
                        marker=dict(color=cat_bar_colors, line=dict(width=0)),
                        text=[f"{v:,}" for v in affected_vals] if show_data_labels else None,
                        textposition="outside" if show_data_labels else "none",
                        cliponaxis=False,
                        hovertemplate="<b>Standard Criteria:</b> %{y}<br><b>Total Affected:</b> %{x:,}<extra></extra>",
                    )
                )
                fig_overall.update_layout(
                    yaxis=dict(autorange="reversed", showgrid=False, tickfont=dict(size=11.5, color="#1e293b", weight=700)),
                    xaxis=dict(showgrid=True, gridcolor="#edf2f7", zeroline=False, title=dict(text="Affected Observations", font=dict(size=12, color="#64748b"))),
                )

            fig_overall.update_layout(
                title=dict(text="<b>Total Affected Scope across All 4 Categories</b>", font=dict(size=12, color="#0f172a"), x=0.01, y=0.96),
                height=240,
                margin=dict(l=25, r=20, t=32, b=20),
                paper_bgcolor="white",
                plot_bgcolor="white",
                font=dict(family="Inter, sans-serif"),
                showlegend=False,
            )
            st.plotly_chart(fig_overall, use_container_width=True, theme=None, config={"displayModeBar": False})

    # --------------------------------------------------------------------------
    # ALL 4 CATEGORIES CHARTS ON HOME PAGE (LEFT: COLUMN-WISE, RIGHT: HORIZONTAL)
    # --------------------------------------------------------------------------
    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)

    # 1. Health and Safety
    render_category_charts_pair("Health and Safety", "🛡️", "#ef4444", filtered_df)

    # 2. Customer Experience
    render_category_charts_pair("Customer Experience", "⭐", "#6366f1", filtered_df)

    # 3. Workplace Organization
    render_category_charts_pair("Workplace Organization", "🏢", "#059669", filtered_df)

    # 4. Logistics
    render_category_charts_pair("Logistics", "🚚", "#0284c7", filtered_df)

# ------------------------------------------------------------------------------
# TAB 2: GLOBAL DATA EXPLORER & EXPORT
# ------------------------------------------------------------------------------
with tab_explorer:
    st.markdown(
        """
        <div style="background:white;border-radius:14px;padding:16px 20px;margin-bottom:18px;border:1.5px solid #cbd5e1;box-shadow:0 3px 12px rgba(15,23,42,0.04);">
          <div style="font-size:18px;font-weight:800;color:#0f172a;">🔍 Global Data Explorer & Export</div>
          <div style="font-size:12.5px;color:#64748b;">Search, inspect raw extracted fields, filter records, and export to CSV or Excel</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if filtered_df.empty:
        st.info("No records to display. Adjust filters in the sidebar.")
    else:
        st.markdown(f"**Showing {len(filtered_df):,} matching observation records**")

        st.dataframe(
            filtered_df,
            use_container_width=True,
            hide_index=True,
            height=460,
        )

        dl_col1, dl_col2 = st.columns([1, 1])
        with dl_col1:
            csv_data = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Download Filtered Data as CSV",
                data=csv_data,
                file_name="wpi_filtered_observations.csv",
                mime="text/csv",
                use_container_width=True,
            )

        with dl_col2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Filtered Observations")
                if "cat_pivot" in locals():
                    cat_pivot.to_excel(writer, index=False, sheet_name="Category Summary")
            st.download_button(
                "⬇️ Download Filtered Data as Excel (.xlsx)",
                data=buffer.getvalue(),
                file_name="wpi_filtered_observations.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# ==============================================================================
# Footer
# ==============================================================================
st.markdown(
    """
    <div class="footer-bar">
      Standard Criteria & Station Assessment Dashboard · Automated Excel Extraction & 2D Clustered Visualizations · Built with Python & Streamlit
    </div>
    """,
    unsafe_allow_html=True,
)
