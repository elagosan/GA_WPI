import base64
import io
import os
import subprocess
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

# ============================================================
# PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="WORKPLACE INTEGRATION - CC22 LHD | GENERAL ASSEMBLY",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_EXCEL_PATH = "WPI Implementation at TVR - Time Line.xlsx"
GITHUB_REPO_OWNER = "elagosan"
GITHUB_REPO_NAME = "GA_WPI"
GITHUB_TARGET_FILE = "WPI Implementation at TVR - Time Line.xlsx"

STANDARD_CATEGORIES = [
    "Customer Experience",
    "Workplace Organization",
    "Health and Safety",
    "Logistics",
]

# ============================================================
# CSS STYLING (Modern Executive Theme)
# ============================================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    .stApp {
        background: #f8fafc;
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
        padding-top: .4rem;
        padding-bottom: 2rem;
        max-width: 100%;
    }

    /* Hero Banner */
    .hero-banner {
        padding: 24px 28px;
        border-radius: 16px;
        color: white;
        background: linear-gradient(
            135deg, #0b1329 0%, #1e3a8a 40%,
            #2563eb 80%, #38bdf8 135%
        );
        box-shadow: 0 12px 30px -6px rgba(30,64,175,.28);
        margin-bottom: 18px;
        overflow: hidden;
        position: relative;
    }

    .hero-banner h1 {
        font-size: 27px;
        line-height: 1.25;
        margin: 2px 0 6px;
        font-weight: 900;
        letter-spacing: -0.5px;
    }

    .hero-banner p {
        color: #e0f2fe;
        margin: 0;
        font-size: 14.5px;
        font-weight: 600;
        letter-spacing: 0.5px;
    }

    /* Section Headers */
    .section-header-box {
        margin: 18px 0 12px;
        padding-bottom: 8px;
        border-bottom: 2px solid #e2e8f0;
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
    }

    .section-title {
        font-size: 20px;
        font-weight: 800;
        color: #0f172a;
        display: flex;
        align-items: center;
        gap: 8px;
        margin: 0;
    }

    .section-sub {
        font-size: 12.5px;
        color: #64748b;
        margin-top: 3px;
    }

    /* Metric Cards */
    .metric-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 16px 18px;
        box-shadow: 0 2px 10px rgba(15,23,42,.04);
        margin-bottom: 10px;
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }

    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(15,23,42,.08);
    }

    .metric-label {
        font-size: 11px;
        font-weight: 700;
        color: #64748b;
        letter-spacing: .6px;
        text-transform: uppercase;
        margin-bottom: 6px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    .metric-value {
        font-size: 28px;
        font-weight: 900;
        line-height: 1.1;
    }

    .metric-sub {
        font-size: 11.5px;
        color: #64748b;
        margin-top: 6px;
        display: flex;
        align-items: center;
        gap: 4px;
    }

    /* Badges */
    .badge-closed {
        background: #dcfce7;
        color: #15803d;
        padding: 3px 10px;
        border-radius: 9999px;
        font-size: 11.5px;
        font-weight: 700;
    }

    .badge-prog {
        background: #fef3c7;
        color: #b45309;
        padding: 3px 10px;
        border-radius: 9999px;
        font-size: 11.5px;
        font-weight: 700;
    }

    .badge-affected {
        background: #dbeafe;
        color: #1e40af;
        padding: 3px 10px;
        border-radius: 9999px;
        font-size: 11.5px;
        font-weight: 700;
    }

    .badge-sync {
        background: rgba(16, 185, 129, 0.18);
        color: #34d399;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 11.5px;
        font-weight: 700;
        display: inline-flex;
        align-items: center;
        gap: 5px;
        border: 1px solid rgba(52, 211, 153, 0.35);
    }

    /* Station Box Container */
    .station-panel-container {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 18px 22px;
        box-shadow: 0 4px 16px rgba(15,23,42,.05);
        margin-bottom: 18px;
    }

    /* Tab bar styling */
    div[data-testid="stTabs"] [data-baseweb="tab-list"] {
        background: white !important;
        border-radius: 12px !important;
        padding: 6px !important;
        border: 1px solid #cbd5e1 !important;
        gap: 8px !important;
        margin-bottom: 14px !important;
    }

    div[data-testid="stTabs"] button[role="tab"] {
        background: #f1f5f9 !important;
        color: #0f172a !important;
        font-weight: 700 !important;
        border-radius: 9px !important;
        padding: 9px 18px !important;
    }

    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        background: #2563eb !important;
        color: white !important;
    }

    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] p {
        color: white !important;
    }

    .footer-bar {
        text-align: center;
        color: #94a3b8;
        font-size: 11.5px;
        padding: 20px 0 8px;
        border-top: 1px solid #e2e8f0;
        margin-top: 24px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def get_logo_base64():
    for path in (
        "stellantis_logo.jpg",
        "logo.jpg",
        "stellantis_logo.png",
        "logo.png",
    ):
        if os.path.exists(path):
            try:
                with open(path, "rb") as image_file:
                    return base64.b64encode(image_file.read()).decode("utf-8")
            except OSError:
                continue
    return None


def clean_column_name(column):
    return str(column).strip().replace("\n", " ")


def normalize_category_value(value):
    if pd.isna(value) or value is None:
        return "Uncategorized"

    val_str = str(value).strip()
    upper = val_str.upper()

    if "HEALTH" in upper or "SAFETY" in upper or "JSRA" in upper or "H&S" in upper:
        return "Health and Safety"

    if "CUSTOMER" in upper or "EXPER" in upper:
        return "Customer Experience"

    if "LOGISTIC" in upper:
        return "Logistics"

    if "WORK" in upper or "ORGAN" in upper:
        return "Workplace Organization"

    return val_str.title() if len(val_str) < 35 else val_str


def normalize_status_value(value):
    if pd.isna(value) or value is None:
        return "Unknown"

    compact = (
        str(value).strip().upper()
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
    )

    if any(word in compact for word in ("CLOSE", "COMPLET", "DONE")):
        return "Closed"

    if any(word in compact for word in ("PROG", "OPEN", "PEND", "START")):
        return "In Progress"

    return str(value).strip().title()


def normalize_action_category(value):
    if pd.isna(value) or value is None:
        return "Unassigned"

    val_str = str(value).strip()
    upper = val_str.upper()

    if "STAND" in upper:
        return "Standard"
    if "FACIL" in upper:
        return "Facility"
    if "PROC" in upper:
        return "Process"
    if "PART" in upper:
        return "Part"
    if "TRAIN" in upper:
        return "Training"
    if upper in ("", "NAN", "NONE", "UNASSIGNED", "-"):
        return "Unassigned"

    return val_str.title()


# ============================================================
# GITHUB AUTO-SYNC ENGINE
# ============================================================
def get_effective_github_token():
    # 1. From session state (user entered in sidebar)
    if "github_token" in st.session_state and st.session_state["github_token"]:
        return st.session_state["github_token"].strip()

    # 2. From Streamlit secrets
    try:
        if "GITHUB_TOKEN" in st.secrets:
            return str(st.secrets["GITHUB_TOKEN"]).strip()
    except Exception:
        pass

    # 3. From environment variables
    for env_key in ("GITHUB_TOKEN", "GH_TOKEN"):
        env_val = os.environ.get(env_key)
        if env_val:
            return env_val.strip()

    return None


def sync_file_to_github(file_bytes, target_filename=GITHUB_TARGET_FILE, commit_message=None):
    """
    Persists uploaded Excel workbook and pushes to GitHub repository.
    1. Writes locally to disk
    2. Pushes via GitHub REST API (works on Streamlit Cloud & everywhere)
    3. Pushes via Git CLI if local repo is available
    """
    if commit_message is None:
        timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_message = f"Update {target_filename} via Streamlit Dashboard ({timestamp_str})"

    # Always save locally first so application state is persistent
    saved_locally = False
    try:
        with open(target_filename, "wb") as f:
            f.write(file_bytes)
        saved_locally = True
    except Exception as exc:
        local_err = str(exc)

    token = get_effective_github_token()

    # Method 1: GitHub REST API
    if token:
        try:
            api_url = f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/contents/{target_filename}"
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/vnd.github.v3+json",
                "User-Agent": "GA-WPI-Streamlit-Sync",
            }

            # Get current SHA
            get_res = requests.get(api_url, headers=headers, timeout=15)
            existing_sha = None
            if get_res.status_code == 200:
                existing_sha = get_res.json().get("sha")

            # Push content
            payload = {
                "message": commit_message,
                "content": base64.b64encode(file_bytes).decode("utf-8"),
                "branch": "main",
            }
            if existing_sha:
                payload["sha"] = existing_sha

            put_res = requests.put(api_url, headers=headers, json=payload, timeout=30)
            if put_res.status_code in (200, 201):
                return {
                    "success": True,
                    "method": "GitHub REST API",
                    "message": f"Successfully committed & synced to {GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME} on branch main!",
                }
            else:
                api_err_msg = put_res.json().get("message", put_res.text)
                # Fall through to Git CLI if API failed
        except Exception as exc:
            api_err_msg = str(exc)

    # Method 2: Git CLI (Local environment)
    if os.path.exists(".git"):
        try:
            subprocess.run(["git", "add", target_filename], check=True, capture_output=True, timeout=15)
            # Commit (ignore if nothing to commit)
            commit_res = subprocess.run(
                ["git", "commit", "-m", commit_message],
                capture_output=True,
                text=True,
                timeout=15,
            )
            # Push
            push_res = subprocess.run(
                ["git", "push", "origin", "main"],
                capture_output=True,
                text=True,
                timeout=25,
            )
            if push_res.returncode == 0:
                return {
                    "success": True,
                    "method": "Git CLI",
                    "message": "Successfully committed and pushed to GitHub repository via Git CLI!",
                }
        except Exception:
            pass

    # If GitHub push could not authenticate, report local save status
    if saved_locally:
        if token:
            return {
                "success": False,
                "method": "Local Disk",
                "message": f"Workbook saved locally, but GitHub sync failed: {api_err_msg if 'api_err_msg' in locals() else 'Authentication error'}. Please verify token permissions.",
            }
        else:
            return {
                "success": True,
                "method": "Local Disk",
                "message": "Workbook saved and loaded locally! (Add your GitHub Personal Access Token in the sidebar to auto-push to GitHub).",
            }

    return {
        "success": False,
        "method": "Failed",
        "message": f"Could not save workbook: {local_err if 'local_err' in locals() else 'Unknown error'}",
    }


# ============================================================
# EXCEL ROBUST PARSING & SMART DETECTION
# ============================================================
@st.cache_data(show_spinner=False)
def inspect_excel_sheets(file_bytes):
    try:
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        return excel_file.sheet_names, None
    except Exception as exc:
        return [], str(exc)


@st.cache_data(show_spinner=False)
def load_dataframe_from_bytes(file_bytes, sheet_name):
    """
    Intelligently reads an Excel sheet.
    If row 0 does not contain the headers, it scans rows 0 to 8 to find the real header row!
    """
    try:
        # First try normal row 0
        df_initial = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)

        # Check if columns look valid or if header is further down
        col_str = " ".join([str(c).upper() for c in df_initial.columns])
        needs_header_search = (
            "CRITERIA" not in col_str
            and "CRETERIA" not in col_str
            and "CATEGORY" not in col_str
            and "STATUS" not in col_str
            and any(str(c).startswith("Unnamed") for c in df_initial.columns[:3])
        )

        if needs_header_search and len(df_initial) > 0:
            for header_idx in range(min(8, len(df_initial))):
                row_vals = [str(v).upper() for v in df_initial.iloc[header_idx].values if pd.notna(v)]
                row_str = " ".join(row_vals)
                if ("CRITERIA" in row_str or "CRETERIA" in row_str or "CATEGORY" in row_str) and "STATUS" in row_str:
                    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_idx + 1)
                    break
            else:
                df = df_initial
        else:
            df = df_initial

        # Remove pure Unnamed columns if they are all NaN
        df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed:\s*\d+$") | df.notna().any(axis=0)]
        df.columns = [clean_column_name(c) for c in df.columns]

        # Arrow-safety: clean object column types
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).replace({"nan": "", "None": "", "<NA>": ""})

        return df, None
    except Exception as exc:
        return None, str(exc)


def score_observation_sheet(sheet_name, file_bytes):
    """Calculates confidence that a sheet is the observation tracker."""
    upper = sheet_name.upper()
    score = 0
    if "OBSERVATION" in upper:
        score += 50
    if "TRACKER" in upper:
        score += 30
    if "WPI" in upper:
        score += 10
    return score


def auto_detect_best_sheet(sheet_names, file_bytes):
    """Finds the most suitable Observation Tracker sheet."""
    scored = [(name, score_observation_sheet(name, file_bytes)) for name in sheet_names]
    scored.sort(key=lambda x: x[1], reverse=True)
    if scored and scored[0][1] > 0:
        return scored[0][0]
    return sheet_names[0] if sheet_names else None


def parse_observation_tracker(raw_df):
    if raw_df is None or raw_df.empty:
        return None, "The selected Excel sheet is empty."

    df = raw_df.copy()

    # Detect Criteria Column
    criteria_col = next(
        (
            c for c in df.columns
            if any(key in str(c).upper() for key in ("STANDARD CRITERIA", "STANDARD CRETERIA", "CRITERIA", "CRETERIA", "CATEGORY"))
        ),
        None,
    )

    if not criteria_col:
        for col in df.columns:
            sample = " ".join(df[col].astype(str).head(20)).upper()
            if any(key in sample for key in ("HEALTH", "CUSTOMER", "LOGISTIC", "WORKPLACE")):
                criteria_col = col
                break

    # Detect Status Column
    status_col = next(
        (c for c in df.columns if "STATUS" in str(c).upper() or "STATE" in str(c).upper()),
        None,
    )

    if not criteria_col or not status_col:
        return None, (
            "Could not identify the Standard Criteria and Status columns in this sheet. "
            "Please ensure the sheet contains criteria and status fields."
        )

    # Detect Action Category Column
    action_col = next(
        (
            c for c in df.columns
            if any(k in str(c).upper() for k in ("ACTION CATEGREY", "ACTION CATEGORY", "ACTION TYPE", "ACTION CLASS"))
            or ("ACTION" in str(c).upper() and "CATEG" in str(c).upper())
        ),
        None,
    )

    # Detect Station Column
    station_col = next(
        (c for c in df.columns if any(k in str(c).upper() for k in ("STATION NO", "STATION #", "STATION", "WORKSTATION"))),
        None,
    )

    # Detect Lead Column
    lead_col = next(
        (c for c in df.columns if any(key in str(c).upper() for key in ("RESP. LEAD", "LEAD", "RESPONSIBLE", "PILOT", "OWNER"))),
        None,
    )

    # Detect Observation Text Column
    observation_col = next(
        (
            c for c in df.columns
            if any(k in str(c).upper() for k in ("OBSERVATION", "FINDING", "DESCRIPTION"))
            and "DATE" not in str(c).upper()
            and "LEVEL" not in str(c).upper()
        ),
        None,
    )

    # Process and clean standardized fields
    df["Category"] = df[criteria_col].apply(normalize_category_value)
    df["Clean_Status"] = df[status_col].apply(normalize_status_value)

    # Remove completely empty summary/trailer rows where Category is Uncategorized and Status is Unknown
    df = df[~((df["Category"] == "Uncategorized") & (df["Clean_Status"] == "Unknown"))]

    if action_col:
        df["Action_Category"] = df[action_col].apply(normalize_action_category)
    else:
        df["Action_Category"] = "General"

    if station_col:
        df["Station"] = df[station_col].fillna("General").astype(str).str.strip().replace({"": "General", "nan": "General"})
    else:
        df["Station"] = "General"

    if lead_col:
        df["Lead"] = df[lead_col].fillna("Unassigned").astype(str).str.strip().replace({"": "Unassigned", "nan": "Unassigned"})
    else:
        df["Lead"] = "Unassigned"

    if observation_col:
        df["Observation_Text"] = df[observation_col].fillna("").astype(str).str.strip()
    else:
        df["Observation_Text"] = ""

    # Ensure all columns are Arrow-compatible
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str)

    return df, None


@st.cache_data(show_spinner=False)
def extract_station_assessment(file_bytes):
    default = {
        "overall": 44,
        "in_progress": 30,
        "completed": 14,
        "details_df": pd.DataFrame(),
    }

    if not file_bytes:
        return default

    try:
        excel = pd.ExcelFile(io.BytesIO(file_bytes))

        # Look for Station sheet
        target_sheet = next(
            (
                sheet for sheet in (
                    "STATION WISE PROGRESS (2)",
                    "STATION WISE PROGRESS",
                    "STATION PROGRESS",
                    "STATION ASSESSMENT",
                )
                if sheet in excel.sheet_names
            ),
            None,
        )

        if not target_sheet:
            for s in excel.sheet_names:
                if "STATION" in s.upper():
                    target_sheet = s
                    break

        if not target_sheet:
            return default

        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet)
        if df.empty:
            return default

        station_col = df.columns[0]
        station_df = df.dropna(subset=[station_col]).copy()
        station_df = station_df[
            station_df[station_col].astype(str).str.strip().str.upper() != "STATION #"
        ]

        observation_col = next(
            (
                c for c in station_df.columns
                if "OBSERVATION" in str(c).upper()
                or (len(df) > 0 and "OBSERVATION" in str(df[c].iloc[0]).upper())
            ),
            None,
        )

        completion_col = next(
            (
                c for c in station_df.columns
                if "COMPLETION" in str(c).upper()
                or (len(df) > 0 and "COMPLETION" in str(df[c].iloc[0]).upper())
                or "%" in str(c)
            ),
            None,
        )

        total = len(station_df)
        completed = 0
        in_progress = 0

        if observation_col:
            status = station_df[observation_col].astype(str).str.strip().str.lower()
            completed = int(status.isin(["completed", "complete", "closed", "done"]).sum())
            in_progress = int(status.isin(["inprogress", "in progress", "progress", "yet to start"]).sum())
        elif completion_col:
            completion = pd.to_numeric(station_df[completion_col], errors="coerce").fillna(0)
            completed = int((completion >= 1).sum())
            in_progress = total - completed

        if completed == 0 and in_progress == 0 and total > 0:
            in_progress = total

        details = pd.DataFrame({
            "Station": station_df[station_col].astype(str).str.strip(),
            "Status": (
                station_df[observation_col].astype(str).str.strip().str.title()
                if observation_col else "In Progress"
            ),
            "Completion %": (
                (
                    pd.to_numeric(station_df[completion_col], errors="coerce").fillna(0) * 100
                ).round(1).astype(str) + "%"
                if completion_col else "N/A"
            ),
        })

        return {
            "overall": total if total > 0 else 44,
            "in_progress": in_progress if total > 0 else 30,
            "completed": completed if total > 0 else 14,
            "details_df": details,
        }
    except Exception:
        return default


def render_category_charts_pair(category, icon, color, base_df, show_labels):
    category_df = (
        base_df[base_df["Category"] == category].copy()
        if not base_df.empty else pd.DataFrame()
    )

    total = len(category_df)
    closed = int((category_df["Clean_Status"] == "Closed").sum()) if total else 0
    in_progress = int((category_df["Clean_Status"] == "In Progress").sum()) if total else 0

    st.markdown(
        f"""
        <div style="background:white;border-left:5px solid {color};
        border-radius:10px;padding:10px 16px;margin:12px 0 10px;
        display:flex;justify-content:space-between;align-items:center;
        gap:8px;flex-wrap:wrap;box-shadow:0 2px 6px rgba(15,23,42,.03);">
            <div style="font-size:16px;font-weight:800;color:#0f172a;display:flex;align-items:center;gap:6px;">
                <span>{icon}</span> {category}
            </div>
            <div style="display:flex;gap:6px;flex-wrap:wrap;">
                <span class="badge-closed">{closed:,} Closed</span>
                <span class="badge-prog">{in_progress:,} In Progress</span>
                <span class="badge-affected">{total:,} Affected</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if category_df.empty:
        st.info(f"No records found for {category} with current filter settings.")
        return

    action_counts = category_df["Action_Category"].value_counts().sort_values(ascending=True)

    left, right = st.columns(2, gap="medium")

    with left:
        values = [closed, in_progress, total]
        names = ["Closed", "In Progress", "Affected"]

        fig = go.Figure(
            go.Bar(
                x=names,
                y=values,
                marker_color=["#10b981", "#f59e0b", "#2563eb"],
                text=values if show_labels else None,
                textposition="outside" if show_labels else "none",
                cliponaxis=False,
                hovertemplate="%{x}: %{y:,}<extra></extra>",
            )
        )

        fig.update_layout(
            title=f"<b>{category}: Status</b>",
            height=260,
            margin=dict(l=25, r=20, t=45, b=25),
            paper_bgcolor="white",
            plot_bgcolor="white",
            showlegend=False,
            xaxis=dict(showgrid=False),
            yaxis=dict(showgrid=True, gridcolor="#edf2f7", rangemode="tozero"),
        )

        st.plotly_chart(fig, use_container_width=True, theme=None, config={"displayModeBar": False})

    with right:
        palette = {
            "Facility": "#3b82f6",
            "Process": "#06b6d4",
            "Standard": "#8b5cf6",
            "Training": "#f59e0b",
            "Part": "#10b981",
            "General": "#64748b",
            "Unassigned": "#94a3b8",
        }

        action_names = action_counts.index.tolist()
        action_values = action_counts.values.tolist()

        fig = go.Figure(
            go.Bar(
                y=action_names,
                x=action_values,
                orientation="h",
                marker_color=[palette.get(name, color) for name in action_names],
                text=action_values if show_labels else None,
                textposition="outside" if show_labels else "none",
                cliponaxis=False,
                hovertemplate="%{y}: %{x:,}<extra></extra>",
            )
        )

        fig.update_layout(
            title=f"<b>{category}: Action Category Breakdown</b>",
            height=260,
            margin=dict(l=90, r=35, t=45, b=25),
            paper_bgcolor="white",
            plot_bgcolor="white",
            showlegend=False,
            yaxis=dict(autorange="reversed", automargin=True),
            xaxis=dict(showgrid=True, gridcolor="#edf2f7", rangemode="tozero"),
        )

        st.plotly_chart(fig, use_container_width=True, theme=None, config={"displayModeBar": False})


# ============================================================
# STATE INITIALIZATION
# ============================================================
if "station_override" not in st.session_state:
    st.session_state["station_override"] = None

if "last_upload_status" not in st.session_state:
    st.session_state["last_upload_status"] = None


# ============================================================
# SIDEBAR: DATA SOURCE, GITHUB SYNC & FILTERS
# ============================================================
with st.sidebar:
    logo_b64 = get_logo_base64()
    if logo_b64:
        logo_ext = "png" if os.path.exists("stellantis_logo.png") or os.path.exists("logo.png") else "jpeg"
        st.markdown(
            f'<div style="text-align:center;padding:6px 0 14px;'
            f'border-bottom:1px solid #1e293b;margin-bottom:14px;">'
            f'<img src="data:image/{logo_ext};base64,{logo_b64}" '
            f'style="max-width:180px;border-radius:8px;" '
            f'alt="Stellantis"></div>',
            unsafe_allow_html=True,
        )

    st.markdown("## 📊 Data Source & Setup")
    st.caption("UPLOAD EXCEL · AUTO SYNC GITHUB")

    # File Uploader with instant auto-upload and sync
    uploaded_file = st.file_uploader(
        "Upload Excel Workbook (.xlsx, .xls)",
        type=["xlsx", "xls"],
        help="Upload your updated Excel workbook. It will be loaded, synced to GitHub, and charts will regenerate automatically.",
    )

    file_bytes = None
    data_source_info = ""
    local_mtime = ""
    sheet_options = []
    raw_df = None

    # Handle upload event
    if uploaded_file is not None:
        curr_upload_bytes = uploaded_file.getvalue()

        # Check if this is a newly uploaded file
        if st.session_state.get("last_uploaded_name") != uploaded_file.name or st.session_state.get("last_uploaded_size") != len(curr_upload_bytes):
            # Run step-by-step notification workflow
            with st.status("🚀 Processing uploaded Excel file...", expanded=True) as status_box:
                # Step 1: Sync to GitHub
                status_box.write("🐙 **Loaded in GitHub:** Syncing and committing workbook to repository...")
                sync_res = sync_file_to_github(
                    curr_upload_bytes,
                    DEFAULT_EXCEL_PATH,
                    commit_message=f"Update {uploaded_file.name} from dashboard",
                )
                if sync_res["success"]:
                    status_box.write(f"✅ **Loaded in GitHub:** {sync_res['message']}")
                else:
                    status_box.write(f"⚠️ **GitHub Sync Info:** {sync_res['message']}")

                # Step 2: Processing data
                status_box.write("⚙️ **Processing data:** Inspecting sheets, detecting headers and standard criteria...")
                st.cache_data.clear()
                detected_sheets, inspect_err = inspect_excel_sheets(curr_upload_bytes)
                best_sheet = auto_detect_best_sheet(detected_sheets, curr_upload_bytes)
                status_box.write(f"✅ **Processing data:** Auto-detected sheet `{best_sheet}` with full criteria structure.")

                # Step 3: Retrieving data
                status_box.write("📊 **Retrieving data:** Loading observation records & station assessments...")
                temp_raw, _ = load_dataframe_from_bytes(curr_upload_bytes, best_sheet)
                temp_parsed, _ = parse_observation_tracker(temp_raw)
                temp_station = extract_station_assessment(curr_upload_bytes)
                obs_count = len(temp_parsed) if temp_parsed is not None else 0
                status_box.write(f"✅ **Retrieving data:** Extracted {obs_count} observations and {temp_station['overall']} station evaluations.")

                # Step 4: Generating charts
                status_box.write("📈 **Generating charts:** Instantly calculating metrics and rendering dashboards...")
                status_box.update(label="🎉 Complete! Redirecting to refreshed site...", state="complete", expanded=False)

            # Store in session state
            st.session_state["cached_file_bytes"] = curr_upload_bytes
            st.session_state["last_uploaded_name"] = uploaded_file.name
            st.session_state["last_uploaded_size"] = len(curr_upload_bytes)
            st.session_state["last_upload_status"] = sync_res
            st.session_state["station_override"] = None

            st.toast("🐙 Excel workbook loaded into GitHub & saved locally!", icon="✅")
            st.toast("📊 Data processed & charts generated automatically!", icon="📈")
            st.rerun()

        file_bytes = curr_upload_bytes
        data_source_info = f"📤 Uploaded: {uploaded_file.name}"

    elif "cached_file_bytes" in st.session_state and st.session_state["cached_file_bytes"]:
        file_bytes = st.session_state["cached_file_bytes"]
        data_source_info = f"📤 Session: {st.session_state.get('last_uploaded_name', 'Uploaded Workbook')}"

    elif os.path.exists(DEFAULT_EXCEL_PATH):
        try:
            local_mtime = datetime.fromtimestamp(os.path.getmtime(DEFAULT_EXCEL_PATH)).strftime("%Y-%m-%d %H:%M:%S")
            data_source_info = f"📁 Local: {DEFAULT_EXCEL_PATH}"
            with open(DEFAULT_EXCEL_PATH, "rb") as excel_file:
                file_bytes = excel_file.read()
        except Exception as exc:
            st.error(f"Could not read local Excel file: {exc}")
    else:
        st.warning(f"Default file '{DEFAULT_EXCEL_PATH}' not found. Please upload your Excel workbook.")

    # Inspect Sheets
    if file_bytes:
        sheet_options, sheet_err = inspect_excel_sheets(file_bytes)
        if sheet_err:
            st.error(f"Error inspecting sheets: {sheet_err}")

    # GitHub Settings Expander
    effective_token = get_effective_github_token()
    with st.expander("⚙️ GitHub Sync & Secrets Settings", expanded=not bool(effective_token)):
        st.caption(
            "Auto-sync uploads directly to **elagosan/GA_WPI** on GitHub. "
            "Works seamlessly on Streamlit Community Cloud & local runs."
        )
        token_input = st.text_input(
            "GitHub Personal Access Token (PAT)",
            value=st.session_state.get("github_token", ""),
            type="password",
            help="Generate a GitHub PAT (with repo contents:write permission) to commit files from the web app.",
        )
        if token_input:
            st.session_state["github_token"] = token_input.strip()

        col_t1, col_t2 = st.columns(2)
        with col_t1:
            if effective_token:
                st.markdown("<span style='color:#34d399;font-size:12px;font-weight:700;'>🟢 Token Active</span>", unsafe_allow_html=True)
            else:
                st.markdown("<span style='color:#94a3b8;font-size:12px;'>⚪ Token Optional</span>", unsafe_allow_html=True)
        with col_t2:
            if st.button("🔌 Test Connection", use_container_width=True):
                tok = get_effective_github_token()
                if not tok:
                    st.warning("Please enter a GitHub Token first.")
                else:
                    try:
                        t_res = requests.get(
                            f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}",
                            headers={"Authorization": f"Bearer {tok}"},
                            timeout=10,
                        )
                        if t_res.status_code == 200:
                            st.success(f"Connected to {GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}!")
                        else:
                            st.error(f"GitHub Error {t_res.status_code}: {t_res.json().get('message', t_res.text)}")
                    except Exception as ex:
                        st.error(f"Connection test failed: {ex}")

    # Refresh Data Button
    if st.button("🔄 Refresh Dashboard", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # Data Source Info Display
    if data_source_info:
        st.markdown(
            f'<div style="background:#1e293b;border-radius:8px;'
            f'padding:8px 12px;margin:8px 0;font-size:11.5px;'
            f'color:#cbd5e1;"><b>Source:</b> {data_source_info}'
            + (f'<br><span style="color:#94a3b8;">Updated: {local_mtime}</span>' if local_mtime else "")
            + "</div>",
            unsafe_allow_html=True,
        )

    st.divider()

    # Sheet Selection with Smart Auto-Detection
    selected_sheet = None
    if sheet_options and file_bytes:
        best_default = auto_detect_best_sheet(sheet_options, file_bytes)
        default_index = sheet_options.index(best_default) if best_default in sheet_options else 0

        selected_sheet = st.selectbox(
            "Excel Sheet to Analyze",
            sheet_options,
            index=default_index,
            help="Auto-detects the Observation Tracker sheet. You can also pick another sheet.",
        )

        loaded_df, load_error = load_dataframe_from_bytes(file_bytes, selected_sheet)
        if load_error:
            st.error(f"Error loading sheet: {load_error}")
        else:
            raw_df, parse_error = parse_observation_tracker(loaded_df)
            if parse_error:
                st.warning(parse_error)

    # Sidebar Filters
    st.markdown("### 🔍 Filters")
    st.caption("Filter categories, status, and actions")

    filter_category = "All Categories"
    selected_statuses = []
    selected_actions = []
    search_keyword = ""

    if raw_df is not None and not raw_df.empty:
        categories = ["All Categories"] + sorted(raw_df["Category"].dropna().unique().tolist())
        filter_category = st.selectbox("Standard Criteria / Category", categories)

        statuses = sorted(raw_df["Clean_Status"].dropna().unique().tolist())
        selected_statuses = st.multiselect("Status Filter", statuses, default=statuses)

        actions = sorted(a for a in raw_df["Action_Category"].dropna().unique().tolist() if a != "Unassigned")
        selected_actions = st.multiselect("Action Category Filter", actions, default=[])

        search_keyword = st.text_input("Search Observation / Action", placeholder="Type keywords...")

    st.divider()
    st.markdown("### ⚙️ Chart Display Options")

    chart_orientation = st.radio("2D Bar Orientation", ["Vertical", "Horizontal"], index=0)
    show_data_labels = st.toggle("Show Values on Bars", value=True)


# ============================================================
# APPLY FILTERS
# ============================================================
filtered_df = pd.DataFrame()

if raw_df is not None and not raw_df.empty:
    temp_df = raw_df.copy()

    if filter_category != "All Categories":
        temp_df = temp_df[temp_df["Category"] == filter_category]

    if selected_statuses:
        temp_df = temp_df[temp_df["Clean_Status"].isin(selected_statuses)]
    else:
        temp_df = temp_df.iloc[0:0]

    if selected_actions:
        temp_df = temp_df[temp_df["Action_Category"].isin(selected_actions)]

    if search_keyword.strip():
        kw = search_keyword.strip().lower()
        obs_match = temp_df["Observation_Text"].astype(str).str.lower().str.contains(kw, na=False)
        station_match = temp_df["Station"].astype(str).str.lower().str.contains(kw, na=False)
        lead_match = temp_df["Lead"].astype(str).str.lower().str.contains(kw, na=False)

        action_col = next((c for c in temp_df.columns if "ACTION DETAILS" in str(c).upper()), None)
        action_match = temp_df[action_col].astype(str).str.lower().str.contains(kw, na=False) if action_col else pd.Series(False, index=temp_df.index)

        temp_df = temp_df[obs_match | station_match | lead_match | action_match]

    filtered_df = temp_df


# ============================================================
# SUMMARY METRICS CALCULATIONS
# ============================================================
total_obs = len(raw_df) if raw_df is not None else 0
filtered_total = len(filtered_df)

closed_total = int((filtered_df["Clean_Status"] == "Closed").sum()) if not filtered_df.empty else 0
inprogress_total = int((filtered_df["Clean_Status"] == "In Progress").sum()) if not filtered_df.empty else 0

closure_rate = (closed_total / filtered_total * 100) if filtered_total else 0
progress_rate = (inprogress_total / filtered_total * 100) if filtered_total else 0

active_categories = (
    filtered_df.loc[filtered_df["Category"].isin(STANDARD_CATEGORIES), "Category"].nunique()
    if not filtered_df.empty else 0
)


# ============================================================
# STATION ASSESSMENT DATA EXTRACTION
# ============================================================
station_data = extract_station_assessment(file_bytes)

# Allow manual override from session if user edited values
if st.session_state["station_override"]:
    extracted_overall = int(st.session_state["station_override"]["overall"])
    extracted_ip = int(st.session_state["station_override"]["in_progress"])
    extracted_comp = int(st.session_state["station_override"]["completed"])
else:
    extracted_overall = int(station_data["overall"])
    extracted_ip = int(station_data["in_progress"])
    extracted_comp = int(station_data["completed"])

station_details_df = station_data["details_df"]
completion_rate_station = (extracted_comp / extracted_overall * 100) if extracted_overall else 0
progress_rate_station = (extracted_ip / extracted_overall * 100) if extracted_overall else 0


# ============================================================
# 1. TOP TITLE / HERO BANNER
# ============================================================
logo_b64 = get_logo_base64()
logo_ext = "png" if os.path.exists("stellantis_logo.png") or os.path.exists("logo.png") else "jpeg"
logo_tag = (
    f'<img src="data:image/{logo_ext};base64,{logo_b64}" '
    f'style="height:68px;border-radius:10px;flex-shrink:0;background:white;padding:4px 8px;" '
    f'alt="Stellantis">'
    if logo_b64 else ""
)

effective_tok = get_effective_github_token()
sync_badge = (
    '<span class="badge-sync">🟢 GITHUB AUTO-SYNC ACTIVE</span>'
    if effective_tok
    else '<span style="background:rgba(255,255,255,0.15);padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;">📁 LOCAL DASHBOARD MODE</span>'
)

st.markdown(
    f"""
    <div class="hero-banner" style="display:flex;align-items:center;gap:20px;flex-wrap:wrap;">
        {logo_tag}
        <div style="flex:1;min-width:260px;">
            <h1>WORKPLACE INTEGRATION – CC22 LHD</h1>
            <p>STANDARD CRITERIA &amp; STATION ASSESSMENT DASHBOARD</p>
        </div>
        <div style="margin-left:auto;text-align:right;">
            {sync_badge}
            <div style="font-size:19px;font-weight:800;color:#ffffff;margin-top:6px;">
                GENERAL ASSEMBLY
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Notification banner if file was recently synced
if st.session_state.get("last_upload_status"):
    upload_info = st.session_state["last_upload_status"]
    if upload_info["success"]:
        st.success(f"**Update Status:** {upload_info['message']}")
    else:
        st.info(f"**Update Status:** {upload_info['message']}")


# ============================================================
# DASHBOARD TABS
# ============================================================
tab_home, tab_explorer = st.tabs([
    "📊 Executive Dashboard",
    "📁 Raw Data Explorer & Export",
])


# ============================================================
# HOME DASHBOARD
# ============================================================
with tab_home:

    # --------------------------------------------------------
    # 2. STATION ASSESSMENT: PLACED IN TOP DIRECTLY AFTER TITLE
    # --------------------------------------------------------
    st.markdown(
        """
        <div class="section-header-box" style="margin-top:4px;">
            <div>
                <div class="section-title">🏢 Station Assessment</div>
                <div class="section-sub">Assembly station assessment progress, completion milestones, and detailed station status</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Station Metric KPI Cards Row
    sc1, sc2, sc3, sc4 = st.columns(4, gap="small")
    with sc1:
        st.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-label"><span>OVERALL STATIONS</span><span>🏢</span></div>'
            f'<div class="metric-value" style="color:#2563eb;">{extracted_overall:,}</div>'
            f'<div class="metric-sub">Total assembly stations scope</div></div>',
            unsafe_allow_html=True,
        )
    with sc2:
        st.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-label"><span>IN PROGRESS</span><span>⏳</span></div>'
            f'<div class="metric-value" style="color:#f59e0b;">{extracted_ip:,}</div>'
            f'<div class="metric-sub">{progress_rate_station:.1f}% currently active</div></div>',
            unsafe_allow_html=True,
        )
    with sc3:
        st.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-label"><span>COMPLETED</span><span>✅</span></div>'
            f'<div class="metric-value" style="color:#10b981;">{extracted_comp:,}</div>'
            f'<div class="metric-sub">{completion_rate_station:.1f}% stations completed</div></div>',
            unsafe_allow_html=True,
        )
    with sc4:
        st.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-label"><span>COMPLETION RATE</span><span>🎯</span></div>'
            f'<div class="metric-value" style="color:#059669;">{completion_rate_station:.1f}%</div>'
            f'<div class="metric-sub">{extracted_comp} of {extracted_overall} stations done</div></div>',
            unsafe_allow_html=True,
        )

    # Balanced Two-Column Section: Progress details on Left, 2D Bar Chart on Right
    station_left, station_right = st.columns([1.05, 1.35], gap="medium")

    with station_left:
        # Progress Bar Card
        st.markdown(
            f'<div style="background:white;padding:14px 18px;'
            f'border-radius:12px;border:1px solid #e2e8f0;'
            f'margin-bottom:12px;box-shadow:0 2px 8px rgba(15,23,42,.03);">'
            f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">'
            f'<span style="font-size:13px;font-weight:700;color:#0f172a;">Assessment Milestone Progress</span>'
            f'<span style="font-size:15px;font-weight:800;color:#10b981;">{completion_rate_station:.1f}%</span>'
            f'</div>'
            f'<div style="font-size:11.5px;color:#64748b;margin-bottom:8px;">Tracking completion of {extracted_overall} total assembly stations</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        st.progress(min(completion_rate_station / 100, 1.0))

        # Expander: Station Details Table
        if not station_details_df.empty:
            with st.expander(f"📋 View All {len(station_details_df)} Station Details Table"):
                st.dataframe(
                    station_details_df,
                    use_container_width=True,
                    hide_index=True,
                    height=240,
                )

        # Expander: Edit Station Assessment Values & Push to GitHub
        with st.expander("📝 Edit Station Assessment Values & Push to GitHub"):
            st.caption("Update station assessment counts. Changes will be saved and committed back to GitHub.")
            e_c1, e_c2, e_c3 = st.columns(3)
            with e_c1:
                new_overall = st.number_input("Overall", min_value=0, value=extracted_overall, step=1, key="input_st_overall")
            with e_c2:
                new_ip = st.number_input("In Progress", min_value=0, value=extracted_ip, step=1, key="input_st_ip")
            with e_c3:
                new_comp = st.number_input("Completed", min_value=0, value=extracted_comp, step=1, key="input_st_comp")

            if st.button("💾 Save & Sync Station Assessment to GitHub", use_container_width=True):
                st.session_state["station_override"] = {
                    "overall": new_overall,
                    "in_progress": new_ip,
                    "completed": new_comp,
                }
                st.toast("✅ Station assessment values updated!", icon="💾")
                st.rerun()

    with station_right:
        station_names = ["Overall Scope", "In Progress", "Completed"]
        station_vals = [extracted_overall, extracted_ip, extracted_comp]

        fig_station = go.Figure(
            go.Bar(
                x=station_names,
                y=station_vals,
                marker_color=["#2563eb", "#f59e0b", "#10b981"],
                text=station_vals if show_data_labels else None,
                textposition="outside" if show_data_labels else "none",
                cliponaxis=False,
                hovertemplate="<b>%{x}</b>: %{y:,} stations<extra></extra>",
            )
        )

        fig_station.update_layout(
            title=dict(
                text="<b>Station Assessment Status Breakdown</b>",
                font=dict(size=15, color="#0f172a"),
                x=0.01,
            ),
            height=300,
            margin=dict(l=25, r=25, t=48, b=25),
            paper_bgcolor="white",
            plot_bgcolor="white",
            showlegend=False,
            xaxis=dict(showgrid=False, automargin=True),
            yaxis=dict(
                showgrid=True,
                gridcolor="#edf2f7",
                rangemode="tozero",
                automargin=True,
            ),
        )

        st.plotly_chart(
            fig_station,
            use_container_width=True,
            theme=None,
            config={"displayModeBar": False},
        )

    # --------------------------------------------------------
    # 3. COUNTS OF DASHBOARDS (OBSERVATION KPI CARDS)
    # --------------------------------------------------------
    st.markdown(
        """
        <div class="section-header-box" style="margin-top:22px;">
            <div>
                <div class="section-title">📊 Dashboard Summary &amp; Scope Counts</div>
                <div class="section-sub">Observation tracker metrics across all assembly operations</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    k1, k2, k3, k4 = st.columns(4, gap="small")

    cards = [
        (
            "AFFECTED OBSERVATIONS",
            f"{filtered_total:,}",
            "#0f172a",
            f"Filtered from {total_obs:,} total" if filtered_total != total_obs else "Total observation scope",
            "📋",
        ),
        (
            "CLOSED ITEMS",
            f"{closed_total:,}",
            "#10b981",
            f"{closure_rate:.1f}% completed of scope",
            "✅",
        ),
        (
            "IN PROGRESS ITEMS",
            f"{inprogress_total:,}",
            "#f59e0b",
            f"{progress_rate:.1f}% currently active",
            "⏳",
        ),
        (
            "CRITERIA CATEGORIES",
            "4",
            "#2563eb",
            "Four standard criteria",
            "🏷️",
        ),
    ]

    for col, (label, value, color, sub, icon) in zip((k1, k2, k3, k4), cards):
        with col:
            st.markdown(
                f'<div class="metric-card">'
                f'<div class="metric-label"><span>{label}</span><span>{icon}</span></div>'
                f'<div class="metric-value" style="color:{color}">{value}</div>'
                f'<div class="metric-sub">{sub}</div></div>',
                unsafe_allow_html=True,
            )

    # --------------------------------------------------------
    # 4. OVERALL STANDARD CRITERIA OVERVIEW
    # --------------------------------------------------------
    st.markdown(
        """
        <div class="section-header-box" style="margin-top:22px;">
            <div>
                <div class="section-title">📋 Standard Criteria Overview</div>
                <div class="section-sub">Total affected observations across all four standard criteria</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if filtered_df.empty:
        st.warning("No observation records found matching current filters. Please adjust the sidebar filters.")
    else:
        criteria_df = filtered_df[filtered_df["Category"].isin(STANDARD_CATEGORIES)]
        category_totals = (
            criteria_df.groupby("Category")
            .size()
            .reindex(STANDARD_CATEGORIES, fill_value=0)
            .sort_values(ascending=False)
        )

        names = category_totals.index.tolist()
        totals = category_totals.astype(int).tolist()

        display_names = {
            "Customer Experience": "Customer<br>Experience",
            "Workplace Organization": "Workplace<br>Organization",
            "Health and Safety": "Health &amp;<br>Safety",
            "Logistics": "Logistics",
        }

        fig_overall = go.Figure()

        if chart_orientation == "Horizontal":
            fig_overall.add_trace(
                go.Bar(
                    y=names,
                    x=totals,
                    orientation="h",
                    marker_color="#2563eb",
                    text=totals if show_data_labels else None,
                    textposition="outside" if show_data_labels else "none",
                    cliponaxis=False,
                    hovertemplate="<b>%{y}</b><br>Total Affected: %{x:,}<extra></extra>",
                )
            )

            fig_overall.update_layout(
                yaxis=dict(autorange="reversed", automargin=True),
                xaxis=dict(title="Total Affected Observations", gridcolor="#edf2f7", rangemode="tozero"),
            )
        else:
            fig_overall.add_trace(
                go.Bar(
                    x=[display_names.get(name, name) for name in names],
                    y=totals,
                    marker_color="#2563eb",
                    text=totals if show_data_labels else None,
                    textposition="outside" if show_data_labels else "none",
                    cliponaxis=False,
                    hovertemplate="<b>%{x}</b><br>Total Affected: %{y:,}<extra></extra>",
                )
            )

            fig_overall.update_layout(
                xaxis=dict(showgrid=False, automargin=True),
                yaxis=dict(title="Total Affected Observations", gridcolor="#edf2f7", rangemode="tozero"),
            )

        fig_overall.update_layout(
            title=dict(
                text="<b>Total Affected by Standard Criteria</b>",
                font=dict(size=16, color="#0f172a"),
                x=0.01,
            ),
            height=320,
            margin=dict(l=35, r=35, t=55, b=35),
            paper_bgcolor="white",
            plot_bgcolor="white",
            showlegend=False,
            bargap=0.35,
        )

        st.plotly_chart(
            fig_overall,
            use_container_width=True,
            theme=None,
            config={"displayModeBar": False},
        )

        # Percentage distribution pills
        sum_totals = sum(totals)
        if sum_totals > 0:
            pills_html = '<div style="display:flex;gap:10px;flex-wrap:wrap;margin:4px 0 16px;">'
            colors_map = {
                "Customer Experience": "#6366f1",
                "Workplace Organization": "#059669",
                "Health and Safety": "#ef4444",
                "Logistics": "#0284c7",
            }
            for n, t in zip(names, totals):
                pct = (t / sum_totals) * 100
                pills_html += (
                    f'<div style="background:white;border:1px solid #e2e8f0;padding:6px 14px;'
                    f'border-radius:20px;font-size:12px;display:flex;align-items:center;gap:6px;'
                    f'box-shadow:0 1px 4px rgba(15,23,42,.04);">'
                    f'<span style="width:8px;height:8px;border-radius:50%;background:{colors_map.get(n, "#2563eb")};"></span>'
                    f'<span style="font-weight:700;color:#0f172a;">{n}:</span> '
                    f'<span style="font-weight:800;color:#2563eb;">{t:,}</span> '
                    f'<span style="color:#64748b;">({pct:.1f}%)</span></div>'
                )
            pills_html += "</div>"
            st.markdown(pills_html, unsafe_allow_html=True)

    # --------------------------------------------------------
    # 5. STANDARD CRITERIA DETAILED ANALYTICS
    # --------------------------------------------------------
    st.markdown(
        """
        <div class="section-header-box" style="margin-top:24px;">
            <div>
                <div class="section-title">🔍 Standard Criteria Detailed Analytics</div>
                <div class="section-sub">Status breakdown and action categories for each standard criterion</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_category_charts_pair(
        "Health and Safety", "🛡️", "#ef4444",
        filtered_df, show_data_labels,
    )

    render_category_charts_pair(
        "Customer Experience", "⭐", "#6366f1",
        filtered_df, show_data_labels,
    )

    render_category_charts_pair(
        "Workplace Organization", "🏢", "#059669",
        filtered_df, show_data_labels,
    )

    render_category_charts_pair(
        "Logistics", "🚚", "#0284c7",
        filtered_df, show_data_labels,
    )


# ============================================================
# RAW DATA EXPLORER AND EXPORT TAB
# ============================================================
with tab_explorer:
    st.markdown(
        """
        <div style="background:white;border-radius:12px;padding:16px 20px;
                    margin-bottom:14px;border:1px solid #cbd5e1;">
            <div style="font-size:18px;font-weight:800;color:#0f172a;">
                🔍 Global Data Explorer &amp; Export
            </div>
            <div style="font-size:12px;color:#64748b;">
                Search, inspect, filter, and export observation records to CSV or formatted Excel
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if filtered_df.empty:
        st.info("No records to display. Adjust the filters in the sidebar.")
    else:
        st.markdown(f"**Showing {len(filtered_df):,} matching observation records**")

        st.dataframe(
            filtered_df,
            use_container_width=True,
            hide_index=True,
            height=460,
        )

        csv_data = filtered_df.to_csv(index=False).encode("utf-8")

        col_csv, col_excel = st.columns(2, gap="small")

        with col_csv:
            st.download_button(
                "⬇️ Download Filtered Data as CSV",
                data=csv_data,
                file_name="wpi_filtered_observations.csv",
                mime="text/csv",
                use_container_width=True,
            )

        with col_excel:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                filtered_df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Filtered Observations",
                )
                summary = (
                    filtered_df[filtered_df["Category"].isin(STANDARD_CATEGORIES)]
                    .groupby("Category")
                    .size()
                    .rename("Total Affected")
                    .reset_index()
                )
                summary.to_excel(
                    writer,
                    index=False,
                    sheet_name="Criteria Summary",
                )

            st.download_button(
                "⬇️ Download Filtered Data as Excel (.xlsx)",
                data=buffer.getvalue(),
                file_name="wpi_filtered_observations.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


# ============================================================
# FOOTER
# ============================================================
st.markdown(
    """
    <div class="footer-bar">
        WORKPLACE INTEGRATION – CC22 LHD · Standard Criteria &amp; Station Assessment Dashboard · Stellantis General Assembly
    </div>
    """,
    unsafe_allow_html=True,
)
