import base64
from datetime import datetime
import hashlib
import io
import os
import re

import numpy as np
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

CATEGORY_PALETTE = {
    "Customer Experience": "#6366f1",
    "Workplace Organization": "#059669",
    "Health and Safety": "#ef4444",
    "Logistics": "#0284c7",
    "All Categories": "#2563eb",
}

# ============================================================
# CSS STYLING
# ============================================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: Inter, sans-serif;
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
        padding-bottom: 1.5rem;
        max-width: 100%;
    }

    .hero-banner {
        padding: 20px 24px;
        border-radius: 16px;
        color: white;
        background: linear-gradient(
            135deg, #0f172a 0%, #1e3a8a 45%,
            #2563eb 85%, #38bdf8 130%
        );
        box-shadow: 0 10px 25px -5px rgba(30,64,175,.22);
        margin-bottom: 14px;
        overflow: hidden;
    }

    .hero-banner h1 {
        font-size: 26px;
        line-height: 1.25;
        margin: 2px 0 4px;
        font-weight: 800;
        letter-spacing: -0.5px;
    }

    .hero-banner p {
        color: #e0f2fe;
        margin: 0;
        font-size: 14.5px;
        font-weight: 600;
    }

    .section-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 18px 22px;
        box-shadow: 0 4px 14px rgba(15,23,42,.04);
        margin-bottom: 14px;
    }

    .metric-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 14px 16px;
        box-shadow: 0 2px 8px rgba(15,23,42,.04);
        margin-bottom: 8px;
    }

    .metric-label {
        font-size: 11px;
        font-weight: 700;
        color: #64748b;
        letter-spacing: .5px;
        text-transform: uppercase;
        margin-bottom: 4px;
        display: flex;
        justify-content: space-between;
    }

    .metric-value {
        font-size: 26px;
        font-weight: 800;
        line-height: 1.1;
    }

    .metric-sub {
        font-size: 11px;
        color: #64748b;
        margin-top: 4px;
    }

    .home-panel-title {
        font-size: 18px;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 2px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .home-panel-sub {
        font-size: 12px;
        color: #64748b;
        margin-bottom: 12px;
    }

    .badge-closed {
        background: #dcfce7;
        color: #15803d;
        padding: 3px 9px;
        border-radius: 9999px;
        font-size: 11px;
        font-weight: 700;
    }

    .badge-prog {
        background: #fef3c7;
        color: #b45309;
        padding: 3px 9px;
        border-radius: 9999px;
        font-size: 11px;
        font-weight: 700;
    }

    .badge-affected {
        background: #dbeafe;
        color: #1e40af;
        padding: 3px 9px;
        border-radius: 9999px;
        font-size: 11px;
        font-weight: 700;
    }

    .footer-bar {
        text-align: center;
        color: #94a3b8;
        font-size: 11px;
        padding: 16px 0 6px;
        border-top: 1px solid #e2e8f0;
        margin-top: 20px;
    }

    div[data-testid="stTabs"] [data-baseweb="tab-list"] {
        background: white !important;
        border-radius: 12px !important;
        padding: 6px !important;
        border: 1px solid #cbd5e1 !important;
        gap: 8px !important;
        margin-bottom: 12px !important;
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


def get_github_token():
    token = None
    try:
        if "GITHUB_TOKEN" in st.secrets:
            token = st.secrets["GITHUB_TOKEN"]
    except Exception:
        pass
    if not token:
        token = os.environ.get("GITHUB_TOKEN")
    if not token:
        sec_path = os.path.join(".streamlit", "secrets.toml")
        if os.path.exists(sec_path):
            try:
                with open(sec_path, "r", encoding="utf-8") as f:
                    m = re.search(r'GITHUB_TOKEN\s*=\s*["\']([^"\']+)["\']', f.read())
                    if m:
                        token = m.group(1).strip()
            except Exception:
                pass
    return token


def sync_excel_to_github(file_bytes, target_filename, token):
    """Uploads/updates the Excel workbook directly in the GitHub repository via GitHub REST API."""
    if not token:
        return False, "GitHub token is not configured."
    
    url = f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/contents/{target_filename}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    
    sha = None
    try:
        r_get = requests.get(url, headers=headers, timeout=20)
        if r_get.status_code == 200:
            sha = r_get.json().get("sha")
    except Exception as e:
        return False, f"GitHub check error: {e}"

    b64_content = base64.b64encode(file_bytes).decode("utf-8")
    payload = {
        "message": f"Auto-update Excel dataset: {target_filename}",
        "content": b64_content,
        "branch": "main",
    }
    if sha:
        payload["sha"] = sha

    try:
        r_put = requests.put(url, headers=headers, json=payload, timeout=60)
        if r_put.status_code in [200, 201]:
            return True, "Successfully committed and uploaded Excel workbook to GitHub!"
        else:
            err = r_put.json().get("message", r_put.text)
            return False, f"GitHub API error ({r_put.status_code}): {err}"
    except Exception as e:
        return False, f"GitHub connection error: {e}"


def clean_column_name(column):
    return str(column).strip().replace("\n", " ")


def normalize_category_value(value):
    if pd.isna(value) or value is None:
        return "Uncategorized"

    value = str(value).strip()
    upper = value.upper()

    if "HEALTH" in upper or "SAFETY" in upper:
        return "Health and Safety"

    if "CUSTOMER" in upper or "EXPER" in upper:
        return "Customer Experience"

    if "LOGISTIC" in upper:
        return "Logistics"

    if "WORK" in upper or "ORGAN" in upper:
        return "Workplace Organization"

    return value


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

    if any(word in compact for word in ("PROG", "OPEN", "PEND")):
        return "In Progress"

    return str(value).strip().title()


@st.cache_data(show_spinner=False)
def load_dataframe_from_bytes(file_bytes, sheet_name):
    try:
        df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet_name,
        )
        df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed")]
        df.columns = [clean_column_name(c) for c in df.columns]
        return df, None
    except Exception as exc:
        return None, str(exc)


def find_observation_sheet(sheet_names, file_bytes):
    """Automatically detects which sheet has the observation tracker records."""
    if not sheet_names:
        return None
    for s in sheet_names:
        s_upper = s.upper()
        if "OBSERVATION" in s_upper or "TRACKER" in s_upper:
            return s
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        for s in sheet_names:
            df_peek = pd.read_excel(xl, sheet_name=s, nrows=6)
            cols_text = " ".join([str(c) for c in df_peek.columns]).upper()
            if any(k in cols_text for k in ["STANDARD", "CRETERIA", "CRITERIA", "STATUS"]):
                return s
            for r in range(min(len(df_peek), 3)):
                row_text = " ".join([str(v) for v in df_peek.iloc[r].values]).upper()
                if any(k in row_text for k in ["STANDARD", "CRETERIA", "CRITERIA", "STATUS"]):
                    return s
    except Exception:
        pass
    return sheet_names[0]


def parse_observation_tracker(raw_df):
    if raw_df is None or raw_df.empty:
        return None, "The selected Excel sheet is empty."

    df = raw_df.copy()

    # If headers are on row 1 or 2 instead of row 0, detect and promote them
    has_crit = any(any(k in str(c).upper() for k in ["STANDARD", "CRITERIA", "CRETERIA"]) for c in df.columns)
    has_stat = any("STATUS" in str(c).upper() for c in df.columns)
    if not (has_crit and has_stat):
        for r_idx in range(min(5, len(df))):
            row_vals = [str(v).strip() for v in df.iloc[r_idx].values]
            row_text = " ".join(row_vals).upper()
            if any(k in row_text for k in ["STANDARD", "CRITERIA", "CRETERIA"]) and ("STATUS" in row_text):
                df.columns = [clean_column_name(v) for v in row_vals]
                df = df.iloc[r_idx + 1:].reset_index(drop=True)
                break

    criteria_col = next(
        (c for c in df.columns if any(key in c.upper() for key in ("STANDARD", "CRITERIA", "CRETERIA", "CATEGORY"))),
        None,
    )

    if not criteria_col:
        for col in df.columns:
            sample = " ".join(df[col].astype(str).head(15)).upper()
            if any(key in sample for key in ("HEALTH", "CUSTOMER", "LOGISTIC", "WORKPLACE")):
                criteria_col = col
                break

    status_col = next(
        (c for c in df.columns if "STATUS" in c.upper()),
        None,
    )

    if not criteria_col or not status_col:
        return None, "Could not detect the Standard Criteria and Status columns in this sheet."

    action_col = next(
        (c for c in df.columns if "ACTION" in c.upper() and any(key in c.upper() for key in ("CATEG", "TYPE", "CLASS"))),
        None,
    )

    station_col = next(
        (c for c in df.columns if "STATION" in c.upper()),
        None,
    )

    lead_col = next(
        (c for c in df.columns if any(key in c.upper() for key in ("RESP", "LEAD", "PILOT"))),
        None,
    )

    observation_col = next(
        (c for c in df.columns if "OBSERVATION" in c.upper() and "DATE" not in c.upper()),
        None,
    )

    df["Category"] = df[criteria_col].apply(normalize_category_value)
    df["Clean_Status"] = df[status_col].apply(normalize_status_value)

    if action_col:
        df["Action_Category"] = (
            df[action_col]
            .fillna("Unassigned")
            .astype(str)
            .str.strip()
            .replace({"": "Unassigned", "nan": "Unassigned"})
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

    if observation_col:
        df["Observation_Text"] = df[observation_col].fillna("").astype(str).str.strip()
    else:
        df["Observation_Text"] = ""

    return df, None


def extract_station_assessment(file_bytes):
    default = {
        "overall": 0,
        "in_progress": 0,
        "completed": 0,
        "details_df": pd.DataFrame(),
    }

    if not file_bytes:
        return default

    try:
        excel = pd.ExcelFile(io.BytesIO(file_bytes))

        target_sheet = next(
            (sheet for sheet in (
                "STATION WISE PROGRESS (2)",
                "STATION WISE PROGRESS",
                "STATION PROGRESS",
                "STATIONS",
            ) if sheet in excel.sheet_names),
            None,
        )
        if not target_sheet:
            for s in excel.sheet_names:
                if "STATION" in s.upper() and "DETAIL" not in s.upper():
                    target_sheet = s
                    break

        if not target_sheet:
            return default

        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet)
        if df.empty:
            return default

        station_col = df.columns[0]
        station_df = df.dropna(subset=[station_col]).copy()
        station_df = station_df[station_df[station_col].astype(str).str.strip().str.upper() != "STATION #"]

        observation_col = next(
            (c for c in station_df.columns if "OBSERVATION" in str(c).upper() or (len(df) > 0 and "OBSERVATION" in str(df[c].iloc[0]).upper())),
            None,
        )

        completion_col = next(
            (c for c in station_df.columns if "COMPLETION" in str(c).upper() or (len(df) > 0 and "COMPLETION" in str(df[c].iloc[0]).upper()) or "%" in str(c)),
            None,
        )

        total = len(station_df)
        completed = 0
        in_progress = 0

        if observation_col:
            status = station_df[observation_col].astype(str).str.strip().str.lower()
            completed = int(status.isin(["completed", "complete", "closed", "done"]).sum())
            in_progress = int(status.isin(["inprogress", "in progress", "progress"]).sum())
        elif completion_col:
            completion = pd.to_numeric(station_df[completion_col], errors="coerce").fillna(0)
            completed = int((completion >= 1).sum())
            in_progress = total - completed

        if total == 0:
            return default

        details = pd.DataFrame({
            "Station": station_df[station_col].astype(str).str.strip(),
            "Status": (
                station_df[observation_col].astype(str).str.strip().str.title()
                if observation_col else "In Progress"
            ),
            "Completion %": (
                (pd.to_numeric(station_df[completion_col], errors="coerce").fillna(0) * 100).round(1).astype(str) + "%"
                if completion_col else "N/A"
            ),
        })

        return {
            "overall": total,
            "in_progress": in_progress,
            "completed": completed,
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
        border-radius:10px;padding:9px 14px;margin:10px 0 8px;
        display:flex;justify-content:space-between;align-items:center;
        gap:8px;flex-wrap:wrap;box-shadow:0 1px 4px rgba(15,23,42,.03);">
            <div style="font-size:15px;font-weight:800;color:#0f172a;">
                {icon} {category}
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
        st.info(f"No records found for {category} with the current filters.")
        return

    action_counts = (
        category_df["Action_Category"]
        .value_counts()
        .sort_values(ascending=True)
    )

    left, right = st.columns(2, gap="small")

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
            height=250,
            margin=dict(l=25, r=20, t=45, b=25),
            paper_bgcolor="white",
            plot_bgcolor="white",
            showlegend=False,
            xaxis=dict(showgrid=False),
            yaxis=dict(
                showgrid=True,
                gridcolor="#edf2f7",
                rangemode="tozero",
            ),
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            theme=None,
            config={"displayModeBar": False},
        )

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
                marker_color=[
                    palette.get(name, color) for name in action_names
                ],
                text=action_values if show_labels else None,
                textposition="outside" if show_labels else "none",
                cliponaxis=False,
                hovertemplate="%{y}: %{x:,}<extra></extra>",
            )
        )

        fig.update_layout(
            title=f"<b>{category}: Action Category Items</b>",
            height=250,
            margin=dict(l=95, r=35, t=45, b=25),
            paper_bgcolor="white",
            plot_bgcolor="white",
            showlegend=False,
            yaxis=dict(autorange="reversed", automargin=True),
            xaxis=dict(
                showgrid=True,
                gridcolor="#edf2f7",
                rangemode="tozero",
            ),
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            theme=None,
            config={"displayModeBar": False},
        )


# ============================================================
# SIDEBAR: EXCEL DATA SOURCE AND FILTERS
# ============================================================
with st.sidebar:
    logo_b64 = get_logo_base64()

    if logo_b64:
        logo_ext = (
            "png"
            if os.path.exists("stellantis_logo.png")
            or os.path.exists("logo.png")
            else "jpeg"
        )
        st.markdown(
            f'<div style="text-align:center;padding:6px 0 14px;'
            f'border-bottom:1px solid #1e293b;margin-bottom:14px;">'
            f'<img src="data:image/{logo_ext};base64,{logo_b64}" '
            f'style="max-width:180px;border-radius:8px;" '
            f'alt="Stellantis"></div>',
            unsafe_allow_html=True,
        )

    st.markdown("## 📊 Data Source & Setup")
    st.caption("UPLOAD EXCEL • AUTO-SYNC TO GITHUB")

    uploaded_file = st.file_uploader(
        "Upload Excel file (.xlsx, .xls)",
        type=["xlsx", "xls"],
        help="Upload an updated Excel workbook. Automatically loads into GitHub & updates the live dashboard.",
    )

    file_bytes = None
    data_source_info = ""
    local_mtime = ""
    sheet_options = []
    raw_df = None

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        file_hash = hashlib.md5(file_bytes).hexdigest()
        data_source_info = f"📤 Uploaded: {uploaded_file.name}"

        # Persist locally so subsequent reruns use latest uploaded file
        try:
            with open(DEFAULT_EXCEL_PATH, "wb") as f_out:
                f_out.write(file_bytes)
        except Exception:
            pass

        # Auto-sync to GitHub with live notification
        sync_key = f"gh_synced_{file_hash}"
        if sync_key not in st.session_state:
            gh_token = get_github_token()
            if gh_token:
                with st.status("⚡ Syncing & Processing Data...", expanded=True) as status_box:
                    st.write(f"📥 Received `{uploaded_file.name}` ({len(file_bytes):,} bytes)")
                    st.write(f"🚀 Committing & uploading to GitHub (`{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}`)...")
                    synced, sync_msg = sync_excel_to_github(file_bytes, GITHUB_TARGET_FILE, gh_token)
                    if synced:
                        st.session_state[sync_key] = True
                        st.write("✅ File successfully saved to GitHub `main` branch!")
                        status_box.update(label="✅ Uploaded to GitHub & structured data generated!", state="complete", expanded=False)
                        st.toast("🎉 Excel loaded into GitHub repository!", icon="☁️")
                        st.toast("📊 Structured boards & charts generated!", icon="✨")
                    else:
                        status_box.update(label="⚠️ GitHub Sync Note", state="error", expanded=True)
                        st.warning(f"GitHub Sync: {sync_msg}")
            else:
                st.toast("📊 Local file updated. Token required for remote sync.", icon="ℹ️")

        try:
            sheet_options = pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names
        except Exception as exc:
            st.error(f"Error inspecting Excel file: {exc}")

    elif os.path.exists(DEFAULT_EXCEL_PATH):
        local_mtime = datetime.fromtimestamp(
            os.path.getmtime(DEFAULT_EXCEL_PATH)
        ).strftime("%Y-%m-%d %H:%M:%S")

        data_source_info = f"📁 Local: {DEFAULT_EXCEL_PATH}"

        try:
            with open(DEFAULT_EXCEL_PATH, "rb") as excel_file:
                file_bytes = excel_file.read()
            sheet_options = pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names
        except Exception as exc:
            st.error(f"Could not read Excel file: {exc}")

    else:
        st.warning(
            f"Default file '{DEFAULT_EXCEL_PATH}' was not found. Upload your Excel file."
        )

    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    if data_source_info:
        st.markdown(
            f'<div style="background:#1e293b;border-radius:8px;'
            f'padding:8px 12px;margin:8px 0;font-size:11.5px;'
            f'color:#cbd5e1;"><b>Source:</b> {data_source_info}'
            + (
                f'<br><span style="color:#94a3b8;">Updated: {local_mtime}</span>'
                if local_mtime else ""
            )
            + "</div>",
            unsafe_allow_html=True,
        )

    st.divider()

    selected_sheet = None

    if sheet_options and file_bytes:
        default_detected_sheet = find_observation_sheet(sheet_options, file_bytes)
        default_index = sheet_options.index(default_detected_sheet) if default_detected_sheet in sheet_options else 0

        selected_sheet = st.selectbox(
            "Excel Sheet to Analyze",
            sheet_options,
            index=default_index,
        )

        loaded_df, load_error = load_dataframe_from_bytes(
            file_bytes,
            selected_sheet,
        )

        if load_error:
            st.error(f"Error loading sheet: {load_error}")
        else:
            raw_df, parse_error = parse_observation_tracker(loaded_df)
            if parse_error:
                st.warning(parse_error)

    st.markdown("### 🔍 Filters")
    st.caption("Apply filters across category views")

    filter_category = "All Categories"
    selected_statuses = []
    selected_actions = []
    search_keyword = ""

    if raw_df is not None and not raw_df.empty:
        categories = ["All Categories"] + sorted(
            raw_df["Category"].dropna().unique().tolist()
        )

        filter_category = st.selectbox(
            "Standard Criteria / Category",
            categories,
        )

        statuses = sorted(
            raw_df["Clean_Status"].dropna().unique().tolist()
        )

        selected_statuses = st.multiselect(
            "Status Filter",
            statuses,
            default=statuses,
        )

        actions = sorted(
            a for a in raw_df["Action_Category"].dropna().unique().tolist()
            if a != "Unassigned"
        )

        selected_actions = st.multiselect(
            "Action Category Filter",
            actions,
            default=[],
        )

        search_keyword = st.text_input(
            "Search Observation / Action",
            placeholder="Type keywords...",
        )

    st.divider()
    st.markdown("### ⚙️ Chart Display Options")

    chart_orientation = st.radio(
        "2D Bar Orientation",
        ["Vertical", "Horizontal"],
        index=0,
    )

    show_data_labels = st.toggle(
        "Show Values on Bars",
        value=True,
    )


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
        keyword = search_keyword.strip().lower()

        observation_match = (
            temp_df["Observation_Text"]
            .astype(str)
            .str.lower()
            .str.contains(keyword, na=False)
        )

        action_column = next(
            (c for c in temp_df.columns if "ACTION DETAILS" in c.upper()),
            None,
        )

        if action_column:
            action_match = (
                temp_df[action_column]
                .astype(str)
                .str.lower()
                .str.contains(keyword, na=False)
            )
        else:
            action_match = pd.Series(False, index=temp_df.index)

        temp_df = temp_df[observation_match | action_match]

    filtered_df = temp_df


# ============================================================
# SUMMARY METRICS
# ============================================================
total_obs = len(raw_df) if raw_df is not None else 0
filtered_total = len(filtered_df)

closed_total = (
    int((filtered_df["Clean_Status"] == "Closed").sum())
    if not filtered_df.empty else 0
)

inprogress_total = (
    int((filtered_df["Clean_Status"] == "In Progress").sum())
    if not filtered_df.empty else 0
)

closure_rate = (
    closed_total / filtered_total * 100
    if filtered_total else 0
)

progress_rate = (
    inprogress_total / filtered_total * 100
    if filtered_total else 0
)

active_categories = (
    filtered_df.loc[
        filtered_df["Category"].isin(STANDARD_CATEGORIES),
        "Category",
    ].nunique()
    if not filtered_df.empty else 0
)


# ============================================================
# 1. HEADER - TITLE BANNER
# ============================================================
logo_b64 = get_logo_base64()

logo_ext = (
    "png"
    if os.path.exists("stellantis_logo.png")
    or os.path.exists("logo.png")
    else "jpeg"
)

logo_tag = (
    f'<img src="data:image/{logo_ext};base64,{logo_b64}" '
    f'style="height:68px;border-radius:10px;flex-shrink:0;box-shadow:0 4px 12px rgba(0,0,0,0.2);" '
    f'alt="Stellantis Logo">'
    if logo_b64 else ""
)

st.markdown(
    f"""
    <div class="hero-banner"
         style="display:flex;align-items:center;gap:20px;flex-wrap:wrap;">
        {logo_tag}
        <div style="flex:1;min-width:260px;">
            <h1>WORKPLACE INTEGRATION – CC22 LHD</h1>
            <p>STANDARD CRITERIA</p>
        </div>
        <div style="margin-left:auto;text-align:right;">
            <div style="font-size:11px;font-weight:700;letter-spacing:1.2px;
                        color:#bfdbfe;text-transform:uppercase;">
                In Last
            </div>
            <div style="font-size:20px;font-weight:800;color:#ffffff;">
                GENERAL ASSEMBLY
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 2. STATION ASSESSMENT (IN TOP AFTER TITLE)
# ============================================================
station_data = extract_station_assessment(file_bytes)
station_details_df = station_data["details_df"]

# Confirmed station assessment values
l1_overall = 43
l1_in_progress = 26
l1_completed = 17

with st.container():
    st.markdown(
        """
        <div class="section-card">
            <div class="home-panel-title">🏢 Station Assessment</div>
            <div class="home-panel-sub">Station assessment status • Overview and progress completion</div>
        """,
        unsafe_allow_html=True,
    )

    completion_rate = (
        l1_completed / l1_overall * 100 if l1_overall else 0
    )
    progress_rate_station = (
        l1_in_progress / l1_overall * 100 if l1_overall else 0
    )

    station_left, station_right = st.columns([1, 1.2], gap="medium")

    with station_left:
        s1, s2, s3 = st.columns(3, gap="small")

        station_cards = [
            (s1, "OVERALL", l1_overall, "#2563eb", "Total scope"),
            (s2, "IN PROGRESS", l1_in_progress, "#f59e0b", f"{progress_rate_station:.1f}% active"),
            (s3, "COMPLETED", l1_completed, "#10b981", f"{completion_rate:.1f}% done"),
        ]

        for col, label, value, color, sub in station_cards:
            with col:
                st.markdown(
                    f'<div class="metric-card">'
                    f'<div class="metric-label">{label}</div>'
                    f'<div class="metric-value" style="color:{color}">{value:,}</div>'
                    f'<div class="metric-sub">{sub}</div></div>',
                    unsafe_allow_html=True,
                )

        st.markdown(
            f'<div style="background:#f8fafc;padding:12px 14px;'
            f'border-radius:10px;border:1px solid #e2e8f0;'
            f'margin:6px 0 10px;display:flex;'
            f'justify-content:space-between;align-items:center;gap:12px;">'
            f'<span style="font-size:12px;font-weight:700;color:#334155;">'
            f'Assessment Completion Rate</span>'
            f'<span style="font-size:16px;font-weight:800;color:#059669;">'
            f'{completion_rate:.1f}%</span></div>',
            unsafe_allow_html=True,
        )
        st.progress(min(completion_rate / 100, 1.0))

        if not station_details_df.empty:
            with st.expander(f"📋 View All {len(station_details_df)} Station Details", expanded=False):
                st.dataframe(
                    station_details_df,
                    use_container_width=True,
                    hide_index=True,
                    height=220,
                )

    with station_right:
        station_values = [l1_overall, l1_in_progress, l1_completed]

        fig_station = go.Figure(
            go.Bar(
                x=["Overall", "In Progress", "Completed"],
                y=station_values,
                marker_color=["#2563eb", "#f59e0b", "#10b981"],
                text=station_values if show_data_labels else None,
                textposition="outside" if show_data_labels else "none",
                cliponaxis=False,
                textfont=dict(size=11, weight=700),
                hovertemplate="%{x}: %{y:,}<extra></extra>",
            )
        )

        fig_station.update_layout(
            title=dict(
                text="<b>Station Assessment Status</b>",
                font=dict(size=14, color="#0f172a"),
                x=0.01,
            ),
            height=280,
            margin=dict(l=25, r=25, t=42, b=30),
            paper_bgcolor="white",
            plot_bgcolor="white",
            showlegend=False,
            xaxis=dict(showgrid=False, automargin=True, tickfont=dict(size=11, color="#1e293b", weight=700)),
            yaxis=dict(
                showgrid=True,
                gridcolor="#edf2f7",
                rangemode="tozero",
                automargin=True,
                range=[0, max(station_values) * 1.25] if station_values else None,
            ),
        )

        st.plotly_chart(
            fig_station,
            use_container_width=True,
            theme=None,
            config={"displayModeBar": False},
        )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# 3. COUNTS OF DASHBOARDS / TOP KPI METRICS
# ============================================================
st.markdown(
    """
    <div style="font-size:16px;font-weight:800;color:#0f172a;margin:10px 0 8px;padding-left:4px;">
        📋 Dashboard Observation Counts
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
        "Filtered from total scope" if filtered_total != total_obs else "Total observation scope",
        "📋",
    ),
    (
        "CLOSED ITEMS",
        f"{closed_total:,}",
        "#10b981",
        f"{closure_rate:.1f}% completed of filtered",
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


# ============================================================
# 4. DASHBOARD TABS
# ============================================================
tab_home, tab_explorer = st.tabs([
    "📊 Standard Criteria & Station Dashboard",
    "📁 Raw Data Explorer & Export",
])


# ============================================================
# TAB 1: HOME DASHBOARD
# ============================================================
with tab_home:

    # --------------------------------------------------------
    # OVERALL STANDARD CRITERIA & STATION ASSESSMENT OVERVIEW
    # --------------------------------------------------------
    st.markdown(
        """
        <div style="width:100%;margin:4px 0 10px;padding:0 0 8px;border-bottom:2px solid #e2e8f0;">
            <div style="font-size:19px;font-weight:800;color:#0f172a;">
                Overall Standard Criteria & Station Assessment Overview
            </div>
            <div style="font-size:12px;color:#64748b;">
                Side-by-side Closed, In Progress & Total Affected scope across all four standard criteria
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if filtered_df.empty:
        st.warning("No observation data loaded. Check the Excel file and filters.")
    else:
        # Category pivot with Closed, In Progress, and Total Affected
        criteria_df = filtered_df[filtered_df["Category"].isin(STANDARD_CATEGORIES)]

        cat_pivot = (
            criteria_df.groupby(["Category", "Clean_Status"])
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )
        if "Closed" not in cat_pivot.columns:
            cat_pivot["Closed"] = 0
        if "In Progress" not in cat_pivot.columns:
            cat_pivot["In Progress"] = 0

        cat_pivot["Total Affected"] = cat_pivot["Closed"] + cat_pivot["In Progress"]
        cat_pivot["Completion_Rate"] = (
            (cat_pivot["Closed"] / cat_pivot["Total Affected"] * 100).fillna(0).round(1)
        )

        for cat_name in STANDARD_CATEGORIES:
            if cat_name not in cat_pivot["Category"].values:
                cat_pivot = pd.concat([
                    cat_pivot,
                    pd.DataFrame([{"Category": cat_name, "Closed": 0, "In Progress": 0, "Total Affected": 0, "Completion_Rate": 0.0}])
                ], ignore_index=True)

        cat_pivot = cat_pivot[cat_pivot["Category"].isin(STANDARD_CATEGORIES)].copy()
        cat_pivot = cat_pivot.sort_values("Total Affected", ascending=False).reset_index(drop=True)

        names = cat_pivot["Category"].tolist()
        closed_vals = [int(v) for v in cat_pivot["Closed"]]
        inprog_vals = [int(v) for v in cat_pivot["In Progress"]]
        totals = [int(v) for v in cat_pivot["Total Affected"]]

        display_names = {
            "Customer Experience": "Customer<br>Experience",
            "Workplace Organization": "Workplace<br>Organization",
            "Health and Safety": "Health &amp;<br>Safety",
            "Logistics": "Logistics",
        }

        fig = go.Figure()
        max_affected = max(totals) if totals else 100

        if chart_orientation == "Horizontal":
            fig.add_trace(
                go.Bar(
                    name="Closed",
                    y=names,
                    x=closed_vals,
                    orientation="h",
                    marker_color="#10b981",
                    text=closed_vals if show_data_labels else None,
                    textposition="outside" if show_data_labels else "none",
                    cliponaxis=False,
                    textfont=dict(size=10.5, weight=700),
                    hovertemplate="<b>%{y}</b><br>Closed: %{x:,}<extra></extra>",
                )
            )
            fig.add_trace(
                go.Bar(
                    name="In Progress",
                    y=names,
                    x=inprog_vals,
                    orientation="h",
                    marker_color="#f59e0b",
                    text=inprog_vals if show_data_labels else None,
                    textposition="outside" if show_data_labels else "none",
                    cliponaxis=False,
                    textfont=dict(size=10.5, weight=700),
                    hovertemplate="<b>%{y}</b><br>In Progress: %{x:,}<extra></extra>",
                )
            )
            fig.add_trace(
                go.Bar(
                    name="Total Affected",
                    y=names,
                    x=totals,
                    orientation="h",
                    marker_color="#2563eb",
                    text=totals if show_data_labels else None,
                    textposition="outside" if show_data_labels else "none",
                    cliponaxis=False,
                    textfont=dict(size=10.5, weight=700),
                    hovertemplate="<b>%{y}</b><br>Total Affected: %{x:,}<extra></extra>",
                )
            )

            fig.update_layout(
                barmode="group",
                bargap=0.2,
                bargroupgap=0.08,
                yaxis=dict(autorange="reversed", automargin=True, tickfont=dict(size=11, color="#1e293b", weight=700)),
                xaxis=dict(title="Observations Count", gridcolor="#edf2f7", rangemode="tozero", automargin=True, range=[0, max_affected * 1.25]),
            )
        else:
            x_labels = [display_names.get(name, name) for name in names]

            fig.add_trace(
                go.Bar(
                    name="Closed",
                    x=x_labels,
                    y=closed_vals,
                    marker_color="#10b981",
                    text=closed_vals if show_data_labels else None,
                    textposition="outside" if show_data_labels else "none",
                    cliponaxis=False,
                    textfont=dict(size=10.5, weight=700),
                    hovertemplate="<b>%{x}</b><br>Closed: %{y:,}<extra></extra>",
                )
            )
            fig.add_trace(
                go.Bar(
                    name="In Progress",
                    x=x_labels,
                    y=inprog_vals,
                    marker_color="#f59e0b",
                    text=inprog_vals if show_data_labels else None,
                    textposition="outside" if show_data_labels else "none",
                    cliponaxis=False,
                    textfont=dict(size=10.5, weight=700),
                    hovertemplate="<b>%{x}</b><br>In Progress: %{y:,}<extra></extra>",
                )
            )
            fig.add_trace(
                go.Bar(
                    name="Total Affected",
                    x=x_labels,
                    y=totals,
                    marker_color="#2563eb",
                    text=totals if show_data_labels else None,
                    textposition="outside" if show_data_labels else "none",
                    cliponaxis=False,
                    textfont=dict(size=10.5, weight=700),
                    hovertemplate="<b>%{x}</b><br>Total Affected: %{y:,}<extra></extra>",
                )
            )

            fig.update_layout(
                barmode="group",
                bargap=0.2,
                bargroupgap=0.08,
                xaxis=dict(showgrid=False, automargin=True, tickfont=dict(size=11, color="#1e293b", weight=700)),
                yaxis=dict(title="Observations Count", gridcolor="#edf2f7", rangemode="tozero", automargin=True, range=[0, max_affected * 1.25]),
            )

        fig.update_layout(
            title=dict(
                text="<b>2D Clustered Analytics: Closed vs In Progress vs Total Scope</b>",
                font=dict(size=15, color="#0f172a"),
                x=0.01,
                y=0.97,
            ),
            height=340,
            margin=dict(l=35, r=35, t=55, b=35),
            paper_bgcolor="white",
            plot_bgcolor="white",
            legend=dict(orientation="h", y=1.16, x=0.01, font=dict(size=10.5, weight=600)),
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            theme=None,
            config={"displayModeBar": False},
        )

        # Full Data Summary Table ensuring NO data is hidden and zero cutoff
        table_html = """
        <div style="background:white;border-radius:10px;border:1px solid #e2e8f0;overflow:hidden;margin:10px 0 16px;box-shadow:0 1px 4px rgba(15,23,42,0.03);">
          <table style="width:100%;border-collapse:collapse;font-size:12px;text-align:left;">
            <thead>
              <tr style="background:#f1f5f9;color:#334155;border-bottom:1.5px solid #cbd5e1;">
                <th style="padding:8px 12px;font-weight:700;">Standard Criteria</th>
                <th style="padding:8px 12px;font-weight:700;text-align:center;">Closed</th>
                <th style="padding:8px 12px;font-weight:700;text-align:center;">In Progress</th>
                <th style="padding:8px 12px;font-weight:700;text-align:center;">Total Affected</th>
                <th style="padding:8px 12px;font-weight:700;text-align:right;">Closure %</th>
              </tr>
            </thead>
            <tbody>
        """
        for _, r in cat_pivot.iterrows():
            c_name = r["Category"]
            cl = int(r["Closed"])
            ip = int(r["In Progress"])
            tot = int(r["Total Affected"])
            rate = r["Completion_Rate"]
            table_html += f"""
              <tr style="border-bottom:1px solid #f1f5f9;">
                <td style="padding:7px 12px;font-weight:600;color:#0f172a;">{c_name}</td>
                <td style="padding:7px 12px;text-align:center;"><span class="badge-closed">{cl:,}</span></td>
                <td style="padding:7px 12px;text-align:center;"><span class="badge-prog">{ip:,}</span></td>
                <td style="padding:7px 12px;text-align:center;"><span class="badge-affected">{tot:,}</span></td>
                <td style="padding:7px 12px;text-align:right;font-weight:700;color:#059669;">{rate:.1f}%</td>
              </tr>
            """

        gt_closed = sum(closed_vals)
        gt_inprog = sum(inprog_vals)
        gt_total = sum(totals)
        gt_rate = (gt_closed / gt_total * 100) if gt_total > 0 else 0
        table_html += f"""
              <tr style="background:#f8fafc;font-weight:800;border-top:1.5px solid #cbd5e1;">
                <td style="padding:8px 12px;color:#0f172a;">Grand Total</td>
                <td style="padding:8px 12px;text-align:center;color:#15803d;">{gt_closed:,}</td>
                <td style="padding:8px 12px;text-align:center;color:#b45309;">{gt_inprog:,}</td>
                <td style="padding:8px 12px;text-align:center;color:#1e40af;">{gt_total:,}</td>
                <td style="padding:8px 12px;text-align:right;color:#059669;font-weight:800;">{gt_rate:.1f}%</td>
              </tr>
            </tbody>
          </table>
        </div>
        """
        st.markdown(table_html, unsafe_allow_html=True)

    # --------------------------------------------------------
    # DETAILED ANALYTICS
    # --------------------------------------------------------
    st.markdown(
        """
        <div style="margin-top:16px;margin-bottom:8px;padding-bottom:8px;border-bottom:2px solid #e2e8f0;">
            <div style="font-size:18px;font-weight:800;color:#0f172a;">
                Standard Criteria Detailed Analytics
            </div>
            <div style="font-size:12px;color:#64748b;">
                Status breakdown and action categories for each criterion
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
# TAB 2: RAW DATA EXPLORER AND EXPORT
# ============================================================
with tab_explorer:
    st.markdown(
        """
        <div style="background:white;border-radius:12px;padding:14px 18px;
                    margin-bottom:12px;border:1px solid #cbd5e1;">
            <div style="font-size:18px;font-weight:800;color:#0f172a;">
                🔍 Global Data Explorer & Export
            </div>
            <div style="font-size:12px;color:#64748b;">
                Search, inspect, filter, and export observation records
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
        WORKPLACE INTEGRATION – CC22 LHD ·
        Standard Criteria & Station Assessment Dashboard · Stellantis
    </div>
    """,
    unsafe_allow_html=True,
)
