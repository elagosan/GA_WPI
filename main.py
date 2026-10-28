import base64
import io
import os
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import requests

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
        padding-top: .35rem;
        padding-bottom: 1rem;
        max-width: 100%;
    }

    .hero-banner {
        padding: 22px 26px;
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
        margin: 2px 0 6px;
        font-weight: 800;
    }

    .hero-banner p {
        color: #e0f2fe;
        margin: 0;
        font-size: 15px;
        font-weight: 600;
    }

    .metric-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 15px 18px;
        box-shadow: 0 2px 8px rgba(15,23,42,.04);
        margin-bottom: 8px;
    }

    .metric-label {
        font-size: 11px;
        font-weight: 700;
        color: #64748b;
        letter-spacing: .5px;
        text-transform: uppercase;
        margin-bottom: 5px;
        display: flex;
        justify-content: space-between;
    }

    .metric-value {
        font-size: 27px;
        font-weight: 800;
        line-height: 1.1;
    }

    .metric-sub {
        font-size: 11px;
        color: #64748b;
        margin-top: 5px;
    }

    .home-panel-title {
        font-size: 17px;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 3px;
    }

    .home-panel-sub {
        font-size: 12px;
        color: #64748b;
        margin-bottom: 10px;
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
        margin-top: 18px;
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
        padding: 9px 16px !important;
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
                    return base64.b64encode(
                        image_file.read()
                    ).decode("utf-8")
            except OSError:
                continue
    return None


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

        df = df.loc[
            :, ~df.columns.astype(str).str.contains(r"^Unnamed")
        ]

        df.columns = [clean_column_name(c) for c in df.columns]

        return df, None

    except Exception as exc:
        return None, str(exc)


def parse_observation_tracker(raw_df):
    if raw_df is None or raw_df.empty:
        return None, "The selected Excel sheet is empty."

    df = raw_df.copy()

    criteria_col = next(
        (
            c for c in df.columns
            if any(
                key in c.upper()
                for key in (
                    "STANDARD", "CRITERIA", "CRETERIA", "CATEGORY"
                )
            )
        ),
        None,
    )

    if not criteria_col:
        for col in df.columns:
            sample = " ".join(
                df[col].astype(str).head(15)
            ).upper()

            if any(
                key in sample
                for key in (
                    "HEALTH", "CUSTOMER", "LOGISTIC", "WORKPLACE"
                )
            ):
                criteria_col = col
                break

    status_col = next(
        (c for c in df.columns if "STATUS" in c.upper()),
        None,
    )

    if not criteria_col or not status_col:
        return None, (
            "Could not detect the Standard Criteria and Status "
            "columns in this sheet."
        )

    action_col = next(
        (
            c for c in df.columns
            if "ACTION" in c.upper()
            and any(
                key in c.upper()
                for key in ("CATEG", "TYPE", "CLASS")
            )
        ),
        None,
    )

    station_col = next(
        (c for c in df.columns if "STATION" in c.upper()),
        None,
    )

    lead_col = next(
        (
            c for c in df.columns
            if any(key in c.upper() for key in ("RESP", "LEAD", "PILOT"))
        ),
        None,
    )

    observation_col = next(
        (
            c for c in df.columns
            if "OBSERVATION" in c.upper()
            and "DATE" not in c.upper()
        ),
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
        df["Station"] = (
            df[station_col].fillna("Unknown").astype(str).str.strip()
        )
    else:
        df["Station"] = "General"

    if lead_col:
        df["Lead"] = (
            df[lead_col].fillna("Unassigned").astype(str).str.strip()
        )
    else:
        df["Lead"] = "Unassigned"

    if observation_col:
        df["Observation_Text"] = (
            df[observation_col].fillna("").astype(str).str.strip()
        )
    else:
        df["Observation_Text"] = ""

    return df, None


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

        target_sheet = next(
            (
                sheet for sheet in (
                    "STATION WISE PROGRESS (2)",
                    "STATION WISE PROGRESS",
                )
                if sheet in excel.sheet_names
            ),
            None,
        )

        if not target_sheet:
            return default

        df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=target_sheet,
        )

        if df.empty:
            return default

        station_col = df.columns[0]
        station_df = df.dropna(subset=[station_col]).copy()

        station_df = station_df[
            station_df[station_col].astype(str).str.strip().str.upper()
            != "STATION #"
        ]

        observation_col = next(
            (
                c for c in station_df.columns
                if "OBSERVATION" in str(c).upper()
                or (
                    len(df) > 0
                    and "OBSERVATION" in str(df[c].iloc[0]).upper()
                )
            ),
            None,
        )

        completion_col = next(
            (
                c for c in station_df.columns
                if "COMPLETION" in str(c).upper()
                or (
                    len(df) > 0
                    and "COMPLETION" in str(df[c].iloc[0]).upper()
                )
                or "%" in str(c)
            ),
            None,
        )

        total = len(station_df)
        completed = 0
        in_progress = 0

        if observation_col:
            status = (
                station_df[observation_col]
                .astype(str).str.strip().str.lower()
            )

            completed = int(
                status.isin(
                    ["completed", "complete", "closed", "done"]
                ).sum()
            )

            in_progress = int(
                status.isin(
                    ["inprogress", "in progress", "progress"]
                ).sum()
            )

        elif completion_col:
            completion = pd.to_numeric(
                station_df[completion_col],
                errors="coerce",
            ).fillna(0)

            completed = int((completion >= 1).sum())
            in_progress = total - completed

        details = pd.DataFrame({
            "Station": station_df[station_col].astype(str).str.strip(),
            "Status": (
                station_df[observation_col].astype(str).str.strip().str.title()
                if observation_col else "In Progress"
            ),
            "Completion %": (
                (
                    pd.to_numeric(
                        station_df[completion_col],
                        errors="coerce",
                    ).fillna(0) * 100
                ).round(1).astype(str) + "%"
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


def render_category_charts_pair(
    category, icon, color, base_df, show_labels
):
    category_df = (
        base_df[base_df["Category"] == category].copy()
        if not base_df.empty else pd.DataFrame()
    )

    total = len(category_df)

    closed = (
        int((category_df["Clean_Status"] == "Closed").sum())
        if total else 0
    )

    in_progress = (
        int((category_df["Clean_Status"] == "In Progress").sum())
        if total else 0
    )

    st.markdown(
        f"""
        <div style="background:white;border-left:5px solid {color};
        border-radius:10px;padding:9px 14px;margin:10px 0 8px;
        display:flex;justify-content:space-between;align-items:center;
        gap:8px;flex-wrap:wrap;">
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
        st.info(
            f"No records found for {category} with the current filters."
        )
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
    st.caption("LOAD EXCEL FILE & RUN DASHBOARD")

    uploaded_file = st.file_uploader(
        "Upload Excel file (.xlsx, .xls)",
        type=["xlsx", "xls"],
        help="Upload an updated Excel workbook to generate the dashboard.",
    )

    file_bytes = None
    data_source_info = ""
    local_mtime = ""
    sheet_options = []
    raw_df = None

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        data_source_info = f"📤 Uploaded: {uploaded_file.name}"

        try:
            sheet_options = pd.ExcelFile(
                io.BytesIO(file_bytes)
            ).sheet_names
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

            sheet_options = pd.ExcelFile(
                io.BytesIO(file_bytes)
            ).sheet_names

        except Exception as exc:
            st.error(f"Could not read Excel file: {exc}")

    else:
        st.warning(
            f"Default file '{DEFAULT_EXCEL_PATH}' was not found. "
            "Upload your Excel file."
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
                f'<br><span style="color:#94a3b8;">Updated: '
                f'{local_mtime}</span>'
                if local_mtime else ""
            )
            + "</div>",
            unsafe_allow_html=True,
        )

    st.divider()

    selected_sheet = None

    if sheet_options and file_bytes:
        default_index = next(
            (
                i for i, name in enumerate(sheet_options)
                if "OBSERVATION" in name.upper()
                or "TRACKER" in name.upper()
            ),
            0,
        )

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
        temp_df = temp_df[
            temp_df["Category"] == filter_category
        ]

    if selected_statuses:
        temp_df = temp_df[
            temp_df["Clean_Status"].isin(selected_statuses)
        ]
    else:
        temp_df = temp_df.iloc[0:0]

    if selected_actions:
        temp_df = temp_df[
            temp_df["Action_Category"].isin(selected_actions)
        ]

    if search_keyword.strip():
        keyword = search_keyword.strip().lower()

        observation_match = (
            temp_df["Observation_Text"]
            .astype(str)
            .str.lower()
            .str.contains(keyword, na=False)
        )

        action_column = next(
            (
                c for c in temp_df.columns
                if "ACTION DETAILS" in c.upper()
            ),
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
            action_match = pd.Series(
                False,
                index=temp_df.index,
            )

        temp_df = temp_df[
            observation_match | action_match
        ]

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

# Count ONLY the four standard criteria.
active_categories = (
    filtered_df.loc[
        filtered_df["Category"].isin(STANDARD_CATEGORIES),
        "Category",
    ].nunique()
    if not filtered_df.empty else 0
)


# ============================================================
# HEADER - NO STELLANTIS WPI OPERATIONS LABEL
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
    f'style="height:70px;border-radius:10px;flex-shrink:0;" '
    f'alt="Stellantis Logo">'
    if logo_b64 else ""
)

st.markdown(
    f"""
    <div class="hero-banner"
         style="display:flex;align-items:center;gap:22px;flex-wrap:wrap;">
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
# TOP KPI CARDS
# ============================================================
k1, k2, k3, k4 = st.columns(4, gap="small")

cards = [
    (
        "AFFECTED OBSERVATIONS",
        f"{filtered_total:,}",
        "#0f172a",
        "Filtered from total"
        if filtered_total != total_obs
        else "Total observation scope",
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

for col, (label, value, color, sub, icon) in zip(
    (k1, k2, k3, k4),
    cards,
):
    with col:
        st.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-label"><span>{label}</span>'
            f'<span>{icon}</span></div>'
            f'<div class="metric-value" style="color:{color}">'
            f'{value}</div>'
            f'<div class="metric-sub">{sub}</div></div>',
            unsafe_allow_html=True,
        )


# ============================================================
# DASHBOARD TABS
# ============================================================
tab_home, tab_explorer = st.tabs([
    "📊 Standard Criteria & Station Dashboard",
    "📁 Raw Data Explorer & Export",
])


# ============================================================
# HOME DASHBOARD
# ============================================================
with tab_home:

    # Full-width section heading
    st.markdown(
        """
        <div style="width:100%;margin:0 0 10px;padding:0 0 8px;
                    border-bottom:2px solid #e2e8f0;">
            <div style="font-size:20px;font-weight:800;color:#0f172a;">
                Standard Criteria & Station Assessment Overview
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --------------------------------------------------------
    # OVERALL CRITERIA: FULL WIDTH, TOTAL AFFECTED ONLY
    # --------------------------------------------------------
    st.markdown(
        '<div class="home-panel-title">📊 Overall Criteria</div>'
        '<div class="home-panel-sub">'
        'Total affected observations across all four standard criteria'
        '</div>',
        unsafe_allow_html=True,
    )

    if filtered_df.empty:
        st.warning(
            "No observation data loaded. Check the Excel file and filters."
        )
    else:
        criteria_df = filtered_df[
            filtered_df["Category"].isin(STANDARD_CATEGORIES)
        ]

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

        fig = go.Figure()

        if chart_orientation == "Horizontal":
            fig.add_trace(
                go.Bar(
                    y=names,
                    x=totals,
                    orientation="h",
                    marker_color="#2563eb",
                    text=totals if show_data_labels else None,
                    textposition=(
                        "outside" if show_data_labels else "none"
                    ),
                    cliponaxis=False,
                    hovertemplate=(
                        "<b>%{y}</b><br>"
                        "Total Affected: %{x:,}<extra></extra>"
                    ),
                )
            )

            fig.update_layout(
                yaxis=dict(
                    autorange="reversed",
                    automargin=True,
                ),
                xaxis=dict(
                    title="Total Affected",
                    gridcolor="#edf2f7",
                    rangemode="tozero",
                ),
            )

        else:
            fig.add_trace(
                go.Bar(
                    x=[
                        display_names.get(name, name)
                        for name in names
                    ],
                    y=totals,
                    marker_color="#2563eb",
                    text=totals if show_data_labels else None,
                    textposition=(
                        "outside" if show_data_labels else "none"
                    ),
                    cliponaxis=False,
                    hovertemplate=(
                        "<b>%{x}</b><br>"
                        "Total Affected: %{y:,}<extra></extra>"
                    ),
                )
            )

            fig.update_layout(
                xaxis=dict(
                    showgrid=False,
                    automargin=True,
                ),
                yaxis=dict(
                    title="Total Affected",
                    gridcolor="#edf2f7",
                    rangemode="tozero",
                ),
            )

        fig.update_layout(
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

        # Uses the full available content width.
        st.plotly_chart(
            fig,
            use_container_width=True,
            theme=None,
            config={"displayModeBar": False},
        )

    # --------------------------------------------------------
    # STATION ASSESSMENT: BELOW CHART, 50% WIDTH
    # --------------------------------------------------------
    station_data = extract_station_assessment(file_bytes)

    extracted_overall = int(station_data["overall"])
    extracted_ip = int(station_data["in_progress"])
    extracted_comp = int(station_data["completed"])

    station_details_df = station_data["details_df"]

    st.markdown(
        '<div class="home-panel-title">🏢 Station Assessment</div>'
        '<div class="home-panel-sub">'
        'Station assessment status extracted from Excel'
        '</div>',
        unsafe_allow_html=True,
    )

    with st.expander("📝 View / Edit Station Assessment Values"):
        st.caption("Values extracted from the station progress sheet.")

        l1_overall = st.number_input(
            "Overall assessments",
            min_value=0,
            max_value=1000000,
            value=extracted_overall,
            step=1,
            key="station_overall",
        )

        l1_in_progress = st.number_input(
            "In progress",
            min_value=0,
            max_value=1000000,
            value=extracted_ip,
            step=1,
            key="station_inprogress",
        )

        l1_completed = st.number_input(
            "Completed",
            min_value=0,
            max_value=1000000,
            value=extracted_comp,
            step=1,
            key="station_completed",
        )

        if l1_in_progress + l1_completed > l1_overall:
            st.warning("In Progress + Completed exceeds Overall.")

    completion_rate = (
        l1_completed / l1_overall * 100 if l1_overall else 0
    )
    progress_rate_station = (
        l1_in_progress / l1_overall * 100 if l1_overall else 0
    )

    # Balanced two-panel layout fills the available page width.
    station_left, station_right = st.columns([0.95, 1.35], gap="medium")

    with station_left:
        s1, s2, s3 = st.columns(3, gap="small")

        station_cards = [
            (s1, "OVERALL", l1_overall, "#2563eb", "Total scope"),
            (s2, "IN PROGRESS", l1_in_progress, "#f59e0b",
             f"{progress_rate_station:.1f}% active"),
            (s3, "COMPLETED", l1_completed, "#10b981",
             f"{completion_rate:.1f}% done"),
        ]

        for col, label, value, color, sub in station_cards:
            with col:
                st.markdown(
                    f'<div class="metric-card">'
                    f'<div class="metric-label">{label}</div>'
                    f'<div class="metric-value" style="color:{color}">'
                    f'{value:,}</div>'
                    f'<div class="metric-sub">{sub}</div></div>',
                    unsafe_allow_html=True,
                )

        st.markdown(
            f'<div style="background:white;padding:12px 14px;'
            f'border-radius:10px;border:1px solid #e2e8f0;'
            f'margin:4px 0 10px;display:flex;'
            f'justify-content:space-between;align-items:center;gap:12px;">'
            f'<span style="font-size:12px;font-weight:700;color:#334155;">'
            f'Assessment Completion Rate</span>'
            f'<span style="font-size:16px;font-weight:800;color:#059669;">'
            f'{completion_rate:.1f}%</span></div>',
            unsafe_allow_html=True,
        )
        st.progress(min(completion_rate / 100, 1.0))

        if not station_details_df.empty:
            with st.expander(
                f"📋 View All {len(station_details_df)} Station Details"
            ):
                st.dataframe(
                    station_details_df,
                    use_container_width=True,
                    hide_index=True,
                    height=260,
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
                hovertemplate="%{x}: %{y:,}<extra></extra>",
            )
        )

        fig_station.update_layout(
            title=dict(
                text="<b>Station Assessment Status</b>",
                font=dict(size=15, color="#0f172a"),
                x=0.01,
            ),
            height=350,
            margin=dict(l=25, r=25, t=48, b=35),
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
    # DETAILED ANALYTICS
    # --------------------------------------------------------
    st.markdown(
        """
        <div style="margin-top:14px;margin-bottom:8px;
                    padding-bottom:8px;border-bottom:2px solid #e2e8f0;">
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
# RAW DATA EXPLORER AND EXPORT
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
        st.info(
            "No records to display. Adjust the filters in the sidebar."
        )

    else:
        st.markdown(
            f"**Showing {len(filtered_df):,} matching observation records**"
        )

        st.dataframe(
            filtered_df,
            use_container_width=True,
            hide_index=True,
            height=460,
        )

        csv_data = filtered_df.to_csv(
            index=False
        ).encode("utf-8")

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

            with pd.ExcelWriter(
                buffer,
                engine="openpyxl",
            ) as writer:
                filtered_df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Filtered Observations",
                )

                summary = (
                    filtered_df[
                        filtered_df["Category"].isin(
                            STANDARD_CATEGORIES
                        )
                    ]
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
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
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
