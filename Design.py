# ══════════════════════════════════════════════════════════════════════════════
# DesignPulse — Weekly Design Task Tracker
# Built with Streamlit + Google Sheets (gspread)
# Visual style: ClientPulse aesthetic (DM Sans, cream bg, dark sidebar)
# Deployed on: Streamlit Cloud (secrets via st.secrets)

# ══════════════════════════════════════════════════════════════════════════════
# SHEET COLUMN STRUCTURE (create your Google Sheet with these exact headers)
#
# Each weekly tab should have these columns in Row 1:
#   Project | Platform | Assigned To | Task | Status | Revision No | Start Date | End Date | Comments
#
# - "Assigned To" should match the designer's label exactly
# - Revision rows: same Task name, Status = "Revision (n)", Revision No = n
# - Correction rows: same Task name, Status = "Correction (n)", Revision No = n
# - Weekly tabs named like: "Week 1", "Week Jun 02", etc. (anything except "Design Closure Timeline")
# ══════════════════════════════════════════════════════════════════════════════

import io
import re
import time
import json
import datetime
from typing import Dict, List, Optional, Tuple

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

try:
    from streamlit_autorefresh import st_autorefresh
    HAS_AUTOREFRESH = True
except ImportError:
    HAS_AUTOREFRESH = False


# ──────────────────────────────────────────────────────────────────────────────
# CONFIG — Edit these
# ──────────────────────────────────────────────────────────────────────────────

REFRESH_INTERVAL = 60  # seconds

# Sheet URL — read from Streamlit Cloud secrets
SHEET_URL = st.secrets["SHEET_URL"]

# Accounts — add designers and lead here
ACCOUNTS = {
    "lead_1": {
        "username": "Kumar",
        "password": "leadpass",
        "role": "lead",
        "label": "Kumar",
    },
    "designer_1": {
        "username": "Dharnu",
        "password": "1234",
        "role": "designer",
        "label": "Dharnu",
    },
    "designer_2": {
        "username": "Vignesh",
        "password": "1234",
        "role": "designer",
        "label": "Vignesh",
    },
    "designer_3": {
        "username": "Sanjay",
        "password": "samanthasanjay",
        "role": "designer",
        "label": "Sanjay",
    },
    # Add more designers here following the same pattern
}

DESIGNER_LABELS = [v["label"] for v in ACCOUNTS.values() if v["role"] == "designer"]

STATUSES = [
    "Open",
    "In Progress",
    "In Internal Review",
    "In Client Review",
    "Re Work",
    "Hold",
    "Completed",
]

PLATFORMS = ["Shopify", "Webflow", "Zoketo", "Other"]

# Sheet columns — order matters for appending rows
SHEET_COLUMNS = ["Project", "Platform", "Assigned To", "Task", "Status",
                 "Revision No", "Start Date", "End Date", "Comments"]

# ──────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DesignPulse",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────────────────────
# STYLES — ClientPulse aesthetic
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');
:root {
    --ink:#111111; --cream:#faf8f4; --stone:#ece7df; --muted:#6f6a61;
    --border:#dbd5cc; --white:#ffffff; --soft:#f4f1eb;
    --done-bg:#ebf7ef;   --done-fg:#276749;
    --prog-bg:#fff6de;   --prog-fg:#8a6100;
    --hold-bg:#fdecec;   --hold-fg:#9b2c2c;
    --review-bg:#eaf4fb; --review-fg:#125d85;
    --open-bg:#efefef;   --open-fg:#474747;
    --rework-bg:#fdf3e7; --rework-fg:#b45309;
    --rev-bg:#f3e8ff;    --rev-fg:#6b21a8;
    --corr-bg:#fce7f3;   --corr-fg:#9d174d;
}
html, body, [class*="css"] { font-family:'DM Sans',sans-serif; color:var(--ink); }
.stApp { background:var(--cream); }
#MainMenu, footer, header { visibility:hidden; }
.block-container { padding-top:1rem; padding-bottom:1rem; }

[data-testid="stSidebar"] { background:var(--ink) !important; }
[data-testid="stSidebar"] * { color:#f2efe9 !important; }
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p { color:#bfb8ad !important; font-size:.78rem; letter-spacing:.08em; text-transform:uppercase; }
[data-testid="stSidebar"] .stSelectbox>div>div,
[data-testid="stSidebar"] .stTextInput>div>div>input,
[data-testid="stSidebar"] .stTextArea textarea {
    background:rgba(255,255,255,.08) !important;
    border:1px solid rgba(255,255,255,.15) !important;
    color:#faf8f4 !important; border-radius:8px !important;
}
div[data-baseweb="input"] { border:1px solid transparent !important; box-shadow:none !important; }
div[data-baseweb="input"]:focus-within { border:1px solid #111111 !important; box-shadow:none !important; outline:none !important; }

h1,h2,h3,h4,h5,h6 { font-family:'DM Serif Display',serif; letter-spacing:.01em; }

.stButton>button, .stDownloadButton>button {
    background:var(--ink) !important; color:#ffffff !important;
    border:none !important; border-radius:8px !important;
    padding:.55rem 1.15rem !important; font-weight:500 !important;
    transition:all .2s ease !important;
}
.stButton>button:hover, .stDownloadButton>button:hover {
    background:#000000 !important; transform:translateY(-1px);
}

/* Login */
.login-wrapper { width:100%; display:flex; justify-content:center; align-items:center; min-height:40vh; }
.login-card-wrap { width:100%; max-width:420px; }
.login-badge { display:inline-block; background:#cd0000; color:#ffffff !important; font-size:.68rem; font-weight:600; letter-spacing:.12em; text-transform:uppercase; padding:5px 14px; border-radius:999px; margin-bottom:1.25rem; }
.login-title { font-family:'DM Serif Display',serif; font-size:2.4rem; color:#111111; margin-bottom:.3rem; line-height:1.2; }
.login-sub { color:#6f6a61; font-size:.88rem; margin-bottom:1.75rem; line-height:1.5; }
.login-footer-note { color:#aaa49c; font-size:.72rem; margin-top:1.25rem; letter-spacing:.04em; text-align:center; }
label { color:#111111 !important; font-weight:500; }

/* Banners */
.banner { border-radius:10px; padding:.8rem 1rem; font-size:.84rem; margin-bottom:.8rem; }
.banner-ok   { background:#eef6f1; border-left:3px solid #3f7d5a; color:#25543d; }
.banner-warn { background:#fbebeb; border-left:3px solid #b94545; color:#8a2e2e; }
.banner-info { background:#eef4f8; border-left:3px solid #4e7f9c; color:#244f69; }

/* Stat cards */
.stat-card { background:#fff; border:1.5px solid var(--border); border-radius:12px; padding:.85rem 1rem; text-align:center; }
.stat-val { font-family:'DM Sans'; font-size:1.75rem; font-weight:600; }
.stat-lbl { font-size:.68rem; color:#5a5449; letter-spacing:.08em; text-transform:uppercase; margin-top:.2rem; }

/* Chips */
.chip { display:inline-block; padding:2px 10px; border-radius:999px; font-size:.7rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; margin:1px 4px 1px 0; }
.chip-done    { background:var(--done-bg);   color:var(--done-fg); }
.chip-prog    { background:var(--prog-bg);   color:var(--prog-fg); }
.chip-hold    { background:var(--hold-bg);   color:var(--hold-fg); }
.chip-review  { background:var(--review-bg); color:var(--review-fg); }
.chip-open    { background:var(--open-bg);   color:var(--open-fg); }
.chip-rework  { background:var(--rework-bg); color:var(--rework-fg); }
.chip-rev     { background:var(--rev-bg);    color:var(--rev-fg); }
.chip-corr    { background:var(--corr-bg);   color:var(--corr-fg); }

/* Platform chips */
.platform-chip { display:inline-block; padding:2px 10px; border-radius:999px; font-size:.7rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; margin-left:4px; }
.chip-shopify { background:#e6f4ea; color:#2f6f4e; }
.chip-webflow { background:#eaf2ff; color:#2a4dbf; }
.chip-zoketo  { background:#f3e8ff; color:#6b21a8; }
.chip-default { background:#efefef; color:#474747; }

/* Designer badge */
.designer-tag { font-size:.7rem; color:#5f5a52; font-weight:600; letter-spacing:.06em; text-transform:uppercase; background:#fff; border:1px solid var(--border); border-radius:5px; padding:1px 7px; margin-right:4px; }

/* Project header */
.project-header { font-family:'DM Sans'; font-size:16px; font-weight:bold; color:#121212; border-bottom:1.5px solid var(--border); padding-bottom:4px; margin:1rem 0 .55rem; }

/* Misc */
.divider { border:none; border-top:1.5px solid var(--border); margin:1rem 0; }
.tab-title-chip { display:inline-block; background:#111; color:#fff; font-size:.68rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; padding:4px 10px; border-radius:999px; }
.week-title { display:flex; align-items:center; gap:8px; margin-bottom:.6rem; font-family:'DM Sans'; font-size:1rem; color:#000000; }
.small-note { color:#6b665e; font-size:.8rem; }

/* Section label in task list */
.section-label { font-size:.68rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:#888; margin:.8rem 0 .3rem; }

/* Lead metric card */
.metric-card { background:#fff; border:1.5px solid var(--border); border-radius:14px; padding:1.1rem 1.2rem; }
.metric-val { font-size:2rem; font-weight:700; }
.metric-lbl { font-size:.72rem; color:#6f6a61; letter-spacing:.07em; text-transform:uppercase; margin-top:.15rem; }
.metric-sub { font-size:.75rem; color:#999; margin-top:.2rem; }

/* Welcome toast */
@keyframes toastIn  { from{opacity:0;transform:translateY(-10px)} to{opacity:1;transform:translateY(0)} }
@keyframes toastOut { from{opacity:1;transform:translateY(0)} to{opacity:0;transform:translateY(-10px)} }
.welcome-toast { display:inline-flex; align-items:center; gap:8px; background:#ffffff; border:1.5px solid #d4f0e0; border-left:4px solid #3f7d5a; border-radius:10px; padding:.5rem 1rem; font-size:.83rem; color:#1e5c3e; font-weight:500; box-shadow:0 4px 14px rgba(0,0,0,.06); animation:toastIn .35s ease forwards, toastOut .4s ease 2s forwards; }

button[data-baseweb="tab"] { color:#6f6a61 !important; }
button[data-baseweb="tab"][aria-selected="true"] { color:#111111 !important; }
div[data-baseweb="tab-highlight"] { background-color:#111 !important; }

/* Add task form card */
.form-card { background:#fff; border:1.5px solid var(--border); border-radius:14px; padding:1.4rem 1.5rem; margin-bottom:1rem; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# GOOGLE SHEETS CONNECTION
# ──────────────────────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

@st.cache_resource(show_spinner=False)
def get_gspread_client():
    """Returns an authenticated gspread client using Streamlit Cloud secrets."""
    try:
        service_account_info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        return None


def get_spreadsheet():
    client = get_gspread_client()
    if client is None:
        return None
    try:
        return client.open_by_url(SHEET_URL)
    except Exception:
        return None


def get_all_worksheets():
    """Returns dict of {sheet_name: pd.DataFrame} skipping the timeline tab."""
    sh = get_spreadsheet()
    if sh is None:
        return None, "Could not connect to Google Sheet. Check service account setup."
    skip = {"design closure timeline"}
    result = {}
    try:
        for ws in sh.worksheets():
            if ws.title.strip().lower() in skip:
                continue
            data = ws.get_all_values()
            if not data or len(data) < 2:
                result[ws.title] = pd.DataFrame()
                continue
            df = pd.DataFrame(data[1:], columns=data[0])
            result[ws.title] = df
        return result, None
    except Exception as e:
        return None, str(e)


def append_row_to_sheet(week_tab: str, row_data: list) -> Tuple[bool, str]:
    """Appends a single row to the given week tab."""
    sh = get_spreadsheet()
    if sh is None:
        return False, "Could not connect to Google Sheet."
    try:
        ws = sh.worksheet(week_tab)
        ws.append_row(row_data, value_input_option="USER_ENTERED")
        return True, ""
    except Exception as e:
        return False, str(e)


def update_cell_in_sheet(week_tab: str, task_name: str, assigned_to: str,
                          col_name: str, new_value: str) -> Tuple[bool, str]:
    """
    Finds the first row matching task_name + assigned_to and updates col_name cell.
    Used for status updates on existing tasks.
    """
    sh = get_spreadsheet()
    if sh is None:
        return False, "Could not connect to Google Sheet."
    try:
        ws = sh.worksheet(week_tab)
        data = ws.get_all_values()
        if not data:
            return False, "Sheet is empty."
        headers = data[0]
        if col_name not in headers:
            return False, f"Column '{col_name}' not found in sheet."
        col_idx = headers.index(col_name) + 1  # 1-indexed

        task_col_idx = headers.index("Task") if "Task" in headers else None
        assigned_col_idx = headers.index("Assigned To") if "Assigned To" in headers else None

        if task_col_idx is None:
            return False, "Task column not found."

        for i, row in enumerate(data[1:], start=2):
            row_task = row[task_col_idx] if task_col_idx < len(row) else ""
            row_assigned = row[assigned_col_idx] if (assigned_col_idx is not None and assigned_col_idx < len(row)) else ""
            if row_task == task_name and (assigned_col_idx is None or row_assigned == assigned_to):
                ws.update_cell(i, col_idx, new_value)
                return True, ""
        return False, "Task row not found in sheet."
    except Exception as e:
        return False, str(e)


def get_revision_count_for_task(df: pd.DataFrame, task_name: str) -> int:
    """Count existing revision/correction rows for a task to determine next number."""
    if df is None or df.empty:
        return 0
    if "Task" not in df.columns or "Status" not in df.columns:
        return 0
    mask = (df["Task"].astype(str).str.strip() == task_name.strip()) & \
           (df["Status"].astype(str).str.lower().str.contains("revision|correction"))
    return int(mask.sum())


# ──────────────────────────────────────────────────────────────────────────────
# SHEET CACHE & REFRESH
# ──────────────────────────────────────────────────────────────────────────────

def refresh_sheet_data() -> Optional[Dict[str, pd.DataFrame]]:
    now = time.time()
    store = st.session_state.sheet_store

    needs_fetch = (
        "data" not in store
        or (now - store.get("fetched_at", 0)) >= REFRESH_INTERVAL
    )

    if needs_fetch:
        data, err = get_all_worksheets()
        if err:
            st.session_state.sheet_error = err
            return store.get("data")
        store["data"] = data
        store["fetched_at"] = now
        st.session_state.sheet_error = ""

    return store.get("data")


# ──────────────────────────────────────────────────────────────────────────────
# DATA HELPERS
# ──────────────────────────────────────────────────────────────────────────────

def get_week_tabs(xl: Dict[str, pd.DataFrame]) -> list:
    skip = {"design closure timeline"}
    return [s for s in (xl.keys() if xl else []) if s.strip().lower() not in skip]


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure all expected columns exist and clean up values."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for col in SHEET_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip().replace("nan", "")
    df = df[df["Task"].str.strip() != ""]
    return df


def is_revision_row(status: str) -> bool:
    s = status.lower()
    return "revision" in s or "correction" in s


def status_chip(s: str) -> str:
    if not s or s.strip().lower() in ("nan", ""):
        return ""
    sl = s.strip().lower()
    if sl == "completed":                    cls = "chip-done"
    elif sl in ("in progress",):             cls = "chip-prog"
    elif sl in ("hold",):                    cls = "chip-hold"
    elif sl in ("in internal review", "in client review"): cls = "chip-review"
    elif sl == "re work":                    cls = "chip-rework"
    elif "revision" in sl:                   cls = "chip-rev"
    elif "correction" in sl:                 cls = "chip-corr"
    else:                                    cls = "chip-open"
    return f'<span class="chip {cls}">{s}</span>'


def platform_chip(platform: str) -> str:
    if not platform or platform.strip().lower() in ("nan", ""):
        return ""
    key = platform.strip().lower()
    cls = {"shopify": "chip-shopify", "webflow": "chip-webflow", "zoketo": "chip-zoketo"}.get(key, "chip-default")
    return f'<span class="platform-chip {cls}">{platform}</span>'


def get_current_week_tab(xl: Dict[str, pd.DataFrame]) -> Optional[str]:
    tabs = get_week_tabs(xl)
    return tabs[0] if tabs else None


# ──────────────────────────────────────────────────────────────────────────────
# LEAD METRICS
# ──────────────────────────────────────────────────────────────────────────────

def compute_lead_metrics(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {k: 0 for k in ["total", "completed", "pending", "on_hold",
                                "revision_count", "delayed", "client_review", "designer_breakdown"]}
    df = normalize_df(df)
    # Exclude pure revision/correction rows from primary task count
    base = df[~df["Status"].apply(is_revision_row)]

    today = datetime.date.today()

    def parse_date(s):
        try:
            return datetime.datetime.strptime(str(s).strip(), "%Y-%m-%d").date()
        except Exception:
            try:
                return datetime.datetime.strptime(str(s).strip(), "%d/%m/%Y").date()
            except Exception:
                return None

    total       = len(base)
    completed   = int((base["Status"].str.lower() == "completed").sum())
    on_hold     = int((base["Status"].str.lower() == "hold").sum())
    pending     = total - completed - on_hold
    client_rev  = int((base["Status"].str.lower() == "in client review").sum())

    # Revision/correction rows
    rev_rows    = df[df["Status"].apply(is_revision_row)]
    rev_count   = len(rev_rows)

    # Delayed: end date is in the past and not completed
    delayed = 0
    for _, row in base.iterrows():
        ed = parse_date(row.get("End Date", ""))
        if ed and ed < today and row["Status"].lower() != "completed":
            delayed += 1

    # Designer breakdown
    designer_breakdown = base.groupby("Assigned To").size().to_dict() if "Assigned To" in base.columns else {}

    return {
        "total": total,
        "completed": completed,
        "pending": max(pending, 0),
        "on_hold": on_hold,
        "revision_count": rev_count,
        "delayed": delayed,
        "client_review": client_rev,
        "designer_breakdown": designer_breakdown,
    }


# ──────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ──────────────────────────────────────────────────────────────────────────────

def init_state():
    defaults = {
        "is_logged_in": False,
        "user_key": None,
        "role": None,
        "label": None,
        "auth_error": "",
        "sheet_store": {},
        "sheet_error": "",
        "week_tab": None,
        "show_welcome": False,
        "welcome_label": "",
        "add_task_success": "",
        "add_task_error": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def login_user(username: str, password: str) -> bool:
    for key, info in ACCOUNTS.items():
        if info["username"] == username and info["password"] == password:
            st.session_state.is_logged_in = True
            st.session_state.user_key = key
            st.session_state.role = info["role"]
            st.session_state.label = info["label"]
            st.session_state.auth_error = ""
            st.session_state.show_welcome = True
            st.session_state.welcome_label = info["label"]
            return True
    st.session_state.auth_error = "Invalid username or password."
    return False


def logout_user():
    for k in ["is_logged_in", "user_key", "role", "label", "auth_error",
              "week_tab", "show_welcome", "welcome_label",
              "add_task_success", "add_task_error"]:
        st.session_state[k] = "" if isinstance(st.session_state.get(k), str) else None if k != "is_logged_in" else False
    st.session_state.is_logged_in = False
    st.rerun()


# ──────────────────────────────────────────────────────────────────────────────
# SHARED UI COMPONENTS
# ──────────────────────────────────────────────────────────────────────────────

def render_login():
    st.markdown("""
    <div class="login-wrapper">
      <div class="login-card-wrap">
        <div style="text-align:center;margin-bottom:1.5rem;">
          <div class="login-badge">Design Team</div>
          <div class="login-title">DesignPulse</div>
          <div class="login-sub">Sign in with your Designer or Lead credentials<br>to access your workspace.</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username", placeholder="Enter your username")
            password = st.text_input("Password", type="password", placeholder="Enter your password")
            submitted = st.form_submit_button("Sign In →", use_container_width=True)
            if submitted:
                if login_user(username.strip(), password.strip()):
                    st.rerun()
        if st.session_state.auth_error:
            st.markdown(f'<div class="banner banner-warn" style="text-align:center;">{st.session_state.auth_error}</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-footer-note">✦ DesignPulse · Crafted by Dharnu</div>', unsafe_allow_html=True)


def render_top_header(title: str, subtitle: str):
    left, right = st.columns([6, 1])
    with left:
        st.markdown(f"""
        <div style="display:flex;align-items:baseline;gap:10px;margin-bottom:4px;">
          <span style="font-family:'DM Serif Display',serif;font-size:2.1rem;">{title}</span>
          <span style="font-size:.82rem;color:#000;letter-spacing:.06em;text-transform:uppercase;">· Design Team</span>
        </div>
        <p style="color:#000;font-size:.88rem;margin-bottom:.2rem;">{subtitle}</p>
        """, unsafe_allow_html=True)
        if st.session_state.get("show_welcome"):
            st.markdown(f'<div class="welcome-toast"><span>👋</span><span>Welcome, <strong>{st.session_state.welcome_label}</strong>!</span></div>', unsafe_allow_html=True)
            st.session_state.show_welcome = False
    with right:
        if st.button("Logout", use_container_width=True):
            logout_user()
    st.markdown('<hr class="divider">', unsafe_allow_html=True)


def week_selector(xl: Dict[str, pd.DataFrame], key_prefix: str = "week") -> str:
    tabs = get_week_tabs(xl)
    if not tabs:
        return None
    if st.session_state.week_tab not in tabs:
        st.session_state.week_tab = tabs[0]
    col, _ = st.columns([3, 2])
    with col:
        selected = st.selectbox("📅 Select Week", tabs,
                                index=tabs.index(st.session_state.week_tab),
                                key=f"{key_prefix}_select")
    st.session_state.week_tab = selected
    return selected


def render_footer():
    st.markdown("""
    <div style="display:flex;justify-content:space-between;align-items:center;color:#6b665e;font-size:.75rem;margin-top:1rem;">
        <span>✦ DesignPulse · Crafted By Dharnu</span>
        <span>Internal build · Auto-refresh every 60s</span>
    </div>
    """, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# ADD TASK FORM
# ──────────────────────────────────────────────────────────────────────────────

def render_add_task_form(xl: Dict[str, pd.DataFrame], default_designer: Optional[str] = None, lead_mode: bool = False):
    """
    Form to add a new task or revision row.
    default_designer: pre-fill Assigned To for designer view.
    lead_mode: allows assigning to any designer.
    """
    st.markdown('<div class="form-card">', unsafe_allow_html=True)
    st.markdown("#### ＋ Add Task / Revision")

    week_tabs = get_week_tabs(xl)
    if not week_tabs:
        st.markdown('<div class="banner banner-warn">No week tabs found in sheet.</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        return

    with st.form("add_task_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            project   = st.text_input("Project Name *", placeholder="e.g. Acme Homepage")
            task_name = st.text_input("Task Name *", placeholder="e.g. Banner Design")
            platform  = st.selectbox("Platform", [""] + PLATFORMS)
        with c2:
            if lead_mode:
                assigned_to = st.selectbox("Assigned To *", DESIGNER_LABELS)
            else:
                assigned_to = default_designer
                st.text_input("Assigned To", value=default_designer, disabled=True)

            status    = st.selectbox("Status", STATUSES)
            target_week = st.selectbox("Week Tab", week_tabs)

        c3, c4 = st.columns(2)
        with c3:
            start_date = st.date_input("Start Date", value=datetime.date.today())
        with c4:
            end_date   = st.date_input("End Date", value=datetime.date.today())

        comments = st.text_area("Comments", placeholder="Optional notes...", height=80)

        # Revision mode
        is_revision = st.checkbox("This is a Revision / Correction row")
        rev_type    = None
        if is_revision:
            rev_type = st.radio("Type", ["Revision", "Correction"], horizontal=True)

        submitted = st.form_submit_button("Add to Sheet →", use_container_width=True)

        if submitted:
            if not project.strip() or not task_name.strip():
                st.session_state.add_task_error = "Project Name and Task Name are required."
            else:
                # Determine status label
                final_status = status
                rev_no = ""
                if is_revision and rev_type:
                    current_df = normalize_df(xl.get(target_week, pd.DataFrame()))
                    count = get_revision_count_for_task(current_df, task_name.strip()) + 1
                    final_status = f"{rev_type} ({count})"
                    rev_no = str(count)

                row = [
                    project.strip(),
                    platform if platform else "",
                    assigned_to,
                    task_name.strip(),
                    final_status,
                    rev_no,
                    str(start_date),
                    str(end_date),
                    comments.strip(),
                ]
                ok, err = append_row_to_sheet(target_week, row)
                if ok:
                    st.session_state.add_task_success = f"Task '{task_name.strip()}' added to {target_week}!"
                    st.session_state.add_task_error = ""
                    # Invalidate cache so next read gets fresh data
                    st.session_state.sheet_store = {}
                    st.rerun()
                else:
                    st.session_state.add_task_error = f"Failed to write to sheet: {err}"

    if st.session_state.add_task_success:
        st.markdown(f'<div class="banner banner-ok">✓ {st.session_state.add_task_success}</div>', unsafe_allow_html=True)
        st.session_state.add_task_success = ""
    if st.session_state.add_task_error:
        st.markdown(f'<div class="banner banner-warn">⚠ {st.session_state.add_task_error}</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# TASK LIST VIEW
# ──────────────────────────────────────────────────────────────────────────────

def render_task_list(df: pd.DataFrame, filter_designer: Optional[str] = None,
                     show_designer_col: bool = True):
    """Renders grouped task rows. Optionally filter by designer."""
    if df is None or df.empty:
        st.markdown('<div class="small-note">No tasks found for this week.</div>', unsafe_allow_html=True)
        return

    df = normalize_df(df)
    if filter_designer:
        df = df[df["Assigned To"] == filter_designer]

    if df.empty:
        st.markdown('<div class="small-note">No tasks found.</div>', unsafe_allow_html=True)
        return

    # Group by project
    projects = df["Project"].unique() if "Project" in df.columns else [""]
    for proj in projects:
        proj_df = df[df["Project"] == proj] if proj else df
        platform = proj_df["Platform"].iloc[0] if not proj_df.empty and "Platform" in proj_df.columns else ""
        pb = platform_chip(platform)

        rows_html = []
        for _, row in proj_df.iterrows():
            designer_html = f'<span class="designer-tag">{row["Assigned To"]}</span>' if show_designer_col and row.get("Assigned To") else ""
            status_html   = status_chip(row.get("Status", ""))
            task_txt      = row.get("Task", "")
            comment_html  = f' <span style="font-size:.78rem;color:#666;">— {row["Comments"]}</span>' if row.get("Comments") else ""
            date_html     = ""
            if row.get("Start Date") or row.get("End Date"):
                date_html = f' <span style="font-size:.72rem;color:#aaa;">({row.get("Start Date","")} → {row.get("End Date","")})</span>'
            rev_no = row.get("Revision No", "")
            rev_html = f' <span style="font-size:.68rem;color:var(--rev-fg);background:var(--rev-bg);padding:1px 6px;border-radius:4px;">#{rev_no}</span>' if rev_no else ""

            rows_html.append(
                f'<div style="padding:4px 0;font-size:.88rem;color:#000;">'
                f'{designer_html}{status_html} {task_txt}{rev_html}{comment_html}{date_html}'
                f'</div>'
            )

        st.markdown(
            f'<div class="project-header">{proj}{pb}</div>' + "".join(rows_html),
            unsafe_allow_html=True,
        )


# ──────────────────────────────────────────────────────────────────────────────
# REVISION LOG VIEW
# ──────────────────────────────────────────────────────────────────────────────

def render_revision_log(df: pd.DataFrame):
    """Shows only revision/correction rows."""
    if df is None or df.empty:
        st.markdown('<div class="small-note">No data.</div>', unsafe_allow_html=True)
        return
    df = normalize_df(df)
    rev_df = df[df["Status"].apply(is_revision_row)]
    if rev_df.empty:
        st.markdown('<div class="small-note">No revision or correction rows this week.</div>', unsafe_allow_html=True)
        return
    display_cols = ["Project", "Assigned To", "Task", "Status", "Revision No", "Comments"]
    display_cols = [c for c in display_cols if c in rev_df.columns]
    st.dataframe(rev_df[display_cols].reset_index(drop=True), use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────────────────────────────────────
# LEAD DASHBOARD METRICS
# ──────────────────────────────────────────────────────────────────────────────

def render_lead_metrics(df: pd.DataFrame):
    m = compute_lead_metrics(df)

    cards = [
        ("total",          "Total Tasks",            "#111111"),
        ("completed",      "Completed",              "#276749"),
        ("pending",        "Pending",                "#8a6100"),
        ("on_hold",        "On Hold",                "#9b2c2c"),
        ("revision_count", "Revisions / Corrections","#6b21a8"),
        ("delayed",        "Delayed",                "#b45309"),
        ("client_review",  "Client Review Pending",  "#125d85"),
    ]

    cols = st.columns(len(cards))
    for col, (key, lbl, color) in zip(cols, cards):
        col.markdown(
            f'<div class="stat-card"><div class="stat-val" style="color:{color};">{m[key]}</div>'
            f'<div class="stat-lbl">{lbl}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    # Designer-wise breakdown
    if m["designer_breakdown"]:
        st.markdown("#### 👥 Designer-wise Task Count")
        d_cols = st.columns(len(m["designer_breakdown"]))
        for col, (designer, count) in zip(d_cols, m["designer_breakdown"].items()):
            col.markdown(
                f'<div class="stat-card"><div class="stat-val" style="color:#111;">{count}</div>'
                f'<div class="stat-lbl">{designer}</div></div>',
                unsafe_allow_html=True,
            )
        st.markdown('<hr class="divider">', unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# DESIGNER VIEW
# ──────────────────────────────────────────────────────────────────────────────

def render_designer_view():
    label = st.session_state.label

    render_top_header("DesignPulse", f"{label}'s workspace")

    with st.sidebar:
        st.markdown("## ✦ DesignPulse")
        st.caption(label)
        st.markdown("---")
        st.markdown("### Access")
        st.markdown(f"**Role:** Designer")
        st.markdown(f"**Name:** {label}")
        if st.button("Logout", key="sb_logout_designer", use_container_width=True):
            logout_user()

    xl = refresh_sheet_data()
    if st.session_state.sheet_error:
        st.markdown(f'<div class="banner banner-warn">⚠️ {st.session_state.sheet_error}</div>', unsafe_allow_html=True)
        render_footer()
        return
    if not xl:
        st.markdown('<div class="banner banner-warn">⚠️ Could not load sheet data.</div>', unsafe_allow_html=True)
        render_footer()
        return

    tab_my, tab_all = st.tabs([f"📋 My Tasks", "👁 All Tasks (Read-only)"])

    with tab_my:
        selected_week = week_selector(xl, key_prefix="designer_my")
        if not selected_week:
            st.markdown('<div class="small-note">No week tabs found.</div>', unsafe_allow_html=True)
        else:
            df = normalize_df(xl.get(selected_week, pd.DataFrame()))

            # Quick personal stats
            my_df = df[df["Assigned To"] == label] if not df.empty else pd.DataFrame()
            if not my_df.empty:
                base = my_df[~my_df["Status"].apply(is_revision_row)]
                c1, c2, c3, c4 = st.columns(4)
                for col, val, lbl, color in [
                    (c1, len(base),                                                  "My Tasks",   "#111"),
                    (c2, int((base["Status"].str.lower()=="completed").sum()),        "Done",       "#276749"),
                    (c3, int((base["Status"].str.lower().isin(["in progress"])).sum()), "In Progress","#8a6100"),
                    (c4, int(my_df["Status"].apply(is_revision_row).sum()),           "Revisions",  "#6b21a8"),
                ]:
                    col.markdown(f'<div class="stat-card"><div class="stat-val" style="color:{color};">{val}</div><div class="stat-lbl">{lbl}</div></div>', unsafe_allow_html=True)
                st.markdown('<hr class="divider">', unsafe_allow_html=True)

            left, right = st.columns([1, 1], gap="large")
            with left:
                st.markdown(f'<div class="week-title">📋 My Tasks — {selected_week}</div>', unsafe_allow_html=True)
                render_task_list(df, filter_designer=label, show_designer_col=False)
            with right:
                render_add_task_form(xl, default_designer=label, lead_mode=False)

    with tab_all:
        selected_week_all = week_selector(xl, key_prefix="designer_all")
        if selected_week_all:
            df_all = normalize_df(xl.get(selected_week_all, pd.DataFrame()))
            # Filter controls
            fc1, fc2 = st.columns(2)
            with fc1:
                filter_proj = st.text_input("Filter by Project", placeholder="Type project name...", key="d_filter_proj")
            with fc2:
                designers_in_sheet = ["All"] + sorted(df_all["Assigned To"].unique().tolist()) if not df_all.empty else ["All"]
                filter_designer = st.selectbox("Filter by Designer", designers_in_sheet, key="d_filter_des")

            if filter_proj:
                df_all = df_all[df_all["Project"].str.lower().str.contains(filter_proj.lower())]
            if filter_designer != "All":
                df_all = df_all[df_all["Assigned To"] == filter_designer]

            render_task_list(df_all, show_designer_col=True)

    render_footer()


# ──────────────────────────────────────────────────────────────────────────────
# LEAD VIEW
# ──────────────────────────────────────────────────────────────────────────────

def render_lead_view():
    render_top_header("DesignPulse", "Design Lead · Full team overview")

    with st.sidebar:
        st.markdown("## ✦ DesignPulse")
        st.caption("Design Lead")
        st.markdown("---")
        st.markdown("### Access")
        st.markdown("**Role:** Design Lead")
        st.markdown("**View:** All designers")
        if st.button("Logout", key="sb_logout_lead", use_container_width=True):
            logout_user()

    xl = refresh_sheet_data()
    if st.session_state.sheet_error:
        st.markdown(f'<div class="banner banner-warn">⚠️ {st.session_state.sheet_error}</div>', unsafe_allow_html=True)
        render_footer()
        return
    if not xl:
        st.markdown('<div class="banner banner-warn">⚠️ Could not load sheet data.</div>', unsafe_allow_html=True)
        render_footer()
        return

    tab_dash, tab_tasks, tab_revlog, tab_add = st.tabs([
        "📊 Dashboard",
        "📋 All Tasks",
        "🔄 Revision Log",
        "＋ Add Task",
    ])

    with tab_dash:
        selected_week = week_selector(xl, key_prefix="lead_dash")
        if selected_week:
            df = normalize_df(xl.get(selected_week, pd.DataFrame()))
            render_lead_metrics(df)

            # Full task list below metrics
            st.markdown(f'<div class="week-title">📋 All Tasks — {selected_week}</div>', unsafe_allow_html=True)
            render_task_list(df, show_designer_col=True)

    with tab_tasks:
        selected_week_t = week_selector(xl, key_prefix="lead_tasks")
        if selected_week_t:
            df_t = normalize_df(xl.get(selected_week_t, pd.DataFrame()))
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                fp = st.text_input("Filter by Project", key="l_fp")
            with fc2:
                all_designers = ["All"] + sorted(df_t["Assigned To"].unique().tolist()) if not df_t.empty else ["All"]
                fd = st.selectbox("Filter by Designer", all_designers, key="l_fd")
            with fc3:
                all_statuses = ["All"] + STATUSES
                fs = st.selectbox("Filter by Status", all_statuses, key="l_fs")

            if fp:
                df_t = df_t[df_t["Project"].str.lower().str.contains(fp.lower())]
            if fd != "All":
                df_t = df_t[df_t["Assigned To"] == fd]
            if fs != "All":
                df_t = df_t[df_t["Status"].str.lower() == fs.lower()]

            render_task_list(df_t, show_designer_col=True)

    with tab_revlog:
        selected_week_r = week_selector(xl, key_prefix="lead_rev")
        if selected_week_r:
            df_r = normalize_df(xl.get(selected_week_r, pd.DataFrame()))
            render_revision_log(df_r)

    with tab_add:
        render_add_task_form(xl, lead_mode=True)

    render_footer()


# ──────────────────────────────────────────────────────────────────────────────
# APP ENTRY
# ──────────────────────────────────────────────────────────────────────────────

init_state()

# Auto-refresh
if HAS_AUTOREFRESH:
    st_autorefresh(interval=REFRESH_INTERVAL * 1000, limit=None, key="auto_refresh")
else:
    st.markdown(
        f'<script>setTimeout(function(){{window.location.reload();}},{REFRESH_INTERVAL * 1000});</script>',
        unsafe_allow_html=True,
    )

if not st.session_state.is_logged_in:
    render_login()
else:
    if st.session_state.role == "designer":
        render_designer_view()
    elif st.session_state.role == "lead":
        render_lead_view()
    else:
        logout_user()
