import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import anthropic
import os
import html
import re
import threading
import base64 as _b64
from datetime import datetime, date, timedelta
from jinja2 import Template
from urllib.parse import quote as urlquote
import auth
import email_utils
import gsheets
import github_sync


def _load_logo_b64() -> str:
    try:
        # prefer the transparent PNG; fall back to the original JPG
        for _fname in ("qualesce_logo.png", "qualesce_logo.jpg"):
            _p = os.path.join(os.path.dirname(os.path.abspath(__file__)), _fname)
            if os.path.exists(_p):
                with open(_p, "rb") as f:
                    _data = _b64.b64encode(f.read()).decode()
                    _mime = "image/png" if _fname.endswith(".png") else "image/jpeg"
                    return _mime, _data
    except Exception:
        pass
    return "image/png", ""

_LOGO_MIME, _LOGO_B64 = _load_logo_b64()

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Qualesce AI Project Manager",
    page_icon=os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "qualesce_logo.png" if os.path.exists(os.path.join(os.path.dirname(os.path.abspath(__file__)), "qualesce_logo.png")) else "qualesce_logo.jpg"),
    layout="wide",
    initial_sidebar_state="collapsed",
)
github_sync.ensure_db_downloaded()
github_sync.start_sync_thread()

if "db_initialized" not in st.session_state:
    auth.init_db()
    st.session_state.db_initialized = True
    try:
        import api_server as _api
        _api.start_background(port=8503)
    except ImportError:
        pass

if hasattr(email_utils, "start_license_notification_scheduler"):
    email_utils.start_license_notification_scheduler()

# Cache secrets in main thread so background email threads can access them
email_utils.prime_secrets_cache()


# ── ONEDRIVE FILE STORAGE ─────────────────────────────────────────────────────
def _get_onedrive_base_path() -> str:
    """Return path to 'Qualesce Dashboard' inside the OneDrive folder that
    matches the Outlook email configured in admin settings. Falls back to
    the default 'OneDrive - Qualesce' if no match is found.
    Result is cached in session state for the lifetime of the session."""
    cached = st.session_state.get("_od_base_path_cache")
    if cached:
        return cached
    home = os.path.expanduser("~")
    result = os.path.join(home, "OneDrive - Qualesce", "Qualesce Dashboard")
    try:
        cfg = auth.get_email_settings()
        email = cfg.get("outlook_email", "")
        if email and "@" in email:
            domain = email.split("@")[1].lower()
            org_key = domain.split(".")[0]
            candidate = os.path.join(home, f"OneDrive - {org_key.capitalize()}")
            if os.path.isdir(candidate):
                result = os.path.join(candidate, "Qualesce Dashboard")
            else:
                for entry in os.listdir(home):
                    if entry.lower().startswith("onedrive") and org_key in entry.lower():
                        full = os.path.join(home, entry)
                        if os.path.isdir(full):
                            result = os.path.join(full, "Qualesce Dashboard")
                            break
    except Exception:
        pass
    st.session_state["_od_base_path_cache"] = result
    return result


def _get_file_counts_cached() -> dict:
    """Scan the OneDrive Qualesce Dashboard folder once per session.
    Returns {safe_folder_name: file_count} for all project folders.
    Much faster than calling get_project_files() per row (1 scan vs N)."""
    cache_key = "_file_counts_cache"
    if cache_key in st.session_state:
        return st.session_state[cache_key]
    base = _get_onedrive_base_path()
    counts: dict = {}
    if os.path.isdir(base):
        for entry in os.listdir(base):
            fp = os.path.join(base, entry)
            if os.path.isdir(fp):
                try:
                    counts[entry] = sum(
                        1 for f in os.listdir(fp)
                        if os.path.isfile(os.path.join(fp, f))
                    )
                except OSError:
                    counts[entry] = 0
    st.session_state[cache_key] = counts
    return counts


PROJECT_COLS = ["id","name","client","lead","employee","status","proj_type","start","end","due_date","po","desc",
               "manual_hrs","auto_hrs","cost_per_hr","hours_saved","cost_saved","roi_pct","is_new","is_active",
               "ckpt_pdd_sdd_start","ckpt_pdd_sdd_end",
               "ckpt_development_start","ckpt_development_end",
               "ckpt_uat_start","ckpt_uat_end",
               "ckpt_deployment_start","ckpt_deployment_end",
               "num_bots","manual_run_mins","bot_run_mins","monthly_runs","num_persons"]
EXCEL_COLS = PROJECT_COLS  # alias kept for the Excel download export

PORTAL_URL = "https://q-dashboard.streamlit.app/"

def fmt_date(d: str) -> str:
    """Normalise any date string to DD-MM-YYYY for display. Returns original if unparseable."""
    if not d or not str(d).strip():
        return ""
    s = str(d).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d-%m-%Y")
        except ValueError:
            pass
    return s

# ── BASE DATA ─────────────────────────────────────────────────────────────────
BASE_PROJECTS = []  # All project data is loaded from the database / Excel

ALL_STATUSES  = ["R&M","UAT","In Progress","Completed","PDD","Discontinued","Internal POC","External POC","Important","Presales"]
WS_STATUSES   = ["In Progress", "Completed", "On Hold", "Discontinued"]
STATUS_STYLES = {
    "R&M":          {"bg":"#EFF7F7","text":"#3F8E91","dot":"#5FA9AB"},
    "UAT":          {"bg":"#FBF6E7","text":"#966D17","dot":"#D4A02C"},
    "Completed":    {"bg":"#E5F2EC","text":"#2E7D5B","dot":"#2E7D5B"},
    "In Progress":  {"bg":"#EFF7F7","text":"#2F6F72","dot":"#5FA9AB"},
    "PDD":          {"bg":"#FBF6E7","text":"#966D17","dot":"#D4A02C"},
    "Discontinued": {"bg":"#FCEAEA","text":"#B23A3A","dot":"#B23A3A"},
    "Internal POC": {"bg":"#EFF7F7","text":"#2F6F72","dot":"#4A989B"},
    "External POC": {"bg":"#F7F8F9","text":"#4E5860","dot":"#9BA5AE"},
    "Important":    {"bg":"#FCEAEA","text":"#B23A3A","dot":"#B23A3A"},
    "Presales":     {"bg":"#EFF7F7","text":"#3F8E91","dot":"#5FA9AB"},
}
STATUS_CHART_COLORS = ["#5FA9AB","#D4A02C","#4A989B","#2E7D5B","#B8881F","#B23A3A","#3F8E91","#6E7A84","#966D17","#8FC4C5"]

SYSTEM_PROMPT = """You are an AI Project Manager Agent for Qualesce (RPA automation company).
STATUSES: R&M (Run & Maintain), UAT (User Acceptance Testing), In Progress, Completed, PDD (Pre-Due Diligence), Discontinued, Internal POC, External POC, Important (high-priority flagged tasks needing immediate attention).
ROI FORMULA: Hours Saved = Manual Hrs - Auto Hrs | Cost Saved = Hours Saved x Cost/Hr | ROI% = (Hours Saved / Manual Hrs) x 100

FORMATTING RULES (always follow):
- When listing multiple projects, people, or items with attributes → use a markdown table with | column | headers |
- When explaining steps, reasons, or a summary → use bullet points (- item)
- Never write long prose paragraphs — always break into bullets
- Show ROI formula steps when calculating
- Be concise and data-driven"""

# ── EXCEL HELPERS ─────────────────────────────────────────────────────────────
def build_excel_bytes(df: pd.DataFrame) -> bytes:
    """Generate Excel file in memory from live DB — always up to date."""
    import io
    out = df.copy()
    for col in EXCEL_COLS:
        if col not in out.columns:
            out[col] = ""
    _poc_statuses_excel = {"Presales", "Internal POC", "External POC"}
    presales_df = out[out["status"].str.strip().isin(_poc_statuses_excel)][EXCEL_COLS].reset_index(drop=True)
    license_records = auth.get_all_licenses()
    license_df = pd.DataFrame(license_records) if license_records else pd.DataFrame(
        columns=["id","tool_name","no_of_licenses","start_date","end_date","created_at"])
    user_records = auth.get_all_users()
    user_df = pd.DataFrame(user_records, columns=["id","name","email","role","is_active","created_at"]) \
              if user_records else pd.DataFrame(columns=["id","name","email","role","is_active","created_at"])
    sold_records = auth.get_all_sold_licenses()
    sold_df = pd.DataFrame(sold_records) if sold_records else pd.DataFrame(
        columns=["id","tool_name","client","no_of_licenses","start_date","end_date","notes","created_at"])
    task_records = auth.get_all_tasks_asc()
    task_df = pd.DataFrame(task_records) if task_records else pd.DataFrame(
        columns=["id","title","description","status","progress","due_date","start_date",
                 "created_at","updated_at","comment","assigned_to","assigned_to_email",
                 "assigned_by","assigned_by_email"])
    comment_records = auth.get_all_comments_for_excel()
    comment_df = pd.DataFrame(comment_records) if comment_records else pd.DataFrame(
        columns=["id","task_id","task_title","employee_name","employee_email","comment","week_start","created_at"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out[EXCEL_COLS].to_excel(writer, sheet_name="Project Details", index=False)
        presales_df.to_excel(writer, sheet_name="Presales_POC", index=False)
        license_df.to_excel(writer, sheet_name="License", index=False)
        sold_df.to_excel(writer, sheet_name="Sold_License", index=False)
        user_df.to_excel(writer, sheet_name="Users", index=False)
        task_df.drop(columns=["comment"], errors="ignore").to_excel(writer, sheet_name="Tasks", index=False)
        comment_df.to_excel(writer, sheet_name="Comments", index=False)
    return buf.getvalue()


# ── ONEDRIVE FILE HELPERS ─────────────────────────────────────────────────────
def _safe_folder(name: str) -> str:
    import re as _re
    return _re.sub(r'[<>:"/\\|?*]', '_', str(name)).strip() or "Unknown_Project"

def _project_dir(pname: str) -> str:
    d = os.path.join(_get_onedrive_base_path(), _safe_folder(pname))
    os.makedirs(d, exist_ok=True)
    return d

def get_project_files(pname: str) -> list:
    d = os.path.join(_get_onedrive_base_path(), _safe_folder(pname))
    if not os.path.isdir(d):
        return []
    files = []
    for f in sorted(os.listdir(d)):
        fp = os.path.join(d, f)
        if os.path.isfile(fp):
            files.append({"name": f, "path": fp, "size": os.path.getsize(fp)})
    return files

def save_project_file(pname: str, uploaded_file) -> str:
    import re as _re
    safe = _re.sub(r'[^\w\s\-.]', '_', os.path.basename(uploaded_file.name))
    dest = os.path.join(_project_dir(pname), safe)
    with open(dest, "wb") as out:
        out.write(uploaded_file.getbuffer())
    return safe

def delete_project_file(pname: str, filename: str) -> bool:
    fp = os.path.join(_get_onedrive_base_path(), _safe_folder(pname), os.path.basename(filename))
    if os.path.isfile(fp):
        os.remove(fp)
        return True
    return False

def fmt_file_size(b: int) -> str:
    if b < 1024: return f"{b} B"
    if b < 1024**2: return f"{b/1024:.1f} KB"
    return f"{b/1024**2:.1f} MB"


@st.dialog("📎 Project Files", width="large")
def _file_upload_dialog():
    pname     = st.session_state.get("file_panel_name", "")
    panel_pid = st.session_state.get("file_panel_proj", "")
    if not pname:
        st.error("No project selected.")
        return

    _od_email = auth.get_email_settings().get("outlook_email", "") or "qualesce account"
    _od_path  = _get_onedrive_base_path()
    _od_label = os.path.basename(os.path.dirname(_od_path))   # e.g. "OneDrive - Qualesce"
    st.markdown(
        f"<div style='font-size:12px;color:#64748B;margin-bottom:4px'>"
        f"📁 {_od_label} ({_od_email}) / Qualesce Dashboard / <b>{pname}</b></div>",
        unsafe_allow_html=True,
    )

    # ── Existing files list ───────────────────────────────────────────────────
    existing = get_project_files(pname)
    if existing:
        st.markdown("**Uploaded Files**")
        for fi in existing:
            c1, c2, c3, c4 = st.columns([5, 2, 2, 1])
            c1.markdown(f"📄 {fi['name']}")
            c2.caption(fmt_file_size(fi["size"]))
            with open(fi["path"], "rb") as fh:
                c3.download_button(
                    "⬇ Download", fh.read(),
                    file_name=fi["name"],
                    key=f"dlg_dl_{panel_pid}_{fi['name']}",
                )
            if c4.button("🗑", key=f"dlg_del_{panel_pid}_{fi['name']}", help="Delete"):
                delete_project_file(pname, fi["name"])
                st.session_state.pop("_file_counts_cache", None)
                st.rerun()
        st.divider()

    # ── Upload area ───────────────────────────────────────────────────────────
    uploaded = st.file_uploader(
        "Drag & drop files here, or click to browse",
        accept_multiple_files=True,
        key=f"dlg_up_{panel_pid}",
    )

    if uploaded:
        st.markdown(f"**{len(uploaded)} file(s) ready to upload:**")
        for uf in uploaded:
            st.caption(f"• {uf.name}  ({fmt_file_size(uf.size)})")
        st.markdown("")
        c_save, c_close = st.columns(2)
        if c_save.button("💾 Save to OneDrive", type="primary", use_container_width=True, key="dlg_save"):
            for uf in uploaded:
                save_project_file(pname, uf)
            st.session_state.file_panel_proj = None
            st.session_state.file_panel_name = ""
            st.session_state.pop("_file_counts_cache", None)  # invalidate badge cache
            st.success(f"✅ {len(uploaded)} file(s) saved to OneDrive!")
            st.rerun()
        if c_close.button("Cancel", use_container_width=True, key="dlg_cancel"):
            st.session_state.file_panel_proj = None
            st.session_state.file_panel_name = ""
            st.rerun()
    else:
        if st.button("Close", use_container_width=True, key="dlg_close"):
            st.session_state.file_panel_proj = None
            st.session_state.file_panel_name = ""
            st.rerun()


def save_projects(df: pd.DataFrame) -> bool:
    """Save projects to SQLite (single source of truth)."""
    try:
        auth.upsert_projects(df.to_dict("records"))
        load_projects.clear()  # Bust cache so next load_projects() call is fresh
        return True
    except Exception:
        return False


def save_projects_async(df: pd.DataFrame):
    """Non-blocking DB save — runs in a daemon thread."""
    _df_copy = df.copy()
    def _write():
        save_projects(_df_copy)
    threading.Thread(target=_write, daemon=True).start()


@st.cache_data(ttl=60, show_spinner=False)
def load_project_types() -> list:
    """Load project types from DB. Returns list of dicts with id, name, color."""
    return auth.get_project_types()


@st.cache_data(ttl=30, show_spinner=False)
def load_projects() -> pd.DataFrame:
    """Load projects from SQLite (single source of truth). Cached for 30 seconds."""
    records = auth.get_all_projects()
    if records:
        df = pd.DataFrame(records)
        for col in PROJECT_COLS:
            if col not in df.columns:
                df[col] = ""
        return df
    return pd.DataFrame(columns=PROJECT_COLS)


def compute_roi(manual, auto, cost):
    try:
        m, a, c = float(manual), float(auto), float(cost)
        if m > 0:
            saved = max(0.0, m - a)
            return {"saved": saved, "cost": saved * c, "pct": round((saved / m) * 100)}
    except (ValueError, TypeError):
        pass
    return None

def compute_health_score(row: dict) -> dict:
    """Return health score (0-100) and RAG status for a project row."""
    score = 100
    issues = []
    today = date.today()

    status = str(row.get("status", "")).strip()
    if status in ("Discontinued",):
        return {"score": 0, "color": "#6B7280", "label": "N/A", "issues": []}
    if status == "Completed":
        return {"score": 100, "color": "#10B981", "label": "Excellent", "issues": []}

    due = _parse_dmy(str(row.get("due_date", "") or ""))
    if due:
        days_left = (due - today).days
        if days_left < 0:
            score -= 35
            issues.append(f"Overdue by {abs(days_left)}d")
        elif days_left <= 7:
            score -= 20
            issues.append(f"Due in {days_left}d")
        elif days_left <= 14:
            score -= 10
            issues.append(f"Due in {days_left}d")

    manual = row.get("manual_hrs", "")
    auto = row.get("auto_hrs", "")
    try:
        m, a = float(manual), float(auto)
        if m > 0 and (m - a) / m < 0.2:
            score -= 15
            issues.append("Low ROI (<20%)")
    except (TypeError, ValueError, ZeroDivisionError):
        pass

    ckpt_fields = [
        ("ckpt_pdd_sdd_end", "PDD/SDD"),
        ("ckpt_development_end", "Dev"),
        ("ckpt_uat_end", "UAT"),
        ("ckpt_deployment_end", "Deploy"),
    ]
    overdue_ckpts = 0
    for fld, lbl in ckpt_fields:
        val = str(row.get(fld, "") or "").strip()
        if val and val != "nan":
            d = _parse_dmy(val)
            if d and d < today and status not in ("Completed", "R&M"):
                overdue_ckpts += 1
    if overdue_ckpts > 0:
        score -= overdue_ckpts * 8
        issues.append(f"{overdue_ckpts} overdue checkpoint(s)")

    score = max(0, min(100, score))
    if score >= 80:
        color, label = "#10B981", "Healthy"
    elif score >= 55:
        color, label = "#F59E0B", "At Risk"
    else:
        color, label = "#EF4444", "Critical"

    return {"score": score, "color": color, "label": label, "issues": issues}


def health_badge_html(row: dict) -> str:
    h = compute_health_score(row)
    return (
        f'<span style="display:inline-flex;align-items:center;gap:3px;padding:2px 7px;'
        f'border-radius:12px;font-size:10px;font-weight:700;'
        f'background:{h["color"]}18;color:{h["color"]};border:1px solid {h["color"]}40">'
        f'<span style="width:5px;height:5px;border-radius:50%;background:{h["color"]};display:inline-block"></span>'
        f'{h["label"]}</span>'
    )


def format_relative_time(ts_str: str) -> str:
    if not ts_str:
        return "—"
    try:
        ts = datetime.fromisoformat(str(ts_str)[:19])
        delta = datetime.now() - ts
        secs = int(delta.total_seconds())
        if secs < 60:
            return "just now"
        if secs < 3600:
            return f"{secs // 60}m ago"
        if secs < 86400:
            return f"{secs // 3600}h ago"
        if secs < 604800:
            return f"{secs // 86400}d ago"
        return ts.strftime("%d %b")
    except Exception:
        return str(ts_str)[:10]


def generate_pdf_report(df: pd.DataFrame, stats: dict, company: str = "Qualesce") -> bytes:
    """Generate a PDF project summary report. Returns bytes."""
    try:
        from fpdf import FPDF
    except ImportError:
        return b""

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Header
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_text_color(31, 59, 77)
    pdf.cell(0, 12, f"{company} - Project Portfolio Report", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(100, 116, 139)
    pdf.cell(0, 6, f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}", ln=True)
    pdf.ln(4)

    # KPI summary
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(31, 59, 77)
    pdf.cell(0, 8, "Portfolio Summary", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(55, 65, 81)
    kpis = [
        ("Total Projects", stats.get("total", 0)),
        ("In Progress", stats.get("in_progress", 0)),
        ("Completed", stats.get("completed", 0)),
        ("UAT", stats.get("uat", 0)),
        ("R&M", stats.get("rm", 0)),
        ("Total Hrs Saved", f"{stats.get('total_hrs', 0):,.0f}"),
        ("Total Cost Saved", f"Rs {stats.get('total_cost', 0):,.0f}"),
    ]
    for label, value in kpis:
        pdf.cell(70, 7, f"  {label}:", 0)
        pdf.cell(0, 7, str(value), ln=True)
    pdf.ln(4)

    # Projects table
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(31, 59, 77)
    pdf.cell(0, 8, "Active Projects", ln=True)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(241, 245, 249)
    pdf.set_text_color(71, 85, 105)
    headers = ["#", "Project", "Client", "Status", "ROI%"]
    widths = [8, 72, 50, 28, 22]
    for h, w in zip(headers, widths):
        pdf.cell(w, 7, h, border=1, fill=True)
    pdf.ln()

    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(55, 65, 81)
    active = df[df.get("is_active", pd.Series([True]*len(df))).astype(str).str.lower().isin(["true","1","yes"])] if "is_active" in df.columns else df
    for i, (_, row) in enumerate(active.iterrows()):
        fill = i % 2 == 1
        if fill:
            pdf.set_fill_color(248, 250, 252)
        pdf.cell(8,  6, str(row.get("id", "")), border=1, fill=fill)
        pdf.cell(72, 6, str(row.get("name", ""))[:45], border=1, fill=fill)
        pdf.cell(50, 6, str(row.get("client", ""))[:28], border=1, fill=fill)
        pdf.cell(28, 6, str(row.get("status", ""))[:16], border=1, fill=fill)
        pdf.cell(22, 6, str(row.get("roi_pct", ""))[:8], border=1, fill=fill)
        pdf.ln()

    return bytes(pdf.output())


@st.cache_data(ttl=300, show_spinner=False)
def get_api_key() -> str:
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass
    env_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if env_key:
        return env_key
    try:
        return auth.get_anthropic_api_key()
    except Exception:
        return ""

def is_new(row) -> bool:
    return str(row.get("is_new","")).lower() in ["true","1","yes"]

# ── HTML HELPERS ──────────────────────────────────────────────────────────────
esc = html.escape   # shorthand — always escape user-sourced values before HTML injection

def _parse_dmy(s: str):
    """Parse any supported date format into a date object."""
    v = str(s).strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(v, fmt).date()
        except ValueError:
            pass
    return None

def _parse_ymd(s: str):
    """Parse YYYY-MM-DD or any supported format into a date object."""
    v = str(s).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(v, fmt).date()
        except ValueError:
            pass
    return None

def _due_cell(due_str: str) -> str:
    v = str(due_str).strip()
    if not v:
        return '<span style="font-size:10px;color:#CBD5E1">—</span>'
    d = _parse_dmy(v)
    if not d:
        return f'<span style="font-size:11px;color:#64748B">{esc(v)}</span>'
    diff = (d - date.today()).days
    if diff < 0:
        return f'<span style="font-size:10px;font-weight:700;background:#FEF2F2;color:#991B1B;padding:2px 5px;border-radius:4px">{esc(v)}</span>'
    if diff <= 7:
        return f'<span style="font-size:10px;font-weight:700;background:#FFFBEB;color:#92400E;padding:2px 5px;border-radius:4px">{esc(v)}</span>'
    return f'<span style="font-size:11px;color:#64748B">{esc(v)}</span>'

def badge_html(status: str) -> str:
    s = STATUS_STYLES.get(status, {"bg":"#F1F5F9","text":"#475569","dot":"#94A3B8"})
    return (f'<span style="display:inline-flex;align-items:center;gap:4px;padding:3px 10px;'
            f'border-radius:20px;font-size:11px;font-weight:700;background:{s["bg"]};color:{s["text"]}">'
            f'<span style="width:6px;height:6px;border-radius:50%;background:{s["dot"]};'
            f'display:inline-block"></span>{esc(status)}</span>')

def cell(val, size: str = "11px", color: str = "#374151") -> str:
    """Render a safe, consistently-styled table cell span."""
    return f'<span style="font-size:{size};color:{color}">{esc(str(val))}</span>'

def _inline_md(t: str) -> str:
    t = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', t)
    t = re.sub(r'\*(.+?)\*', r'<em>\1</em>', t)
    t = re.sub(r'`([^`]+)`', r'<code style="background:#F1F5F9;padding:1px 4px;border-radius:3px;font-size:11px;font-family:monospace">\1</code>', t)
    return t

def _is_table_separator(line: str) -> bool:
    return bool(re.match(r'^\|[\s\-|:]+\|$', line.strip()))

def _parse_table_row(line: str) -> list:
    cells = line.strip().strip('|').split('|')
    return [c.strip() for c in cells]

def md_to_html(text: str) -> str:
    raw_lines = str(text).split('\n')
    out, in_list, i = [], False, 0
    while i < len(raw_lines):
        line = raw_lines[i]
        # ── Markdown table detection ──────────────────────────────────────────
        if (line.strip().startswith('|') and
                i + 1 < len(raw_lines) and _is_table_separator(raw_lines[i + 1])):
            if in_list: out.append('</ul>'); in_list = False
            header_cells = _parse_table_row(line)
            i += 2
            body_rows = []
            while i < len(raw_lines) and raw_lines[i].strip().startswith('|'):
                body_rows.append(_parse_table_row(raw_lines[i]))
                i += 1
            th = ''.join(
                f'<th style="padding:6px 12px;text-align:left;font-size:11px;'
                f'font-weight:700;text-transform:uppercase;letter-spacing:.4px;'
                f'color:#475569;background:#F1F5F9;border-bottom:2px solid #E2E8F0">'
                f'{html.escape(c)}</th>' for c in header_cells)
            rows_html = ''
            for ri, row in enumerate(body_rows):
                bg = '#ffffff' if ri % 2 == 0 else '#F8FAFC'
                td = ''.join(
                    f'<td style="padding:6px 12px;font-size:12px;color:#334155;'
                    f'border-bottom:1px solid #F1F5F9">{_inline_md(html.escape(c))}</td>'
                    for c in row)
                rows_html += f'<tr style="background:{bg}">{td}</tr>'
            out.append(
                f'<div style="overflow-x:auto;margin:8px 0">'
                f'<table style="width:100%;border-collapse:collapse;border:1px solid #E2E8F0;'
                f'border-radius:8px;overflow:hidden;font-family:inherit">'
                f'<thead><tr>{th}</tr></thead>'
                f'<tbody>{rows_html}</tbody>'
                f'</table></div>')
            continue
        # ── Everything else ───────────────────────────────────────────────────
        esc = html.escape(line)
        m = re.match(r'^(#{1,3}) (.+)$', esc)
        if m:
            if in_list: out.append('</ul>'); in_list = False
            sz = {1: '15px', 2: '14px', 3: '13px'}[len(m.group(1))]
            out.append(f'<div style="font-size:{sz};font-weight:700;margin:6px 0 2px">{_inline_md(m.group(2))}</div>')
        elif re.match(r'^[-*] ', esc):
            if not in_list: out.append('<ul style="margin:4px 0;padding-left:18px">'); in_list = True
            out.append(f'<li style="margin:2px 0">{_inline_md(esc[2:])}</li>')
        elif re.match(r'^\d+\. ', esc):
            if not in_list: out.append('<ul style="margin:4px 0;padding-left:18px">'); in_list = True
            out.append(f'<li style="margin:2px 0">{_inline_md(re.sub(r"^\d+[.] ", "", esc))}</li>')
        elif not esc.strip():
            if in_list: out.append('</ul>'); in_list = False
            out.append('<br>')
        else:
            if in_list: out.append('</ul>'); in_list = False
            out.append(_inline_md(esc) + '<br>')
        i += 1
    if in_list:
        out.append('</ul>')
    return ''.join(out)


# ── SESSION STATE ─────────────────────────────────────────────────────────────
if "projects" not in st.session_state:
    st.session_state.projects = load_projects()
if "is_active" not in st.session_state.projects.columns:
    st.session_state.projects["is_active"] = True
if "proj_type" not in st.session_state.projects.columns:
    st.session_state.projects["proj_type"] = ""
if "due_date" not in st.session_state.projects.columns:
    st.session_state.projects["due_date"] = ""
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content":
        "Hello! I'm your **AI Project Manager Agent**.\n\n"
        "I have live access to all **46 Qualesce projects** across Raychem, TEPL, "
        "Swagekklok-California, Swagelok-Alabama and internal/external POCs.\n\n"
        "Ask me anything about projects, team workload, status breakdown, or ROI!"}]
if "next_id" not in st.session_state:
    ids = pd.to_numeric(st.session_state.projects.get("id", pd.Series([])), errors="coerce").dropna()
    st.session_state.next_id = int(ids.max()) + 1 if not ids.empty else 1
if "active_tab"           not in st.session_state: st.session_state.active_tab           = "dashboard"
if "dash_slicer"          not in st.session_state: st.session_state.dash_slicer          = None
if "show_modal"           not in st.session_state: st.session_state.show_modal           = None
if "confirm_delete"       not in st.session_state: st.session_state.confirm_delete       = None
if "file_panel_proj"      not in st.session_state: st.session_state.file_panel_proj      = None
if "file_panel_name"      not in st.session_state: st.session_state.file_panel_name      = ""
if "toast"                not in st.session_state: st.session_state.toast                = None
if "dismissed_notifs"     not in st.session_state: st.session_state.dismissed_notifs     = set()
if "show_notif_detail"    not in st.session_state: st.session_state.show_notif_detail    = None
if "project_filter_preset"  not in st.session_state: st.session_state.project_filter_preset  = "All"
if "presales_filter_preset" not in st.session_state: st.session_state.presales_filter_preset = "All"
if "lc_edit_id"            not in st.session_state: st.session_state.lc_edit_id            = None
if "sl_edit_id"            not in st.session_state: st.session_state.sl_edit_id            = None
if "sl_mail_id"            not in st.session_state: st.session_state.sl_mail_id            = None
if "lc_mail_id"            not in st.session_state: st.session_state.lc_mail_id            = None
if "sl_send_all_trigger"  not in st.session_state: st.session_state.sl_send_all_trigger  = False
if "lc_last_notif_check"  not in st.session_state: st.session_state.lc_last_notif_check  = ""
if "dash_client_filter"    not in st.session_state: st.session_state.dash_client_filter    = "All"
if "dash_slicers_expanded" not in st.session_state: st.session_state.dash_slicers_expanded = False
if "current_user"         not in st.session_state: st.session_state.current_user         = None
if "reset_pwd_uid"        not in st.session_state: st.session_state.reset_pwd_uid        = None
if "login_attempts"       not in st.session_state: st.session_state.login_attempts       = 0
if "forgot_step"          not in st.session_state: st.session_state.forgot_step          = None
if "forgot_email"         not in st.session_state: st.session_state.forgot_email         = ""
if "forgot_otp"           not in st.session_state: st.session_state.forgot_otp           = ""
if "forgot_otp_expiry"    not in st.session_state: st.session_state.forgot_otp_expiry    = None
if "forgot_uid"           not in st.session_state: st.session_state.forgot_uid           = None
if "user_edit_id"         not in st.session_state: st.session_state.user_edit_id         = None
if "task_comment_view" not in st.session_state: st.session_state.task_comment_view = None
if "_db_excel_bytes"   not in st.session_state: st.session_state["_db_excel_bytes"]  = None
if "_db_sqlite_bytes"  not in st.session_state: st.session_state["_db_sqlite_bytes"] = None
if "_db_sqlite_name"   not in st.session_state: st.session_state["_db_sqlite_name"]  = "qualesce.db"
if "poc_row_edit"     not in st.session_state: st.session_state.poc_row_edit     = None
if "task_popup"       not in st.session_state: st.session_state.task_popup       = None
if "save_popup"       not in st.session_state: st.session_state.save_popup       = None
if "notif_panel_open" not in st.session_state: st.session_state.notif_panel_open = False
if "bulk_selected"    not in st.session_state: st.session_state.bulk_selected    = set()
if "show_gantt"       not in st.session_state: st.session_state.show_gantt       = False

# ── SELF-HEALING AGENT ────────────────────────────────────────────────────────
_SS_HEAL_DEFAULTS = {
    "active_tab":            "dashboard",
    "dash_slicer":           None,
    "show_modal":            None,
    "confirm_delete":        None,
    "toast":                 None,
    "show_notif_detail":     None,
    "dash_client_filter":    "All",
    "dash_slicers_expanded": False,
    "login_attempts":        0,
    "proj_tracker_open":     None,
    "poc_row_edit":          None,
    "task_popup":            None,
    "save_popup":            None,
    "task_comment_view":     None,
    "forgot_step":           None,
    "user_edit_id":          None,
    "lc_edit_id":            None,
    "sl_edit_id":            None,
    "sl_mail_id":            None,
    "lc_mail_id":            None,
    "reset_pwd_uid":         None,
    "project_filter_preset": "All",
    "presales_filter_preset":"All",
    "sl_send_all_trigger":   False,
    "lc_last_notif_check":   "",
}
_VALID_TABS = {"dashboard","projects","presales","license","agent","users","settings","tasks"}
_VOLATILE_KEYS = [
    "show_modal","confirm_delete","show_notif_detail","task_popup",
    "save_popup","poc_row_edit","proj_tracker_open","task_comment_view",
]

def _self_heal() -> bool:
    """Inspect and repair broken session state every render. Returns True if healing occurred."""
    healed = False
    _now = datetime.now().timestamp()

    # ── Rerun-loop guard: >12 reruns within 4 s → reset all volatile popups ──
    if "_heal_ts" not in st.session_state:
        st.session_state._heal_ts  = _now
        st.session_state._heal_cnt = 0
    else:
        if _now - st.session_state._heal_ts < 4.0:
            st.session_state._heal_cnt += 1
        else:
            st.session_state._heal_ts  = _now
            st.session_state._heal_cnt = 0
    if st.session_state._heal_cnt > 12:
        for _k in _VOLATILE_KEYS:
            st.session_state[_k] = None
        st.session_state._heal_cnt = 0
        healed = True

    # ── Fill any missing session-state keys ───────────────────────────────────
    for _k, _v in _SS_HEAL_DEFAULTS.items():
        if _k not in st.session_state:
            st.session_state[_k] = _v
            healed = True

    # ── Validate active_tab is a real tab ─────────────────────────────────────
    if st.session_state.get("active_tab") not in _VALID_TABS:
        st.session_state.active_tab = "dashboard"
        healed = True

    # ── Ensure dismissed_notifs is always a set ───────────────────────────────
    if not isinstance(st.session_state.get("dismissed_notifs"), set):
        try:
            st.session_state.dismissed_notifs = set(st.session_state.dismissed_notifs or [])
        except Exception:
            st.session_state.dismissed_notifs = set()
        healed = True

    # ── Repair projects DataFrame ─────────────────────────────────────────────
    _proj = st.session_state.get("projects")
    _need_reload = (
        _proj is None
        or not isinstance(_proj, pd.DataFrame)
        or _proj.empty
    )
    if _need_reload:
        try:
            st.session_state.projects = load_projects()
        except Exception:
            st.session_state.projects = pd.DataFrame(columns=EXCEL_COLS)
        healed = True
    else:
        # Ensure schema columns always present
        for _col, _def in [("is_active","True"),("proj_type",""),
                            ("due_date",""),("lead","")]:
            if _col not in st.session_state.projects.columns:
                st.session_state.projects[_col] = _def
                healed = True

    return healed

_healed = _self_heal()

# ── HELPERS ───────────────────────────────────────────────────────────────────
def get_stats(d):
    if d.empty:
        return dict(total=0, rm=0, uat=0, completed=0, in_progress=0, poc=0, pdd=0,
                    discontinued=0, important=0, new_added=0, total_hrs=0.0, total_cost=0.0)
    new_mask = d["is_new"].astype(str).str.lower().isin(["true","1","yes"]) if "is_new" in d.columns else pd.Series([False]*len(d))
    hrs  = pd.to_numeric(d.get("hours_saved", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    cost = pd.to_numeric(d.get("cost_saved",  pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    status_counts = d["status"].value_counts() if "status" in d.columns else pd.Series(dtype=int)
    def c(s): return int(d["status"].str.contains(s, na=False).sum())
    return dict(total=len(d), rm=c("R&M"), uat=c("UAT"), completed=c("Completed"),
                in_progress=c("In Progress"), poc=c("POC"), pdd=c("PDD"),
                discontinued=c("Discontinued"), important=c("Important"), new_added=int(new_mask.sum()),
                total_hrs=float(hrs), total_cost=float(cost))

_anthropic_client: anthropic.Anthropic | None = None

def _get_anthropic_client(api_key: str) -> anthropic.Anthropic:
    global _anthropic_client
    if _anthropic_client is None:
        _anthropic_client = anthropic.Anthropic(api_key=api_key)
    return _anthropic_client

def call_claude(api_key, msgs, df):
    client = _get_anthropic_client(api_key)

    # ── Projects (from Excel / session state) ─────────────────────────────────
    proj_lines = "\n".join(
        f"- {r['name']} | Client: {r['client']} | Employee: {r['employee']} | Status: {r['status']}"
        + (f" | Hours Saved: {r['hours_saved']}" if r.get('hours_saved') else "")
        + (f" | Cost Saved: {r['cost_saved']}" if r.get('cost_saved') else "")
        for r in df.to_dict('records'))

    # ── Tasks (live from DB) ───────────────────────────────────────────────────
    try:
        _tasks = auth.get_all_tasks()
        task_lines = "\n".join(
            f"- [{t['status']} {t['progress']}%] {t['title']} | Assigned to: {t['assigned_to']} | By: {t['assigned_by']}"
            + (f" | Due: {t['due_date']}" if t.get('due_date') else "")
            + (f" | Note: {t['comment']}" if t.get('comment') else "")
            for t in _tasks
        ) if _tasks else "No tasks yet."
    except Exception:
        task_lines = "Task data unavailable."

    # ── Licenses (live from DB) ────────────────────────────────────────────────
    try:
        _lics = auth.get_all_licenses()
        lic_lines = "\n".join(
            f"- {l['tool_name']} | Qty: {l['no_of_licenses']} | {l['start_date']} to {l['end_date']}"
            for l in _lics
        ) if _lics else "No licenses."
    except Exception:
        lic_lines = "License data unavailable."

    # ── Users (live from DB) ───────────────────────────────────────────────────
    try:
        _users = auth.get_all_users()
        user_lines = "\n".join(
            f"- {u[1]} | {u[3].upper()} | {u[2]}"
            for u in _users if u[4]  # only active users
        ) if _users else "No users."
    except Exception:
        user_lines = "User data unavailable."

    live_ctx = f"""

LIVE DATABASE (source: projects.xlsx + qualesce.db):

[PROJECTS — {len(df)} total]
{proj_lines}

[TASKS — {len(_tasks) if isinstance(_tasks, list) else 0} total]
{task_lines}

[LICENSES]
{lic_lines}

[TEAM MEMBERS]
{user_lines}"""

    api_msgs = [{"role": m["role"], "content": m["content"]} for m in msgs[-12:]]
    while api_msgs and api_msgs[0]["role"] != "user":
        api_msgs = api_msgs[1:]
    if not api_msgs:
        return "Please ask me a question to get started!"
    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=500,
        system=SYSTEM_PROMPT + live_ctx,
        messages=api_msgs)
    return resp.content[0].text

# ── STYLES ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url("https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap");

html,body,[class*="css"]{
  font-family:'Inter','Segoe UI',-apple-system,BlinkMacSystemFont,sans-serif!important;
  background:#F7F8F9!important;
  color:#1F3B4D!important;
  -webkit-font-smoothing:antialiased!important;
  -moz-osx-font-smoothing:grayscale!important;
  text-rendering:optimizeLegibility!important;
}
#MainMenu,footer,header{visibility:hidden}
.block-container{padding:0 1.5rem 2rem!important;max-width:100%!important}
section[data-testid="stSidebar"]{display:none!important}

/* ── KPI Cards ── */
.kpi-wrap{
  text-align:center;padding:18px 12px;border-radius:8px;background:#FFFFFF;
  border:1px solid #DFE3E7;
  box-shadow:0 1px 2px rgba(15,30,42,.06),0 1px 1px rgba(15,30,42,.04);
  cursor:pointer;transition:box-shadow 120ms cubic-bezier(0.22,1,0.36,1)}
.kpi-wrap:hover{box-shadow:0 4px 10px rgba(15,30,42,.08),0 2px 4px rgba(15,30,42,.04)}
.kpi-num{font-family:'JetBrains Mono','Courier New',monospace;font-size:28px;font-weight:700;margin:8px 0 4px;letter-spacing:-1px;color:#1F3B4D}
.kpi-lbl{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#9BA5AE}

/* ── Generic content card ── */
.q-card{
  background:#FFFFFF;border:1px solid #DFE3E7;border-radius:8px;
  padding:18px 20px;box-shadow:0 1px 2px rgba(15,30,42,.06)}

/* ── Top Navigation ── */
.q-nav{
  background:#162C3B;
  padding:0 28px;
  display:flex;align-items:center;justify-content:space-between;
  height:62px;position:sticky;top:0;z-index:100;
  box-shadow:0 2px 12px rgba(0,0,0,.30);
  margin:0 -1.5rem 24px}
.q-nav-btns{display:flex;align-items:center;gap:8px}

/* ── Slicer rows ── */
.srow{
  padding:11px 18px;border-bottom:1px solid #F1F5F9;
  display:flex;justify-content:space-between;align-items:center;
  transition:background .12s}
.srow:nth-child(even){background:#F8FAFC}
.srow:hover{background:#EFF7F7}

/* ── Project table rows ── */
.prow{padding:9px 8px;border-bottom:1px solid #F1F5F9;background:#fff;transition:background .12s}
.prow:nth-child(even){background:#F8FAFC}
.prow:hover{background:#EFF7F7}

/* ── Table section header label ── */
.q-tbl-hdr{padding:11px 16px 8px;background:#F8FAFC;border-bottom:2px solid #DFE3E7;
  font-size:11px;font-weight:700;color:#1F3B4D;letter-spacing:-.1px;display:flex;
  align-items:center;justify-content:space-between}
.q-tbl-count{font-size:11px;color:#64748B;font-weight:500}

/* ── Chat bubbles ── */
@keyframes slideInRight{from{opacity:0;transform:translateX(30px)}to{opacity:1;transform:translateX(0)}}
@keyframes slideInLeft{from{opacity:0;transform:translateX(-30px)}to{opacity:1;transform:translateX(0)}}
@keyframes typingPulse{0%,80%,100%{transform:scale(0);opacity:.3}40%{transform:scale(1);opacity:1}}
.chat-row{display:flex;align-items:flex-end;gap:8px;margin:8px 0}
.chat-row.user-row{flex-direction:row-reverse}
.chat-avatar{
  width:30px;height:30px;border-radius:50%;display:flex;align-items:center;
  justify-content:center;font-size:14px;flex-shrink:0;font-weight:700}
.avatar-user{background:#D9ECEC;color:#3F8E91}
.avatar-bot{background:#E5F2EC;color:#2E7D5B}
.chat-user{
  background:#EFF7F7;border:1px solid #B6DADB;
  border-radius:16px 16px 4px 16px;padding:12px 16px;font-size:13px;line-height:1.6;
  max-width:80%;animation:slideInRight .3s cubic-bezier(0.22,1,0.36,1)}
.chat-bot{
  background:#E5F2EC;border:1px solid #A8D5BE;
  border-radius:16px 16px 16px 4px;padding:12px 16px;font-size:13px;line-height:1.6;
  max-width:80%;animation:slideInLeft .3s cubic-bezier(0.22,1,0.36,1)}
.typing-indicator{
  display:flex;align-items:center;gap:10px;
  background:#E5F2EC;border:1px solid #A8D5BE;
  border-radius:16px 16px 16px 4px;padding:12px 16px;
  width:fit-content;animation:slideInLeft .3s cubic-bezier(0.22,1,0.36,1)}
.typing-dots{display:flex;gap:4px;align-items:center}
.typing-dots span{
  width:7px;height:7px;border-radius:50%;background:#2E7D5B;display:inline-block}
.typing-dots span:nth-child(1){animation:typingPulse 1.2s infinite ease-in-out}
.typing-dots span:nth-child(2){animation:typingPulse 1.2s infinite ease-in-out .2s}
.typing-dots span:nth-child(3){animation:typingPulse 1.2s infinite ease-in-out .4s}

/* ── ROI Banner ── */
.roi-banner{
  background:linear-gradient(135deg,#162C3B,#244E51);
  border:1px solid rgba(95,169,171,.30);
  border-radius:8px;padding:18px 26px;
  display:flex;gap:32px;align-items:center;margin-bottom:20px;
  box-shadow:0 4px 10px rgba(15,30,42,.08)}

/* ── Streamlit buttons ── */
div[data-testid="stButton"] > button{
  border-radius:6px!important;
  font-family:'Inter','Segoe UI',sans-serif!important;
  font-weight:600!important;
  font-size:12px!important;
  letter-spacing:.2px!important}

/* ── Notification popup ── */
.notif-popup{
  border-radius:8px;padding:20px 24px;margin-bottom:20px;
  box-shadow:0 4px 10px rgba(15,30,42,.08);}
.notif-alert{}

/* ── Streamlit container borders ── */
div[data-testid="stVerticalBlockBorderWrapper"]{
  border:1px solid #E2E8F0!important;border-radius:10px!important;
  box-shadow:0 2px 8px rgba(15,30,42,.05)!important;
  overflow:hidden!important}

/* ── Keyframes for edit button animation ── */
@keyframes editWiggle{
  0%,100%{transform:rotate(0deg) scale(1)}
  20%    {transform:rotate(-12deg) scale(1.15)}
  40%    {transform:rotate(10deg) scale(1.12)}
  60%    {transform:rotate(-7deg) scale(1.10)}
  80%    {transform:rotate(5deg) scale(1.05)}}

/* ── Table action icon buttons (✏ edit · 🗑 delete · 🔑 warn) ── */
div[data-testid="stMarkdownContainer"]:has(.act-edit-marker) ~ div[data-testid="stButton"] > button{
  background:#EFF7F7!important;border:1.5px solid #B6DADB!important;
  color:#3F8E91!important;font-size:15px!important;
  min-height:30px!important;padding:2px 8px!important;
  border-radius:8px!important;
  transition:background 150ms,border-color 150ms,box-shadow 150ms,transform 150ms!important}
div[data-testid="stMarkdownContainer"]:has(.act-edit-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#D9ECEC!important;border-color:#3F8E91!important;
  animation:editWiggle .45s ease forwards!important;
  box-shadow:0 4px 14px rgba(63,142,145,.35)!important}
div[data-testid="stMarkdownContainer"]:has(.act-del-marker) ~ div[data-testid="stButton"] > button{
  background:#FCEAEA!important;border:1.5px solid #F0BABA!important;
  color:#B23A3A!important;font-size:15px!important;
  min-height:30px!important;padding:2px 8px!important;
  transition:background 120ms,border-color 120ms!important}
div[data-testid="stMarkdownContainer"]:has(.act-del-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#F5CECE!important;border-color:#D98080!important}
div[data-testid="stMarkdownContainer"]:has(.act-warn-marker) ~ div[data-testid="stButton"] > button{
  background:#FBF6E7!important;border:1.5px solid #ECD58A!important;
  color:#966D17!important;font-size:15px!important;
  min-height:30px!important;padding:2px 8px!important;
  transition:background 120ms,border-color 120ms!important}
div[data-testid="stMarkdownContainer"]:has(.act-warn-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#F5E9C2!important;border-color:#D4A02C!important}

/* ── Login ── */
.login-hint{text-align:center;font-size:11px;color:#9BA5AE;margin-top:12px}

/* ── Task progress bar ── */
.progress-bar-outer{background:#DFE3E7;border-radius:99px;height:7px;overflow:hidden;margin:4px 0}
.progress-bar-inner{height:7px;border-radius:99px}

/* ── Role badge ── */
.role-badge{font-size:9px;font-weight:700;padding:2px 8px;border-radius:999px;text-transform:uppercase}

/* ── Tab-switch: smooth fade-in for all main content ── */
@keyframes tabFadeIn{
  0%  {opacity:0;transform:translateY(6px)}
  100%{opacity:1;transform:translateY(0)}
}
section[data-testid="stMain"] > div > div[data-testid="stVerticalBlock"] {
  animation:tabFadeIn .22s cubic-bezier(0.22,1,0.36,1) both;
}

/* ── Hide Streamlit's default bouncing loading bar ── */
[data-testid="stStatusWidget"]{display:none!important}
@keyframes topBarSlide{
  0%  {opacity:0;width:0}
  30% {opacity:1;width:60%}
  100%{opacity:0;width:100%}
}

/* ── KPI expand animation ── */
@keyframes kpi-slide-in{
  0%  {opacity:0;transform:translateY(-10px)}
  100%{opacity:1;transform:translateY(0)}
}
.kpi-anim{animation:kpi-slide-in .28s cubic-bezier(0.22,1,0.36,1) both}

/* ── Toast notification animation ── */
@keyframes toastPop{
  0%  {opacity:0;transform:translateY(-18px) scale(0.88)}
  55% {opacity:1;transform:translateY(5px)   scale(1.03)}
  80% {transform:translateY(-2px) scale(1.00)}
  100%{opacity:1;transform:translateY(0)     scale(1)}
}
.toast-anim{animation:toastPop .45s cubic-bezier(0.34,1.4,0.64,1) both}

/* ── Collapse all marker spans (zero height, no extra spacing) ── */
div[data-testid="stMarkdownContainer"]:has(.act-edit-marker),
div[data-testid="stMarkdownContainer"]:has(.act-del-marker),
div[data-testid="stMarkdownContainer"]:has(.act-warn-marker) {
  height:0!important;overflow:hidden!important;margin:0!important;padding:0!important;
  line-height:0!important;font-size:0!important}

/* ── Project table: collapse marker spans ── */
div[data-testid="stMarkdownContainer"]:has(.proj-act-marker),
div[data-testid="stMarkdownContainer"]:has(.proj-ts-marker){
  height:0!important;overflow:hidden!important;margin:0!important;padding:0!important;
  line-height:0!important;font-size:0!important}

/* Base style for all project action + timeline buttons */
div[data-testid="stMarkdownContainer"]:has(.proj-act-marker) ~ div[data-testid="stButton"] > button,
div[data-testid="stMarkdownContainer"]:has(.proj-ts-marker) ~ div[data-testid="stButton"] > button{
  min-height:30px!important;max-height:34px!important;
  padding:0 4px!important;font-size:15px!important;
  display:flex!important;align-items:center!important;justify-content:center!important;
  transition:background 120ms,border-color 120ms!important}
div[data-testid="stMarkdownContainer"]:has(.proj-act-marker) ~ div[data-testid="stButton"] > button > div,
div[data-testid="stMarkdownContainer"]:has(.proj-ts-marker) ~ div[data-testid="stButton"] > button > div{
  display:flex!important;align-items:center!important;justify-content:center!important;width:100%!important}
div[data-testid="stMarkdownContainer"]:has(.proj-act-marker) ~ div[data-testid="stButton"] > button p,
div[data-testid="stMarkdownContainer"]:has(.proj-ts-marker) ~ div[data-testid="stButton"] > button p{
  margin:0!important;line-height:1!important;font-size:15px!important}

/* Edit button */
div[data-testid="stMarkdownContainer"]:has(.proj-edit-marker) ~ div[data-testid="stButton"] > button{
  background:#EFF7F7!important;border:1.5px solid #B6DADB!important;color:#3F8E91!important}
div[data-testid="stMarkdownContainer"]:has(.proj-edit-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#D9ECEC!important;border-color:#8FC4C5!important}
/* Cancel (active edit) button */
div[data-testid="stMarkdownContainer"]:has(.act-warn-marker.proj-act-marker) ~ div[data-testid="stButton"] > button{
  background:#FBF6E7!important;border:1.5px solid #ECD58A!important;color:#966D17!important}
div[data-testid="stMarkdownContainer"]:has(.act-warn-marker.proj-act-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#F5E9C2!important;border-color:#D4A02C!important}
/* Delete button */
div[data-testid="stMarkdownContainer"]:has(.proj-del-marker) ~ div[data-testid="stButton"] > button{
  background:#FCEAEA!important;border:1.5px solid #F0BABA!important;color:#B23A3A!important}
div[data-testid="stMarkdownContainer"]:has(.proj-del-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#F5CECE!important;border-color:#D98080!important}
/* Files button */
div[data-testid="stMarkdownContainer"]:has(.proj-files-marker) ~ div[data-testid="stButton"] > button{
  background:#EFF7F7!important;border:1.5px solid #B6DADB!important;color:#3F8E91!important}
div[data-testid="stMarkdownContainer"]:has(.proj-files-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#D9ECEC!important;border-color:#8FC4C5!important}
/* Timeline button */
div[data-testid="stMarkdownContainer"]:has(.proj-ts-marker) ~ div[data-testid="stButton"] > button{
  background:#F0F9FF!important;border:1.5px solid #BAE6FD!important;color:#0369A1!important}
div[data-testid="stMarkdownContainer"]:has(.proj-ts-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#E0F2FE!important;border-color:#7DD3FC!important}

/* ── Expand arrow button ── */
.expand-btn button{
  border-radius:50%!important;
  width:38px!important;height:38px!important;
  padding:0!important;font-size:16px!important;
  background:#EFF1F3!important;border:1px solid #C5CCD2!important;
  color:#4E5860!important;font-weight:700!important}
.expand-btn button:hover{background:#DFE3E7!important}

/* ── HD table ── */
.hd-table{width:100%;border-collapse:collapse;font-size:12px;background:#fff}
.hd-table th{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#64748B;padding:9px 10px;border-bottom:2px solid #DFE3E7;white-space:nowrap;background:#F8FAFC;text-align:left}
.hd-table td{padding:10px 10px;border-bottom:1px solid #F1F5F9;vertical-align:middle;color:#1F3B4D;line-height:1.5}
.hd-table tr:last-child td{border-bottom:none}
.hd-table tr:hover td{background:#EFF7F7}
.hd-table tr:nth-child(even) td{background:#F8FAFC}
.hd-table tr:nth-child(even):hover td{background:#EFF7F7}

/* ── Task creation popup overlay ── */
.task-popup-overlay{
  position:fixed;top:0;left:0;width:100vw;height:100vh;
  background:rgba(15,23,42,.55);z-index:99999;
  display:flex;align-items:center;justify-content:center;
  animation:popOverlayIn .22s ease-out forwards;cursor:pointer;}
.task-popup-box{
  width:188px;height:188px;border-radius:26px;background:#fff;
  display:flex;flex-direction:column;align-items:center;justify-content:center;
  gap:16px;box-shadow:0 28px 72px rgba(0,0,0,.5);
  animation:popBoxIn .42s cubic-bezier(.34,1.6,.64,1) forwards;}
.task-popup-lbl{font-size:15px;font-weight:800;letter-spacing:-.3px}
.task-popup-spinner{
  width:60px;height:60px;border:5px solid #EFF7F7;
  border-top-color:#5FA9AB;border-right-color:#5FA9AB;
  border-radius:50%;animation:popSpin .65s linear infinite;}
.task-popup-svg{width:72px;height:72px;overflow:visible}
.tp-circle{fill:none;stroke-width:4;stroke-dasharray:165;stroke-dashoffset:165;
  animation:popDraw .5s cubic-bezier(.4,0,.2,1) .06s forwards;}
.tp-tick-c{stroke:#10B981}.tp-cross-c{stroke:#EF4444}
.tp-tick-p{fill:none;stroke:#10B981;stroke-width:5;stroke-linecap:round;
  stroke-linejoin:round;stroke-dasharray:47;stroke-dashoffset:47;
  animation:popDraw .28s ease-out .55s forwards;}
.tp-cross-p{fill:none;stroke:#EF4444;stroke-width:5;stroke-linecap:round;
  stroke-dasharray:34;stroke-dashoffset:34;}
.tp-cross-p1{animation:popDraw .2s ease-out .52s forwards;}
.tp-cross-p2{animation:popDraw .2s ease-out .7s forwards;}
@keyframes popOverlayIn{from{opacity:0}to{opacity:1}}
@keyframes popBoxIn{
  0%{opacity:0;transform:scale(.3)}
  65%{transform:scale(1.08)}
  100%{opacity:1;transform:scale(1)}}
@keyframes popSpin{to{transform:rotate(360deg)}}
@keyframes popDraw{to{stroke-dashoffset:0}}
/* Auto-dismiss: fade in → stay → fade out, total ~1.2s */
@keyframes popLifecycle{
  0%{opacity:0}
  18%{opacity:1}
  70%{opacity:1}
  100%{opacity:0}}
.task-popup-auto{animation:popLifecycle 1.2s ease forwards!important;pointer-events:none!important;cursor:default!important}

</style>
""", unsafe_allow_html=True)

# ── TASK POPUP HTML CONSTANTS ─────────────────────────────────────────────────
_POPUP_LOADING = (
    '<div class="task-popup-overlay" id="tpop" style="cursor:default">'
    '<div class="task-popup-box">'
    '<div class="task-popup-spinner"></div>'
    '<div class="task-popup-lbl" style="color:#5FA9AB">Creating…</div>'
    '</div></div>'
)
_POPUP_SUCCESS = (
    '<div class="task-popup-overlay task-popup-auto">'
    '<div class="task-popup-box">'
    '<svg class="task-popup-svg" viewBox="0 0 54 54">'
    '<circle class="tp-circle tp-tick-c" cx="27" cy="27" r="23"/>'
    '<path class="tp-tick-p" d="M15 28 L23.5 36.5 L39 18"/>'
    '</svg>'
    '<div class="task-popup-lbl" style="color:#10B981">Assigned!</div>'
    '</div></div>'
)
_POPUP_ERROR = (
    '<div class="task-popup-overlay task-popup-auto">'
    '<div class="task-popup-box">'
    '<svg class="task-popup-svg" viewBox="0 0 54 54">'
    '<circle class="tp-circle tp-cross-c" cx="27" cy="27" r="23"/>'
    '<path class="tp-cross-p tp-cross-p1" d="M17 17 L37 37"/>'
    '<path class="tp-cross-p tp-cross-p2" d="M37 17 L17 37"/>'
    '</svg>'
    '<div class="task-popup-lbl" style="color:#EF4444">Not Assigned</div>'
    '</div></div>'
)
_POPUP_SAVED = (
    '<div class="task-popup-overlay task-popup-auto">'
    '<div class="task-popup-box">'
    '<svg class="task-popup-svg" viewBox="0 0 54 54">'
    '<circle class="tp-circle tp-tick-c" cx="27" cy="27" r="23"/>'
    '<path class="tp-tick-p" d="M15 28 L23.5 36.5 L39 18"/>'
    '</svg>'
    '<div class="task-popup-lbl" style="color:#10B981">Saved!</div>'
    '</div></div>'
)

# ── LOGIN GATE ───────────────────────────────────────────────────────────────
def _render_login():
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        with st.container(border=True):
            st.markdown(
                f'<div style="text-align:center;padding:16px 0 20px">'
                + (f'<img src="data:{_LOGO_MIME};base64,{_LOGO_B64}" style="height:80px;width:auto;object-fit:contain;margin-bottom:8px" alt="Qualesce">'
                   if _LOGO_B64 else
                   f'<div style="font-size:36px;font-weight:900;color:#5FA9AB;letter-spacing:-2px;margin-bottom:8px">Q</div>'
                   f'<div style="font-size:20px;font-weight:800;color:#1F3B4D;letter-spacing:-.3px;font-family:Manrope,sans-serif">QUALESCE</div>') +
                f'<div style="font-size:12px;color:#6E7A84;margin-top:4px">Project Tracking</div>'
                f'</div>',
                unsafe_allow_html=True)

            step = st.session_state.forgot_step

            # ── Step 1: Enter email to receive OTP ────────────────────────────
            if step == "enter_email":
                st.markdown('<p style="font-size:13px;color:#475569;margin-bottom:8px">Enter your account email to receive a reset code.</p>', unsafe_allow_html=True)
                with st.form("fp_email_form"):
                    fp_email = st.text_input("Account Email", placeholder="you@company.com")
                    col_send, col_back = st.columns(2)
                    send_btn = col_send.form_submit_button("Send Code", use_container_width=True, type="primary")
                    back_btn = col_back.form_submit_button("Back to Login", use_container_width=True)
                if send_btn:
                    if not fp_email.strip():
                        st.error("Please enter your email.")
                    else:
                        user = auth.get_user_by_email(fp_email.strip())
                        if not user:
                            st.error("No active account found with that email.")
                        else:
                            otp = email_utils.generate_otp()
                            ok, err = email_utils.send_otp_email(user["email"], user["name"], otp)
                            if ok:
                                st.session_state.forgot_otp        = otp
                                st.session_state.forgot_otp_expiry = datetime.now().timestamp() + 600
                                st.session_state.forgot_uid        = user["id"]
                                st.session_state.forgot_email      = user["email"]
                                st.session_state.forgot_step       = "enter_otp"
                                st.rerun()
                            else:
                                st.error(f"Failed to send email: {err}")
                if back_btn:
                    st.session_state.forgot_step = None
                    st.rerun()

            # ── Step 2: Enter OTP ─────────────────────────────────────────────
            elif step == "enter_otp":
                st.markdown(f'<p style="font-size:13px;color:#475569;margin-bottom:8px">A 6-digit code was sent to <b>{st.session_state.forgot_email}</b>. Enter it below.</p>', unsafe_allow_html=True)
                with st.form("fp_otp_form"):
                    entered_otp = st.text_input("Reset Code", placeholder="123456", max_chars=6)
                    col_verify, col_back = st.columns(2)
                    verify_btn = col_verify.form_submit_button("Verify Code", use_container_width=True, type="primary")
                    back_btn   = col_back.form_submit_button("Back", use_container_width=True)
                if verify_btn:
                    expiry = st.session_state.forgot_otp_expiry or 0
                    if datetime.now().timestamp() > expiry:
                        st.error("Code expired. Please request a new one.")
                        st.session_state.forgot_step = "enter_email"
                        st.rerun()
                    elif entered_otp.strip() != st.session_state.forgot_otp:
                        st.error("Incorrect code. Please try again.")
                    else:
                        st.session_state.forgot_step = "new_pwd"
                        st.rerun()
                if back_btn:
                    st.session_state.forgot_step = "enter_email"
                    st.rerun()

            # ── Step 3: Set new password ──────────────────────────────────────
            elif step == "new_pwd":
                st.markdown('<p style="font-size:13px;color:#475569;margin-bottom:8px">Choose a new password for your account.</p>', unsafe_allow_html=True)
                with st.form("fp_newpwd_form"):
                    new_pwd  = st.text_input("New Password", type="password", placeholder="Min 6 characters")
                    conf_pwd = st.text_input("Confirm Password", type="password", placeholder="Repeat password")
                    save_btn = st.form_submit_button("Reset Password", use_container_width=True, type="primary")
                if save_btn:
                    if len(new_pwd) < 6:
                        st.error("Password must be at least 6 characters.")
                    elif new_pwd != conf_pwd:
                        st.error("Passwords do not match.")
                    else:
                        auth.reset_password(st.session_state.forgot_uid, new_pwd)
                        st.session_state.forgot_step       = None
                        st.session_state.forgot_otp        = ""
                        st.session_state.forgot_otp_expiry = None
                        st.session_state.forgot_uid        = None
                        st.session_state.forgot_email      = ""
                        st.session_state.login_attempts    = 0
                        st.success("Password reset! You can now sign in.")
                        st.rerun()

            # ── Normal login form ─────────────────────────────────────────────
            else:
                with st.form("login_form"):
                    email    = st.text_input("Email Address", placeholder="you@company.com")
                    password = st.text_input("Password", type="password", placeholder="••••••••")
                    submitted = st.form_submit_button("Sign In", use_container_width=True, type="primary")
                if submitted:
                    if not email.strip() or not password:
                        st.error("Email and password are required.")
                    else:
                        user = auth.authenticate(email, password)
                        if user:
                            st.session_state.current_user  = user
                            st.session_state.login_attempts = 0
                            st.session_state.active_tab    = "tasks" if user["role"] == "employee" else "dashboard"
                            auth.log_audit(user["id"], user["name"], "LOGIN", "users",
                                           str(user["id"]), f'User "{user["name"]}" logged in')
                            st.rerun()
                        else:
                            st.session_state.login_attempts += 1
                            attempts = st.session_state.login_attempts
                            remaining = max(0, 3 - attempts)
                            if remaining > 0:
                                st.error(f"Invalid credentials or account is inactive. ({remaining} attempt{'s' if remaining != 1 else ''} left before reset option appears)")
                            else:
                                st.error("Invalid credentials or account is inactive.")
                if st.session_state.login_attempts >= 3:
                    st.markdown('<p style="text-align:center;font-size:12px;color:#94A3B8;margin:8px 0 4px">Too many failed attempts</p>', unsafe_allow_html=True)
                    if st.button("Forgot Password?", use_container_width=True):
                        st.session_state.forgot_step = "enter_email"
                        st.rerun()

            st.markdown('<div class="login-hint"> </div>', unsafe_allow_html=True)

if st.session_state.current_user is None:
    _render_login()
    st.stop()

# ── POST-LOGIN SELF-HEAL + SAFETY CHECKS ─────────────────────────────────────
try:
    cu   = st.session_state.current_user
    if not isinstance(cu, dict) or "role" not in cu:
        raise ValueError("Corrupted user session")
    role = cu["role"]
except Exception:
    st.session_state.current_user = None
    st.rerun()
    st.stop()
    raise SystemExit(0)

# ── NAV ───────────────────────────────────────────────────────────────────────
try:
    df = st.session_state.projects
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        df = load_projects()
        st.session_state.projects = df
    stats = get_stats(df)
except Exception as _df_err:
    st.error(f"Data load error — recovering... ({_df_err})")
    st.session_state.projects = pd.DataFrame(columns=EXCEL_COLS)
    df    = st.session_state.projects
    stats = get_stats(df)

_new_badge = f"&nbsp;<span style='color:#34D399;font-weight:600'>+{stats['new_added']} new</span>" if stats["new_added"] else ""
st.markdown("""
<style>
/* ── Remove Streamlit default top padding ────────────────────────────────── */
.block-container, [data-testid="stMainBlockContainer"],
[data-testid="stAppViewContainer"] > .main > .block-container {
    padding-top:0 !important;
}
/* ── Title bar: fixed full-viewport-width, 62px tall ─────────────────────── */
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) {
    position:fixed !important; top:0 !important; left:0 !important; right:0 !important;
    width:100vw !important; max-width:100vw !important;
    background:#162C3B !important; padding:0 !important; margin:0 !important;
    z-index:1000 !important; box-shadow:0 2px 12px rgba(0,0,0,.30) !important;
    height:62px !important; display:flex !important; align-items:center !important; gap:0 !important;
}
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div {
    background:#162C3B !important; display:flex !important;
    align-items:center !important; padding:0 !important; height:62px !important;
}
/* Logo column inner wrappers */
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:first-child > div,
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:first-child [data-testid="stVerticalBlock"],
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:first-child [class*="element-container"] {
    display:flex !important; align-items:center !important;
    padding:0 !important; margin:0 !important; height:62px !important; width:100% !important;
}
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:first-child { padding-left:20px !important; }
/* Stats column: right-aligned */
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:nth-child(2) {
    justify-content:flex-end !important; padding-right:62px !important;
}
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:nth-child(2) > div,
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:nth-child(2) [data-testid="stVerticalBlock"],
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:nth-child(2) [class*="element-container"] {
    display:flex !important; align-items:center !important; justify-content:flex-end !important;
    padding:0 !important; margin:0 !important; height:62px !important; width:100% !important;
}
/* ⋮ column: independently fixed to top-right corner */
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:last-child {
    position:fixed !important; top:0 !important; right:0 !important;
    width:44px !important; height:62px !important;
    display:flex !important; align-items:center !important; justify-content:center !important;
    padding:0 !important; margin:0 !important; z-index:1001 !important;
    background:#162C3B !important;
}
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:last-child > div,
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:last-child [data-testid="stVerticalBlock"],
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:last-child [class*="element-container"] {
    padding:0 !important; margin:0 !important; gap:0 !important;
    display:flex !important; align-items:center !important; justify-content:center !important;
    width:auto !important; min-width:0 !important; height:62px !important;
}
/* ⋮ button */
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:last-child button,
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:last-child button:focus,
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:last-child button:active {
    background:transparent !important; border:none !important;
    box-shadow:none !important; outline:none !important;
    position:relative !important; overflow:hidden !important;
    width:24px !important; height:24px !important; min-width:0 !important; min-height:0 !important;
    padding:0 !important; cursor:pointer !important;
}
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:last-child button > * {
    visibility:hidden !important;
}
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:last-child button::before {
    content:"⋮" !important; position:absolute !important; inset:0 !important;
    display:flex !important; align-items:center !important; justify-content:center !important;
    font-size:20px !important; font-weight:900 !important;
    color:#DC2626 !important; visibility:visible !important; z-index:10 !important;
}
div[data-testid="stHorizontalBlock"]:has(.q-nav-left) > div:last-child button:hover::before {
    color:#B91C1C !important;
}
/* Collapse the act-add-marker wrapper so it takes no space and doesn't push the button down */
[class*="element-container"]:has(.act-add-marker) {
    display:none !important;
}
/* Add button in Projects filter — align the column so button sits flush with other filter inputs */
[class*="element-container"]:has(.act-add-marker) + [class*="element-container"] {
    display:flex !important; align-items:center !important;
    height:38px !important; margin:0 !important; padding:0 !important;
}
[class*="element-container"]:has(.act-add-marker) + [class*="element-container"] .stButton {
    width:100% !important; height:38px !important;
    display:flex !important; align-items:center !important; margin:0 !important; padding:0 !important;
}
[class*="element-container"]:has(.act-add-marker) + [class*="element-container"] .stButton button {
    background:#DC2626 !important; color:#fff !important; border-color:#DC2626 !important;
    border-radius:4px !important; height:38px !important; width:100% !important;
    display:flex !important; align-items:center !important; justify-content:center !important;
    margin:0 !important; padding:0 6px !important;
}
[class*="element-container"]:has(.act-add-marker) + [class*="element-container"] .stButton button:hover {
    background:#B91C1C !important; border-color:#B91C1C !important;
}
/* Mail buttons in License table rows (purchased + sold) */
div[data-testid="stHorizontalBlock"] button[data-testid^="sl_mail_"],
div[data-testid="stHorizontalBlock"] button[data-testid^="lc_mail_"] {
    background:#2563EB !important; color:#fff !important;
    border-color:#3F8E91 !important; border-radius:3px !important;
    font-size:9px !important; font-weight:700 !important;
    height:24px !important; min-height:24px !important;
    padding:0 6px !important; line-height:1 !important;
    white-space:nowrap !important; overflow:hidden !important;
}
div[data-testid="stHorizontalBlock"] button[data-testid^="sl_mail_"] p,
div[data-testid="stHorizontalBlock"] button[data-testid^="lc_mail_"] p {
    white-space:nowrap !important; font-size:9px !important; margin:0 !important;
}
/* Logout popover */
div[data-testid="stPopover"] .stButton button {
    font-size:12px !important; height:26px !important; padding:0 12px !important;
    width:100% !important; font-weight:600 !important;
    background:#DC2626 !important; color:#fff !important; border-color:#DC2626 !important;
    border-radius:4px !important;
}
div[data-testid="stPopover"] .stButton button:hover {
    background:#B91C1C !important; border-color:#B91C1C !important;
}
</style>
""", unsafe_allow_html=True)
_hdr_l, _hdr_m, _hdr_r = st.columns([4, 7, 1])
with _hdr_l:
    st.markdown(
        f'<div class="q-nav-left" style="display:flex;align-items:center;height:62px">'
        f'<div style="display:flex;align-items:center;gap:14px">'
        + (f'<img src="data:{_LOGO_MIME};base64,{_LOGO_B64}" style="height:44px;width:auto;object-fit:contain" alt="Qualesce">'
           if _LOGO_B64 else
           f'<div style="width:38px;height:38px;background:linear-gradient(135deg,#5FA9AB,#3F8E91);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:800;color:#fff;letter-spacing:-0.5px;box-shadow:0 0 0 1px rgba(255,255,255,.12)">Q</div>') +
        f'<div>'
        f'<div style="font-size:11px;color:#94A3B8;letter-spacing:1.2px;text-transform:uppercase;font-weight:500;margin-top:1px">AI Project Manager</div>'
        f'</div>'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True
    )
with _hdr_m:
    _hdr_live_part = (
        f'<span style="width:8px;height:8px;border-radius:50%;background:#10B981;box-shadow:0 0 8px #10B981;display:inline-block;vertical-align:middle;flex-shrink:0"></span>'
        f'<b style="color:#E2E8F0;font-weight:600;vertical-align:middle">{stats["total"]}</b>'
        f'<span style="vertical-align:middle">projects live</span>'
        f'{_new_badge}'
        f'<span style="color:#475569;vertical-align:middle">|</span>'
    ) if role != "employee" else ""
    st.markdown(
        f'<div style="display:flex;align-items:center;justify-content:flex-end;height:62px;width:100%">'
        f'<div style="font-size:14px;color:#94A3B8;display:flex;align-items:center;gap:10px;line-height:1">'
        f'{_hdr_live_part}'
        f'<span style="color:#E2E8F0;font-weight:600;vertical-align:middle">{esc(cu["name"])}</span>'
        f'<span style="background:#1E3A8A;color:#93C5FD;font-size:11px;font-weight:700;padding:2px 10px;border-radius:10px;text-transform:uppercase;vertical-align:middle">{esc(cu["role"])}</span>'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True
    )
with _hdr_r:
    _unread_count = auth.count_unread_notifications(cu["id"]) if cu else 0
    _bell_label = f"🔔 {_unread_count}" if _unread_count > 0 else "🔔"
    with st.popover(_bell_label, use_container_width=False):
        st.markdown(f'<div style="font-size:12px;font-weight:700;color:#1F3B4D;margin-bottom:8px">Notifications</div>', unsafe_allow_html=True)
        _notifs = auth.get_notifications(cu["id"])
        if not _notifs:
            st.markdown('<div style="font-size:11px;color:#94A3B8">No notifications yet.</div>', unsafe_allow_html=True)
        else:
            for _n in _notifs[:8]:
                _nc = {"info": "#3B82F6", "success": "#10B981", "warning": "#F59E0B", "error": "#EF4444"}.get(_n["type"], "#64748B")
                _fw = "600" if not _n["is_read"] else "400"
                st.markdown(
                    f'<div style="padding:6px 0;border-bottom:1px solid #F1F5F9">'
                    f'<div style="font-size:11px;font-weight:{_fw};color:#111827">{esc(_n["title"])}</div>'
                    f'<div style="font-size:10px;color:#64748B">{esc(_n["body"][:60])}</div>'
                    f'<div style="font-size:9px;color:#94A3B8">{format_relative_time(_n["created_at"])}</div>'
                    f'</div>',
                    unsafe_allow_html=True
                )
            if st.button("Mark all read", key="mark_all_notif_read", use_container_width=True):
                auth.mark_all_notifications_read(cu["id"])
                st.rerun()
        st.divider()
        if st.button("Logout", key="nav_logout_title", use_container_width=True):
            st.session_state.current_user = None
            st.rerun()

# Spacer: 20px above nav row
st.markdown('<div style="height:20px"></div>', unsafe_allow_html=True)

# ── TOAST ─────────────────────────────────────────────────────────────────────
if st.session_state.toast:
    t = st.session_state.toast
    _toast_cfg = {
        "success": ("#064E3B","#10B981","✓"),
        "error":   ("#7F1D1D","#EF4444","✗"),
        "info":    ("#244E51","#5FA9AB","ℹ"),
    }
    bg, border, icon = _toast_cfg.get(t.get("type","success"), ("#064E3B","#10B981","✓"))
    st.markdown(
        f'<div class="toast-anim" style="background:{bg};border:1px solid {border};border-radius:12px;'
        f'padding:12px 20px;color:#fff;font-size:13px;font-weight:600;margin-bottom:12px;'
        f'display:flex;align-items:center;gap:10px;box-shadow:0 6px 20px rgba(0,0,0,.28)">'
        f'<span style="font-size:17px;flex-shrink:0">{icon}</span>'
        f'<span>{html.escape(t["msg"])}</span>'
        f'</div>',
        unsafe_allow_html=True)
    st.session_state.toast = None

# ── BACKGROUND EMAIL ERRORS ───────────────────────────────────────────────────
for _email_err in email_utils.pop_email_errors():
    st.markdown(
        f'<div class="toast-anim" style="background:#7F1D1D;border:1px solid #EF4444;border-radius:12px;'
        f'padding:12px 20px;color:#fff;font-size:13px;font-weight:600;margin-bottom:12px;'
        f'display:flex;align-items:center;gap:10px;box-shadow:0 6px 20px rgba(0,0,0,.28)">'
        f'<span style="font-size:17px;flex-shrink:0">✗</span>'
        f'<span>Email failed: {html.escape(_email_err)}</span>'
        f'</div>',
        unsafe_allow_html=True)

# ── TOP BAR: TABS + ACTIONS ───────────────────────────────────────────────────
if role == "employee":
    _tab_defs = [("tasks", "My Tasks"), ("projects", "Projects")]
elif role == "sales":
    _tab_defs = [("dashboard", "Dashboard"), ("presales", "Presales/POC")]
elif role in ("lead", "manager"):
    _tab_defs = [("dashboard", "Dashboard"), ("projects", "Projects"),
                 ("presales", "Presales/POC"), ("license", "License"),
                 ("agent", "AI Agent"), ("tasks", "Tasks")]
else:
    _tab_defs = [("dashboard", "Dashboard"), ("projects", "Projects"),
                 ("presales", "Presales/POC"), ("license", "License"),
                 ("agent", "AI Agent"), ("users", "Users"), ("tasks", "Tasks")]

if st.session_state.active_tab not in [t[0] for t in _tab_defs]:
    st.session_state.active_tab = _tab_defs[0][0]

_n = len(_tab_defs)
nav_c = st.columns([1] * _n + [1])  # tabs + Refresh (Add moved into Projects tab)

# Inject a hidden marker into the first nav column so CSS can scope to this row only
nav_c[0].markdown('<span class="q-nav-bar" style="display:none"></span>', unsafe_allow_html=True)

# ── Button colour CSS (scoped to nav row via .q-nav-bar marker) ────────────────
st.markdown("""
<style>
/* Kill the element-container gap Streamlit wraps around the nav row */
[class*="element-container"]:has(> div[data-testid="stHorizontalBlock"]:has(.q-nav-bar)) {
    margin:0 !important; padding:0 !important;
    margin-top:18px !important;
}
/* Nav row: equal-width columns, pushed down with top padding */
div[data-testid="stHorizontalBlock"]:has(.q-nav-bar) {
    align-items:flex-end !important; gap:4px !important;
    padding:20px 8px 6px 8px !important; margin:0 !important;
    border-bottom:2px solid #DFE3E7;
    background:linear-gradient(180deg,#f8fafc 0%,#fff 100%);
    border-radius:8px 8px 0 0;
}
div[data-testid="stHorizontalBlock"]:has(.q-nav-bar) > div {
    display:flex !important; align-items:flex-end !important;
    padding:0 !important; margin:0 !important;
    height:36px !important; overflow:visible !important;
}
/* Collapse the marker-span wrapper so it takes zero space in column 0 */
div[data-testid="stHorizontalBlock"]:has(.q-nav-bar) [class*="element-container"]:has(.q-nav-bar) {
    display:none !important;
}
/* Lock every wrapper level to 36px so primary/secondary renders identically */
div[data-testid="stHorizontalBlock"]:has(.q-nav-bar) > div > div,
div[data-testid="stHorizontalBlock"]:has(.q-nav-bar) [data-testid="stVerticalBlock"],
div[data-testid="stHorizontalBlock"]:has(.q-nav-bar) [class*="element-container"],
div[data-testid="stHorizontalBlock"]:has(.q-nav-bar) .stButton {
    padding:0 !important; margin:0 !important; width:100% !important;
    height:36px !important; overflow:hidden !important;
    display:flex !important; align-items:flex-end !important;
}
/* Nav buttons: tab-style, active tab raised slightly */
div[data-testid="stHorizontalBlock"]:has(.q-nav-bar) .stButton button {
    font-size:11px !important; font-weight:700 !important;
    padding:0 10px !important;
    height:34px !important; min-height:0 !important; max-height:36px !important;
    line-height:34px !important; border-radius:6px 6px 0 0 !important; letter-spacing:.3px !important;
    width:100% !important;
    background:#DC2626 !important; color:#fff !important;
    border:1px solid #DC2626 !important; border-bottom:none !important;
    display:flex !important; align-items:center !important; justify-content:center !important;
    white-space:nowrap !important; overflow:hidden !important;
    box-sizing:border-box !important; vertical-align:bottom !important;
    box-shadow:0 -2px 6px rgba(220,38,38,.15) !important;
    transition:background .15s,box-shadow .15s !important;
}
div[data-testid="stHorizontalBlock"]:has(.q-nav-bar) .stButton button[kind="secondary"] {
    background:#fee2e2 !important; color:#DC2626 !important;
    border-color:#fca5a5 !important;
    height:30px !important; line-height:30px !important;
    box-shadow:none !important;
}
div[data-testid="stHorizontalBlock"]:has(.q-nav-bar) .stButton button:hover {
    background:#B91C1C !important; border-color:#B91C1C !important; color:#fff !important;
}
</style>
""", unsafe_allow_html=True)

for _i, (_tid, _tlabel) in enumerate(_tab_defs):
    _active = st.session_state.active_tab == _tid
    if nav_c[_i].button(f"{_tlabel}", key=f"tab_{_tid}",
                        type="primary" if _active else "secondary",
                        use_container_width=True):
        st.session_state.active_tab = _tid
        st.rerun()

if role in ("admin", "lead", "manager"):
    if nav_c[_n].button("Refresh", use_container_width=True, key="nav_sync_admin"):
        st.session_state.projects = load_projects()
        ids = pd.to_numeric(st.session_state.projects.get("id", pd.Series([])), errors="coerce").dropna()
        st.session_state.next_id = int(ids.max()) + 1 if not ids.empty else 1
        st.session_state.toast = {"msg": "Synced!", "type": "success"}
        st.rerun()
else:
    if nav_c[_n].button("Refresh", use_container_width=True, key="nav_sync_emp"):
        st.session_state.projects = load_projects()
        st.session_state.toast = {"msg": "Synced!", "type": "success"}
        st.rerun()

# ── Self-heal recovery banner (shown only when healing was needed this render) ─
if _healed and st.session_state.get("_show_heal_banner", False):
    _hb_c1, _hb_c2 = st.columns([8, 1])
    _hb_c1.markdown(
        '<div style="background:#FFF7ED;border:1px solid #FED7AA;border-radius:8px;'
        'padding:6px 14px;font-size:11px;color:#92400E">'
        'Self-healing agent repaired a session issue. All data is intact.</div>',
        unsafe_allow_html=True
    )
    if _hb_c2.button("Dismiss", key="_heal_banner_dismiss", use_container_width=True):
        st.session_state._show_heal_banner = False
        st.rerun()
else:
    st.session_state._show_heal_banner = _healed

st.markdown('<hr style="margin:2px 0;border:none;border-top:1px solid #E2E8F0">', unsafe_allow_html=True)
df = st.session_state.projects   # re-bind after possible sync
_HDR_STYLE = 'font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;letter-spacing:.5px;padding:8px 4px;border-bottom:2px solid #DFE3E7;white-space:nowrap;background:#F8FAFC'

# ── File upload dialog (project files → OneDrive) ────────────────────────────
if st.session_state.get("file_panel_proj"):
    _file_upload_dialog()

# ══════════════════════════════════════════════════════════════════════════════
# MODAL: ADD / EDIT
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.show_modal is not None and role in ("admin", "lead", "manager"):
    mode     = "add" if st.session_state.show_modal == "add" else "edit"
    edit_row = {} if mode == "add" else st.session_state.show_modal.get("edit", {})

    # Build sorted unique employee list from current data
    all_employees = sorted(set(
        n.strip()
        for raw in st.session_state.projects.get("employee", pd.Series(dtype=str)).dropna()
        for n in str(raw).replace("&", ",").split(",")
        if n.strip()
    ))
    # Include leads in the employee pool for lead selection
    if "lead" in st.session_state.projects.columns:
        all_employees = sorted(set(all_employees) | set(
            str(l).strip()
            for l in st.session_state.projects["lead"].dropna()
            if str(l).strip()
        ))
    # Build sorted unique client list from current data
    all_clients = sorted(set(
        str(c).strip()
        for c in st.session_state.projects.get("client", pd.Series(dtype=str)).dropna()
        if str(c).strip()
    ))
    EMP_NEW    = "── Type new name ──"
    CLIENT_NEW = "── Type new client ──"
    client_options = all_clients + [CLIENT_NEW]

    title = "Add New Project" if mode == "add" else "Edit Project"
    st.markdown(f"### {title}")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        name = c1.text_input("Project Name *", value=edit_row.get("name",""))

        # Client: searchable selectbox + optional free-text override
        current_client = edit_row.get("client","")
        client_idx     = client_options.index(current_client) if current_client in client_options else len(client_options) - 1
        client_select  = c2.selectbox(
            "Client * (search or select)",
            options=client_options,
            index=client_idx,
            help="Start typing to search existing clients. Choose the last option to enter a new client."
        )
        if client_select == CLIENT_NEW:
            client = c2.text_input("Enter new client name *", value="", placeholder="e.g. Acme Corp")
        else:
            client = client_select

        # Lead: searchable selectbox (single person — project lead)
        lead_options_full = [""] + all_employees + [EMP_NEW]
        current_lead = edit_row.get("lead", "")
        lead_idx = lead_options_full.index(current_lead) if current_lead in lead_options_full else 0
        lead_select = c1.selectbox(
            "Lead (search or select)",
            options=lead_options_full,
            index=lead_idx,
            help="Select the project lead. Start typing to search existing team members."
        )
        if lead_select == EMP_NEW:
            lead = c1.text_input("Enter new lead name", value="", placeholder="e.g. Jane Smith")
        else:
            lead = lead_select

        idx    = ALL_STATUSES.index(edit_row["status"]) if edit_row.get("status") in ALL_STATUSES else 0
        status = c2.selectbox("Status", ALL_STATUSES, index=idx)

        _PROJ_TYPES = [""] + [t["name"] for t in load_project_types()]
        _pt_val = (edit_row.get("proj_type", "")
                   or (st.session_state.get("proj_add_type", "") if mode == "add" else ""))
        _pt_idx = _PROJ_TYPES.index(_pt_val) if _pt_val in _PROJ_TYPES else 0
        proj_type = c1.selectbox("Type", _PROJ_TYPES, index=_pt_idx,
                                 format_func=lambda x: "— Select type —" if x == "" else x)

        # Employees: multi-select — one or more team members assigned to the project
        current_emp_raw  = str(edit_row.get("employee",""))
        current_emp_list = [n.strip() for n in current_emp_raw.replace("&", ",").split(",") if n.strip()]
        valid_emp_defaults = [e for e in current_emp_list if e in all_employees]
        selected_emps = st.multiselect(
            "Employees * (select one or more)",
            options=all_employees,
            default=valid_emp_defaults,
            help="Search and select all team members assigned to this project."
        )
        new_emp_name = st.text_input(
            "Add new employee name (optional)",
            value="",
            placeholder="e.g. John Doe — leave blank if not needed"
        )
        if new_emp_name.strip():
            emp = ", ".join(selected_emps + [new_emp_name.strip()])
        else:
            emp = ", ".join(selected_emps)

        _dc1, _dc2, _dc3 = st.columns(3)
        _s_default = _parse_dmy(edit_row.get("start", ""))
        _start_dt = _dc1.date_input("Start Date", value=_s_default, key="modal_start", format="DD/MM/YYYY")
        start = _start_dt.strftime("%d/%m/%Y") if _start_dt else ""

        _e_raw = edit_row.get("end", "").strip()
        _is_ongoing = not bool(_e_raw)
        _ongoing = _dc2.checkbox("Ongoing (no end date)", value=_is_ongoing, key="modal_ongoing")
        if _ongoing:
            end = ""
        else:
            _e_default = _parse_dmy(_e_raw) or date.today()
            _end_dt = _dc2.date_input("End Date", value=_e_default, key="modal_end", format="DD/MM/YYYY")
            end = _end_dt.strftime("%d/%m/%Y") if _end_dt else ""


        _d_raw = edit_row.get("due_date", "").strip()
        _due_dt_default = _parse_dmy(_d_raw) if _d_raw else None
        _due_dt = _dc3.date_input("Due Date (optional)", value=_due_dt_default, key="modal_due", format="DD/MM/YYYY")
        due_date = _due_dt.strftime("%d/%m/%Y") if _due_dt else ""

        # ── Checkpoint phase dates ──────────────────────────────────────────
        st.markdown("**Checkpoint Dates** *(optional — enter start & end for each phase)*")
        _ckpt_form_phases = [
            ("PDD / SDD",   "pdd_sdd"),
            ("Development", "development"),
            ("UAT",         "uat"),
            ("Deployment",  "deployment"),
        ]
        _ckpt_form_vals = {}
        for _cfl, _cfk in _ckpt_form_phases:
            _cfa, _cfb = st.columns(2)
            _cf_rs = str(edit_row.get("ckpt_" + _cfk + "_start", "")).strip()
            _cf_re = str(edit_row.get("ckpt_" + _cfk + "_end",   "")).strip()
            _cf_ds = _parse_dmy(_cf_rs) if _cf_rs and _cf_rs != "nan" else None
            _cf_de = _parse_dmy(_cf_re) if _cf_re and _cf_re != "nan" else None
            _cf_sd = _cfa.date_input(
                _cfl + " Start", value=_cf_ds,
                key="modal_ckpt_" + _cfk + "_s", format="DD/MM/YYYY"
            )
            _cf_ed = _cfb.date_input(
                _cfl + " End", value=_cf_de,
                key="modal_ckpt_" + _cfk + "_e", format="DD/MM/YYYY"
            )
            _ckpt_form_vals["ckpt_" + _cfk + "_start"] = _cf_sd.strftime("%d/%m/%Y") if _cf_sd else ""
            _ckpt_form_vals["ckpt_" + _cfk + "_end"]   = _cf_ed.strftime("%d/%m/%Y") if _cf_ed else ""

        po     = c1.text_input("PO Number",           value=edit_row.get("po",""))
        desc   = c2.text_input("Description",         value=edit_row.get("desc",""))
        _is_active_raw = str(edit_row.get("is_active", "True")).strip().lower()
        is_active = c1.checkbox("Active", value=(_is_active_raw not in ["false","0","no"]))

        st.markdown("**ROI Calculator** *(optional)*")
        r1, r2, r3 = st.columns(3)
        manual_hrs  = r1.text_input("Manual Hrs",  value=edit_row.get("manual_hrs",""))
        auto_hrs    = r2.text_input("Auto Hrs",    value=edit_row.get("auto_hrs",""))
        cost_per_hr = r3.text_input("Cost/Hr (₹)", value=edit_row.get("cost_per_hr",""))

        roi = compute_roi(manual_hrs, auto_hrs, cost_per_hr)
        if roi:
            st.success(f"ROI: **{roi['pct']}%** | Hrs Saved: **{roi['saved']}** | Cost Saved: **₹{roi['cost']:,.0f}**")

        # ── File Upload (OneDrive) — edit mode only ───────────────────────────
        st.markdown("**Project Files** *(stored in OneDrive · Qualesce)*")
        if mode == "edit":
            proj_id   = edit_row.get("id", "")
            proj_name = edit_row.get("name", "")
            if proj_id and proj_name:
                existing_files = get_project_files(proj_name)

                uploaded_files = st.file_uploader(
                    "Upload files for this project",
                    accept_multiple_files=True,
                    key=f"file_upload_{proj_id}",
                    help=f"Saved to OneDrive · Qualesce / Qualesce Dashboard / {proj_name}",
                )
                if uploaded_files:
                    for uf in uploaded_files:
                        saved_name = save_project_file(proj_name, uf)
                        st.success(f"Uploaded: {saved_name}")
                    existing_files = get_project_files(proj_name)

                if existing_files:
                    for finfo in existing_files:
                        fc1, fc2, fc3, fc4 = st.columns([4, 2, 2, 1])
                        fc1.markdown(f"📎 **{finfo['name']}**")
                        fc2.caption(fmt_file_size(finfo["size"]))
                        with open(finfo["path"], "rb") as fh:
                            fc3.download_button(
                                label="⬇ Download",
                                data=fh.read(),
                                file_name=finfo["name"],
                                key=f"dl_{proj_id}_{finfo['name']}",
                            )
                        if fc4.button("🗑", key=f"del_file_{proj_id}_{finfo['name']}", help="Delete"):
                            delete_project_file(proj_name, finfo["name"])
                            st.rerun()
                else:
                    st.caption("No files uploaded yet for this project.")
        else:
            st.caption("Save the project first, then use **Edit** to upload files to OneDrive.")

        s1, s2 = st.columns(2)
        save_clicked   = s1.button("Save",   type="primary", use_container_width=True, key="modal_save")
        cancel_clicked = s2.button("Cancel",  use_container_width=True, key="modal_cancel")

        if cancel_clicked:
            st.session_state.show_modal = None
            st.session_state.pop("proj_add_type", None)
            st.rerun()

        if save_clicked:
            errors = []
            if not name or len(name.strip()) < 3:  errors.append("Project name must be at least 3 characters.")
            if not client.strip():                  errors.append("Client is required.")
            if not emp.strip():                     errors.append("Employee is required.")
            if errors:
                for e in errors: st.error(e)
            else:
                if mode == "add":
                    new_row = {
                        "id": st.session_state.next_id,
                        "name": name.strip(), "client": client.strip(),
                        "lead": lead.strip(), "employee": emp.strip(),
                        "status": status, "proj_type": proj_type,
                        "start": start, "end": end, "due_date": due_date, "po": po, "desc": desc.strip(),
                        "manual_hrs": manual_hrs, "auto_hrs": auto_hrs, "cost_per_hr": cost_per_hr,
                        "hours_saved": str(roi["saved"]) if roi else "",
                        "cost_saved":  str(roi["cost"])  if roi else "",
                        "roi_pct":     str(roi["pct"])   if roi else "",
                        "is_new": True,
                        "is_active": is_active,
                        **_ckpt_form_vals,
                    }
                    st.session_state.projects = pd.concat(
                        [st.session_state.projects, pd.DataFrame([new_row])], ignore_index=True)
                    st.session_state.next_id += 1
                    roi_line = f" | ROI {roi['pct']}%" if roi else ""
                    st.session_state.messages.append({"role":"user","content":
                        f"New project added: {name} | {client} | {emp} | {status}{roi_line}. Confirm and give a brief health insight."})
                    st.session_state.toast = {"msg": f'"{name}" added!', "type": "success"}
                else:
                    eid = str(edit_row.get("id",""))
                    records = []
                    for r in st.session_state.projects.to_dict("records"):
                        if str(r.get("id","")) == eid:
                            r.update({"name":name.strip(),"client":client.strip(),
                                      "lead":lead.strip(),"employee":emp.strip(),
                                      "status":status,"proj_type":proj_type,
                                      "start":start,"end":end,"due_date":due_date,"po":po,"desc":desc.strip(),
                                      "manual_hrs":manual_hrs,"auto_hrs":auto_hrs,"cost_per_hr":cost_per_hr,
                                      "hours_saved":str(roi["saved"]) if roi else r.get("hours_saved",""),
                                      "cost_saved": str(roi["cost"])  if roi else r.get("cost_saved",""),
                                      "roi_pct":    str(roi["pct"])   if roi else r.get("roi_pct",""),
                                      "is_active":  is_active,
                                      **_ckpt_form_vals})
                        records.append(r)
                    st.session_state.projects = pd.DataFrame(records)
                    st.session_state.toast = {"msg": f'"{name}" updated!', "type": "success"}

                save_projects_async(st.session_state.projects)
                _audit_action = "CREATE" if mode == "add" else "UPDATE"
                _audit_desc = f'{"Added" if mode=="add" else "Edited"} project "{name.strip()}"'
                auth.log_audit(cu["id"], cu["name"], _audit_action, "projects",
                               str(st.session_state.get("next_id", "")), _audit_desc)
                st.session_state.show_modal = None
                st.session_state.pop("proj_add_type", None)
                st.rerun()

    st.markdown("---")

# ── CONFIRM DELETE ────────────────────────────────────────────────────────────
if st.session_state.confirm_delete and role in ("admin", "lead", "manager"):
    cd = st.session_state.confirm_delete
    st.warning(f"Delete \"{cd['name']}\"? This cannot be undone.")
    da, db, _ = st.columns([1,1,4])
    if da.button("Yes, Delete", type="primary", use_container_width=True, key="yes_del"):
        # Delete by id (not name) to avoid deleting two projects with the same name
        st.session_state.projects = st.session_state.projects[
            st.session_state.projects["id"].astype(str) != str(cd["id"])].reset_index(drop=True)
        save_projects_async(st.session_state.projects)
        auth.log_audit(cu["id"], cu["name"], "DELETE", "projects",
                       str(cd.get("id","")), f'Deleted project "{cd["name"]}"')
        st.session_state.messages.append({"role":"assistant",
            "content": f'"{cd["name"]}" removed. Dashboard updated.'})
        st.session_state.toast = {"msg": f'"{cd["name"]}" deleted.', "type": "info"}
        st.session_state.confirm_delete = None
        st.rerun()
    if db.button("Cancel", use_container_width=True, key="no_del"):
        st.session_state.confirm_delete = None
        st.rerun()
    st.markdown("---")

df = st.session_state.projects

# ── Jinja2 chat templates ─────────────────────────────────────────────────────
_TMPL_USER_MSG = Template("""
<div class="chat-row user-row">
  <div class="chat-avatar avatar-user">U</div>
  <div class="chat-user">{{ content }}</div>
</div>
""")

_TMPL_BOT_MSG = Template("""
<div class="chat-row">
  <div class="chat-avatar avatar-bot">Q</div>
  <div class="chat-bot">{{ content }}</div>
</div>
""")

_TMPL_TYPING = Template("""
<div class="chat-row">
  <div class="chat-avatar avatar-bot">Q</div>
  <div class="typing-indicator">
    <div class="typing-dots"><span></span><span></span><span></span></div>
  </div>
</div>
""")

# ══════════════════════════════════════════════════════════════════════════════
# TAB: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.active_tab == "dashboard" and role not in ("employee",):
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px;font-family:Manrope,sans-serif;letter-spacing:-.3px">Project Portfolio Dashboard</h2>', unsafe_allow_html=True)


    # Pre-compute df for KPI cards and all panels
    _dash_df_pre = df.copy()
    stats = get_stats(_dash_df_pre)

    dash_df    = _dash_df_pre
    dash_stats = stats

    _dt_rpa, _dt_ws = st.tabs(["🔧 RPA", "⚙️ Worksoft"])
    with _dt_rpa:
        st.markdown(
            '<style>'
            '@keyframes fadeInUp{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:translateY(0)}}'
            '@keyframes slideInLeft{from{opacity:0;transform:translateX(-18px)}to{opacity:1;transform:translateX(0)}}'
            '@keyframes scaleIn{from{opacity:0;transform:scale(0.96)}to{opacity:1;transform:scale(1)}}'
            '[data-testid="stPlotlyChart"]{animation:fadeInUp 0.65s ease both}'
            '[data-testid="stMetric"]{animation:fadeInUp 0.5s ease both}'
            '</style>',
            unsafe_allow_html=True,
        )

        # ══════════════════════════════════════════════════════════════════════════
        # SALES & MARKETING INTELLIGENCE
        # ══════════════════════════════════════════════════════════════════════════
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">'
            '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Sales &amp; Marketing</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#162C3B;color:#5FA9AB;border:1px solid #3F8E9140">Qualesce Company Profile</span>'
            '</div>',
            unsafe_allow_html=True
        )

        # ── COMPANY PROFILE CARD ─────────────────────────────────────────────────
        _cp_left, _cp_right = st.columns([1.2, 1])

        with _cp_left:
            st.markdown(
                '<div style="background:linear-gradient(135deg,#162C3B 0%,#1a3549 50%,#162C3B 100%);'
                'border-radius:16px;padding:28px 32px;height:100%;min-height:260px;position:relative;overflow:hidden">'

                # Subtle background circle decoration
                '<div style="position:absolute;top:-30px;right:-30px;width:160px;height:160px;'
                'border-radius:50%;background:#3F8E9115;pointer-events:none"></div>'
                '<div style="position:absolute;bottom:-20px;left:-20px;width:100px;height:100px;'
                'border-radius:50%;background:#3F8E9110;pointer-events:none"></div>'

                # Logo / Brand mark
                + (f'<div style="margin-bottom:20px">'
                   f'<img src="data:{_LOGO_MIME};base64,{_LOGO_B64}" '
                   f'style="height:64px;width:auto;object-fit:contain" alt="Qualesce">'
                   f'</div>'
                   if _LOGO_B64 else
                   '<div style="display:flex;align-items:center;gap:14px;margin-bottom:20px">'
                   '<div style="width:52px;height:52px;border-radius:14px;'
                   'background:linear-gradient(135deg,#3F8E91,#2F6F72);'
                   'display:flex;align-items:center;justify-content:center;'
                   'font-size:26px;font-weight:900;color:#fff;font-family:Manrope,sans-serif;'
                   'box-shadow:0 4px 12px #3F8E9140">Q</div>'
                   '<div>'
                   '<div style="font-size:22px;font-weight:900;color:#fff;font-family:Manrope,sans-serif;'
                   'letter-spacing:-0.3px;line-height:1.1">Qualesce</div>'
                   '<div style="font-size:10px;color:#5FA9AB;font-weight:600;letter-spacing:1.2px;'
                   'text-transform:uppercase;margin-top:2px">Intelligent Automation</div>'
                   '</div>'
                   '</div>') +

                # Tagline
                '<div style="font-size:16px;font-weight:700;color:#E2E8F0;line-height:1.4;margin-bottom:14px;'
                'font-family:Manrope,sans-serif">'
                'Accelerating Enterprise Growth<br>with RPA &amp; AI Automation'
                '</div>'

                # Description
                '<div style="font-size:11px;color:#94A3B8;line-height:1.6;margin-bottom:20px">'
                'Qualesce delivers end-to-end automation solutions — from Robotic Process Automation '
                'to AI-powered agents — helping enterprises eliminate manual effort, reduce costs, '
                'and scale operations with intelligent technology.'
                '</div>'

                # Tags
                '<div style="display:flex;gap:8px;flex-wrap:wrap">'
                '<span style="font-size:9px;font-weight:700;padding:3px 10px;border-radius:20px;'
                'background:#3F8E9125;color:#5FA9AB;border:1px solid #3F8E9140">RPA</span>'
                '<span style="font-size:9px;font-weight:700;padding:3px 10px;border-radius:20px;'
                'background:#8B5CF615;color:#A78BFA;border:1px solid #8B5CF630">AI Agents</span>'
                '<span style="font-size:9px;font-weight:700;padding:3px 10px;border-radius:20px;'
                'background:#F59E0B15;color:#FCD34D;border:1px solid #F59E0B30">Enterprise</span>'
                '<span style="font-size:9px;font-weight:700;padding:3px 10px;border-radius:20px;'
                'background:#10B98115;color:#6EE7B7;border:1px solid #10B98130">Automation</span>'
                '<span style="font-size:9px;font-weight:700;padding:3px 10px;border-radius:20px;'
                'background:#3B82F615;color:#93C5FD;border:1px solid #3B82F630">www.qualesce.com</span>'
                '</div>'

                '</div>',
                unsafe_allow_html=True
            )

        with _cp_right:
            # RPA Service tile
            st.markdown(
                '<div style="background:#fff;border:1.5px solid #E2E8F0;border-radius:14px;'
                'padding:20px 22px;margin-bottom:12px;border-left:4px solid #3F8E91">'
                '<div style="display:flex;align-items:center;gap:12px;margin-bottom:10px">'
                '<div style="width:36px;height:36px;border-radius:10px;background:#3F8E9115;'
                'display:flex;align-items:center;justify-content:center;font-size:18px">🤖</div>'
                '<div>'
                '<div style="font-size:13px;font-weight:800;color:#1F3B4D">Robotic Process Automation</div>'
                '<div style="font-size:10px;color:#3F8E91;font-weight:600">RPA · Bot Development · Workflow Automation</div>'
                '</div>'
                '</div>'
                '<div style="font-size:11px;color:#64748B;line-height:1.6">'
                'Deploy software robots that mimic human actions to automate high-volume, '
                'rule-based tasks — reducing errors, accelerating processing speed, and freeing '
                'your team for higher-value work.'
                '</div>'
                '<div style="display:flex;gap:8px;margin-top:12px">'
                '<span style="font-size:10px;color:#3F8E91;font-weight:700">✓ Zero manual errors</span>'
                '<span style="font-size:10px;color:#3F8E91;font-weight:700">✓ 24/7 operation</span>'
                '<span style="font-size:10px;color:#3F8E91;font-weight:700">✓ Fast ROI</span>'
                '</div>'
                '</div>',
                unsafe_allow_html=True
            )
            # AI Agents tile
            st.markdown(
                '<div style="background:#fff;border:1.5px solid #E2E8F0;border-radius:14px;'
                'padding:20px 22px;border-left:4px solid #8B5CF6">'
                '<div style="display:flex;align-items:center;gap:12px;margin-bottom:10px">'
                '<div style="width:36px;height:36px;border-radius:10px;background:#8B5CF615;'
                'display:flex;align-items:center;justify-content:center;font-size:18px">🧠</div>'
                '<div>'
                '<div style="font-size:13px;font-weight:800;color:#1F3B4D">AI Agent Solutions</div>'
                '<div style="font-size:10px;color:#8B5CF6;font-weight:600">Agentic AI · LLM Integration · Smart Automation</div>'
                '</div>'
                '</div>'
                '<div style="font-size:11px;color:#64748B;line-height:1.6">'
                'Build intelligent AI agents that reason, plan, and act autonomously — handling '
                'complex, unstructured workflows that traditional automation cannot. Powered by '
                'large language models and real-time data.'
                '</div>'
                '<div style="display:flex;gap:8px;margin-top:12px">'
                '<span style="font-size:10px;color:#8B5CF6;font-weight:700">✓ Understands context</span>'
                '<span style="font-size:10px;color:#8B5CF6;font-weight:700">✓ Self-improving</span>'
                '<span style="font-size:10px;color:#8B5CF6;font-weight:700">✓ Scales instantly</span>'
                '</div>'
                '</div>',
                unsafe_allow_html=True
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── LIVE IMPACT STATS (marketing slide numbers) ───────────────────────────
        st.markdown(
            '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
            'letter-spacing:.8px;margin-bottom:12px">Live Business Impact — Delivered to Clients</div>',
            unsafe_allow_html=True
        )

        # ── 1. MARKETING IMPACT BANNER ────────────────────────────────────────────
        _all_proj = st.session_state.projects.copy()
        _mi_hrs   = pd.to_numeric(_all_proj.get("hours_saved",  pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
        _mi_cost  = pd.to_numeric(_all_proj.get("cost_saved",   pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
        _mi_done  = int((_all_proj["status"] == "Completed").sum()) if "status" in _all_proj.columns else 0
        _mi_live  = int(_all_proj["status"].isin(["In Progress","UAT","R&M"]).sum()) if "status" in _all_proj.columns else 0
        _mi_clients = _all_proj["client"].dropna().nunique() if "client" in _all_proj.columns else 0
        _mi_cost_disp = (f"₹{_mi_cost/10_000_000:.2f} Cr" if _mi_cost >= 10_000_000
                         else f"₹{int(_mi_cost/1000):,}K" if _mi_cost >= 1000
                         else "—")
        _mi_hrs_disp  = f"{int(_mi_hrs):,}" if _mi_hrs > 0 else "—"

        st.markdown(
            f'<div style="background:linear-gradient(135deg,#162C3B 0%,#1F3B4D 100%);'
            f'border-radius:14px;padding:20px 28px;margin-bottom:16px;display:flex;'
            f'align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px">'
            f'<div>'
            f'<div style="font-size:11px;font-weight:700;color:#5FA9AB;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px">Total Impact Delivered to Clients</div>'
            f'<div style="font-size:10px;color:#94A3B8">Automation ROI generated across all Qualesce projects</div>'
            f'</div>'
            f'<div style="display:flex;gap:32px;flex-wrap:wrap">'
            f'<div style="text-align:center">'
            f'<div style="font-size:28px;font-weight:900;color:#10B981;line-height:1">{_mi_hrs_disp}</div>'
            f'<div style="font-size:10px;color:#94A3B8;margin-top:2px">Hours Saved</div>'
            f'</div>'
            f'<div style="text-align:center">'
            f'<div style="font-size:28px;font-weight:900;color:#F59E0B;line-height:1">{_mi_cost_disp}</div>'
            f'<div style="font-size:10px;color:#94A3B8;margin-top:2px">Cost Saved</div>'
            f'</div>'
            f'<div style="text-align:center">'
            f'<div style="font-size:28px;font-weight:900;color:#3B82F6;line-height:1">{_mi_done}</div>'
            f'<div style="font-size:10px;color:#94A3B8;margin-top:2px">Projects Delivered</div>'
            f'</div>'
            f'<div style="text-align:center">'
            f'<div style="font-size:28px;font-weight:900;color:#8B5CF6;line-height:1">{_mi_live}</div>'
            f'<div style="font-size:10px;color:#94A3B8;margin-top:2px">Live Now</div>'
            f'</div>'
            f'<div style="text-align:center">'
            f'<div style="font-size:28px;font-weight:900;color:#EC4899;line-height:1">{_mi_clients}</div>'
            f'<div style="font-size:10px;color:#94A3B8;margin-top:2px">Clients Served</div>'
            f'</div>'
            f'</div>'
            f'</div>',
            unsafe_allow_html=True
        )

        _sm_col1, _sm_col2 = st.columns(2)

        # ── 2. PRESALES CONVERSION PIPELINE ──────────────────────────────────────
        with _sm_col1:
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                    'text-transform:uppercase;letter-spacing:.8px;margin-bottom:12px">'
                    'Presales → Live Conversion Pipeline</div>',
                    unsafe_allow_html=True
                )
                _ap = _all_proj.copy()
                _conv_presales = int(_ap["status"].isin(["Presales"]).sum())                     if "status" in _ap.columns else 0
                _conv_poc      = int(_ap["status"].str.contains("POC", na=False).sum())           if "status" in _ap.columns else 0
                _conv_pdd      = int((_ap["status"] == "PDD").sum())                              if "status" in _ap.columns else 0
                _conv_dev      = int(_ap["status"].isin(["In Progress"]).sum())                   if "status" in _ap.columns else 0
                _conv_uat      = int(_ap["status"].isin(["UAT","R&M"]).sum())                     if "status" in _ap.columns else 0
                _conv_done     = int((_ap["status"] == "Completed").sum())                        if "status" in _ap.columns else 0
                _conv_disc     = int((_ap["status"] == "Discontinued").sum())                     if "status" in _ap.columns else 0
                _total_engaged = _conv_presales + _conv_poc + _conv_pdd + _conv_dev + _conv_uat + _conv_done + _conv_disc
                _live_or_done  = _conv_dev + _conv_uat + _conv_done

                # Conversion rate: how many deals became live projects
                _conv_rate = round((_live_or_done / max(_total_engaged, 1)) * 100)

                _pipeline_stages = [
                    ("Presales",    _conv_presales, "#0EA5E9"),
                    ("POC",         _conv_poc,      "#8B5CF6"),
                    ("PDD / Design",_conv_pdd,      "#F59E0B"),
                    ("Development", _conv_dev,      "#06B6D4"),
                    ("UAT / R&M",   _conv_uat,      "#3B82F6"),
                    ("Completed",   _conv_done,     "#10B981"),
                ]
                _conv_html = '<div style="display:flex;flex-direction:column;gap:8px">'
                for _ps_label, _ps_count, _ps_color in _pipeline_stages:
                    _bar_pct = round((_ps_count / max(_total_engaged, 1)) * 100)
                    _conv_html += (
                        f'<div>'
                        f'<div style="display:flex;justify-content:space-between;margin-bottom:3px">'
                        f'<span style="font-size:11px;font-weight:600;color:#374151">{_ps_label}</span>'
                        f'<span style="font-size:11px;font-weight:700;color:{_ps_color}">{_ps_count}</span>'
                        f'</div>'
                        f'<div style="height:8px;background:#F1F5F9;border-radius:4px;overflow:hidden">'
                        f'<div style="height:100%;width:{_bar_pct}%;background:{_ps_color};border-radius:4px;'
                        f'transition:width .4s ease"></div>'
                        f'</div>'
                        f'</div>'
                    )
                _conv_html += '</div>'
                st.markdown(_conv_html, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown(
                    f'<div style="display:flex;gap:20px;padding:10px 14px;background:#F0FDF4;'
                    f'border-radius:8px;border:1px solid #BBF7D0">'
                    f'<div><span style="font-size:20px;font-weight:900;color:#10B981">{_conv_rate}%</span>'
                    f'<div style="font-size:10px;color:#64748B">Conversion Rate</div></div>'
                    f'<div><span style="font-size:20px;font-weight:900;color:#1F3B4D">{_live_or_done}</span>'
                    f'<div style="font-size:10px;color:#64748B">Became Live Projects</div></div>'
                    f'<div><span style="font-size:20px;font-weight:900;color:#EF4444">{_conv_disc}</span>'
                    f'<div style="font-size:10px;color:#64748B">Discontinued</div></div>'
                    f'</div>',
                    unsafe_allow_html=True
                )

        # ── 3. SERVICE MIX & LICENSE RENEWAL OPPORTUNITIES ───────────────────────
        with _sm_col2:
            # Service mix donut
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                    'text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">'
                    'Service Mix — RPA vs AI Agent vs Presales</div>',
                    unsafe_allow_html=True
                )
                _pt_col = _ap.get("proj_type", pd.Series(dtype=str)).fillna("").str.strip()
                _rpa_n   = int((_pt_col == "RPA").sum())
                _ai_n    = int((_pt_col == "AI Agent").sum())
                _pre_n   = int((_pt_col == "Presales").sum())
                _other_n = int(len(_pt_col) - _rpa_n - _ai_n - _pre_n)
                _sm_labels = ["RPA", "AI Agent", "Presales", "Other"]
                _sm_values = [_rpa_n, _ai_n, _pre_n, _other_n]
                _sm_colors = ["#3F8E91", "#8B5CF6", "#F59E0B", "#94A3B8"]
                _sm_fig = go.Figure(go.Pie(
                    labels=_sm_labels, values=_sm_values,
                    hole=0.6,
                    marker=dict(colors=_sm_colors),
                    textinfo="percent+label",
                    textfont=dict(size=10),
                ))
                _sm_fig.update_layout(
                    margin=dict(t=0, b=0, l=0, r=0), height=180,
                    showlegend=False,
                    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    annotations=[dict(text=f"<b>{sum(_sm_values)}</b><br>Projects", x=0.5, y=0.5,
                                      font=dict(size=12), showarrow=False)]
                )
                st.plotly_chart(_sm_fig, use_container_width=True)

            # License renewal alerts
            st.markdown("<br>", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                    'text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px">'
                    'License Renewal Opportunities — Upsell Alerts</div>',
                    unsafe_allow_html=True
                )
                _sold_lic = auth.get_all_sold_licenses()
                _today_sm = date.today()
                _renewal_rows = []
                for _sl in _sold_lic:
                    _ed = str(_sl.get("end_date","")).strip()
                    if not _ed:
                        continue
                    for _fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                        try:
                            _exp_d = datetime.strptime(_ed, _fmt).date()
                            _days  = (_exp_d - _today_sm).days
                            _renewal_rows.append((_days, _sl, _exp_d))
                            break
                        except ValueError:
                            pass
                _renewal_rows.sort(key=lambda x: x[0])
                _upcoming = [r for r in _renewal_rows if r[0] <= 90]
                if _upcoming:
                    for _days_left, _sl, _exp_d in _upcoming:
                        if _days_left < 0:
                            _rc, _rt, _rbg = "#B23A3A", "Expired", "#FCEAEA"
                        elif _days_left <= 30:
                            _rc, _rt, _rbg = "#DC2626", f"{_days_left}d left", "#FEF2F2"
                        elif _days_left <= 60:
                            _rc, _rt, _rbg = "#B45309", f"{_days_left}d left", "#FFFBEB"
                        else:
                            _rc, _rt, _rbg = "#92400E", f"{_days_left}d left", "#FEF3C7"
                        st.markdown(
                            f'<div style="display:flex;align-items:center;justify-content:space-between;'
                            f'padding:8px 12px;background:{_rbg};border-radius:8px;margin-bottom:6px;'
                            f'border-left:3px solid {_rc}">'
                            f'<div>'
                            f'<div style="font-size:12px;font-weight:700;color:#1F3B4D">{esc(str(_sl["tool_name"]))}</div>'
                            f'<div style="font-size:10px;color:#64748B">{esc(str(_sl["client"]))} · '
                            f'{_sl["no_of_licenses"]} licence(s) · expires {_exp_d.strftime("%d %b %Y")}</div>'
                            f'</div>'
                            f'<span style="font-size:10px;font-weight:700;color:{_rc};'
                            f'background:{"#ffffff80"};padding:3px 10px;border-radius:12px;'
                            f'border:1px solid {_rc}40;white-space:nowrap">{_rt}</span>'
                            f'</div>',
                            unsafe_allow_html=True
                        )
                else:
                    st.markdown(
                        '<div style="text-align:center;padding:18px;color:#94A3B8;font-size:11px">'
                        'No licenses expiring in the next 90 days.<br>'
                        '<span style="font-size:10px">Add sold licenses in the License tab to track renewals.</span>'
                        '</div>',
                        unsafe_allow_html=True
                    )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── RPA BOT EFFICIENCY ────────────────────────────────────────────────────
        _rpa_bm_all = (
            _all_proj[_all_proj["proj_type"].fillna("").str.strip() == "RPA"].copy()
            if "proj_type" in _all_proj.columns else pd.DataFrame(columns=_all_proj.columns)
        )
        for _bmc in ["num_bots", "num_persons", "manual_run_mins", "bot_run_mins"]:
            if _bmc not in _rpa_bm_all.columns:
                _rpa_bm_all[_bmc] = 0
            _rpa_bm_all[_bmc] = pd.to_numeric(_rpa_bm_all[_bmc], errors="coerce").fillna(0)
        _rpa_bm_configured = _rpa_bm_all[_rpa_bm_all["num_bots"] > 0].copy()

        if not _rpa_bm_configured.empty:
            with st.container(border=True):
                # ── Filter row ────────────────────────────────────────────────────
                _bm_hdr_col, _bm_client_col, _bm_proj_col = st.columns([2, 2, 3])
                _bm_hdr_col.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                    'letter-spacing:.8px;margin-top:6px">RPA Bot Efficiency — Time Saved</div>',
                    unsafe_allow_html=True
                )
                _bm_all_clients = sorted(_rpa_bm_configured["client"].fillna("").unique().tolist())
                _bm_sel_clients = _bm_client_col.multiselect(
                    "Filter by Client", options=_bm_all_clients, default=[],
                    key="bm_dash_client_filter", label_visibility="collapsed",
                    placeholder="Filter by client…"
                )
                _bm_after_client = (
                    _rpa_bm_configured[_rpa_bm_configured["client"].isin(_bm_sel_clients)]
                    if _bm_sel_clients else _rpa_bm_configured
                )
                _bm_all_names = _bm_after_client["name"].tolist()
                _bm_selected  = _bm_proj_col.multiselect(
                    "Filter Projects", options=_bm_all_names, default=_bm_all_names,
                    key="bm_dash_filter", label_visibility="collapsed",
                    placeholder="Select projects to display…"
                )
                _bm_chart_src = (
                    _bm_after_client[_bm_after_client["name"].isin(_bm_selected)]
                    if _bm_selected else _bm_after_client
                )

                # ── Date range row ────────────────────────────────────────────────
                _bm_dr1, _bm_dr2, _ = st.columns([2, 2, 3])
                _bm_start = _bm_dr1.date_input("From", value=date.today().replace(day=1),
                                                format="DD/MM/YYYY", key="bm_dash_start")
                _bm_end   = _bm_dr2.date_input("To",   value=date.today(),
                                                format="DD/MM/YYYY", key="bm_dash_end")

                # Load logs for selected date range
                _bm_dash_logs = auth.get_bot_metric_logs(
                    start_date=str(_bm_start), end_date=str(_bm_end)
                )
                _bm_dash_qty: dict = {}
                for _dl in _bm_dash_logs:
                    _dpid = int(_dl.get("project_id", 0))
                    _bm_dash_qty[_dpid] = _bm_dash_qty.get(_dpid, 0) + int(_dl.get("qty", 0) or 0)

                # Build per-project values
                _bm_proj_names, _bm_manual_hrs, _bm_bot_hrs, _bm_saved_hrs = [], [], [], []
                for _, _pr in _bm_chart_src.iterrows():
                    _dpid  = int(float(_pr.get("id", 0) or 0))
                    _qty   = _bm_dash_qty.get(_dpid, 0)
                    _mh    = float(_pr.get("manual_run_mins", 0) or 0) * float(_pr.get("num_persons", 0) or 0) * _qty / 60
                    _bh    = float(_pr.get("bot_run_mins",    0) or 0) * float(_pr.get("num_bots",    0) or 0) * _qty / 60
                    _bm_proj_names.append(str(_pr.get("name", "")))
                    _bm_manual_hrs.append(round(_mh, 2))
                    _bm_bot_hrs.append(round(_bh, 2))
                    _bm_saved_hrs.append(round(max(_mh - _bh, 0), 2))

                _bm_chart_left, _bm_chart_right = st.columns([3, 1])
                with _bm_chart_left:
                    _bm_fig = go.Figure()
                    _bm_fig.add_trace(go.Bar(
                        name="Manual Process", x=_bm_proj_names, y=_bm_manual_hrs,
                        marker_color="#EF4444", opacity=0.85,
                        text=[f"{v:.1f}h" for v in _bm_manual_hrs],
                        textposition="outside", textfont=dict(size=9)
                    ))
                    _bm_fig.add_trace(go.Bar(
                        name="Bot Process", x=_bm_proj_names, y=_bm_bot_hrs,
                        marker_color="#10B981", opacity=0.85,
                        text=[f"{v:.1f}h" for v in _bm_bot_hrs],
                        textposition="outside", textfont=dict(size=9)
                    ))
                    _bm_fig.update_layout(
                        barmode="group", height=260,
                        margin=dict(t=30, b=0, l=0, r=0),
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                                    xanchor="right", x=1, font=dict(size=10)),
                        xaxis=dict(tickfont=dict(size=10)),
                        yaxis=dict(title="Hours", tickfont=dict(size=10)),
                    )
                    st.plotly_chart(_bm_fig, use_container_width=True)
                with _bm_chart_right:
                    _bm_total_s = sum(_bm_saved_hrs)
                    _bm_total_b = int(_bm_chart_src["num_bots"].sum())
                    st.markdown(
                        f'<div style="background:linear-gradient(135deg,#F0FDF4,#DCFCE7);border-radius:12px;'
                        f'padding:16px;border:1px solid #BBF7D0;text-align:center;margin-bottom:8px">'
                        f'<div style="font-size:10px;color:#64748B;font-weight:600;margin-bottom:4px">Total Saved</div>'
                        f'<div style="font-size:30px;font-weight:900;color:#10B981;line-height:1">{_bm_total_s:,.1f}</div>'
                        f'<div style="font-size:10px;color:#94A3B8;margin-top:2px">hours</div>'
                        f'<div style="font-size:20px;font-weight:700;color:#1F3B4D;margin-top:8px">{_bm_total_s * 60:,.0f}</div>'
                        f'<div style="font-size:10px;color:#94A3B8">minutes</div>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    st.markdown(
                        f'<div style="background:linear-gradient(135deg,#EFF6FF,#DBEAFE);border-radius:12px;'
                        f'padding:14px;border:1px solid #BFDBFE;text-align:center">'
                        f'<div style="font-size:10px;color:#64748B;font-weight:600;margin-bottom:4px">Active Bots</div>'
                        f'<div style="font-size:28px;font-weight:900;color:#3B82F6">{_bm_total_b}</div>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
            st.markdown("<br>", unsafe_allow_html=True)

        # ── PROJECT TYPE BREAKDOWN ────────────────────────────────────────────────
        # ══════════════════════════════════════════════════════════════════════════
        # QUALESCE INDIA — COMPANY OVERVIEW
        # ══════════════════════════════════════════════════════════════════════════
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">'
            '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Qualesce India</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#F0FDF4;color:#10B981;border:1px solid #BBF7D0">Company Overview</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#EFF6FF;color:#3B82F6;border:1px solid #BFDBFE">Live Portfolio Data</span>'
            '</div>',
            unsafe_allow_html=True
        )

        # Hero company card
        _co_total_projs = int(_mi_done + _mi_live)
        _co_client_cnt  = int(_mi_clients)
        _co_stats_html  = ""
        for _csv, _csl, _csc in [
            ("RPA + AI",
             "Service Lines",
             "#5FA9AB"),
            (str(_co_client_cnt) if _co_client_cnt else "4+",
             "Enterprise Clients",
             "#A78BFA"),
            (str(_co_total_projs) if _co_total_projs else "46+",
             "Total Projects",
             "#6EE7B7"),
            ("412%",
             "Portfolio ROI",
             "#FCD34D"),
        ]:
            _co_stats_html += (
                f'<div style="background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);'
                f'border-radius:12px;padding:14px;text-align:center">'
                f'<div style="font-size:20px;font-weight:900;color:{_csc};line-height:1">{_csv}</div>'
                f'<div style="font-size:9px;color:#94A3B8;margin-top:4px;font-weight:600;'
                f'text-transform:uppercase;letter-spacing:.6px">{_csl}</div>'
                f'</div>'
            )
        st.markdown(
            '<div style="background:linear-gradient(135deg,#0F172A 0%,#1E293B 60%,#0F2027 100%);'
            'border-radius:18px;padding:28px 32px;margin-bottom:16px;'
            'border:1px solid rgba(95,169,171,0.25);position:relative;overflow:hidden">'
            '<div style="position:absolute;top:-50px;right:-50px;width:220px;height:220px;'
            'border-radius:50%;background:rgba(95,169,171,0.07);pointer-events:none"></div>'
            '<div style="position:absolute;bottom:-30px;left:30px;width:140px;height:140px;'
            'border-radius:50%;background:rgba(139,92,246,0.05);pointer-events:none"></div>'
            '<div style="display:flex;align-items:center;gap:16px;margin-bottom:14px">'
            '<div style="width:54px;height:54px;border-radius:14px;'
            'background:linear-gradient(135deg,#3F8E91,#2F6F72);display:flex;align-items:center;'
            'justify-content:center;font-size:28px;font-weight:900;color:#fff;'
            'font-family:Manrope,sans-serif;box-shadow:0 4px 16px rgba(63,142,145,0.35)">Q</div>'
            '<div>'
            '<div style="font-size:22px;font-weight:900;color:#fff;font-family:Manrope,sans-serif;'
            'letter-spacing:-0.4px;line-height:1.1">Qualesce India</div>'
            '<div style="font-size:10px;color:#5FA9AB;font-weight:600;letter-spacing:1.2px;'
            'text-transform:uppercase;margin-top:3px">Intelligent Automation · RPA · AI Agents</div>'
            '</div>'
            '</div>'
            '<div style="font-size:13px;font-weight:600;color:#CBD5E1;line-height:1.6;margin-bottom:4px">'
            'Enterprise-grade automation through RPA and AI Agents — eliminating manual effort, '
            'accelerating ROI, and scaling operations across India.'
            '</div>'
            '<div style="font-size:10px;color:#64748B;margin-bottom:22px">'
            'End-to-end automation consulting, development &amp; managed services &nbsp;·&nbsp; '
            '<span style="color:#5FA9AB;font-weight:600">www.qualesce.com</span>'
            '</div>'
            f'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px">'
            f'{_co_stats_html}'
            f'</div>'
            '</div>',
            unsafe_allow_html=True
        )

        # Services deep-dive (2 columns)
        _svc_c1, _svc_c2 = st.columns(2)
        _svc_rpa_caps = [
            "Invoice & PO processing automation",
            "ERP / SAP data entry & reconciliation",
            "Compliance & audit report generation",
            "Vendor master & data management",
            "Cross-system data migration & sync",
        ]
        _svc_ai_caps = [
            "Intelligent document understanding (IDP)",
            "Conversational AI & chatbot integration",
            "Predictive analytics & decision support",
            "Multi-step agentic workflow orchestration",
            "LLM-powered data extraction & classification",
        ]
        with _svc_c1:
            _rpa_caps_html = "".join(
                f'<div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:7px">'
                f'<div style="width:6px;height:6px;border-radius:50%;background:#3F8E91;'
                f'flex-shrink:0;margin-top:5px"></div>'
                f'<span style="font-size:11px;color:#475569">{cap}</span></div>'
                for cap in _svc_rpa_caps
            )
            st.markdown(
                '<div style="background:#fff;border:1.5px solid #E2E8F0;border-radius:16px;'
                'padding:24px;border-top:4px solid #3F8E91">'
                '<div style="display:flex;align-items:center;gap:12px;margin-bottom:14px">'
                '<div style="width:44px;height:44px;border-radius:12px;background:#3F8E9115;'
                'display:flex;align-items:center;justify-content:center;font-size:22px">🤖</div>'
                '<div>'
                '<div style="font-size:15px;font-weight:800;color:#1F3B4D">'
                'Robotic Process Automation</div>'
                '<div style="font-size:10px;color:#3F8E91;font-weight:700;letter-spacing:.4px">'
                'RPA · BOT DEV · WORKFLOW AUTOMATION</div>'
                '</div></div>'
                '<div style="font-size:12px;color:#475569;line-height:1.7;margin-bottom:14px">'
                'Software robots that replicate human actions across enterprise systems — automating '
                'high-volume, rule-based tasks with zero errors, 24/7 availability, and measurable ROI.'
                '</div>'
                '<div style="font-size:10px;font-weight:700;color:#1F3B4D;text-transform:uppercase;'
                'letter-spacing:.6px;margin-bottom:8px">Key Capabilities</div>'
                f'<div>{_rpa_caps_html}</div>'
                '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:14px">'
                '<span style="font-size:10px;color:#3F8E91;font-weight:700;padding:3px 10px;'
                'background:#3F8E9110;border-radius:20px;border:1px solid #3F8E9130">✓ Zero errors</span>'
                '<span style="font-size:10px;color:#3F8E91;font-weight:700;padding:3px 10px;'
                'background:#3F8E9110;border-radius:20px;border:1px solid #3F8E9130">✓ 24/7 uptime</span>'
                '<span style="font-size:10px;color:#3F8E91;font-weight:700;padding:3px 10px;'
                'background:#3F8E9110;border-radius:20px;border:1px solid #3F8E9130">✓ Fast ROI</span>'
                '</div></div>',
                unsafe_allow_html=True
            )
        with _svc_c2:
            _ai_caps_html = "".join(
                f'<div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:7px">'
                f'<div style="width:6px;height:6px;border-radius:50%;background:#8B5CF6;'
                f'flex-shrink:0;margin-top:5px"></div>'
                f'<span style="font-size:11px;color:#475569">{cap}</span></div>'
                for cap in _svc_ai_caps
            )
            st.markdown(
                '<div style="background:#fff;border:1.5px solid #E2E8F0;border-radius:16px;'
                'padding:24px;border-top:4px solid #8B5CF6">'
                '<div style="display:flex;align-items:center;gap:12px;margin-bottom:14px">'
                '<div style="width:44px;height:44px;border-radius:12px;background:#8B5CF615;'
                'display:flex;align-items:center;justify-content:center;font-size:22px">🧠</div>'
                '<div>'
                '<div style="font-size:15px;font-weight:800;color:#1F3B4D">AI Agent Solutions</div>'
                '<div style="font-size:10px;color:#8B5CF6;font-weight:700;letter-spacing:.4px">'
                'AGENTIC AI · LLM INTEGRATION · SMART AUTOMATION</div>'
                '</div></div>'
                '<div style="font-size:12px;color:#475569;line-height:1.7;margin-bottom:14px">'
                'Intelligent AI agents that reason, plan, and act autonomously — handling complex, '
                'unstructured workflows powered by large language models and real-time enterprise data.'
                '</div>'
                '<div style="font-size:10px;font-weight:700;color:#1F3B4D;text-transform:uppercase;'
                'letter-spacing:.6px;margin-bottom:8px">Key Capabilities</div>'
                f'<div>{_ai_caps_html}</div>'
                '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:14px">'
                '<span style="font-size:10px;color:#8B5CF6;font-weight:700;padding:3px 10px;'
                'background:#8B5CF615;border-radius:20px;border:1px solid #8B5CF630">'
                '✓ Context-aware</span>'
                '<span style="font-size:10px;color:#8B5CF6;font-weight:700;padding:3px 10px;'
                'background:#8B5CF615;border-radius:20px;border:1px solid #8B5CF630">'
                '✓ Self-improving</span>'
                '<span style="font-size:10px;color:#8B5CF6;font-weight:700;padding:3px 10px;'
                'background:#8B5CF615;border-radius:20px;border:1px solid #8B5CF630">'
                '✓ Scales instantly</span>'
                '</div></div>',
                unsafe_allow_html=True
            )

        # Client portfolio cards (live DB)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
            'letter-spacing:.8px;margin-bottom:12px">Client Portfolio · Enterprise Accounts</div>',
            unsafe_allow_html=True
        )
        _co_df = st.session_state.projects.copy()
        _co_cg: dict = {}
        if "client" in _co_df.columns:
            for _, _crow in _co_df.iterrows():
                _cl = str(_crow.get("client","")).strip()
                if not _cl or _cl in ("nan","None",""):
                    continue
                if _cl not in _co_cg:
                    _co_cg[_cl] = {"rpa":0,"ai":0,"presales":0,"other":0,
                                    "total":0,"done":0,"live":0}
                _co_cg[_cl]["total"] += 1
                _cpt2 = str(_crow.get("proj_type","")).strip().lower()
                _cst2 = str(_crow.get("status","")).strip().lower()
                if   _cpt2 == "rpa":                            _co_cg[_cl]["rpa"]      += 1
                elif _cpt2 == "ai agent":                       _co_cg[_cl]["ai"]       += 1
                elif "presales" in _cpt2 or "poc" in _cst2:     _co_cg[_cl]["presales"] += 1
                else:                                            _co_cg[_cl]["other"]    += 1
                if   _cst2 == "completed":                      _co_cg[_cl]["done"] += 1
                elif _cst2 in ("in progress","uat","r&m"):       _co_cg[_cl]["live"] += 1

        if _co_cg:
            _cg_srt    = sorted(_co_cg.items(), key=lambda x: x[1]["total"], reverse=True)
            _cl_show   = _cg_srt[:4]
            _cl_n      = max(len(_cl_show), 1)
            _cl_cols2  = st.columns(_cl_n)
            _cl_pal    = ["#3F8E91","#8B5CF6","#F59E0B","#10B981","#3B82F6","#EC4899"]
            for _ci2, (_cname2, _cinfo2) in enumerate(_cl_show):
                _ca2   = _cl_pal[_ci2 % len(_cl_pal)]
                _clin2 = esc(_cname2[0].upper())
                _cnes2 = esc(_cname2)
                with _cl_cols2[_ci2]:
                    st.markdown(
                        f'<div style="background:linear-gradient(135deg,{_ca2}15,{_ca2}06);'
                        f'border:1.5px solid {_ca2}44;border-radius:14px;padding:18px 16px;'
                        f'border-top:3px solid {_ca2}">'
                        f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:14px">'
                        f'<div style="width:38px;height:38px;border-radius:10px;background:{_ca2}25;'
                        f'border:1px solid {_ca2}50;display:flex;align-items:center;'
                        f'justify-content:center;font-size:16px;font-weight:900;color:{_ca2}">'
                        f'{_clin2}</div>'
                        f'<div style="min-width:0">'
                        f'<div style="font-size:13px;font-weight:800;color:#1F3B4D;'
                        f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{_cnes2}</div>'
                        f'<div style="font-size:9px;color:#64748B;font-weight:600">Enterprise Client</div>'
                        f'</div></div>'
                        f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">'
                        f'<div style="text-align:center;background:#fff;border-radius:8px;'
                        f'padding:8px 4px;border:1px solid #F1F5F9">'
                        f'<div style="font-size:20px;font-weight:900;color:{_ca2}">'
                        f'{_cinfo2["total"]}</div>'
                        f'<div style="font-size:9px;color:#94A3B8;margin-top:2px">Total</div>'
                        f'</div>'
                        f'<div style="text-align:center;background:#fff;border-radius:8px;'
                        f'padding:8px 4px;border:1px solid #F1F5F9">'
                        f'<div style="font-size:20px;font-weight:900;color:#10B981">'
                        f'{_cinfo2["done"]}</div>'
                        f'<div style="font-size:9px;color:#94A3B8;margin-top:2px">Done</div>'
                        f'</div>'
                        f'<div style="text-align:center;background:#fff;border-radius:8px;'
                        f'padding:8px 4px;border:1px solid #F1F5F9">'
                        f'<div style="font-size:20px;font-weight:900;color:#3B82F6">'
                        f'{_cinfo2["rpa"]}</div>'
                        f'<div style="font-size:9px;color:#94A3B8;margin-top:2px">RPA</div>'
                        f'</div>'
                        f'<div style="text-align:center;background:#fff;border-radius:8px;'
                        f'padding:8px 4px;border:1px solid #F1F5F9">'
                        f'<div style="font-size:20px;font-weight:900;color:#8B5CF6">'
                        f'{_cinfo2["ai"]}</div>'
                        f'<div style="font-size:9px;color:#94A3B8;margin-top:2px">AI Agent</div>'
                        f'</div>'
                        f'</div></div>',
                        unsafe_allow_html=True
                    )
        else:
            st.markdown(
                '<div style="text-align:center;padding:18px;color:#94A3B8;font-size:11px;'
                'background:#F8FAFC;border-radius:12px;border:1px solid #E2E8F0">'
                'No client data. Add projects with client names to populate this panel.</div>',
                unsafe_allow_html=True
            )

        # Active projects table (live DB)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
            'letter-spacing:.8px;margin-bottom:12px">Active Portfolio · Live Projects</div>',
            unsafe_allow_html=True
        )
        _proj_sc2 = {
            "In Progress":"#3B82F6","UAT":"#8B5CF6","R&M":"#F59E0B",
            "PDD":"#0EA5E9","Completed":"#10B981","Presales":"#64748B","Discontinued":"#EF4444",
        }
        _live_pf2 = (
            _co_df[_co_df["status"].isin(["In Progress","UAT","R&M","PDD"])]
            if "status" in _co_df.columns else pd.DataFrame()
        )
        if not _live_pf2.empty:
            _pt_html2 = (
                '<div style="background:#fff;border:1px solid #E2E8F0;border-radius:14px;'
                'overflow:hidden">'
                '<div style="display:grid;grid-template-columns:2fr 1.2fr 0.8fr 0.8fr 1fr;'
                'padding:10px 16px;background:#F8FAFC;border-bottom:1px solid #E2E8F0">'
            )
            for _hdr2 in ("Project","Client","Type","Status","Lead"):
                _pt_html2 += (
                    f'<div style="font-size:10px;font-weight:700;color:#64748B;'
                    f'text-transform:uppercase;letter-spacing:.5px">{_hdr2}</div>'
                )
            _pt_html2 += '</div>'
            for _pi2, (_, _pr2) in enumerate(_live_pf2.head(8).iterrows()):
                _prn2  = esc(str(_pr2.get("name",     "—")))
                _prc3  = esc(str(_pr2.get("client",   "—")))
                _prt2  = esc(str(_pr2.get("proj_type","—")))
                _prs2  = str(_pr2.get("status","—"))
                _prl2  = esc(str(_pr2.get("lead",     "—")))
                _prc4  = _proj_sc2.get(_prs2, "#94A3B8")
                _pbg2  = "#FAFAFA" if _pi2 % 2 == 0 else "#fff"
                _pt_html2 += (
                    f'<div style="display:grid;grid-template-columns:2fr 1.2fr 0.8fr 0.8fr 1fr;'
                    f'padding:10px 16px;background:{_pbg2};border-bottom:1px solid #F1F5F9;'
                    f'align-items:center">'
                    f'<div style="font-size:12px;font-weight:600;color:#1F3B4D;overflow:hidden;'
                    f'text-overflow:ellipsis;white-space:nowrap">{_prn2}</div>'
                    f'<div style="font-size:11px;color:#475569">{_prc3}</div>'
                    f'<div style="font-size:10px;font-weight:600;color:#64748B;background:#F1F5F9;'
                    f'border-radius:6px;padding:2px 8px;display:inline-block">{_prt2}</div>'
                    f'<div><span style="font-size:10px;font-weight:700;color:{_prc4};'
                    f'background:{_prc4}18;border-radius:6px;padding:2px 8px;display:inline-block">'
                    f'{esc(_prs2)}</span></div>'
                    f'<div style="font-size:11px;color:#475569">{_prl2}</div>'
                    f'</div>'
                )
            _pt_html2 += '</div>'
            st.markdown(_pt_html2, unsafe_allow_html=True)
        else:
            st.markdown(
                '<div style="text-align:center;padding:20px;color:#94A3B8;font-size:11px;'
                'background:#F8FAFC;border-radius:12px;border:1px solid #E2E8F0">'
                'No active projects. Projects in progress will appear here.</div>',
                unsafe_allow_html=True
            )

        st.markdown("<br>", unsafe_allow_html=True)

        if role == "sales":
            # ══════════════════════════════════════════════════════════════════════
            # SALES ROLE — DATA CHARTS (replaces static marketing cards)
            # ══════════════════════════════════════════════════════════════════════
            st.markdown(
                '<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">'
                '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Portfolio Analytics</span>'
                '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
                'background:#EFF6FF;color:#3B82F6;border:1px solid #BFDBFE">Live Data</span>'
                '</div>',
                unsafe_allow_html=True
            )

            _sc1, _sc2 = st.columns(2)

            # Chart 1 — Conversion Funnel
            with _sc1:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                        'letter-spacing:.8px;margin-bottom:8px">Presales → Delivery Conversion Funnel</div>',
                        unsafe_allow_html=True
                    )
                    _fap = _all_proj.copy()
                    _funnel_stages = [
                        ("Presales",    int(_fap["status"].isin(["Presales"]).sum())          if "status" in _fap.columns else 0),
                        ("POC",         int(_fap["status"].str.contains("POC", na=False).sum()) if "status" in _fap.columns else 0),
                        ("PDD/Design",  int((_fap["status"] == "PDD").sum())                  if "status" in _fap.columns else 0),
                        ("Development", int(_fap["status"].isin(["In Progress"]).sum())       if "status" in _fap.columns else 0),
                        ("UAT / R&M",   int(_fap["status"].isin(["UAT","R&M"]).sum())         if "status" in _fap.columns else 0),
                        ("Completed",   int((_fap["status"] == "Completed").sum())            if "status" in _fap.columns else 0),
                    ]
                    _funnel_labels = [s for s, _ in _funnel_stages]
                    _funnel_values = [v for _, v in _funnel_stages]
                    _funnel_colors = ["#0EA5E9","#8B5CF6","#F59E0B","#06B6D4","#3B82F6","#10B981"]
                    _f_fig = go.Figure(go.Funnel(
                        y=_funnel_labels, x=_funnel_values,
                        textinfo="value+percent initial",
                        textfont=dict(size=11),
                        marker=dict(color=_funnel_colors),
                        connector=dict(line=dict(color="#E2E8F0", width=1)),
                    ))
                    _f_fig.update_layout(
                        margin=dict(t=0, b=0, l=80, r=20), height=260,
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                        font=dict(size=10),
                    )
                    st.plotly_chart(_f_fig, use_container_width=True)

            # Chart 2 — Top Clients by Project Count
            with _sc2:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                        'letter-spacing:.8px;margin-bottom:8px">Top Clients by Project Count</div>',
                        unsafe_allow_html=True
                    )
                    if "client" in _all_proj.columns:
                        _cl_counts = (
                            _all_proj.groupby("client")
                            .size().reset_index(name="count")
                            .sort_values("count", ascending=True)
                            .tail(10)
                        )
                        _cl_fig = go.Figure(go.Bar(
                            x=_cl_counts["count"], y=_cl_counts["client"],
                            orientation="h",
                            marker=dict(
                                color=_cl_counts["count"],
                                colorscale=[[0,"rgba(63,142,145,0.18)"],[1,"#3F8E91"]],
                                showscale=False,
                            ),
                            text=_cl_counts["count"], textposition="outside",
                            textfont=dict(size=10),
                        ))
                        _cl_fig.update_layout(
                            margin=dict(t=0, b=0, l=10, r=30), height=260,
                            xaxis=dict(visible=False),
                            yaxis=dict(tickfont=dict(size=10), autorange="reversed" if len(_cl_counts) > 1 else True),
                            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                        )
                        st.plotly_chart(_cl_fig, use_container_width=True)
                    else:
                        st.info("No client data available.")

            st.markdown("<br>", unsafe_allow_html=True)
            _sc3, _sc4 = st.columns(2)

            # Chart 3 — Cost Saved by Client
            with _sc3:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                        'letter-spacing:.8px;margin-bottom:8px">Cost Savings Delivered — by Client (₹K)</div>',
                        unsafe_allow_html=True
                    )
                    if "client" in _all_proj.columns and "cost_saved" in _all_proj.columns:
                        _cs_df = _all_proj.copy()
                        _cs_df["cost_saved_n"] = pd.to_numeric(_cs_df["cost_saved"], errors="coerce").fillna(0)
                        _cs_grp = (
                            _cs_df.groupby("client")["cost_saved_n"]
                            .sum().reset_index()
                            .sort_values("cost_saved_n", ascending=True)
                            .tail(10)
                        )
                        _cs_grp["label"] = (_cs_grp["cost_saved_n"] / 1000).round(1).astype(str) + "K"
                        _has_data = _cs_grp["cost_saved_n"].sum() > 0
                        if _has_data:
                            _cs_fig = go.Figure(go.Bar(
                                x=_cs_grp["cost_saved_n"] / 1000,
                                y=_cs_grp["client"],
                                orientation="h",
                                marker=dict(color="#F59E0B", opacity=0.85),
                                text=_cs_grp["label"], textposition="outside",
                                textfont=dict(size=10),
                            ))
                            _cs_fig.update_layout(
                                margin=dict(t=0, b=0, l=10, r=50), height=260,
                                xaxis=dict(visible=False),
                                yaxis=dict(tickfont=dict(size=10)),
                                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                            )
                            st.plotly_chart(_cs_fig, use_container_width=True)
                        else:
                            st.info("No cost savings data yet. Add cost_saved values in the Projects tab.")
                    else:
                        st.info("No cost savings data available.")

            # Chart 4 — Monthly Project Activity
            with _sc4:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                        'letter-spacing:.8px;margin-bottom:8px">Monthly Project Activity</div>',
                        unsafe_allow_html=True
                    )
                    if "start" in _all_proj.columns:
                        _ma_df = _all_proj.copy()
                        _ma_df["start_dt"] = pd.to_datetime(_ma_df["start"], dayfirst=True, errors="coerce")
                        _ma_df = _ma_df.dropna(subset=["start_dt"])
                        if not _ma_df.empty:
                            _ma_df["ym"] = _ma_df["start_dt"].dt.to_period("M").astype(str)
                            _ma_grp = _ma_df.groupby("ym").size().reset_index(name="count").tail(12)
                            _ma_fig = go.Figure()
                            _ma_fig.add_trace(go.Bar(
                                x=_ma_grp["ym"], y=_ma_grp["count"],
                                marker=dict(color="#8B5CF6", opacity=0.85),
                                name="Projects Started",
                            ))
                            _ma_fig.update_layout(
                                margin=dict(t=10, b=30, l=20, r=10), height=240,
                                xaxis=dict(tickfont=dict(size=9), tickangle=-30),
                                yaxis=dict(tickfont=dict(size=10), gridcolor="#F1F5F9"),
                                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                showlegend=False,
                            )
                            st.plotly_chart(_ma_fig, use_container_width=True)
                        else:
                            st.info("No start date data to plot activity.")
                    else:
                        st.info("No start date column found.")

            st.markdown("<br>", unsafe_allow_html=True)
            _sc5, _sc6 = st.columns([1, 1.4])

            # Chart 5 — Win Rate Gauge
            with _sc5:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                        'letter-spacing:.8px;margin-bottom:8px">Presales Win Rate</div>',
                        unsafe_allow_html=True
                    )
                    _wr_total   = int(_all_proj["status"].notna().sum()) if "status" in _all_proj.columns else 0
                    _wr_won     = int(_all_proj["status"].isin(["Completed","In Progress","UAT","R&M","PDD"]).sum()) if "status" in _all_proj.columns else 0
                    _wr_disc    = int((_all_proj["status"] == "Discontinued").sum()) if "status" in _all_proj.columns else 0
                    _wr_rate    = round((_wr_won / max(_wr_total, 1)) * 100)
                    _wr_fig = go.Figure(go.Indicator(
                        mode="gauge+number+delta",
                        value=_wr_rate,
                        number=dict(suffix="%", font=dict(size=32, color="#1F3B4D")),
                        delta=dict(reference=60, valueformat=".0f", suffix="%"),
                        gauge=dict(
                            axis=dict(range=[0, 100], tickfont=dict(size=9)),
                            bar=dict(color="#10B981"),
                            bgcolor="#F8FAFC",
                            bordercolor="#E2E8F0",
                            steps=[
                                dict(range=[0, 40],   color="#FEE2E2"),
                                dict(range=[40, 70],  color="#FEF3C7"),
                                dict(range=[70, 100], color="#DCFCE7"),
                            ],
                            threshold=dict(line=dict(color="#3B82F6", width=3), thickness=0.75, value=60),
                        ),
                    ))
                    _wr_fig.update_layout(
                        margin=dict(t=10, b=0, l=10, r=10), height=210,
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    )
                    st.plotly_chart(_wr_fig, use_container_width=True)
                    st.markdown(
                        f'<div style="display:flex;gap:16px;justify-content:center;padding:6px 0">'
                        f'<span style="font-size:11px;color:#10B981;font-weight:700">✓ {_wr_won} Active/Done</span>'
                        f'<span style="font-size:11px;color:#EF4444;font-weight:700">✗ {_wr_disc} Discontinued</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )

            # Chart 6 — ROI by Project Type
            with _sc6:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                        'letter-spacing:.8px;margin-bottom:8px">Avg ROI % by Project Type</div>',
                        unsafe_allow_html=True
                    )
                    if "proj_type" in _all_proj.columns and "roi_pct" in _all_proj.columns:
                        _roi_df = _all_proj.copy()
                        _roi_df["roi_n"] = pd.to_numeric(_roi_df["roi_pct"], errors="coerce")
                        _roi_grp = (
                            _roi_df.dropna(subset=["roi_n"])
                            .groupby("proj_type")["roi_n"]
                            .mean().reset_index()
                            .sort_values("roi_n", ascending=False)
                        )
                        if not _roi_grp.empty:
                            _roi_colors = {"RPA":"#3F8E91","AI Agent":"#8B5CF6","Presales":"#F59E0B"}
                            _roi_bar_colors = [_roi_colors.get(t, "#94A3B8") for t in _roi_grp["proj_type"]]
                            _roi_fig = go.Figure(go.Bar(
                                x=_roi_grp["proj_type"],
                                y=_roi_grp["roi_n"].round(1),
                                marker=dict(color=_roi_bar_colors, opacity=0.9),
                                text=_roi_grp["roi_n"].round(1).astype(str) + "%",
                                textposition="outside",
                                textfont=dict(size=11),
                            ))
                            _roi_fig.update_layout(
                                margin=dict(t=20, b=20, l=20, r=10), height=240,
                                xaxis=dict(tickfont=dict(size=11)),
                                yaxis=dict(tickfont=dict(size=10), gridcolor="#F1F5F9",
                                           ticksuffix="%"),
                                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                showlegend=False,
                            )
                            st.plotly_chart(_roi_fig, use_container_width=True)
                        else:
                            st.info("Add ROI % values in the Projects tab to see this chart.")
                    else:
                        st.info("No ROI data available.")

            st.markdown("<br>", unsafe_allow_html=True)

        if role != "sales":
            # ══════════════════════════════════════════════════════════════════════
            # NON-SALES — STATIC MARKETING INTELLIGENCE CARDS
            # ══════════════════════════════════════════════════════════════════════

            # ════════════════════════════════════════════════════════════════════
            # MARKETING INTELLIGENCE — INDUSTRIES, USPs, TECH STACK, ENGAGEMENT
            # ════════════════════════════════════════════════════════════════════
            st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">'
            '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Marketing Intelligence</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#FFF7ED;color:#EA580C;border:1px solid #FED7AA">Go-to-Market</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#F0FDF4;color:#16A34A;border:1px solid #BBF7D0">Qualesce India</span>'
            '</div>',
            unsafe_allow_html=True
            )

            _mkt_c1, _mkt_c2 = st.columns(2)

            with _mkt_c1:
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                    'letter-spacing:.8px;margin-bottom:12px">Industries We Serve</div>',
                    unsafe_allow_html=True
                )
                _industries = [
                    ("💰", "BFSI",              "Banking, Financial Services & Insurance", "#3B82F6",
                     "Automating loan processing, compliance reporting, reconciliation & fraud detection."),
                    ("🏥", "Healthcare",        "Pharma & Life Sciences",                  "#10B981",
                     "Streamlining claims, patient onboarding, regulatory submissions & lab data pipelines."),
                    ("🏭", "Manufacturing",     "Production & Supply Chain",               "#F59E0B",
                     "Digitising shop-floor reports, inventory sync, quality checks & supplier workflows."),
                    ("🛒", "Retail & FMCG",     "Consumer & Distribution",                 "#8B5CF6",
                     "Automating order management, demand forecasting, invoice processing & reconciliation."),
                ]
                _ind_html = '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">'
                for _iico, _iname, _isub, _iclr, _idesc in _industries:
                    _ind_html += (
                        f'<div style="background:linear-gradient(135deg,{_iclr}12,{_iclr}06);'
                        f'border:1.5px solid {_iclr}40;border-radius:12px;padding:14px 16px">'
                        f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">'
                        f'<span style="font-size:20px">{_iico}</span>'
                        f'<div>'
                        f'<div style="font-size:12px;font-weight:800;color:#1F3B4D">{_iname}</div>'
                        f'<div style="font-size:9px;color:{_iclr};font-weight:600">{_isub}</div>'
                        f'</div></div>'
                        f'<div style="font-size:10px;color:#64748B;line-height:1.5">{_idesc}</div>'
                        f'</div>'
                    )
                _ind_html += '</div>'
                st.markdown(_ind_html, unsafe_allow_html=True)

            with _mkt_c2:
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                    'letter-spacing:.8px;margin-bottom:12px">Why Choose Qualesce</div>',
                    unsafe_allow_html=True
                )
                _usps = [
                    ("⚡", "Rapid Time-to-Value",  "#F59E0B",
                     "Live in weeks, not months — phased delivery from discovery to production."),
                    ("🎯", "End-to-End Ownership", "#3B82F6",
                     "Single partner for design, build, deploy & support — zero handoff risk."),
                    ("🔓", "Platform Agnostic",    "#8B5CF6",
                     "UiPath, Power Automate, Python bots — best-fit technology for your stack."),
                    ("📈", "Proven ROI",           "#10B981",
                     "412%+ average portfolio ROI with measurable cost savings from day one."),
                ]
                _usp_html = '<div style="display:flex;flex-direction:column;gap:10px">'
                for _uico, _utitle, _uclr, _udesc in _usps:
                    _usp_html += (
                        f'<div style="display:flex;align-items:flex-start;gap:12px;padding:12px 14px;'
                        f'background:#fff;border:1.5px solid {_uclr}30;border-radius:10px;'
                        f'border-left:4px solid {_uclr}">'
                        f'<span style="font-size:22px;line-height:1.2">{_uico}</span>'
                        f'<div>'
                        f'<div style="font-size:12px;font-weight:800;color:#1F3B4D;margin-bottom:3px">{_utitle}</div>'
                        f'<div style="font-size:10px;color:#64748B;line-height:1.5">{_udesc}</div>'
                        f'</div></div>'
                    )
                _usp_html += '</div>'
                st.markdown(_usp_html, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            _mkt_c3, _mkt_c4 = st.columns([1.2, 1])

            with _mkt_c3:
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                    'letter-spacing:.8px;margin-bottom:12px">Technology Stack &amp; Partners</div>',
                    unsafe_allow_html=True
                )
                _tech_groups = [
                    ("RPA Platforms",  [("UiPath","#F6511D"),("Power Automate","#0078D4"),
                                        ("Blue Prism","#5A48A5"),("Automation Anywhere","#FF6D00")]),
                    ("AI &amp; LLM",   [("OpenAI GPT","#10A37F"),("LangChain","#1C3553"),
                                        ("Azure OpenAI","#0078D4"),("Hugging Face","#F5A623")]),
                    ("Infrastructure", [("Azure","#0078D4"),("AWS","#FF9900"),
                                        ("Docker","#2496ED"),("Python","#3776AB")]),
                ]
                for _tgrp, _techs in _tech_groups:
                    _tg_html = (
                        f'<div style="margin-bottom:14px">'
                        f'<div style="font-size:10px;font-weight:700;color:#374151;margin-bottom:7px">{_tgrp}</div>'
                        f'<div style="display:flex;flex-wrap:wrap;gap:6px">'
                    )
                    for _tn, _tc in _techs:
                        _tg_html += (
                            f'<span style="font-size:10px;font-weight:700;padding:4px 12px;border-radius:20px;'
                            f'background:{_tc}18;color:{_tc};border:1.5px solid {_tc}50">{_tn}</span>'
                        )
                    _tg_html += '</div></div>'
                    st.markdown(_tg_html, unsafe_allow_html=True)

            with _mkt_c4:
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                    'letter-spacing:.8px;margin-bottom:12px">Engagement Models</div>',
                    unsafe_allow_html=True
                )
                _eng_models = [
                    ("🎯", "Fixed-Price Projects",  "#3B82F6", "Most Popular",
                     "Defined scope, timeline & cost — ideal for new automation implementations."),
                    ("🔄", "Managed Support & R&M", "#8B5CF6", "Recurring Revenue",
                     "Monthly retainer for ongoing bot maintenance, monitoring & enhancements."),
                    ("📦", "License + Consulting",  "#F59E0B", "High Value",
                     "Vendor license procurement bundled with implementation & training services."),
                ]
                for _eico, _ename, _ec, _ebadge, _edesc in _eng_models:
                    st.markdown(
                        f'<div style="padding:14px 16px;background:linear-gradient(135deg,{_ec}08,{_ec}04);'
                        f'border:1.5px solid {_ec}35;border-radius:12px;margin-bottom:10px">'
                        f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">'
                        f'<div style="display:flex;align-items:center;gap:8px">'
                        f'<span style="font-size:18px">{_eico}</span>'
                        f'<span style="font-size:12px;font-weight:800;color:#1F3B4D">{_ename}</span>'
                        f'</div>'
                        f'<span style="font-size:9px;font-weight:700;padding:2px 8px;border-radius:10px;'
                        f'background:{_ec}20;color:{_ec};border:1px solid {_ec}40">{_ebadge}</span>'
                        f'</div>'
                        f'<div style="font-size:10px;color:#64748B;line-height:1.5">{_edesc}</div>'
                        f'</div>',
                        unsafe_allow_html=True
                    )

            st.markdown("<br>", unsafe_allow_html=True)

            # Key Achievements Banner
            st.markdown(
                '<div style="background:linear-gradient(135deg,#1E1B4B 0%,#312E81 50%,#1E1B4B 100%);'
                'border-radius:14px;padding:20px 28px;margin-bottom:16px">'
                '<div style="font-size:10px;font-weight:700;color:#A5B4FC;text-transform:uppercase;'
                'letter-spacing:1px;margin-bottom:14px">Qualesce — Key Achievements &amp; Recognitions</div>'
                '<div style="display:flex;gap:24px;flex-wrap:wrap;justify-content:space-between">'

                '<div style="text-align:center;min-width:80px">'
                '<div style="font-size:26px;font-weight:900;color:#818CF8;line-height:1">7+</div>'
                '<div style="font-size:10px;color:#C7D2FE;margin-top:3px">Years in RPA</div>'
                '</div>'

                '<div style="text-align:center;min-width:80px">'
                '<div style="font-size:26px;font-weight:900;color:#34D399;line-height:1">100+</div>'
                '<div style="font-size:10px;color:#C7D2FE;margin-top:3px">Bots Deployed</div>'
                '</div>'

                '<div style="text-align:center;min-width:80px">'
                '<div style="font-size:26px;font-weight:900;color:#FCD34D;line-height:1">₹50Cr+</div>'
                '<div style="font-size:10px;color:#C7D2FE;margin-top:3px">Client Value Created</div>'
                '</div>'

                '<div style="text-align:center;min-width:80px">'
                '<div style="font-size:26px;font-weight:900;color:#F472B6;line-height:1">98%</div>'
                '<div style="font-size:10px;color:#C7D2FE;margin-top:3px">Client Retention</div>'
                '</div>'

                '<div style="text-align:center;min-width:80px">'
                '<div style="font-size:26px;font-weight:900;color:#38BDF8;line-height:1">UiPath</div>'
                '<div style="font-size:10px;color:#C7D2FE;margin-top:3px">Gold Partner</div>'
                '</div>'

                '<div style="text-align:center;min-width:80px">'
                '<div style="font-size:26px;font-weight:900;color:#A3E635;line-height:1">ISO</div>'
                '<div style="font-size:10px;color:#C7D2FE;margin-top:3px">27001 Aligned</div>'
                '</div>'

                '</div>'
                '</div>',
                unsafe_allow_html=True
            )

            st.markdown("<br>", unsafe_allow_html=True)

            st.markdown(
            '<h3 style="font-size:15px;font-weight:700;color:#1F3B4D;margin:0 0 10px;letter-spacing:-.2px">'
            'Project Type Breakdown</h3>',
            unsafe_allow_html=True
        )
        _pt_df = df.copy()
        _pt_type_col = _pt_df["proj_type"].str.strip().str.lower() if "proj_type" in _pt_df.columns else pd.Series([], dtype=str)
        _pt_stat_col = _pt_df["status"].str.strip().str.lower()    if "status"    in _pt_df.columns else pd.Series([], dtype=str)

        _ai_count       = int((_pt_type_col == "ai agent").sum())
        _rpa_count      = int((_pt_type_col == "rpa").sum())
        _presales_count = int(_pt_stat_col.isin(["presales", "internal poc", "external poc"]).sum())
        _other_count    = int(len(_pt_df) - _ai_count - _rpa_count - _presales_count)
        _other_count    = max(_other_count, 0)

        _type_cards = [
            ("AI Agent", "AI",  "#6366F1", _ai_count),
            ("RPA",      "RPA", "#F59E0B", _rpa_count),
            ("Presales", "PRE", "#10B981", _presales_count),
            ("Other",    "OTH", "#94A3B8", _other_count),
        ]
        _tc_cols = st.columns(4)
        for _tc_col, (_tc_lbl, _tc_ico, _tc_clr, _tc_cnt) in zip(_tc_cols, _type_cards):
            _tc_col.markdown(
                f'<div style="background:linear-gradient(135deg,{_tc_clr}18,{_tc_clr}06);'
                f'border:1.5px solid {_tc_clr}66;border-radius:12px;padding:18px 10px 14px;'
                f'text-align:center;min-height:90px">'
                f'<div style="font-size:11px;font-weight:700;color:{_tc_clr};letter-spacing:.6px;'
                f'text-transform:uppercase;margin-bottom:6px">{_tc_lbl}</div>'
                f'<div style="font-size:28px;font-weight:900;color:{_tc_clr};line-height:1">{_tc_cnt}</div>'
                f'<div style="font-size:10px;color:#94A3B8;margin-top:4px">'
                f'project{"s" if _tc_cnt != 1 else ""}</div>'
                f'</div>',
                unsafe_allow_html=True
            )
        st.markdown("<br>", unsafe_allow_html=True)

        # ── MONTH-OVER-MONTH KPI DELTAS ───────────────────────────────────────────
        _mom_today      = date.today()
        _mom_cur_start  = _mom_today.replace(day=1)
        _mom_prev_end   = _mom_cur_start - timedelta(days=1)
        _mom_prev_start = _mom_prev_end.replace(day=1)
        _mom_cur_lbl    = _mom_today.strftime("%b %Y")
        _mom_prev_lbl   = _mom_prev_end.strftime("%b %Y")

        def _mom_parse(val):
            if not val or str(val).strip() in ("", "nan", "None"):
                return None
            return _parse_dmy(str(val).strip())

        _mom_start_cur = _mom_start_prev = _mom_end_cur = _mom_end_prev = 0
        if not df.empty:
            for _, _mr in df.iterrows():
                _ms = _mom_parse(_mr.get("start", ""))
                _me = _mom_parse(_mr.get("end", ""))
                if _ms:
                    if _mom_cur_start <= _ms <= _mom_today:
                        _mom_start_cur += 1
                    elif _mom_prev_start <= _ms <= _mom_prev_end:
                        _mom_start_prev += 1
                if _me:
                    if _mom_cur_start <= _me <= _mom_today:
                        _mom_end_cur += 1
                    elif _mom_prev_start <= _me <= _mom_prev_end:
                        _mom_end_prev += 1

        _mom_active_cur   = int(df["status"].isin(["In Progress", "UAT", "R&M"]).sum()) if not df.empty else 0
        _mom_presales_cur = int(df["status"].str.contains("Presales|POC|PDD", na=False).sum()) if not df.empty else 0

        def _mom_delta_html(cur, prev, label, color):
            diff = cur - prev
            diff_str = (f'+{diff}' if diff > 0 else str(diff)) if diff != 0 else "–"
            diff_color = "#10B981" if diff > 0 else ("#EF4444" if diff < 0 else "#94A3B8")
            return (
                f'<div style="background:#F8FAFC;border:1.5px solid #E2E8F0;border-radius:12px;'
                f'padding:14px 16px;border-top:3px solid {color}">'
                f'<div style="font-size:10px;font-weight:700;color:{color};text-transform:uppercase;'
                f'letter-spacing:.6px;margin-bottom:4px">{label}</div>'
                f'<div style="display:flex;align-items:baseline;gap:8px">'
                f'<span style="font-size:24px;font-weight:800;color:#1F3B4D;line-height:1">{cur}</span>'
                f'<span style="font-size:12px;font-weight:700;color:{diff_color};'
                f'background:{diff_color}18;padding:1px 7px;border-radius:10px">{diff_str} vs {_mom_prev_lbl}</span>'
                f'</div>'
                f'<div style="font-size:10px;color:#94A3B8;margin-top:3px">{_mom_cur_lbl} · '
                f'prev: {prev}</div>'
                f'</div>'
            )

        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">'
            '<span style="font-size:13px;font-weight:700;color:#1F3B4D">Month-over-Month Trends</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#EFF6FF;color:#3B82F6;border:1px solid #BFDBFE">vs last month</span>'
            '</div>',
            unsafe_allow_html=True
        )
        _mom_c1, _mom_c2, _mom_c3, _mom_c4 = st.columns(4)
        _mom_c1.markdown(_mom_delta_html(_mom_start_cur, _mom_start_prev, "New Projects", "#3B82F6"), unsafe_allow_html=True)
        _mom_c2.markdown(_mom_delta_html(_mom_end_cur, _mom_end_prev, "Delivered", "#10B981"), unsafe_allow_html=True)
        _mom_c3.markdown(_mom_delta_html(_mom_active_cur, 0, "Active (IP+UAT+R&M)", "#F59E0B"), unsafe_allow_html=True)
        _mom_c4.markdown(_mom_delta_html(_mom_presales_cur, 0, "In Pipeline", "#8B5CF6"), unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # ── NOTIFICATION ALERT PANEL ──────────────────────────────────────────────
        NOTIF_DEFS = [
            {
                "key":    "Important",
                "label":  "Important Tasks",
                "icon":   "!",
                "color":  "#F43F5E",
                "bg":     "#FFF1F2",
                "border": "#FDA4AF",
                "note":   "High-priority tasks requiring immediate attention",
            },
        ]

        # Compute project lists for each alert type (respects lead filter)
        def get_alert_projects(status_key):
            mask = dash_df["status"].str.contains(status_key, na=False)
            return dash_df[mask]

        active_notifs = [n for n in NOTIF_DEFS if n["key"] not in st.session_state.dismissed_notifs]
        notif_data    = {n["key"]: get_alert_projects(n["key"]) for n in active_notifs}
        visible_notifs = [n for n in active_notifs if len(notif_data[n["key"]]) > 0]

        if visible_notifs:
            total_important = len(notif_data.get("Important", pd.DataFrame()))
            if total_important > 0 and "Important" not in st.session_state.dismissed_notifs:
                st.markdown(
                    f'<div class="notif-alert" style="background:#FFF1F2;border:2px solid #F43F5E;'
                    f'border-radius:10px;padding:10px 16px;display:flex;align-items:center;'
                    f'gap:10px;margin-bottom:8px">'
                    f'<span style="font-size:13px;font-weight:900;color:#F43F5E;padding:2px 7px;background:#FEE2E2;border-radius:6px">!</span>'
                    f'<span style="font-weight:800;color:#BE123C;font-size:13px">ALERT:</span>'
                    f'<span style="color:#9F1239;font-size:12px">'
                    f'<b>{total_important}</b> project(s) marked as <b>Important</b> require immediate attention!</span>'
                    f'</div>',
                    unsafe_allow_html=True
                )

            notif_cols = st.columns(len(visible_notifs))
            for col, notif in zip(notif_cols, visible_notifs):
                proj_list   = notif_data[notif["key"]]
                proj_count  = len(proj_list)
                preview     = proj_list["name"].head(3).tolist()
                preview_str = "  •  ".join(preview) + ("  …" if proj_count > 3 else "")
                is_active   = st.session_state.show_notif_detail == notif["key"]

                col.markdown(f"""
                <div class="notif-alert" style="background:{notif['bg']};border:1.5px solid {notif['border']};
                  border-radius:12px;padding:12px 14px;min-height:90px">
                  <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px">
                    <span style="font-size:16px">{notif['icon']}</span>
                    <span style="font-size:10px;font-weight:800;color:{notif['color']};
                      background:{notif['color']}18;padding:2px 8px;border-radius:20px">{proj_count} projects</span>
                  </div>
                  <div style="font-size:12px;font-weight:800;color:#1E293B;margin-bottom:2px">{notif['label']}</div>
                  <div style="font-size:10px;color:#64748B;margin-bottom:6px">{notif['note']}</div>
                  <div style="font-size:9.5px;color:{notif['color']};font-style:italic;
                    white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{preview_str}</div>
                </div>""".strip(), unsafe_allow_html=True)

                btn_c1, btn_c2 = col.columns(2)
                detail_label = "Hide" if is_active else "Details"
                if btn_c1.button(detail_label, key=f"notif_detail_{notif['key']}", use_container_width=True,
                                 type="primary" if is_active else "secondary"):
                    st.session_state.show_notif_detail = None if is_active else notif["key"]
                    st.rerun()
                if btn_c2.button("Dismiss", key=f"notif_dismiss_{notif['key']}", use_container_width=True,
                                 help="Dismiss this notification"):
                    st.session_state.dismissed_notifs.add(notif["key"])
                    if st.session_state.show_notif_detail == notif["key"]:
                        st.session_state.show_notif_detail = None
                    st.rerun()

            # ── NOTIFICATION POPUP DETAIL ─────────────────────────────────────────
            if st.session_state.show_notif_detail:
                nd_key   = st.session_state.show_notif_detail
                nd_info  = next((n for n in NOTIF_DEFS if n["key"] == nd_key), None)
                nd_projs = notif_data.get(nd_key, pd.DataFrame())
                if nd_info and not nd_projs.empty:
                    st.markdown(f"""
                    <div class="notif-popup" style="background:{nd_info['bg']};
                      border:2px solid {nd_info['border']}">
                      <div style="display:flex;align-items:center;justify-content:space-between;
                        margin-bottom:14px">
                        <div style="display:flex;align-items:center;gap:10px">
                          <span style="font-size:13px;font-weight:900;padding:2px 8px;background:{nd_info['border']};color:{nd_info['color']};border-radius:6px">{nd_info['icon']}</span>
                          <div>
                            <div style="font-size:14px;font-weight:800;color:#1E293B">
                              {nd_info['label']} — {len(nd_projs)} Projects</div>
                            <div style="font-size:11px;color:#64748B">{nd_info['note']}</div>
                          </div>
                        </div>
                      </div>
                    </div>""".strip(), unsafe_allow_html=True)

                    # Project table inside popup
                    pop_hdr = st.columns([0.4, 3.0, 2.0, 2.2, 1.4, 1.2, 1.2])
                    for ph, pl in zip(pop_hdr, ["ID","Project Name","Client","Employee","Status","Start","End"]):
                        ph.markdown(f'<div style="font-size:9px;font-weight:700;text-transform:uppercase;'
                                    f'color:{nd_info["color"]};letter-spacing:.5px;padding:3px 0;'
                                    f'border-bottom:2px solid {nd_info["border"]}">{pl}</div>',
                                    unsafe_allow_html=True)

                    for _, prow in nd_projs.iterrows():
                        pc = st.columns([0.4, 3.0, 2.0, 2.2, 1.4, 1.2, 1.2])
                        pc[0].markdown(cell(prow.get("id",""), size="10px", color="#94A3B8"), unsafe_allow_html=True)
                        pc[1].markdown(f'<span style="font-size:11px;font-weight:700;color:#111827">'
                                       f'{esc(str(prow.get("name","")))}</span>', unsafe_allow_html=True)
                        pc[2].markdown(cell(prow.get("client",""), size="11px"), unsafe_allow_html=True)
                        pc[3].markdown(cell(prow.get("employee",""), size="11px"), unsafe_allow_html=True)
                        pc[4].markdown(badge_html(str(prow.get("status",""))), unsafe_allow_html=True)
                        pc[5].markdown(cell(prow.get("start",""), size="10px", color="#64748B"), unsafe_allow_html=True)
                        pc[6].markdown(cell(prow.get("end","") or "Ongoing", size="10px", color="#64748B"), unsafe_allow_html=True)

                    st.markdown("<br>", unsafe_allow_html=True)
                    pa, pb, pc_col = st.columns([1.5, 1.5, 3])
                    if pa.button("Open in Projects Tab", key="notif_goto_projects",
                                 type="primary", use_container_width=True):
                        st.session_state.project_filter_preset = nd_key
                        st.session_state.active_tab            = "projects"
                        st.session_state.show_notif_detail     = None
                        st.rerun()
                    if pb.button("Set Dashboard Filter", key="notif_set_slicer",
                                 use_container_width=True):
                        st.session_state.dash_slicer       = nd_key
                        st.session_state.show_notif_detail = None
                        st.rerun()
                    if pc_col.button("Close Panel", key="notif_close_popup",
                                     use_container_width=True):
                        st.session_state.show_notif_detail = None
                        st.rerun()
                    st.markdown("---")

        if st.session_state.dismissed_notifs:
            if st.button("Restore Notifications", key="restore_notifs",
                         help="Re-show all dismissed alerts"):
                st.session_state.dismissed_notifs = set()
                st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)


        # ── CHARTS ────────────────────────────────────────────────────────────────
        _client_label = st.session_state.dash_client_filter
        _slicer_key   = st.session_state.dash_slicer

        # Apply slicer filter on top of client filter so the pie reflects both
        _pie_df = dash_df
        _slicer_label = None
        if _slicer_key is not None:
            if _slicer_key == "__active_dev__":
                _sm = dash_df["status"].isin({"In Progress"})
                if "is_active" in dash_df.columns:
                    _sm = _sm & (~dash_df["is_active"].astype(str).str.strip().str.lower().isin(["false","0","no"]))
                _pie_df, _slicer_label = dash_df[_sm], "Active Development"
            elif _slicer_key == "__new__":
                _sm = dash_df["is_new"].astype(str).str.lower().isin(["true","1","yes"]) if "is_new" in dash_df.columns else pd.Series([False]*len(dash_df))
                _pie_df, _slicer_label = dash_df[_sm], "New Added"
            elif _slicer_key == "POC":
                _pie_df, _slicer_label = dash_df[dash_df["status"].str.contains("POC", na=False)], "POC"
            else:
                _pie_df, _slicer_label = dash_df[dash_df["status"].str.contains(_slicer_key, na=False)], _slicer_key

        _cl_part = f" — {_client_label}" if _client_label != "All" else " — All Clients"
        _sl_part = f" · {_slicer_label}" if _slicer_label else ""

        _chart_c1, _chart_c2 = st.columns(2)

        with _chart_c1:
            with st.container(border=True):
                _tm_title = f"Status Breakdown{_cl_part}{_sl_part}"
                st.markdown(f'<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px">{_tm_title}</div>', unsafe_allow_html=True)
                if not _pie_df.empty:
                    _sc = _pie_df["status"].value_counts().reset_index()
                    _sc.columns = ["status", "count"]
                    _color_map = {s: STATUS_STYLES.get(s, {"dot": "#94A3B8"})["dot"] for s in _sc["status"]}
                    _tm_fig = px.treemap(
                        _sc, path=["status"], values="count",
                        color="status", color_discrete_map=_color_map,
                    )
                    _tm_fig.update_traces(
                        textinfo="label+value",
                        textfont=dict(size=12),
                        hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Share: %{percentRoot:.1%}<extra></extra>",
                    )
                    _tm_fig.update_layout(
                        margin=dict(t=0, b=0, l=0, r=0), height=240,
                        paper_bgcolor="rgba(0,0,0,0)",
                    )
                    st.plotly_chart(_tm_fig, use_container_width=True)
                else:
                    st.info("No data for current filter.")

        with _chart_c2:
            with st.container(border=True):
                _bar_title = f"Projects by Client{_cl_part}{_sl_part}"
                st.markdown(f'<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px">{_bar_title}</div>', unsafe_allow_html=True)
                _bar_src = _pie_df if not _pie_df.empty else dash_df
                if not _bar_src.empty and "client" in _bar_src.columns:
                    _ccounts = (_bar_src.groupby("client").size()
                                .reset_index(name="count")
                                .sort_values("count", ascending=True))
                    _bar_fig = go.Figure(go.Bar(
                        x=_ccounts["count"], y=_ccounts["client"],
                        orientation="h",
                        marker=dict(color="#5FA9AB", opacity=0.85),
                        text=_ccounts["count"], textposition="outside",
                        textfont=dict(size=10),
                    ))
                    _bar_fig.update_layout(
                        margin=dict(t=0, b=0, l=0, r=30), height=240,
                        xaxis=dict(visible=False),
                        yaxis=dict(tickfont=dict(size=10)),
                        plot_bgcolor="rgba(0,0,0,0)",
                        paper_bgcolor="rgba(0,0,0,0)",
                    )
                    st.plotly_chart(_bar_fig, use_container_width=True)
                else:
                    st.info("No client data available.")

        # ── SALES PERSPECTIVE SECTION ────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">'
            '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Sales Perspective</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#0EA5E920;color:#0EA5E9;border:1px solid #0EA5E940">Deals &amp; Conversion</span>'
            '</div>',
            unsafe_allow_html=True
        )

        # Compute sales-specific slices from the current client-filtered df
        _sp_presales   = dash_df[dash_df["status"] == "Presales"]
        _sp_poc        = dash_df[dash_df["status"].str.contains("POC", na=False)]
        _sp_pdd        = dash_df[dash_df["status"] == "PDD"]
        _sp_inprog     = dash_df[dash_df["status"] == "In Progress"]
        _sp_uat        = dash_df[dash_df["status"] == "UAT"]
        _sp_rm         = dash_df[dash_df["status"] == "R&M"]
        _sp_completed  = dash_df[dash_df["status"] == "Completed"]
        _sp_disc       = dash_df[dash_df["status"] == "Discontinued"]
        _sp_pipeline   = len(_sp_presales) + len(_sp_poc) + len(_sp_pdd)
        _sp_won        = len(_sp_inprog) + len(_sp_rm) + len(_sp_completed) + len(_sp_uat)
        _sp_win_rate   = round((_sp_won / max(_sp_won + len(_sp_disc), 1)) * 100)

        # At-risk: active projects whose end date has passed
        _today_sp = date.today()
        def _sp_is_overdue(row):
            end_val = row.get("end", "")
            if not end_val or not str(end_val).strip():
                return False
            s = str(end_val).strip()
            for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                try:
                    return datetime.strptime(s, fmt).date() < _today_sp
                except ValueError:
                    pass
            return False

        _sp_at_risk_statuses = {"In Progress", "UAT", "R&M", "PDD", "Presales"}
        _sp_at_risk_df = dash_df[
            dash_df.apply(_sp_is_overdue, axis=1) &
            dash_df["status"].isin(_sp_at_risk_statuses)
        ]
        _sp_at_risk_count = len(_sp_at_risk_df)

        # Sales KPI cards
        _sk1, _sk2, _sk3, _sk4, _sk5 = st.columns(5)
        _sales_kpi_data = [
            (_sk1, "Presales",    len(_sp_presales),   "#0EA5E9", "🎯"),
            (_sk2, "POC Active",  len(_sp_poc),        "#8B5CF6", "🔬"),
            (_sk3, "In Deals", _sp_pipeline,        "#F59E0B", "📋"),
            (_sk4, "Win Rate",    f"{_sp_win_rate}%",  "#10B981", "🏆"),
            (_sk5, "At Risk",     _sp_at_risk_count,   "#EF4444", "⚠ï¸"),
        ]
        for _sc, _sl, _sv, _scolor, _sico in _sales_kpi_data:
            _sc.markdown(
                f'<div style="background:#F8FAFC;border:1.5px solid {_scolor}33;border-radius:12px;'
                f'padding:14px 16px;text-align:center;border-top:3px solid {_scolor}">'
                f'<div style="font-size:16px;margin-bottom:4px">{_sico}</div>'
                f'<div style="font-size:22px;font-weight:800;color:{_scolor};line-height:1">{_sv}</div>'
                f'<div style="font-size:11px;color:#64748B;margin-top:4px;font-weight:600">{_sl}</div>'
                f'</div>',
                unsafe_allow_html=True
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # Sales funnel + stacked client breakdown
        _sf_col, _cb_col = st.columns(2)

        with _sf_col:
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                    'text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">'
                    'Sales Deals Funnel</div>',
                    unsafe_allow_html=True
                )
                _funnel_stages = ["Presales / PDD", "POC", "In Progress / UAT", "R&M / Completed"]
                _funnel_counts = [
                    len(_sp_presales) + len(_sp_pdd),
                    len(_sp_poc),
                    len(_sp_inprog) + len(_sp_uat),
                    len(_sp_rm) + len(_sp_completed),
                ]
                _funnel_fig = go.Figure(go.Funnel(
                    y=_funnel_stages,
                    x=_funnel_counts,
                    textinfo="value+percent initial",
                    marker=dict(color=["#0EA5E9", "#8B5CF6", "#06B6D4", "#10B981"]),
                ))
                _funnel_fig.update_layout(
                    margin=dict(t=0, b=0, l=0, r=0), height=240,
                    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    font=dict(size=10),
                )
                st.plotly_chart(_funnel_fig, use_container_width=True)

        with _cb_col:
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                    'text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">'
                    'Project Status per Client</div>',
                    unsafe_allow_html=True
                )
                if not dash_df.empty and "client" in dash_df.columns:
                    _cb_statuses = ["R&M", "In Progress", "UAT", "Completed", "Presales"]
                    _cb_src = dash_df[dash_df["status"].isin(_cb_statuses)].copy()
                    _cb_grouped = _cb_src.groupby(["client", "status"]).size().reset_index(name="count")
                    _cb_colors = {
                        "R&M": "#3B82F6", "In Progress": "#06B6D4",
                        "UAT": "#F59E0B", "Completed": "#10B981", "Presales": "#0EA5E9",
                    }
                    _cb_fig = go.Figure()
                    for _cst in _cb_statuses:
                        _cst_data = _cb_grouped[_cb_grouped["status"] == _cst]
                        if not _cst_data.empty:
                            _cb_fig.add_trace(go.Bar(
                                name=_cst,
                                y=_cst_data["client"],
                                x=_cst_data["count"],
                                orientation="h",
                                marker_color=_cb_colors.get(_cst, "#94A3B8"),
                            ))
                    _cb_fig.update_layout(
                        barmode="stack",
                        margin=dict(t=0, b=0, l=0, r=0), height=240,
                        legend=dict(font=dict(size=8), orientation="h", y=-0.15),
                        xaxis=dict(visible=False),
                        yaxis=dict(tickfont=dict(size=9)),
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    )
                    st.plotly_chart(_cb_fig, use_container_width=True)
                else:
                    st.info("No client data available.")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── CRM LEADS & OPPORTUNITIES DASHBOARD ──────────────────────────────────────
        _crm_leads_dash = auth.get_all_leads()
        _crm_opps_dash  = auth.get_all_opportunities()
        _crm_acts_dash  = auth.get_all_activities()

        if _crm_leads_dash or _crm_opps_dash:
            st.markdown(
                '<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">'
                '<span style="font-size:14px;font-weight:800;color:#1F3B4D">CRM Pipeline</span>'
                '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
                'background:#8B5CF620;color:#8B5CF6;border:1px solid #8B5CF640">Leads &amp; Opportunities</span>'
                '</div>',
                unsafe_allow_html=True
            )

            # CRM KPI row
            _ck1, _ck2, _ck3, _ck4, _ck5 = st.columns(5)
            _crm_total_leads  = len(_crm_leads_dash)
            _crm_qual         = sum(1 for l in _crm_leads_dash if l["status"] in ("Qualified", "Proposal", "Negotiation"))
            _crm_won          = sum(1 for l in _crm_leads_dash if l["status"] == "Won")
            _crm_pipe_val     = sum(o["value"] for o in _crm_opps_dash if o["stage"] not in ("Closed Won", "Closed Lost"))
            _crm_open_acts    = sum(1 for a in _crm_acts_dash if not a["is_done"])
            _crm_kpis = [
                (_ck1, "Total Leads",     _crm_total_leads,              "#0EA5E9", "👥"),
                (_ck2, "Qualified",       _crm_qual,                     "#8B5CF6", "✅"),
                (_ck3, "Pipeline Value",  f"₹{_crm_pipe_val:,.0f}",     "#F59E0B", "💰"),
                (_ck4, "Won",             _crm_won,                      "#10B981", "🏆"),
                (_ck5, "Open Activities", _crm_open_acts,                "#EF4444", "📋"),
            ]
            for _kc, _kl, _kv, _kcol, _kico in _crm_kpis:
                _kc.markdown(
                    f'<div style="background:#F8FAFC;border:1.5px solid {_kcol}33;border-radius:12px;'
                    f'padding:14px 16px;text-align:center;border-top:3px solid {_kcol}">'
                    f'<div style="font-size:16px;margin-bottom:4px">{_kico}</div>'
                    f'<div style="font-size:22px;font-weight:800;color:{_kcol};line-height:1">{_kv}</div>'
                    f'<div style="font-size:11px;color:#64748B;margin-top:4px;font-weight:600">{_kl}</div>'
                    f'</div>',
                    unsafe_allow_html=True
                )

            st.markdown("<br>", unsafe_allow_html=True)
            _crm_left, _crm_right = st.columns(2)

            # Lead status breakdown chart
            with _crm_left:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                        'text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">'
                        'Leads by Status</div>',
                        unsafe_allow_html=True
                    )
                    _lead_status_counts = {}
                    for _l in _crm_leads_dash:
                        _lead_status_counts[_l["status"]] = _lead_status_counts.get(_l["status"], 0) + 1
                    _lead_status_order = [s for s in auth.CRM_LEAD_STATUSES if s in _lead_status_counts]
                    if _lead_status_order:
                        _lead_colors_map = {
                            "New": "#0EA5E9", "Contacted": "#06B6D4", "Qualified": "#8B5CF6",
                            "Proposal": "#F59E0B", "Negotiation": "#EF4444",
                            "Won": "#10B981", "Lost": "#6B7280",
                        }
                        _ls_fig = go.Figure(go.Bar(
                            x=_lead_status_order,
                            y=[_lead_status_counts[s] for s in _lead_status_order],
                            marker_color=[_lead_colors_map.get(s, "#94A3B8") for s in _lead_status_order],
                            text=[_lead_status_counts[s] for s in _lead_status_order],
                            textposition="outside",
                        ))
                        _ls_fig.update_layout(
                            margin=dict(t=10, b=0, l=0, r=0), height=220,
                            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                            xaxis=dict(tickfont=dict(size=9)),
                            yaxis=dict(visible=False),
                            font=dict(size=10),
                            showlegend=False,
                        )
                        st.plotly_chart(_ls_fig, use_container_width=True)
                    else:
                        st.info("No lead data.")

            # Top open opportunities
            with _crm_right:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                        'text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">'
                        'Top Open Opportunities</div>',
                        unsafe_allow_html=True
                    )
                    _open_opps = [o for o in _crm_opps_dash if o["stage"] not in ("Closed Won", "Closed Lost")]
                    _open_opps_sorted = sorted(_open_opps, key=lambda o: o["value"], reverse=True)[:6]
                    if _open_opps_sorted:
                        _OPP_STAGE_DASH_COLORS = {
                            "Prospecting": "#0EA5E9", "Qualification": "#06B6D4",
                            "Proposal": "#F59E0B", "Negotiation": "#EF4444",
                        }
                        _oh_s = "font-size:9px;font-weight:700;text-transform:uppercase;color:#94A3B8;padding:3px 0;border-bottom:1px solid #E2E8F0"
                        _op_hc = st.columns([4, 2, 2])
                        _op_hc[0].markdown(f'<div style="{_oh_s}">Title</div>', unsafe_allow_html=True)
                        _op_hc[1].markdown(f'<div style="{_oh_s}">Value</div>', unsafe_allow_html=True)
                        _op_hc[2].markdown(f'<div style="{_oh_s}">Stage</div>', unsafe_allow_html=True)
                        for _oo in _open_opps_sorted:
                            _op_rc = st.columns([4, 2, 2])
                            _oo_co_html = (f'<br><span style="font-size:9px;color:#94A3B8">{esc(_oo["company_name"])}</span>'
                                           if _oo["company_name"] else "")
                            _op_rc[0].markdown(
                                f'<div style="font-size:11px;font-weight:600;color:#111827;padding:4px 0">'
                                f'{esc(_oo["title"])}{_oo_co_html}</div>',
                                unsafe_allow_html=True
                            )
                            _op_rc[1].markdown(
                                f'<div style="font-size:12px;font-weight:700;color:#0F172A;padding:4px 0">₹{_oo["value"]:,.0f}</div>',
                                unsafe_allow_html=True
                            )
                            _ost_color = _OPP_STAGE_DASH_COLORS.get(_oo["stage"], "#6B7280")
                            _op_rc[2].markdown(
                                f'<span style="font-size:9px;font-weight:700;background:{_ost_color}22;color:{_ost_color};'
                                f'padding:2px 7px;border-radius:10px">{esc(_oo["stage"])}</span>',
                                unsafe_allow_html=True
                            )
                    else:
                        st.info("No open opportunities.")

            # Recent activities strip
            if _crm_acts_dash:
                st.markdown("<br>", unsafe_allow_html=True)
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                        'text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">'
                        'Recent Activities</div>',
                        unsafe_allow_html=True
                    )
                    _ACT_TYPE_COLS = {
                        "Call": "#0EA5E9", "Email": "#8B5CF6", "Meeting": "#F59E0B",
                        "Demo": "#10B981", "Follow-up": "#EF4444", "Other": "#6B7280",
                    }
                    _recent_acts = sorted(_crm_acts_dash, key=lambda a: a["activity_date"] or "", reverse=True)[:5]
                    _ra_h = st.columns([1.2, 3, 2.5, 1.5, 1.2])
                    for _rh, _rl in zip(_ra_h, ["Type", "Subject", "Lead / Opp", "Date", "Status"]):
                        _rh.markdown(
                            f'<div style="font-size:9px;font-weight:700;text-transform:uppercase;'
                            f'color:#94A3B8;padding:3px 0;border-bottom:1px solid #E2E8F0">{_rl}</div>',
                            unsafe_allow_html=True
                        )
                    for _ra in _recent_acts:
                        _ra_c = st.columns([1.2, 3, 2.5, 1.5, 1.2])
                        _ratc = _ACT_TYPE_COLS.get(_ra["type"], "#6B7280")
                        _ra_c[0].markdown(
                            f'<span style="font-size:9px;font-weight:700;background:{_ratc}22;color:{_ratc};'
                            f'padding:2px 6px;border-radius:8px">{esc(_ra["type"])}</span>',
                            unsafe_allow_html=True
                        )
                        _ra_c[1].markdown(f'<span style="font-size:11px;color:#111827">{esc(_ra["subject"])}</span>', unsafe_allow_html=True)
                        _ra_ref = " / ".join(filter(None, [_ra["company_name"], _ra["opportunity_title"]]))
                        _ra_c[2].markdown(f'<span style="font-size:10px;color:#64748B">{esc(_ra_ref) if _ra_ref else "—"}</span>', unsafe_allow_html=True)
                        _ra_c[3].markdown(f'<span style="font-size:10px;color:#64748B">{esc(fmt_date(_ra["activity_date"])) if _ra["activity_date"] else "—"}</span>', unsafe_allow_html=True)
                        _done_c = "#10B981" if _ra["is_done"] else "#F59E0B"
                        _done_t = "Done" if _ra["is_done"] else "Pending"
                        _ra_c[4].markdown(
                            f'<span style="font-size:9px;font-weight:700;background:{_done_c}22;color:{_done_c};'
                            f'padding:2px 6px;border-radius:8px">{_done_t}</span>',
                            unsafe_allow_html=True
                        )

            st.markdown("<br>", unsafe_allow_html=True)

        # ── MONTHLY PROJECTS WON & ROI TREND ─────────────────────────────────────────
        with st.container(border=True):
            st.markdown(
                '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                'text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px">'
                'Monthly Projects Won &amp; ROI Trend — Completion velocity over time</div>',
                unsafe_allow_html=True
            )
            # Build month-wise counts from end date (projects that moved to Completed or R&M)
            _won_statuses = {"Completed", "R&M"}
            _trend_df = dash_df[dash_df["status"].isin(_won_statuses)].copy()
            _month_won = {}
            _month_cost = {}
            for _, _tr in _trend_df.iterrows():
                _ev = str(_tr.get("end", "") or "").strip()
                for _fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                    try:
                        _ep = datetime.strptime(_ev, _fmt)
                        _mk = _ep.strftime("%b %Y")
                        _month_won[_mk]  = _month_won.get(_mk, 0) + 1
                        _cs = pd.to_numeric(_tr.get("cost_saved", 0), errors="coerce") or 0
                        _month_cost[_mk] = _month_cost.get(_mk, 0) + float(_cs)
                        break
                    except ValueError:
                        pass

            if _month_won:
                _all_months = sorted(
                    set(list(_month_won.keys()) + list(_month_cost.keys())),
                    key=lambda x: datetime.strptime(x, "%b %Y")
                )[-10:]
                _won_vals  = [_month_won.get(m, 0)  for m in _all_months]
                _cost_vals = [_month_cost.get(m, 0) for m in _all_months]

                _trend_fig = go.Figure()

                # Line 1 — Projects Won (bars + line)
                _trend_fig.add_trace(go.Bar(
                    x=_all_months, y=_won_vals,
                    name="Projects Won",
                    marker_color="#C7D2FE",
                    opacity=0.6,
                    yaxis="y1",
                ))
                _trend_fig.add_trace(go.Scatter(
                    x=_all_months, y=_won_vals,
                    mode="lines+markers",
                    name="Projects Won",
                    line=dict(color="#4F46E5", width=2.5),
                    marker=dict(size=7, color="#4F46E5"),
                    yaxis="y1",
                    showlegend=False,
                ))

                # Line 2 — Cost Saved (right axis, only if data exists)
                if any(v > 0 for v in _cost_vals):
                    _trend_fig.add_trace(go.Scatter(
                        x=_all_months, y=_cost_vals,
                        mode="lines+markers",
                        name="Cost Saved (₹)",
                        line=dict(color="#10B981", width=2, dash="dot"),
                        marker=dict(size=6, color="#10B981"),
                        yaxis="y2",
                    ))
                    _trend_fig.update_layout(
                        yaxis2=dict(
                            overlaying="y", side="right",
                            tickfont=dict(size=8, color="#10B981"),
                            showgrid=False,
                            title=dict(text="Cost Saved (₹)", font=dict(size=8, color="#10B981")),
                        )
                    )

                _trend_fig.update_layout(
                    margin=dict(t=10, b=0, l=0, r=60), height=240,
                    xaxis=dict(tickfont=dict(size=9), tickangle=-30),
                    yaxis=dict(
                        tickfont=dict(size=9),
                        title=dict(text="Projects", font=dict(size=8)),
                        gridcolor="#F1F5F9",
                    ),
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    legend=dict(font=dict(size=9), orientation="h", y=-0.25),
                    barmode="overlay",
                )
                st.plotly_chart(_trend_fig, use_container_width=True)
            else:
                st.info("No completed project end dates found. Update end dates in Projects tab to see the trend.")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── CLIENT PORTFOLIO SCORECARD + TOP PROJECTS BY VALUE ───────────────────────
        _sc_col, _tp_col = st.columns([1.1, 1])

        with _sc_col:
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                    'text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px">'
                    'Client Portfolio Scorecard</div>',
                    unsafe_allow_html=True
                )
                _score_clients = sorted(dash_df["client"].dropna().unique()) if "client" in dash_df.columns else []
                _sc_html = (
                    '<table style="width:100%;border-collapse:collapse">'
                    '<thead><tr>'
                    '<th style="font-size:9px;color:#94A3B8;font-weight:700;text-transform:uppercase;'
                    'padding:5px 8px;border-bottom:2px solid #F1F5F9;text-align:left">Client</th>'
                    '<th style="font-size:9px;color:#94A3B8;font-weight:700;text-transform:uppercase;'
                    'padding:5px 8px;border-bottom:2px solid #F1F5F9;text-align:center">Active</th>'
                    '<th style="font-size:9px;color:#94A3B8;font-weight:700;text-transform:uppercase;'
                    'padding:5px 8px;border-bottom:2px solid #F1F5F9;text-align:center">Done</th>'
                    '<th style="font-size:9px;color:#94A3B8;font-weight:700;text-transform:uppercase;'
                    'padding:5px 8px;border-bottom:2px solid #F1F5F9;text-align:center">Deals</th>'
                    '<th style="font-size:9px;color:#94A3B8;font-weight:700;text-transform:uppercase;'
                    'padding:5px 8px;border-bottom:2px solid #F1F5F9;text-align:center">Health</th>'
                    '</tr></thead><tbody>'
                )
                for _sc_client in _score_clients:
                    _cdf = dash_df[dash_df["client"] == _sc_client]
                    _c_active   = int(_cdf["status"].isin(["In Progress","UAT","R&M"]).sum())
                    _c_done     = int((_cdf["status"] == "Completed").sum())
                    _c_disc     = int((_cdf["status"] == "Discontinued").sum())
                    _c_pipeline = int(_cdf["status"].isin(["Presales","POC","PDD","Internal POC","External POC"]).sum())
                    _c_total    = len(_cdf)
                    # Health: ratio of active+done vs discontinued
                    _health_score = round(((_c_active + _c_done) / max(_c_total, 1)) * 100)
                    if _health_score >= 75:
                        _hcolor, _hlabel = "#10B981", "Good"
                    elif _health_score >= 50:
                        _hcolor, _hlabel = "#F59E0B", "Fair"
                    else:
                        _hcolor, _hlabel = "#EF4444", "Weak"
                    _sc_html += (
                        f'<tr style="border-bottom:1px solid #F8FAFC">'
                        f'<td style="font-size:11px;font-weight:700;color:#1F3B4D;padding:7px 8px">{_sc_client}</td>'
                        f'<td style="text-align:center;padding:7px 8px">'
                        f'<span style="font-size:12px;font-weight:800;color:#3B82F6">{_c_active}</span></td>'
                        f'<td style="text-align:center;padding:7px 8px">'
                        f'<span style="font-size:12px;font-weight:800;color:#10B981">{_c_done}</span></td>'
                        f'<td style="text-align:center;padding:7px 8px">'
                        f'<span style="font-size:12px;font-weight:800;color:#0EA5E9">{_c_pipeline}</span></td>'
                        f'<td style="text-align:center;padding:7px 8px">'
                        f'<span style="font-size:10px;font-weight:700;padding:2px 9px;border-radius:20px;'
                        f'background:{_hcolor}20;color:{_hcolor};border:1px solid {_hcolor}40">{_hlabel}</span>'
                        f'</td></tr>'
                    )
                _sc_html += '</tbody></table>'
                st.markdown(_sc_html, unsafe_allow_html=True)
                st.markdown(
                    '<div style="display:flex;gap:14px;margin-top:10px">'
                    '<span style="font-size:9px;color:#64748B"><b style="color:#3B82F6">Active</b> = In Progress + UAT + R&amp;M</span>'
                    '<span style="font-size:9px;color:#64748B"><b style="color:#10B981">Done</b> = Completed</span>'
                    '<span style="font-size:9px;color:#64748B"><b style="color:#0EA5E9">Deals</b> = Presales + POC + PDD</span>'
                    '</div>',
                    unsafe_allow_html=True
                )

        with _tp_col:
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                    'text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px">'
                    'Top Projects by Value Delivered</div>',
                    unsafe_allow_html=True
                )
                # Prefer hours_saved, fall back to cost_saved
                _val_df = dash_df.copy()
                _val_df["_hrs"]  = pd.to_numeric(_val_df.get("hours_saved",  pd.Series(dtype=float)), errors="coerce").fillna(0)
                _val_df["_cost"] = pd.to_numeric(_val_df.get("cost_saved",   pd.Series(dtype=float)), errors="coerce").fillna(0)
                _val_df["_roi"]  = pd.to_numeric(_val_df.get("roi_pct",      pd.Series(dtype=float)), errors="coerce").fillna(0)
                _val_df["_score"] = _val_df["_hrs"] + (_val_df["_cost"] / 1000) + _val_df["_roi"]
                _top_proj = _val_df[_val_df["_score"] > 0].nlargest(8, "_score")

                if not _top_proj.empty:
                    _tp_names  = [str(n)[:30] + ("…" if len(str(n)) > 30 else "") for n in _top_proj["name"]]
                    _tp_scores = _top_proj["_hrs"].tolist()
                    _tp_labels = [
                        f"{int(h)}h saved" if h > 0 else
                        (f"₹{int(c):,}" if c > 0 else f"{int(r)}% ROI")
                        for h, c, r in zip(_top_proj["_hrs"], _top_proj["_cost"], _top_proj["_roi"])
                    ]
                    _tp_colors = [
                        STATUS_STYLES.get(str(s), {"dot": "#94A3B8"})["dot"]
                        for s in _top_proj["status"]
                    ]
                    _tp_fig = go.Figure(go.Bar(
                        y=_tp_names, x=_tp_scores,
                        orientation="h",
                        marker=dict(color=_tp_colors, opacity=0.85),
                        text=_tp_labels,
                        textposition="outside",
                        textfont=dict(size=9),
                    ))
                    _tp_fig.update_layout(
                        margin=dict(t=0, b=0, l=0, r=70), height=260,
                        xaxis=dict(visible=False),
                        yaxis=dict(tickfont=dict(size=9), autorange="reversed"),
                        plot_bgcolor="rgba(0,0,0,0)",
                        paper_bgcolor="rgba(0,0,0,0)",
                    )
                    st.plotly_chart(_tp_fig, use_container_width=True)
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;margin-top:4px">'
                        'Bar length = hours saved. Label shows hours saved / cost saved / ROI %.</div>',
                        unsafe_allow_html=True
                    )
                else:
                    st.info("Add hours_saved or cost_saved data in the Projects tab to see top performers.")

        if role not in ("sales",):
            st.markdown("<br>", unsafe_allow_html=True)

            # ── STATUS FILTER (segmented control above project list) ──────────────
            _seg_base = ["All", "In Progress", "Completed", "R&M", "UAT", "POC"]
            _seg_counts = {}
            for _s in _seg_base:
                if _s == "All":
                    _seg_counts[_s] = len(dash_df)
                elif _s == "POC":
                    _seg_counts[_s] = int(dash_df["status"].str.contains("POC", na=False).sum()) if "status" in dash_df.columns else 0
                elif _s == "R&M":
                    _seg_counts[_s] = int(dash_df["status"].str.contains("R&M|Maintenance", na=False).sum()) if "status" in dash_df.columns else 0
                else:
                    _seg_counts[_s] = int(dash_df["status"].str.contains(_s, na=False).sum()) if "status" in dash_df.columns else 0
            _seg_labels = [f"{_s}  ({_seg_counts[_s]})" for _s in _seg_base]
            _cur_label = "All" if st.session_state.dash_slicer is None else st.session_state.dash_slicer
            _cur_seg_label = next((l for l in _seg_labels if l.startswith(_cur_label)), _seg_labels[0])
            _sel_seg = st.radio("Filter", _seg_labels, index=_seg_labels.index(_cur_seg_label),
                                horizontal=True, label_visibility="collapsed", key="dash_seg_radio")
            _sel_seg_base = _sel_seg.split("  (")[0].strip()
            _new_slicer = None if _sel_seg_base == "All" else _sel_seg_base
            if _new_slicer != st.session_state.dash_slicer:
                st.session_state.dash_slicer = _new_slicer
                st.rerun()

            # ── PROJECT DETAIL PANEL (slicer narrows the view) ───────────────────
            _detail_key = st.session_state.dash_slicer

            if _detail_key is not None:
                key = _detail_key
                if key == "__active_dev__":
                    _ad_mask = dash_df["status"].isin({"In Progress"})
                    if "is_active" in dash_df.columns:
                        _ad_mask = _ad_mask & (
                            ~dash_df["is_active"].astype(str).str.strip().str.lower().isin(["false","0","no"])
                        )
                    sliced, slicer_label = dash_df[_ad_mask], "Active Development"
                elif key == "__new__":
                    new_mask = dash_df["is_new"].astype(str).str.lower().isin(["true","1","yes"]) if "is_new" in dash_df.columns else pd.Series([False]*len(dash_df))
                    sliced, slicer_label = dash_df[new_mask], "New Added"
                elif key == "POC":
                    sliced, slicer_label = dash_df[dash_df["status"].str.contains("POC", na=False)], "POC (Internal + External)"
                else:
                    sliced, slicer_label = dash_df[dash_df["status"].str.contains(key, na=False)], key
                _dev_keys = {"In Progress", "PDD", "Important"}
                if key in _dev_keys and "is_active" in sliced.columns:
                    sliced = sliced[~sliced["is_active"].astype(str).str.strip().str.lower().isin(["false","0","no"])]
            else:
                sliced = dash_df
                _cl = st.session_state.dash_client_filter
                slicer_label = f"All — {_cl}" if _cl != "All" else "All Projects"

            emp_map = {}
            for _, row in sliced.iterrows():
                for n in str(row.get("employee","")).replace("&",",").split(","):
                    n = n.strip()
                    if not n: continue
                    if n not in emp_map: emp_map[n] = {"projects":[], "clients":set()}
                    emp_map[n]["projects"].append(row["name"])
                    emp_map[n]["clients"].add(str(row.get("client","")))
            team_list = sorted(emp_map.items(), key=lambda x: -len(x[1]["projects"]))

            st.markdown("<br>", unsafe_allow_html=True)
            hc1, hc2 = st.columns([5, 1])
            _style_key = _detail_key if _detail_key not in [None, "__new__", "POC", "__active_dev__"] else (
                "In Progress" if _detail_key == "__active_dev__" else
                "Completed"   if _detail_key == "__new__"        else
                "Internal POC" if _detail_key == "POC"           else "R&M"
            )
            hc1.markdown(f"""
            <div style="display:flex;align-items:center;gap:12px;padding:10px 16px;
              background:#fff;border:1px solid #E2E8F0;border-radius:10px">
              {badge_html(slicer_label if slicer_label in STATUS_STYLES else "R&M")}
              <span style="color:#64748B;font-size:12px;font-weight:500">
                <b style="color:#1F3B4D">{len(sliced)}</b> projects &nbsp;·&nbsp;
                <b style="color:#1F3B4D">{len(team_list)}</b> team members assigned
                &nbsp;·&nbsp; <b style="color:#64748B">{slicer_label}</b></span>
            </div>""".strip(), unsafe_allow_html=True)

            pl, pr = st.columns([1.6, 1])

            with pl:
                with st.container(border=True):
                    st.markdown(f'<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;padding-bottom:8px;border-bottom:1px solid #E2E8F0">Project Details — {len(sliced)} records</div>', unsafe_allow_html=True)
                    if sliced.empty:
                        st.info("No projects in this category.")
                    else:
                        for i, (_, row) in enumerate(sliced.iterrows()):
                            roi_badge  = ""
                            if str(row.get("roi_pct","")).strip():
                                roi_badge = f'<span style="font-size:10px;background:#064E3B;color:#10B981;border-radius:4px;padding:2px 8px;font-weight:800;margin-left:6px">ROI {esc(str(row["roi_pct"]))}%</span>'
                            new_badge  = '<span style="font-size:9px;background:#10B981;color:#fff;border-radius:4px;padding:1px 5px;font-weight:800;margin-left:4px">NEW</span>' if is_new(row) else ""
                            _lead      = esc(str(row.get("lead","")).strip())
                            _start     = esc(fmt_date(str(row.get("start",""))))
                            _end       = esc(fmt_date(str(row.get("end",""))) or "Ongoing")
                            _due_raw   = str(row.get("due_date","")).strip()
                            _po        = esc(str(row.get("po","")))
                            _desc      = esc(str(row.get("desc","")))
                            meta_spans = [f'<span>{esc(str(row.get("client","")))} </span>']
                            if _lead:
                                meta_spans.append(f'<span>Lead: <b style="color:#3F8E91">{_lead}</b></span>')
                            meta_spans.append(f'<span>{esc(str(row.get("employee","")))} </span>')
                            if _start:
                                meta_spans.append(f'<span>{_start} to {_end}</span>')
                            if _due_raw:
                                _due_d = _parse_dmy(_due_raw)
                                _due_color = "#DC2626" if (_due_d and (_due_d - date.today()).days < 0) else "#92400E" if (_due_d and (_due_d - date.today()).days <= 7) else "#64748B"
                                meta_spans.append(f'<span>Due: <b style="color:{_due_color}">{esc(fmt_date(_due_raw))}</b></span>')
                            if _po:
                                meta_spans.append(f'<span>PO #{_po}</span>')
                            meta_html = "".join(meta_spans)
                            desc_html = f'<div style="font-size:10px;color:#64748B;font-style:italic">{_desc}</div>' if _desc else ""
                            row_bg    = "#fff" if i % 2 == 0 else "#F8FAFC"
                            st.markdown(
                                f'<div class="srow" style="background:{row_bg}">'
                                f'<div style="flex:1">'
                                f'<div style="font-size:12px;font-weight:700;color:#111827;margin-bottom:4px">{esc(str(row.get("name","")))}{new_badge}</div>'
                                f'<div style="display:flex;flex-wrap:wrap;gap:10px;font-size:10px;color:#64748B;margin-bottom:3px">{meta_html}</div>'
                                f'{desc_html}{roi_badge}'
                                f'</div>'
                                f'<div style="flex-shrink:0;margin-left:10px">{badge_html(str(row.get("status","")))}</div>'
                                f'</div>',
                                unsafe_allow_html=True,
                            )

            AVATAR_COLS = [("#1E3A8A","#3B82F6"),("#451A03","#F59E0B"),("#064E3B","#10B981"),
                           ("#1E1B4B","#8B5CF6"),("#7F1D1D","#EF4444"),("#0C4A6E","#06B6D4"),
                           ("#78350F","#F97316"),("#500724","#EC4899")]
            with pr:
                with st.container(border=True):
                    st.markdown('<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;padding-bottom:8px;border-bottom:1px solid #E2E8F0">Team Responsible</div>', unsafe_allow_html=True)
                    if not team_list:
                        st.info("No team members.")
                    else:
                        for i, (name, info) in enumerate(team_list):
                            bg_c, ac = AVATAR_COLS[i % len(AVATAR_COLS)]
                            clients_str = " · ".join(esc(c) for c in sorted(info["clients"]))
                            st.markdown(f"""
                            <div style="display:flex;align-items:center;gap:10px;padding:10px 4px;border-bottom:1px solid #F1F5F9">
                              <div style="width:36px;height:36px;border-radius:10px;flex-shrink:0;
                                background:linear-gradient(135deg,{bg_c},{ac}44);border:1px solid {ac}55;
                                display:flex;align-items:center;justify-content:center;
                                font-size:14px;font-weight:800;color:{ac}">{esc(name[0].upper())}</div>
                              <div style="flex:1;min-width:0">
                                <div style="font-size:13px;font-weight:700;color:#111827">{esc(name)}</div>
                                <div style="font-size:10px;color:#64748B;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">
                                  {clients_str}</div>
                              </div>
                              <div style="width:26px;height:26px;border-radius:7px;flex-shrink:0;
                                background:{ac}22;border:1px solid {ac}44;display:flex;align-items:center;
                                justify-content:center;font-size:13px;font-weight:800;color:{ac};
                                font-family:'JetBrains Mono',monospace">{len(info["projects"])}</div>
                            </div>""".strip(), unsafe_allow_html=True)


        # ── GANTT CHART ───────────────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">'
            '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Project Timeline (Gantt)</span>'
            '</div>',
            unsafe_allow_html=True
        )
        _gantt_df = df.copy()
        _gantt_df["start_dt"] = pd.to_datetime(_gantt_df["start"].apply(_parse_dmy), errors="coerce")
        _gantt_df["end_dt"]   = pd.to_datetime(_gantt_df["end"].apply(_parse_dmy), errors="coerce")
        _gantt_active = _gantt_df[
            _gantt_df["start_dt"].notna() & _gantt_df["end_dt"].notna() &
            _gantt_df["status"].isin(["In Progress","UAT","R&M","PDD","Completed"])
        ].copy()
        if not _gantt_active.empty:
            _gantt_active = _gantt_active.sort_values("start_dt").head(25)
            _gantt_fig = px.timeline(
                _gantt_active,
                x_start="start_dt",
                x_end="end_dt",
                y="name",
                color="status",
                color_discrete_map={s: STATUS_STYLES.get(s, {}).get("dot","#94A3B8") for s in ALL_STATUSES},
                hover_data=["client","employee","roi_pct"],
                title="",
            )
            _gantt_fig.update_layout(
                height=max(300, len(_gantt_active) * 28 + 80),
                margin=dict(l=0, r=10, t=10, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(family="Inter", size=11),
                yaxis=dict(autorange="reversed", title=""),
                xaxis=dict(title=""),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            )
            _gantt_fig.update_traces(marker_line_width=0)
            st.plotly_chart(_gantt_fig, use_container_width=True)
        else:
            st.markdown(
                '<div style="background:#F8FAFC;border:1px dashed #CBD5E1;border-radius:10px;'
                'padding:32px;text-align:center;color:#94A3B8;font-size:13px">'
                '📅 Add start/end dates to projects to see the Gantt chart</div>',
                unsafe_allow_html=True
            )

        # ── TEAM WORKLOAD HEATMAP ─────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:14px;font-weight:800;color:#1F3B4D;margin-bottom:12px">Team Workload Heatmap</div>',
            unsafe_allow_html=True
        )
        _wl_df = df[~df["status"].isin(["Discontinued","Completed"])].copy() if not df.empty else df.copy()
        if not _wl_df.empty and "employee" in _wl_df.columns:
            _wl_rows = []
            for _, _wr in _wl_df.iterrows():
                for _emp_part in str(_wr.get("employee","")).replace("&",",").split(","):
                    _emp_part = _emp_part.strip()
                    if _emp_part:
                        _wl_rows.append({"employee": _emp_part, "project": str(_wr.get("name",""))[:28], "status": str(_wr.get("status",""))})
            if _wl_rows:
                _wl_pivot = pd.DataFrame(_wl_rows)
                _wl_counts = _wl_pivot.groupby("employee").size().reset_index(name="project_count")
                _wl_heatmap = go.Figure(go.Bar(
                    x=_wl_counts["project_count"],
                    y=_wl_counts["employee"],
                    orientation="h",
                    marker=dict(
                        color=_wl_counts["project_count"],
                        colorscale=[[0,"#EFF7F7"],[0.5,"#5FA9AB"],[1,"#162C3B"]],
                        showscale=True,
                        colorbar=dict(title="Projects", tickfont=dict(size=9)),
                    ),
                    text=_wl_counts["project_count"],
                    textposition="outside",
                ))
                _wl_heatmap.update_layout(
                    height=max(280, len(_wl_counts) * 30 + 60),
                    margin=dict(l=10, r=60, t=10, b=10),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(title="Active Projects", showgrid=True, gridcolor="#F1F5F9"),
                    yaxis=dict(autorange="reversed", title=""),
                    font=dict(family="Inter", size=11),
                )
                st.plotly_chart(_wl_heatmap, use_container_width=True)
            else:
                st.info("No active project-employee assignments to display.")
        else:
            st.info("No active projects to build workload heatmap.")

        # ── BURNDOWN / TASK COMPLETION TREND ─────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">'
            '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Task Completion Trend (8 Weeks)</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#ECFDF5;color:#065F46;border:1px solid #A7F3D0">Burndown</span>'
            '</div>',
            unsafe_allow_html=True
        )
        try:
            _bd_tasks = auth.get_all_tasks()
            _bd_today = date.today()
            _bd_weeks = []
            for _wi in range(7, -1, -1):
                _wend   = _bd_today - timedelta(days=_bd_today.weekday()) - timedelta(weeks=_wi - 1) - timedelta(days=1)
                _wstart = _wend - timedelta(days=6)
                _created = sum(
                    1 for t in _bd_tasks
                    if t.get("created_at") and
                    _wstart <= datetime.fromisoformat(str(t["created_at"])[:10]).date() <= _wend
                )
                _done = sum(
                    1 for t in _bd_tasks
                    if t.get("status") == "Completed" and t.get("updated_at") and
                    _wstart <= datetime.fromisoformat(str(t["updated_at"])[:10]).date() <= _wend
                )
                _bd_weeks.append({
                    "week": _wstart.strftime("%d %b"),
                    "created": _created,
                    "completed": _done,
                })
            if any(w["created"] > 0 or w["completed"] > 0 for w in _bd_weeks):
                _bd_df = pd.DataFrame(_bd_weeks)
                _bd_fig = go.Figure()
                _bd_fig.add_trace(go.Bar(
                    name="Created", x=_bd_df["week"], y=_bd_df["created"],
                    marker_color="#3B82F6", opacity=0.7,
                ))
                _bd_fig.add_trace(go.Bar(
                    name="Completed", x=_bd_df["week"], y=_bd_df["completed"],
                    marker_color="#10B981", opacity=0.85,
                ))
                _bd_fig.update_layout(
                    barmode="group",
                    height=280,
                    margin=dict(l=10, r=10, t=10, b=10),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
                    xaxis=dict(tickfont=dict(size=10)),
                    yaxis=dict(title="Tasks", tickfont=dict(size=10), gridcolor="#F1F5F9"),
                    font=dict(family="Inter", size=11),
                )
                st.plotly_chart(_bd_fig, use_container_width=True)
            else:
                st.markdown(
                    '<div style="background:#F8FAFC;border:1px dashed #CBD5E1;border-radius:10px;'
                    'padding:32px;text-align:center;color:#94A3B8;font-size:13px">'
                    '📉 No task activity in the past 8 weeks. Assign and complete tasks to see trends.</div>',
                    unsafe_allow_html=True
                )
        except Exception:
            st.info("Task data unavailable for burndown chart.")

    # ══════════════════════════════════════════════════════════════════════════════
    # TAB: PROJECTS — EMPLOYEE VIEW (assigned Worksoft projects only)
    # ══════════════════════════════════════════════════════════════════════════════

    with _dt_ws:
        # ── Worksoft data ─────────────────────────────────────────────────────────
        _wsd_df = df[df["proj_type"].fillna("").str.strip() == "Worksoft"] if not df.empty else pd.DataFrame(columns=df.columns if not df.empty else [])
        _wsd_all_hours = auth.get_all_worksoft_total_hours()

        _wsd_proj_names_all = (["All Projects"] + sorted(_wsd_df["name"].dropna().tolist())) if not _wsd_df.empty else ["All Projects"]

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Section header ────────────────────────────────────────────────────────
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">'
            '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Worksoft Portfolio</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#162C3B;color:#5FA9AB;border:1px solid #3F8E9140">Live Dashboard</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#EFF6FF;color:#3B82F6;border:1px solid #BFDBFE">Worksoft Projects</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        # ── Project filter ─────────────────────────────────────────────────────────
        _wsd_proj_filter = st.selectbox(
            "Filter by Project",
            _wsd_proj_names_all,
            key="wsd_proj_filter",
            label_visibility="collapsed",
        )
        if _wsd_proj_filter != "All Projects":
            _wsd_df = _wsd_df[_wsd_df["name"] == _wsd_proj_filter].copy()

        # ── KPI computations (after filter so metrics reflect selection) ──────────
        _wsd_total_projs  = len(_wsd_df)
        _WSD_ACTIVE_ST    = {"In Progress", "UAT"}
        _wsd_active_projs = (
            len(_wsd_df[_wsd_df["status"].fillna("").isin(_WSD_ACTIVE_ST)])
            if not _wsd_df.empty else 0
        )

        _wsd_emp_hrs      = {}
        _wsd_proj_hrs     = {}
        _wsd_ent_count    = {}
        _wsd_monthly      = {}
        _wsd_all_punches  = []
        _wsd_emp_proj_hrs = {}  # {emp_name: {proj_name: hours}}

        if not _wsd_df.empty:
            for _, _wdr in _wsd_df.iterrows():
                _wdpid     = int(float(str(_wdr.get("id", 0) or 0)))
                _wdpname   = str(_wdr.get("name", ""))
                _wdpunches = auth.get_project_punches(_wdpid)
                _wsd_proj_hrs[_wdpname] = sum(p["hours_worked"] for p in _wdpunches)
                for _p in _wdpunches:
                    _n = _p["user_name"]
                    _wsd_emp_hrs[_n]   = _wsd_emp_hrs.get(_n, 0)   + _p["hours_worked"]
                    _wsd_ent_count[_n] = _wsd_ent_count.get(_n, 0) + 1
                    _wd = str(_p.get("work_date", "")).strip()
                    if _wd and len(_wd) >= 7:
                        _ym = _wd[:7]
                        _wsd_monthly[_ym] = _wsd_monthly.get(_ym, 0) + _p["hours_worked"]
                    _wsd_all_punches.append(_p)
                    _ep = _wsd_emp_proj_hrs.setdefault(_n, {})
                    _ep[_wdpname] = _ep.get(_wdpname, 0.0) + _p["hours_worked"]

        _wsd_total_emps   = len(_wsd_emp_hrs)
        _wsd_total_logged = sum(_wsd_proj_hrs.values())

        # ── Hero KPI dark gradient card ───────────────────────────────────────────
        _wsd_kpi_html = ""
        for _wv, _wl, _wc in [
            (str(_wsd_total_projs),        "Total Projects",  "#5FA9AB"),
            (str(_wsd_active_projs),       "Active Projects", "#A78BFA"),
            (str(_wsd_total_emps),         "Team Members",    "#6EE7B7"),
            (f"{_wsd_total_logged:.0f}h",  "Hours Logged",    "#FCD34D"),
        ]:
            _wsd_kpi_html += (
                f'<div style="background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);'
                f'border-radius:12px;padding:16px;text-align:center;animation:fadeInUp 0.5s ease both">'
                f'<div style="font-size:26px;font-weight:900;color:{_wc};line-height:1">{_wv}</div>'
                f'<div style="font-size:9px;color:#94A3B8;margin-top:5px;font-weight:600;'
                f'text-transform:uppercase;letter-spacing:.6px">{_wl}</div>'
                f'</div>'
            )
        st.markdown(
            '<div style="background:linear-gradient(135deg,#0F172A 0%,#1E293B 60%,#0F2027 100%);'
            'border-radius:18px;padding:28px 32px;margin-bottom:16px;'
            'border:1px solid rgba(95,169,171,0.25);position:relative;overflow:hidden;'
            'animation:scaleIn 0.4s ease both">'
            '<div style="position:absolute;top:-50px;right:-50px;width:220px;height:220px;'
            'border-radius:50%;background:rgba(95,169,171,0.07);pointer-events:none"></div>'
            '<div style="position:absolute;bottom:-30px;left:30px;width:140px;height:140px;'
            'border-radius:50%;background:rgba(139,92,246,0.05);pointer-events:none"></div>'
            '<div style="display:flex;align-items:center;gap:16px;margin-bottom:20px">'
            '<div style="width:54px;height:54px;border-radius:14px;'
            'background:linear-gradient(135deg,#0EA5E9,#0369A1);display:flex;align-items:center;'
            'justify-content:center;font-size:28px;font-weight:900;color:#fff;'
            'font-family:Manrope,sans-serif;box-shadow:0 4px 16px rgba(14,165,233,0.35)">W</div>'
            '<div>'
            '<div style="font-size:22px;font-weight:900;color:#fff;font-family:Manrope,sans-serif;'
            'letter-spacing:-0.4px;line-height:1.1">Worksoft Portfolio</div>'
            '<div style="font-size:10px;color:#5FA9AB;font-weight:600;letter-spacing:1.2px;'
            'text-transform:uppercase;margin-top:3px">Project Time Tracking · Team Management</div>'
            '</div>'
            '</div>'
            f'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px">'
            f'{_wsd_kpi_html}'
            f'</div>'
            '</div>',
            unsafe_allow_html=True,
        )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Status Distribution donut ─────────────────────────────────────────────
        _, _wpd_col, _ = st.columns([1, 2, 1])
        with _wpd_col:
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                    'letter-spacing:.8px;margin-bottom:8px">Status Distribution</div>',
                    unsafe_allow_html=True,
                )
                if not _wsd_df.empty and "status" in _wsd_df.columns:
                    _wsd_sc = _wsd_df["status"].fillna("Unknown").value_counts()
                    _wsd_pie_map = {
                        "In Progress":"#3B82F6","UAT":"#8B5CF6",
                        "PDD":"#0EA5E9","Completed":"#10B981","Discontinued":"#EF4444",
                    }
                    _wsd_pcols = [_wsd_pie_map.get(s, "#94A3B8") for s in _wsd_sc.index]
                    _wsd_pie = go.Figure(go.Pie(
                        labels=list(_wsd_sc.index),
                        values=list(_wsd_sc.values),
                        hole=0.6,
                        marker=dict(colors=_wsd_pcols, line=dict(color="#fff", width=2)),
                        textinfo="label+percent",
                        textfont=dict(size=10),
                        hovertemplate="%{label}: %{value}<extra></extra>",
                    ))
                    _wsd_pie.update_layout(
                        margin=dict(t=10, b=10, l=10, r=10), height=260,
                        showlegend=False,
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                        annotations=[dict(
                            text=f"<b>{_wsd_total_projs}</b><br>Projects",
                            x=0.5, y=0.5, showarrow=False,
                            font=dict(size=14, color="#1F3B4D"),
                        )],
                    )
                    st.plotly_chart(_wsd_pie, use_container_width=True)
                else:
                    st.markdown(
                        '<div style="text-align:center;padding:50px;color:#94A3B8;font-size:11px">'
                        'No projects found</div>',
                        unsafe_allow_html=True,
                    )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Bar charts: Hours by Project + Hours by Employee ──────────────────────
        _wdc1, _wdc2 = st.columns(2)

        with _wdc1:
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                    'letter-spacing:.8px;margin-bottom:8px">Hours Logged by Project</div>',
                    unsafe_allow_html=True,
                )
                if _wsd_proj_hrs:
                    _sp = sorted(_wsd_proj_hrs.items(), key=lambda x: x[1], reverse=True)[:10]
                    _ph_fig = go.Figure(go.Bar(
                        x=[p[1] for p in _sp],
                        y=[p[0] for p in _sp],
                        orientation="h",
                        marker=dict(
                            color=[p[1] for p in _sp],
                            colorscale=[[0, "rgba(14,165,233,0.18)"], [1, "#0EA5E9"]],
                            showscale=False,
                        ),
                        text=[f"{p[1]:.1f}h" for p in _sp],
                        textposition="outside",
                        textfont=dict(size=10),
                    ))
                    _ph_fig.update_layout(
                        margin=dict(t=0, b=0, l=10, r=50), height=260,
                        xaxis=dict(visible=False),
                        yaxis=dict(tickfont=dict(size=10), autorange="reversed" if len(_sp) > 1 else True),
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    )
                    st.plotly_chart(_ph_fig, use_container_width=True)
                else:
                    st.markdown(
                        '<div style="text-align:center;padding:50px;color:#94A3B8;font-size:11px">'
                        'No project hours logged yet</div>',
                        unsafe_allow_html=True,
                    )

        with _wdc2:
            with st.container(border=True):
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                    'letter-spacing:.8px;margin-bottom:8px">Hours by Employee</div>',
                    unsafe_allow_html=True,
                )
                if _wsd_emp_hrs:
                    _se = sorted(_wsd_emp_hrs.items(), key=lambda x: x[1], reverse=True)[:10]
                    _eh_fig = go.Figure(go.Bar(
                        x=[e[1] for e in _se],
                        y=[e[0] for e in _se],
                        orientation="h",
                        marker=dict(
                            color=[e[1] for e in _se],
                            colorscale=[[0, "rgba(16,185,129,0.18)"], [1, "#10B981"]],
                            showscale=False,
                        ),
                        text=[f"{e[1]:.1f}h" for e in _se],
                        textposition="outside",
                        textfont=dict(size=10),
                    ))
                    _eh_fig.update_layout(
                        margin=dict(t=0, b=0, l=10, r=50), height=260,
                        xaxis=dict(visible=False),
                        yaxis=dict(tickfont=dict(size=10), autorange="reversed" if len(_se) > 1 else True),
                        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    )
                    st.plotly_chart(_eh_fig, use_container_width=True)
                else:
                    st.markdown(
                        '<div style="text-align:center;padding:50px;color:#94A3B8;font-size:11px">'
                        'No employee hours yet</div>',
                        unsafe_allow_html=True,
                    )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Projects table (styled like RPA active portfolio) ─────────────────────
        st.markdown(
            '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
            'letter-spacing:.8px;margin-bottom:12px">Active Portfolio · Worksoft Projects</div>',
            unsafe_allow_html=True,
        )
        _wsd_stcols = {
            "In Progress":"#3B82F6","UAT":"#8B5CF6",
            "PDD":"#0EA5E9","Completed":"#10B981","Discontinued":"#EF4444",
        }
        if not _wsd_df.empty:
            _wpt_html = (
                '<div style="background:#fff;border:1px solid #E2E8F0;border-radius:14px;'
                'overflow:hidden;animation:fadeInUp 0.5s ease both">'
                '<div style="display:grid;grid-template-columns:0.5fr 2fr 1.2fr 1.2fr 1fr 0.9fr 0.9fr 1fr 1fr;'
                'background:#F8FAFC;padding:10px 16px;border-bottom:2px solid #E2E8F0">'
            )
            for _wph_lbl in ["#", "Project", "Client", "Lead", "Status", "Start", "End", "Logged", "Allocated"]:
                _wpt_html += (
                    f'<div style="font-size:9px;font-weight:700;color:#94A3B8;'
                    f'text-transform:uppercase;letter-spacing:.6px">{_wph_lbl}</div>'
                )
            _wpt_html += '</div>'
            for _wi, (_, _wprow) in enumerate(_wsd_df.iterrows()):
                _wp_pid   = int(float(str(_wprow.get("id", 0) or 0)))
                _wp_alloc = float(_wprow.get("allocated_hours") or 0)
                _wp_log   = _wsd_all_hours.get(_wp_pid, 0.0)
                _wp_st    = str(_wprow.get("status", ""))
                _wp_stc   = _wsd_stcols.get(_wp_st, "#94A3B8")
                _wp_rc    = "#DC2626" if _wp_alloc > 0 and _wp_log >= _wp_alloc else "#16A34A" if _wp_alloc > 0 else "#64748B"
                _wp_bg    = "#FAFAFA" if _wi % 2 == 1 else "#FFFFFF"
                _wp_start = fmt_date(str(_wprow.get("start", "") or "")) or "—"
                _wp_end   = fmt_date(str(_wprow.get("end", "") or "")) or "—"
                _wpt_html += (
                    f'<div style="display:grid;grid-template-columns:0.5fr 2fr 1.2fr 1.2fr 1fr 0.9fr 0.9fr 1fr 1fr;'
                    f'padding:10px 16px;border-bottom:1px solid #F1F5F9;background:{_wp_bg};align-items:center">'
                    f'<div style="font-size:10px;color:#94A3B8">{esc(str(_wprow.get("id","")))}</div>'
                    f'<div style="font-size:12px;font-weight:700;color:#1F3B4D">{esc(str(_wprow.get("name","")))}</div>'
                    f'<div style="font-size:11px;color:#374151">{esc(str(_wprow.get("client","")))}</div>'
                    f'<div style="font-size:11px;font-weight:600;color:#3F8E91">{esc(str(_wprow.get("lead","")))}</div>'
                    f'<div><span style="font-size:10px;font-weight:700;color:{_wp_stc};'
                    f'background:{_wp_stc}15;padding:2px 8px;border-radius:12px;border:1px solid {_wp_stc}30">'
                    f'{esc(_wp_st)}</span></div>'
                    f'<div style="font-size:11px;color:#475569">{esc(_wp_start)}</div>'
                    f'<div style="font-size:11px;color:#475569">{esc(_wp_end)}</div>'
                    f'<div style="font-size:12px;font-weight:700;color:{_wp_rc}">{_wp_log:.1f}h</div>'
                    f'<div style="font-size:11px;color:#64748B">{"&mdash;" if _wp_alloc == 0 else f"{_wp_alloc:.0f}h"}</div>'
                    f'</div>'
                )
            _wpt_html += '</div>'
            st.markdown(_wpt_html, unsafe_allow_html=True)

        # ── Individual Hours by Project & Employee ────────────────────────────────
        if _wsd_all_punches:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(
                '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                'letter-spacing:.8px;margin-bottom:12px">Individual Hours · Per Project Per Employee</div>',
                unsafe_allow_html=True,
            )
            # build {project_name: {employee: hours}}
            _iph_data = {}
            for _ip in _wsd_all_punches:
                _ipn = _ip.get("project_name", "")
                _ien = _ip.get("user_name", "")
                _iph_data.setdefault(_ipn, {})
                _iph_data[_ipn][_ien] = _iph_data[_ipn].get(_ien, 0.0) + float(_ip.get("hours_worked", 0))
            _iph_pal = ["#3F8E91","#8B5CF6","#F59E0B","#10B981","#3B82F6","#EC4899","#0EA5E9","#EF4444"]
            for _ipname, _ip_emp_hrs in _iph_data.items():
                _ip_total = sum(_ip_emp_hrs.values())
                _ip_max   = max(_ip_emp_hrs.values(), default=1)
                _iph_html = (
                    f'<div style="background:#fff;border:1px solid #E2E8F0;border-radius:14px;'
                    f'overflow:hidden;margin-bottom:14px;animation:fadeInUp 0.5s ease both">'
                    f'<div style="background:#F8FAFC;padding:10px 16px;border-bottom:2px solid #E2E8F0;'
                    f'display:flex;justify-content:space-between;align-items:center">'
                    f'<span style="font-size:12px;font-weight:800;color:#1F3B4D">{esc(_ipname)}</span>'
                    f'<span style="font-size:11px;font-weight:700;color:#0369A1">{_ip_total:.1f}h total</span>'
                    f'</div>'
                    f'<div style="display:grid;grid-template-columns:2fr 1fr 1fr 3fr;'
                    f'background:#F8FAFC;padding:8px 16px;border-bottom:1px solid #E2E8F0">'
                )
                for _lbl in ["Employee", "Hours", "Share", "Breakdown"]:
                    _iph_html += (
                        f'<div style="font-size:9px;font-weight:700;color:#94A3B8;'
                        f'text-transform:uppercase;letter-spacing:.6px">{_lbl}</div>'
                    )
                _iph_html += '</div>'
                for _iei, (_emp, _ehrs) in enumerate(sorted(_ip_emp_hrs.items(), key=lambda x: x[1], reverse=True)):
                    _share_pct = (_ehrs / max(_ip_total, 1)) * 100
                    _bar_pct   = (_ehrs / max(_ip_max, 1)) * 100
                    _ec        = _iph_pal[_iei % len(_iph_pal)]
                    _ebg       = "#FAFAFA" if _iei % 2 == 1 else "#FFFFFF"
                    _einit     = esc(_emp[0].upper()) if _emp else "?"
                    _iph_html += (
                        f'<div style="display:grid;grid-template-columns:2fr 1fr 1fr 3fr;'
                        f'padding:10px 16px;border-bottom:1px solid #F1F5F9;background:{_ebg};align-items:center">'
                        f'<div style="display:flex;align-items:center;gap:8px">'
                        f'<div style="width:28px;height:28px;border-radius:50%;background:{_ec}20;'
                        f'border:1px solid {_ec}40;display:flex;align-items:center;justify-content:center;'
                        f'font-size:11px;font-weight:800;color:{_ec}">{_einit}</div>'
                        f'<span style="font-size:12px;font-weight:600;color:#1F3B4D">{esc(_emp)}</span>'
                        f'</div>'
                        f'<div style="font-size:13px;font-weight:700;color:{_ec}">{_ehrs:.1f}h</div>'
                        f'<div style="font-size:11px;color:#64748B">{_share_pct:.0f}%</div>'
                        f'<div style="background:#F1F5F9;border-radius:6px;height:8px;overflow:hidden">'
                        f'<div style="width:{_bar_pct:.0f}%;background:{_ec};height:8px;border-radius:6px;'
                        f'transition:width 0.8s cubic-bezier(.4,0,.2,1)"></div>'
                        f'</div>'
                        f'</div>'
                    )
                _iph_html += '</div>'
                st.markdown(_iph_html, unsafe_allow_html=True)
        else:
            st.markdown(
                '<div style="text-align:center;padding:24px;color:#94A3B8;font-size:11px;'
                'background:#F8FAFC;border-radius:12px;border:1px solid #E2E8F0">'
                'No Worksoft projects found. Add projects with type "Worksoft" to populate this panel.</div>',
                unsafe_allow_html=True,
            )

        # ── Employee Hours table ──────────────────────────────────────────────────
        if _wsd_emp_hrs:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(
                '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                'letter-spacing:.8px;margin-bottom:12px">Team · Hours by Employee</div>',
                unsafe_allow_html=True,
            )
            _weh_pal = ["#3F8E91","#8B5CF6","#F59E0B","#10B981","#3B82F6","#EC4899","#0EA5E9","#EF4444"]
            _weh_max = max(_wsd_emp_hrs.values(), default=1)
            _weh_html = (
                '<div style="background:#fff;border:1px solid #E2E8F0;border-radius:14px;'
                'overflow:hidden;animation:fadeInUp 0.5s ease both">'
                '<div style="display:grid;grid-template-columns:2fr 1.5fr 1fr 3fr;'
                'background:#F8FAFC;padding:10px 16px;border-bottom:2px solid #E2E8F0">'
            )
            for _wehl in ["Employee", "Total Hours", "Entries", "Activity"]:
                _weh_html += (
                    f'<div style="font-size:9px;font-weight:700;color:#94A3B8;'
                    f'text-transform:uppercase;letter-spacing:.6px">{_wehl}</div>'
                )
            _weh_html += '</div>'
            for _wei, (_en, _eh) in enumerate(sorted(_wsd_emp_hrs.items(), key=lambda x: x[1], reverse=True)):
                _we_pct = round((_eh / max(_weh_max, 1)) * 100)
                _we_c   = _weh_pal[_wei % len(_weh_pal)]
                _we_bg  = "#FAFAFA" if _wei % 2 == 1 else "#FFFFFF"
                _en_init = esc(_en[0].upper()) if _en else "?"
                _weh_html += (
                    f'<div style="display:grid;grid-template-columns:2fr 1.5fr 1fr 3fr;'
                    f'padding:10px 16px;border-bottom:1px solid #F1F5F9;background:{_we_bg};align-items:center">'
                    f'<div style="display:flex;align-items:center;gap:8px">'
                    f'<div style="width:30px;height:30px;border-radius:50%;background:{_we_c}20;'
                    f'border:1px solid {_we_c}40;display:flex;align-items:center;justify-content:center;'
                    f'font-size:12px;font-weight:800;color:{_we_c}">{_en_init}</div>'
                    f'<span style="font-size:12px;font-weight:600;color:#1F3B4D">{esc(_en)}</span>'
                    f'</div>'
                    f'<div style="font-size:13px;font-weight:700;color:#0369A1">{_eh:.1f}h</div>'
                    f'<div style="font-size:11px;color:#64748B">{_wsd_ent_count.get(_en, 0)}</div>'
                    f'<div style="background:#F1F5F9;border-radius:6px;height:8px;overflow:hidden">'
                    f'<div style="width:{_we_pct}%;background:{_we_c};height:8px;border-radius:6px;'
                    f'transition:width 0.8s cubic-bezier(.4,0,.2,1)"></div>'
                    f'</div>'
                    f'</div>'
                )
            _weh_html += '</div>'
            st.markdown(_weh_html, unsafe_allow_html=True)

        # ── Budget Utilisation (animated progress bars) ───────────────────────────
        if not _wsd_df.empty:
            _wsd_alloc_rows = [r for _, r in _wsd_df.iterrows() if float(r.get("allocated_hours") or 0) > 0]
            if _wsd_alloc_rows:
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                    'letter-spacing:.8px;margin-bottom:12px">Budget Utilisation · Hours</div>',
                    unsafe_allow_html=True,
                )
                _wbu_html = '<div style="animation:fadeInUp 0.5s ease both">'
                for _br in _wsd_alloc_rows:
                    _bp_id    = int(float(str(_br.get("id", 0) or 0)))
                    _bp_alloc = float(_br.get("allocated_hours") or 0)
                    _bp_log   = _wsd_all_hours.get(_bp_id, 0.0)
                    _bp_pct   = min((_bp_log / _bp_alloc) * 100, 100)
                    _bp_rem   = _bp_alloc - _bp_log
                    _bp_bar_c = "#DC2626" if _bp_pct >= 100 else ("#F59E0B" if _bp_pct >= 75 else "#10B981")
                    _bp_rem_c = "#DC2626" if _bp_rem <= 0 else "#64748B"
                    _wbu_html += (
                        f'<div style="background:#fff;border:1px solid #E2E8F0;border-radius:12px;'
                        f'padding:14px 18px;margin-bottom:8px;border-left:4px solid {_bp_bar_c}">'
                        f'<div style="display:flex;justify-content:space-between;margin-bottom:6px;align-items:center">'
                        f'<span style="font-size:13px;font-weight:700;color:#1F3B4D">{esc(str(_br.get("name","")))}</span>'
                        f'<span style="font-size:11px;color:{_bp_rem_c};font-weight:600">'
                        f'{"Over by" if _bp_rem < 0 else "Remaining:"} {abs(_bp_rem):.1f}h'
                        f' &nbsp;&middot;&nbsp; '
                        f'<span style="color:#64748B;font-weight:400">{_bp_log:.1f} / {_bp_alloc:.0f}h</span>'
                        f'</span>'
                        f'</div>'
                        f'<div style="background:#E2E8F0;border-radius:6px;height:10px;overflow:hidden">'
                        f'<div style="width:{_bp_pct:.0f}%;background:{_bp_bar_c};height:10px;border-radius:6px;'
                        f'transition:width 0.8s cubic-bezier(.4,0,.2,1)"></div>'
                        f'</div>'
                        f'<div style="font-size:9px;color:#94A3B8;margin-top:4px">'
                        f'{_bp_pct:.0f}% used of {_bp_alloc:.0f}h budget</div>'
                        f'</div>'
                    )
                _wbu_html += '</div>'
                st.markdown(_wbu_html, unsafe_allow_html=True)

        # ── Monthly Hours Trend ───────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">'
            '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Monthly Trends</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#F0FDF4;color:#10B981;border:1px solid #BBF7D0">Hours Over Time</span>'
            '</div>',
            unsafe_allow_html=True,
        )
        with st.container(border=True):
            st.markdown(
                '<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;'
                'letter-spacing:.8px;margin-bottom:8px">Hours Logged per Month</div>',
                unsafe_allow_html=True,
            )
            if _wsd_monthly:
                _wmt_sorted = sorted(_wsd_monthly.items())
                _wmt_x = [m[0] for m in _wmt_sorted]
                _wmt_y = [m[1] for m in _wmt_sorted]
                _wmt_fig = go.Figure()
                _wmt_fig.add_trace(go.Scatter(
                    x=_wmt_x, y=_wmt_y,
                    mode="lines+markers+text",
                    line=dict(color="#0EA5E9", width=3, shape="spline"),
                    marker=dict(size=8, color="#0EA5E9", line=dict(color="#fff", width=2)),
                    fill="tozeroy",
                    fillcolor="rgba(14,165,233,0.10)",
                    text=[f"{v:.1f}h" for v in _wmt_y],
                    textposition="top center",
                    textfont=dict(size=10, color="#0369A1"),
                    hovertemplate="%{x}: %{y:.1f}h<extra></extra>",
                ))
                _wmt_fig.update_layout(
                    margin=dict(t=20, b=20, l=20, r=20), height=240,
                    xaxis=dict(tickfont=dict(size=10), gridcolor="#F1F5F9", showgrid=True),
                    yaxis=dict(tickfont=dict(size=10), gridcolor="#F1F5F9",
                               ticksuffix="h", zeroline=False),
                    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    showlegend=False,
                    hovermode="x unified",
                )
                st.plotly_chart(_wmt_fig, use_container_width=True)
            else:
                st.markdown(
                    '<div style="text-align:center;padding:40px;color:#94A3B8;font-size:11px">'
                    'No time entries yet — hours will appear here once employees log work.</div>',
                    unsafe_allow_html=True,
                )

        # ── Employee Insights Line Chart ──────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">'
            '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Employee Insights</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#EFF6FF;color:#3B82F6;border:1px solid #BFDBFE">Line Chart</span>'
            '</div>',
            unsafe_allow_html=True,
        )
        _eic_mode = st.radio(
            "View mode",
            ["By Employee — Projects & Hours", "By Project — Employee Hours"],
            horizontal=True,
            key="wsd_eic_mode",
            label_visibility="collapsed",
        )

        if _eic_mode == "By Employee — Projects & Hours":
            # Employee filter — Worksoft users only (from DB dept + punch history)
            _ws_db_users  = [u["name"] for u in auth.get_all_users()
                             if (u.get("department") or "").strip().lower() == "worksoft"]
            _ws_emp_list  = sorted(set(_ws_db_users) | set(_wsd_emp_proj_hrs.keys()))
            if not _ws_emp_list:
                st.info("No Worksoft employees found. Assign employees to Worksoft projects or set their department to 'Worksoft'.")
            else:
                _sel_emp       = st.selectbox("Select Employee", _ws_emp_list, key="wsd_emp_sel")
                _emp_proj_data = _wsd_emp_proj_hrs.get(_sel_emp, {})
                if not _emp_proj_data:
                    st.info(f"{esc(_sel_emp)} has not logged any hours on Worksoft projects yet.")
                else:
                    _emp_projs = sorted(_emp_proj_data.keys())
                    _emp_hrs   = [_emp_proj_data[p] for p in _emp_projs]
                    _emp_total = sum(_emp_hrs)
                    with st.container(border=True):
                        st.markdown(
                            f'<div style="font-size:22px;font-weight:800;color:#8B5CF6;margin-bottom:14px">'
                            f'{_emp_total:.1f}h '
                            f'<span style="font-size:12px;font-weight:400;color:#64748B">total hours · {esc(_sel_emp)}</span>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                        _eic_fig = go.Figure()
                        _eic_fig.add_trace(go.Scatter(
                            x=_emp_projs,
                            y=_emp_hrs,
                            mode="lines+markers+text",
                            line=dict(color="#8B5CF6", width=3, shape="spline"),
                            marker=dict(size=9, color="#8B5CF6", line=dict(color="#fff", width=2)),
                            fill="tozeroy",
                            fillcolor="rgba(139,92,246,0.10)",
                            text=[f"{h:.1f}h" for h in _emp_hrs],
                            textposition="top center",
                            textfont=dict(size=10, color="#6D28D9"),
                            hovertemplate="%{x}: %{y:.1f}h<extra></extra>",
                        ))
                        _eic_fig.update_layout(
                            title=dict(text=f"Hours per Project · {esc(_sel_emp)}", font=dict(size=13, color="#1F3B4D")),
                            margin=dict(t=40, b=70, l=20, r=20), height=300,
                            xaxis=dict(tickfont=dict(size=10), gridcolor="#F1F5F9", showgrid=True, tickangle=-30),
                            yaxis=dict(tickfont=dict(size=10), gridcolor="#F1F5F9", ticksuffix="h", zeroline=False),
                            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                            showlegend=False,
                            hovermode="x unified",
                        )
                        st.plotly_chart(_eic_fig, use_container_width=True)
                        # Project breakdown table
                        st.markdown(
                            '<div style="font-size:9px;color:#94A3B8;font-weight:700;'
                            'text-transform:uppercase;letter-spacing:.8px;margin:4px 0 6px">'
                            'Project Breakdown</div>',
                            unsafe_allow_html=True,
                        )
                        _ep_hdr = st.columns([3, 1.2, 1])
                        for _eh_lbl, _eh_col in zip(["Project", "Hours", "Share"], _ep_hdr):
                            _eh_col.markdown(f'<div style="{_HDR_STYLE}">{_eh_lbl}</div>', unsafe_allow_html=True)
                        for _epj, _eph in sorted(_emp_proj_data.items(), key=lambda x: x[1], reverse=True):
                            _ep_share = (_eph / max(_emp_total, 1)) * 100
                            _ep_row   = st.columns([3, 1.2, 1])
                            _ep_row[0].markdown(f'<span style="font-size:12px;font-weight:600;color:#1F3B4D">{esc(_epj)}</span>', unsafe_allow_html=True)
                            _ep_row[1].markdown(f'<span style="font-size:12px;font-weight:700;color:#8B5CF6">{_eph:.1f}h</span>', unsafe_allow_html=True)
                            _ep_row[2].markdown(f'<span style="font-size:11px;color:#64748B">{_ep_share:.0f}%</span>', unsafe_allow_html=True)
                        st.markdown(
                            f'<div style="border-top:2px solid #E2E8F0;margin-top:8px;padding-top:8px;'
                            f'font-size:13px;font-weight:800;color:#8B5CF6">Total: {_emp_total:.1f}h</div>',
                            unsafe_allow_html=True,
                        )

        else:  # By Project — Employee Hours
            if not _wsd_proj_hrs:
                st.info("No Worksoft projects with logged hours found.")
            else:
                _proj_emp_list = sorted(_wsd_proj_hrs.keys())
                _sel_proj      = st.selectbox("Select Project", _proj_emp_list, key="wsd_proj_emp_sel")
                _pe_hrs = {
                    _en: _wsd_emp_proj_hrs[_en][_sel_proj]
                    for _en in _wsd_emp_proj_hrs
                    if _wsd_emp_proj_hrs[_en].get(_sel_proj, 0.0) > 0
                }
                if not _pe_hrs:
                    st.info(f"No hours logged for '{esc(_sel_proj)}' yet.")
                else:
                    _pe_emps  = sorted(_pe_hrs.keys())
                    _pe_vals  = [_pe_hrs[e] for e in _pe_emps]
                    _pe_total = sum(_pe_vals)
                    with st.container(border=True):
                        st.markdown(
                            f'<div style="font-size:22px;font-weight:800;color:#10B981;margin-bottom:14px">'
                            f'{_pe_total:.1f}h '
                            f'<span style="font-size:12px;font-weight:400;color:#64748B">total hours · {esc(_sel_proj)}</span>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                        _pec_fig = go.Figure()
                        _pec_fig.add_trace(go.Scatter(
                            x=_pe_emps,
                            y=_pe_vals,
                            mode="lines+markers+text",
                            line=dict(color="#10B981", width=3, shape="spline"),
                            marker=dict(size=9, color="#10B981", line=dict(color="#fff", width=2)),
                            fill="tozeroy",
                            fillcolor="rgba(16,185,129,0.10)",
                            text=[f"{h:.1f}h" for h in _pe_vals],
                            textposition="top center",
                            textfont=dict(size=10, color="#059669"),
                            hovertemplate="%{x}: %{y:.1f}h<extra></extra>",
                        ))
                        _pec_fig.update_layout(
                            title=dict(text=f"Employee Hours · {esc(_sel_proj)}", font=dict(size=13, color="#1F3B4D")),
                            margin=dict(t=40, b=70, l=20, r=20), height=300,
                            xaxis=dict(tickfont=dict(size=10), gridcolor="#F1F5F9", showgrid=True, tickangle=-30),
                            yaxis=dict(tickfont=dict(size=10), gridcolor="#F1F5F9", ticksuffix="h", zeroline=False),
                            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                            showlegend=False,
                            hovermode="x unified",
                        )
                        st.plotly_chart(_pec_fig, use_container_width=True)
                        # Employee breakdown table
                        st.markdown(
                            '<div style="font-size:9px;color:#94A3B8;font-weight:700;'
                            'text-transform:uppercase;letter-spacing:.8px;margin:4px 0 6px">'
                            'Employee Breakdown</div>',
                            unsafe_allow_html=True,
                        )
                        _peh_hdr = st.columns([3, 1.2, 1])
                        for _peh_lbl, _peh_col in zip(["Employee", "Hours", "Share"], _peh_hdr):
                            _peh_col.markdown(f'<div style="{_HDR_STYLE}">{_peh_lbl}</div>', unsafe_allow_html=True)
                        for _pee, _pev in sorted(_pe_hrs.items(), key=lambda x: x[1], reverse=True):
                            _pe_share = (_pev / max(_pe_total, 1)) * 100
                            _pe_row   = st.columns([3, 1.2, 1])
                            _pe_row[0].markdown(f'<span style="font-size:12px;font-weight:600;color:#1F3B4D">{esc(_pee)}</span>', unsafe_allow_html=True)
                            _pe_row[1].markdown(f'<span style="font-size:12px;font-weight:700;color:#10B981">{_pev:.1f}h</span>', unsafe_allow_html=True)
                            _pe_row[2].markdown(f'<span style="font-size:11px;color:#64748B">{_pe_share:.0f}%</span>', unsafe_allow_html=True)
                        st.markdown(
                            f'<div style="border-top:2px solid #E2E8F0;margin-top:8px;padding-top:8px;'
                            f'font-size:13px;font-weight:800;color:#10B981">Total: {_pe_total:.1f}h</div>',
                            unsafe_allow_html=True,
                        )

        # ── Project Health Cards ──────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">'
            '<span style="font-size:14px;font-weight:800;color:#1F3B4D">Project Health</span>'
            '<span style="font-size:10px;font-weight:600;padding:2px 10px;border-radius:20px;'
            'background:#FFF7ED;color:#EA580C;border:1px solid #FED7AA">Budget Status</span>'
            '</div>',
            unsafe_allow_html=True,
        )
        if not _wsd_df.empty:
            _wph_rows = list(_wsd_df.iterrows())
            _wph_chunk = 4
            for _wph_i in range(0, len(_wph_rows), _wph_chunk):
                _wph_batch = _wph_rows[_wph_i:_wph_i + _wph_chunk]
                _wph_cols  = st.columns(len(_wph_batch))
                _wph_pal   = ["#3F8E91","#8B5CF6","#F59E0B","#10B981","#3B82F6","#EC4899"]
                _wph_stmap = {
                    "In Progress":"#3B82F6","UAT":"#8B5CF6","PDD":"#0EA5E9",
                    "Completed":"#10B981","Discontinued":"#EF4444",
                }
                for _wph_ci, (_, _wph_row) in enumerate(_wph_batch):
                    _wph_pid   = int(float(str(_wph_row.get("id", 0) or 0)))
                    _wph_alloc = float(_wph_row.get("allocated_hours") or 0)
                    _wph_log   = _wsd_all_hours.get(_wph_pid, 0.0)
                    _wph_pct   = min((_wph_log / _wph_alloc) * 100, 100) if _wph_alloc > 0 else 0
                    _wph_rem   = _wph_alloc - _wph_log if _wph_alloc > 0 else None
                    _wph_st    = str(_wph_row.get("status", ""))
                    _wph_stc   = _wph_stmap.get(_wph_st, "#94A3B8")
                    _wph_ac    = _wph_pal[_wph_ci % len(_wph_pal)]
                    if _wph_alloc == 0:
                        _wph_health     = "No Budget Set"
                        _wph_health_c   = "#94A3B8"
                        _wph_health_bg  = "#F8FAFC"
                        _wph_bar_c      = "#CBD5E1"
                    elif _wph_pct >= 100:
                        _wph_health     = "Over Budget"
                        _wph_health_c   = "#DC2626"
                        _wph_health_bg  = "#FEF2F2"
                        _wph_bar_c      = "#DC2626"
                    elif _wph_pct >= 75:
                        _wph_health     = "At Risk"
                        _wph_health_c   = "#D97706"
                        _wph_health_bg  = "#FFFBEB"
                        _wph_bar_c      = "#F59E0B"
                    else:
                        _wph_health     = "On Track"
                        _wph_health_c   = "#059669"
                        _wph_health_bg  = "#F0FDF4"
                        _wph_bar_c      = "#10B981"
                    _wph_alloc_str = f"{_wph_alloc:.0f}h" if _wph_alloc > 0 else "—"
                    _wph_rem_str   = (
                        f'<span style="color:#DC2626">Over {abs(_wph_rem):.1f}h</span>'
                        if _wph_rem is not None and _wph_rem < 0
                        else (f'{_wph_rem:.1f}h left' if _wph_rem is not None else '—')
                    )
                    _wph_start = fmt_date(str(_wph_row.get("start", "") or "")) or "—"
                    _wph_end   = fmt_date(str(_wph_row.get("end",   "") or "")) or "—"
                    with _wph_cols[_wph_ci]:
                        st.markdown(
                            f'<div style="background:#fff;border:1.5px solid #E2E8F0;'
                            f'border-radius:16px;padding:18px 16px;border-top:4px solid {_wph_ac};'
                            f'animation:fadeInUp 0.5s ease both">'
                            # project initial avatar
                            f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">'
                            f'<div style="width:38px;height:38px;border-radius:10px;background:{_wph_ac}20;'
                            f'border:1px solid {_wph_ac}40;display:flex;align-items:center;'
                            f'justify-content:center;font-size:16px;font-weight:900;color:{_wph_ac};flex-shrink:0">'
                            f'{esc(str(_wph_row.get("name","?"))[0].upper())}</div>'
                            f'<div style="min-width:0">'
                            f'<div style="font-size:12px;font-weight:800;color:#1F3B4D;'
                            f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'
                            f'{esc(str(_wph_row.get("name","")))}</div>'
                            f'<div><span style="font-size:9px;font-weight:700;color:{_wph_stc};'
                            f'background:{_wph_stc}15;padding:1px 7px;border-radius:10px;'
                            f'border:1px solid {_wph_stc}30">{esc(_wph_st)}</span></div>'
                            f'<div style="font-size:9px;color:#94A3B8;margin-top:3px">'
                            f'{esc(_wph_start)} → {esc(_wph_end)}</div>'
                            f'</div></div>'
                            # health badge
                            f'<div style="display:inline-flex;align-items:center;gap:5px;'
                            f'background:{_wph_health_bg};border-radius:20px;padding:4px 10px;'
                            f'margin-bottom:12px">'
                            f'<div style="width:7px;height:7px;border-radius:50%;background:{_wph_health_c}"></div>'
                            f'<span style="font-size:10px;font-weight:700;color:{_wph_health_c}">{_wph_health}</span>'
                            f'</div>'
                            # budget bar
                            f'<div style="background:#F1F5F9;border-radius:6px;height:8px;overflow:hidden;margin-bottom:6px">'
                            f'<div style="width:{_wph_pct:.0f}%;background:{_wph_bar_c};height:8px;border-radius:6px;'
                            f'transition:width 0.8s cubic-bezier(.4,0,.2,1)"></div>'
                            f'</div>'
                            # stats row
                            f'<div style="display:flex;justify-content:space-between;margin-top:6px">'
                            f'<div style="text-align:center">'
                            f'<div style="font-size:14px;font-weight:800;color:#1F3B4D">{_wph_log:.1f}h</div>'
                            f'<div style="font-size:8px;color:#94A3B8;text-transform:uppercase;letter-spacing:.4px">Logged</div>'
                            f'</div>'
                            f'<div style="text-align:center">'
                            f'<div style="font-size:14px;font-weight:800;color:#64748B">{_wph_alloc_str}</div>'
                            f'<div style="font-size:8px;color:#94A3B8;text-transform:uppercase;letter-spacing:.4px">Budget</div>'
                            f'</div>'
                            f'<div style="text-align:center">'
                            f'<div style="font-size:13px;font-weight:700;color:{_wph_health_c}">{_wph_pct:.0f}%</div>'
                            f'<div style="font-size:8px;color:#94A3B8;text-transform:uppercase;letter-spacing:.4px">Used</div>'
                            f'</div>'
                            f'</div>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
        else:
            st.markdown(
                '<div style="text-align:center;padding:24px;color:#94A3B8;font-size:11px;'
                'background:#F8FAFC;border-radius:12px;border:1px solid #E2E8F0">'
                'No Worksoft projects found.</div>',
                unsafe_allow_html=True,
            )


elif st.session_state.active_tab == "projects" and role == "employee":
    st.markdown(
        '<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px;'
        'font-family:Manrope,sans-serif">My Projects</h2>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p style="color:#64748B;font-size:12px;margin-bottom:16px">'
        'Worksoft projects assigned to you — log hours and add notes</p>',
        unsafe_allow_html=True,
    )

    _emp_projs = auth.get_user_worksoft_projects(cu["id"])
    if not _emp_projs:
        st.info("No projects assigned to you yet. Contact your admin or lead.")
    else:
        _ws_badge = (
            '<span style="font-size:10px;font-weight:700;background:#E0F2FE;'
            'color:#0369A1;padding:2px 8px;border-radius:4px">Worksoft</span>'
        )
        _emp_active_projs    = [p for p in _emp_projs if str(p.get("status", "")).strip() != "Completed"]
        _emp_complete_projs  = [p for p in _emp_projs if str(p.get("status", "")).strip() == "Completed"]

        _ep_ptab1, _ep_ptab2 = st.tabs([
            f"Active ({len(_emp_active_projs)})",
            f"Completed ({len(_emp_complete_projs)})",
        ])

        with _ep_ptab2:
            if not _emp_complete_projs:
                st.info("No completed projects yet.")
            else:
                for _cp in _emp_complete_projs:
                    with st.container(border=True):
                        _cph1, _cph2 = st.columns([4, 1])
                        _cph1.markdown(
                            f'<div style="font-size:14px;font-weight:700;color:#1F3B4D">'
                            f'{esc(_cp["name"])}</div>'
                            f'<div style="font-size:11px;color:#64748B;margin-top:2px">'
                            f'Client: {esc(_cp.get("client") or "—")}'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                        _cph2.markdown(
                            '<span style="font-size:10px;font-weight:700;background:#D1FAE5;'
                            'color:#065F46;padding:3px 10px;border-radius:12px">✓ Completed</span>',
                            unsafe_allow_html=True,
                        )
                        _cp_total = auth.get_project_total_hours(_cp["id"])
                        if _cp_total > 0:
                            st.markdown(
                                f'<div style="font-size:11px;color:#64748B;margin-top:4px">'
                                f'Total logged: <b>{_cp_total:.2f}h</b></div>',
                                unsafe_allow_html=True,
                            )

        with _ep_ptab1:
            if not _emp_active_projs:
                st.info("No active projects assigned to you.")
        for _eproj in _emp_active_projs:
            with st.container(border=True):
                # ── Project header ────────────────────────────────────────────
                _ep_alloc      = float(_eproj.get("allocated_hours") or 0)
                _ep_daily      = float(_eproj.get("daily_hours") or 0)
                _ep_total      = auth.get_project_total_hours(_eproj["id"])
                _ep_remain     = _ep_alloc - _ep_total if _ep_alloc > 0 else None

                _ph1, _ph2 = st.columns([4, 1])
                _ep_daily_tag = (
                    f'&nbsp;&nbsp;·&nbsp;&nbsp;Your allocation: <b>{_ep_daily:.1f}h/day</b>'
                    if _ep_daily > 0 else ""
                )
                _ph1.markdown(
                    f'<div style="font-size:15px;font-weight:700;color:#1F3B4D">'
                    f'{esc(_eproj["name"])}</div>'
                    f'<div style="font-size:11px;color:#64748B;margin-top:2px">'
                    f'Client: {esc(_eproj["client"] or "—")}&nbsp;&nbsp;·&nbsp;&nbsp;'
                    f'Status: {esc(_eproj["status"] or "—")}'
                    f'{_ep_daily_tag}</div>',
                    unsafe_allow_html=True,
                )
                _ph2.markdown(_ws_badge, unsafe_allow_html=True)

                # ── Budget progress bar ───────────────────────────────────────
                if _ep_alloc > 0:
                    _ep_pct    = min((_ep_total / _ep_alloc) * 100, 100)
                    _ep_rem_c  = "#DC2626" if (_ep_remain or 0) <= 0 else "#16A34A"
                    _ep_rem_tx = f'{abs(_ep_remain):.1f}h over' if (_ep_remain or 0) < 0 else f'{_ep_remain:.1f}h remaining'
                    _ep_bar_c  = "#DC2626" if _ep_pct >= 100 else ("#F59E0B" if _ep_pct >= 50 else "#3B82F6")
                    st.markdown(
                        f'<div style="margin:6px 0 2px">'
                        f'<div style="display:flex;justify-content:space-between;margin-bottom:3px">'
                        f'<span style="font-size:11px;font-weight:700;color:{_ep_rem_c}">'
                        f'⏰ {_ep_rem_tx}</span>'
                        f'<span style="font-size:10px;color:#64748B">'
                        f'Logged <b>{_ep_total:.2f}h</b> / Budget <b>{_ep_alloc:.1f}h</b></span>'
                        f'</div>'
                        f'<div style="background:#E2E8F0;border-radius:4px;height:7px">'
                        f'<div style="width:{_ep_pct:.0f}%;background:{_ep_bar_c};height:7px;border-radius:4px"></div>'
                        f'</div>'
                        f'<div style="font-size:9px;color:#94A3B8;margin-top:2px">{_ep_pct:.0f}% of project budget used</div>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

                st.markdown('<div style="height:4px"></div>', unsafe_allow_html=True)

                # ── Log Work Hours ────────────────────────────────────────────
                st.markdown(
                    '<span style="font-size:11px;font-weight:700;color:#0369A1">'
                    '📅 Log Work Hours</span>',
                    unsafe_allow_html=True,
                )
                _ep_lf1, _ep_lf2 = st.columns([1.2, 1])
                _ep_date_in = _ep_lf1.date_input(
                    "Work Date", value=date.today(), format="DD/MM/YYYY",
                    min_value=date(2000, 1, 1), key=f"ep_date_{_eproj['id']}",
                )
                _ep_lf3, _ep_lf4 = st.columns(2)
                _ep_from = _ep_lf3.time_input(
                    "From", value=datetime(2000, 1, 1, 9, 0).time(),
                    key=f"ep_from_{_eproj['id']}",
                )
                _ep_to = _ep_lf4.time_input(
                    "To", value=datetime(2000, 1, 1, 17, 0).time(),
                    key=f"ep_to_{_eproj['id']}",
                )
                _ep_from_dt  = datetime.combine(_ep_date_in, _ep_from)
                _ep_to_dt    = datetime.combine(_ep_date_in, _ep_to)
                _ep_calc_hrs = (_ep_to_dt - _ep_from_dt).total_seconds() / 3600
                if _ep_calc_hrs > 0:
                    st.markdown(
                        f'<div style="font-size:11px;color:#3B82F6;font-weight:600;margin:2px 0">'
                        f'⏱ {_ep_calc_hrs:.2f}h &nbsp;({_ep_from.strftime("%H:%M")} → {_ep_to.strftime("%H:%M")})</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        '<div style="font-size:11px;color:#DC2626;margin:2px 0">'
                        '⚠ End time must be after start time</div>',
                        unsafe_allow_html=True,
                    )
                _ep_desc_in = st.text_input(
                    "Description", placeholder="What did you work on?",
                    key=f"ep_desc_{_eproj['id']}", label_visibility="collapsed",
                )
                if st.button("📥 Log Hours", key=f"ep_log_{_eproj['id']}", type="primary",
                             use_container_width=True, disabled=(_ep_calc_hrs <= 0)):
                    auth.add_worksoft_manual_entry(
                        _eproj["id"], _eproj["name"], cu["id"], cu["name"],
                        _ep_date_in.strftime("%Y-%m-%d"), _ep_calc_hrs, _ep_desc_in,
                    )
                    _ep_new_total = auth.get_project_total_hours(_eproj["id"])
                    _ep_al = float(_eproj.get("allocated_hours") or 0)
                    _ep_le = _eproj.get("project_lead_email", "")
                    _ep_new_thresholds = auth.check_and_log_hours_alert(_eproj["id"], _ep_new_total, _ep_al)
                    if 50 in _ep_new_thresholds and _ep_le:
                        threading.Thread(
                            target=email_utils.send_worksoft_50pct_alert,
                            args=(_ep_le, _eproj.get("lead", ""), _eproj["name"], _ep_new_total, _ep_al),
                            daemon=True,
                        ).start()
                    if 100 in _ep_new_thresholds:
                        if _ep_le:
                            threading.Thread(
                                target=email_utils.send_worksoft_hours_alert,
                                args=(_ep_le, _eproj.get("lead", ""), _eproj["name"], _ep_new_total, _ep_al),
                                daemon=True,
                            ).start()
                        for _ep_assign in auth.get_worksoft_project_assignments(_eproj["id"]):
                            if _ep_assign.get("email"):
                                threading.Thread(
                                    target=email_utils.send_worksoft_100pct_employee_alert,
                                    args=(_ep_assign["email"], _ep_assign["name"], _eproj["name"], _ep_new_total, _ep_al),
                                    daemon=True,
                                ).start()
                    st.session_state.toast = {
                        "msg": f"{_ep_calc_hrs:.2f}h logged ({_ep_from.strftime('%H:%M')}–{_ep_to.strftime('%H:%M')}, {_ep_date_in.strftime('%d/%m/%Y')})!",
                        "type": "success",
                    }
                    st.rerun()
                _ep_my_entries = auth.get_user_worksoft_entries(_eproj["id"], cu["id"])
                if _ep_my_entries:
                    with st.expander(f"My Recent Entries ({len(_ep_my_entries)})", expanded=False):
                        for _epe in _ep_my_entries[:10]:
                            _epe_c1, _epe_c2, _epe_c3 = st.columns([1.2, 0.6, 2.5])
                            _epe_c1.markdown(f'<span style="font-size:11px;color:#64748B">{fmt_date(_epe["work_date"]) or _epe["work_date"]}</span>', unsafe_allow_html=True)
                            _epe_c2.markdown(f'<span style="font-size:12px;font-weight:700;color:#1F3B4D">{_epe["hours_worked"]:.1f}h</span>', unsafe_allow_html=True)
                            _epe_c3.markdown(f'<span style="font-size:11px;color:#475569">{esc(_epe["description"])}</span>', unsafe_allow_html=True)

                st.markdown(
                    '<div style="border-top:1px solid #E2E8F0;margin:8px 0 8px"></div>',
                    unsafe_allow_html=True,
                )

                # ── Project Notes ─────────────────────────────────────────────
                st.markdown(
                    '<span style="font-size:11px;font-weight:700;color:#475569">'
                    '💬 Project Notes</span>',
                    unsafe_allow_html=True,
                )
                _ep_comments = auth.get_worksoft_project_comments(_eproj["id"])
                if _ep_comments:
                    with st.container():
                        for _ec in _ep_comments:
                            _ec_ts = str(_ec["created_at"])
                            _ec_disp = (
                                fmt_date(_ec_ts[:10]) + " " + _ec_ts[11:16]
                                if len(_ec_ts) >= 16 else _ec_ts
                            )
                            _is_mine = _ec["user_id"] == cu["id"]
                            _name_color = "#0369A1" if _is_mine else "#374151"
                            _bg_color   = "#EFF6FF" if _is_mine else "#F8FAFC"
                            st.markdown(
                                f'<div style="background:{_bg_color};border:1px solid #E2E8F0;'
                                f'border-radius:8px;padding:8px 12px;margin-bottom:6px">'
                                f'<div style="font-size:10px;font-weight:700;color:{_name_color};'
                                f'margin-bottom:3px">{esc(_ec["user_name"])}'
                                f'<span style="font-weight:400;color:#94A3B8;margin-left:8px">'
                                f'{esc(_ec_disp)}</span></div>'
                                f'<div style="font-size:12px;color:#1F3B4D;white-space:pre-wrap">'
                                f'{esc(_ec["comment"])}</div>'
                                f'</div>',
                                unsafe_allow_html=True,
                            )
                else:
                    st.markdown(
                        '<span style="font-size:11px;color:#94A3B8">'
                        'No notes yet — add the first one below.</span>',
                        unsafe_allow_html=True,
                    )
                _new_note = st.text_input(
                    "Add a note", placeholder="Type your note here…",
                    key=f"ep_note_{_eproj['id']}", label_visibility="collapsed",
                )
                if st.button("💬 Add Note", key=f"ep_addnote_{_eproj['id']}",
                             use_container_width=True,
                             disabled=not bool(_new_note.strip())):
                    auth.add_worksoft_comment(_eproj["id"], _eproj["name"], cu["id"], cu["name"], _new_note.strip())
                    st.session_state.toast = {"msg": "Note saved!", "type": "success"}
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB: PROJECTS — ADMIN / LEAD / MANAGER VIEW
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "projects" and role not in ("employee",):

    if "proj_tracker_open" not in st.session_state:
        st.session_state["proj_tracker_open"] = None
    if "proj_timesheet_open" not in st.session_state:
        st.session_state["proj_timesheet_open"] = None

    # ── Tracker detail page (full-screen) ──────────────────────────────────────
    if st.session_state.get("proj_tracker_open"):
        _trk_sel = st.session_state["proj_tracker_open"]

        _bk_col, _ttl_col = st.columns([1, 9])
        with _bk_col:
            if st.button("← Back", key="proj_tracker_back", use_container_width=True):
                st.session_state["proj_tracker_open"] = None
                st.rerun()
        with _ttl_col:
            st.markdown(
                '<div style="display:flex;align-items:center;gap:10px;margin-bottom:2px">'
                '<span style="font-size:16px;font-weight:800;color:#1F3B4D">Project Tracker</span>'
                '<span style="font-size:11px;font-weight:600;padding:3px 12px;border-radius:20px;'
                'background:#8B5CF620;color:#7C3AED;border:1px solid #8B5CF640">'
                + esc(_trk_sel) + '</span>'
                '</div>',
                unsafe_allow_html=True
            )

        _trk_match = df[df["name"] == _trk_sel]
        if _trk_match.empty:
            st.warning("Project not found.")
        else:
            _trk_row = _trk_match.iloc[0]

            def _trk_v(key, fallback="—"):
                v = _trk_row.get(key, fallback)
                s = str(v).strip()
                return fallback if s in ("", "nan", "None", "NaN") else s

            _trk_status = _trk_v("status", "")
            _trk_color  = STATUS_STYLES.get(_trk_status, {"dot": "#94A3B8"})["dot"]
            _trk_start  = _trk_v("start", "")
            _trk_end    = _trk_v("end", "")
            _trk_due    = _trk_v("due_date", "")
            _trk_hrs    = _trk_v("hours_saved", "")
            _trk_cost   = _trk_v("cost_saved", "")
            _trk_roi    = _trk_v("roi_pct", "")

            _tc1, _tc2 = st.columns([1, 1.3])

            with _tc1:
                with st.container(border=True):
                    st.markdown(
                        '<div style="display:flex;align-items:center;gap:10px;margin-bottom:14px">'
                        '<div style="width:10px;height:10px;border-radius:50%;background:'
                        + _trk_color + ';box-shadow:0 0 6px ' + _trk_color + '40"></div>'
                        '<span style="font-size:13px;font-weight:800;color:#1F3B4D">'
                        + esc(_trk_sel) + '</span></div>',
                        unsafe_allow_html=True
                    )
                    def _trk_field(lbl, val, col="#374151"):
                        return (
                            '<div style="display:flex;gap:8px;padding:6px 0;'
                            'border-bottom:1px solid #F8FAFC">'
                            '<div style="font-size:10px;font-weight:700;color:#94A3B8;'
                            'min-width:80px;flex-shrink:0">' + lbl + '</div>'
                            '<div style="font-size:11px;color:' + col + ';font-weight:600;'
                            'word-break:break-word">' + esc(str(val)) + '</div>'
                            '</div>'
                        )
                    _trk_details = (
                        _trk_field("Client",   _trk_v("client"))
                      + _trk_field("Lead",     _trk_v("lead"))
                      + _trk_field("Employee", _trk_v("employee"))
                      + _trk_field("Status",   _trk_status, _trk_color)
                      + _trk_field("PO No.",   _trk_v("po"))
                      + _trk_field("Start",    fmt_date(_trk_start) or "—")
                      + _trk_field("End",      fmt_date(_trk_end)   or "Ongoing")
                      + _trk_field("Due Date", fmt_date(_trk_due)   or "—")
                    )
                    _desc = _trk_v("desc", "")
                    if _desc and _desc != "—":
                        _trk_details += _trk_field("Description", _desc)
                    st.markdown(_trk_details, unsafe_allow_html=True)

                    # ── Timeline Progress Bar ─────────────────────────────────
                    _tl_s = _parse_dmy(_trk_start) if _trk_start not in ("", "—") else None
                    _tl_e = _parse_dmy(_trk_end)   if _trk_end   not in ("", "—") else None
                    if _tl_s and _tl_e and _tl_e > _tl_s:
                        _tl_today   = date.today()
                        _tl_total   = (_tl_e - _tl_s).days
                        _tl_elapsed = max(0, min(_tl_total, (_tl_today - _tl_s).days))
                        _tl_pct     = round((_tl_elapsed / _tl_total) * 100)
                        if _trk_status == "Completed":
                            _tl_pct = 100
                        _tl_color = "#10B981" if _tl_pct < 70 else ("#F59E0B" if _tl_pct < 90 else "#EF4444")
                        if _trk_status == "Completed":
                            _tl_color = "#10B981"
                        _days_left = (_tl_e - _tl_today).days
                        _days_lbl  = (f"{_days_left}d left" if _days_left > 0
                                      else ("Completed" if _trk_status == "Completed"
                                            else f"{abs(_days_left)}d overdue"))
                        _days_c    = "#10B981" if _days_left > 7 else ("#F59E0B" if _days_left >= 0 else "#EF4444")
                        st.markdown(
                            f'<div style="margin-top:14px;padding-top:12px;border-top:1px solid #F1F5F9">'
                            f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px">'
                            f'<span style="font-size:10px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:.5px">Timeline</span>'
                            f'<span style="font-size:10px;font-weight:700;color:{_days_c}">{_days_lbl}</span>'
                            f'</div>'
                            f'<div style="background:#E2E8F0;border-radius:6px;height:10px;overflow:hidden">'
                            f'<div style="width:{_tl_pct}%;background:linear-gradient(90deg,{_tl_color}cc,{_tl_color});height:100%;border-radius:6px"></div>'
                            f'</div>'
                            f'<div style="display:flex;justify-content:space-between;margin-top:4px">'
                            f'<span style="font-size:9px;color:#94A3B8">{fmt_date(_trk_start)}</span>'
                            f'<span style="font-size:10px;font-weight:700;color:{_tl_color}">{_tl_pct}%</span>'
                            f'<span style="font-size:9px;color:#94A3B8">{fmt_date(_trk_end)}</span>'
                            f'</div>'
                            f'</div>',
                            unsafe_allow_html=True
                        )

                    # ── Checkpoint Track ──────────────────────────────────────
                    _ckpt_defs = [
                        ("ckpt_pdd_sdd",     "PDD/SDD"),
                        ("ckpt_development", "Dev"),
                        ("ckpt_uat",         "UAT"),
                        ("ckpt_deployment",  "Deploy"),
                    ]
                    _ckpt_today = date.today()
                    _ckpt_items = []
                    for _ck_key, _ck_lbl in _ckpt_defs:
                        _ck_s = _parse_dmy(str(_trk_row.get(f"{_ck_key}_start", "") or ""))
                        _ck_e = _parse_dmy(str(_trk_row.get(f"{_ck_key}_end",   "") or ""))
                        if _ck_s or _ck_e:
                            _ck_done  = bool(_ck_e and _ck_e <= _ckpt_today)
                            _ck_late  = bool(_ck_e and _ck_e < _ckpt_today and _trk_status not in ("Completed","R&M","Discontinued"))
                            _ck_color = "#EF4444" if _ck_late else ("#10B981" if _ck_done else "#3B82F6")
                            _ckpt_items.append((_ck_lbl, _ck_color, _ck_done, _ck_late,
                                                fmt_date(str(_ck_s or "")) or "—",
                                                fmt_date(str(_ck_e or "")) or "—"))
                    if _ckpt_items:
                        _ckpt_html = (
                            '<div style="margin-top:12px;padding-top:10px;border-top:1px solid #F1F5F9">'
                            '<div style="font-size:10px;font-weight:700;color:#64748B;text-transform:uppercase;'
                            'letter-spacing:.5px;margin-bottom:8px">Checkpoints</div>'
                            '<div style="display:flex;gap:0;align-items:center">'
                        )
                        for _i, (_lbl, _col, _done, _late, _sd, _ed) in enumerate(_ckpt_items):
                            _dot_bg = _col
                            _txt_c  = _col
                            _ckpt_html += (
                                f'<div style="display:flex;flex-direction:column;align-items:center;flex:1">'
                                f'<div style="width:22px;height:22px;border-radius:50%;background:{_dot_bg};'
                                f'display:flex;align-items:center;justify-content:center;font-size:10px;color:#fff;font-weight:800">'
                                f'{"✓" if _done else ("!" if _late else str(_i+1))}</div>'
                                f'<div style="font-size:9px;font-weight:700;color:{_txt_c};margin-top:3px;text-align:center">{_lbl}</div>'
                                f'<div style="font-size:8px;color:#94A3B8;text-align:center">{_ed}</div>'
                                f'</div>'
                            )
                            if _i < len(_ckpt_items) - 1:
                                _ckpt_html += f'<div style="flex:1;height:2px;background:{_col}40;margin-bottom:16px"></div>'
                        _ckpt_html += '</div></div>'
                        st.markdown(_ckpt_html, unsafe_allow_html=True)

            with _tc2:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                        'text-transform:uppercase;letter-spacing:.8px;margin-bottom:14px">'
                        'Metrics</div>',
                        unsafe_allow_html=True
                    )

                    _roi_html = ""
                    if _trk_hrs not in ("", "—", "0", "0.0"):
                        _roi_html += (
                            '<div style="flex:1;background:#F0FDF4;border-radius:10px;'
                            'padding:12px 14px;border-top:3px solid #10B981;text-align:center">'
                            '<div style="font-size:18px;font-weight:900;color:#059669">'
                            + _trk_hrs + '</div>'
                            '<div style="font-size:10px;color:#14532D;margin-top:3px">Hours Saved</div></div>'
                        )
                    if _trk_cost not in ("", "—", "0", "0.0"):
                        _roi_html += (
                            '<div style="flex:1;background:#EFF7F7;border-radius:10px;'
                            'padding:12px 14px;border-top:3px solid #5FA9AB;text-align:center">'
                            '<div style="font-size:18px;font-weight:900;color:#3F8E91">'
                            + chr(8377) + _trk_cost + '</div>'
                            '<div style="font-size:10px;color:#1E40AF;margin-top:3px">Cost Saved</div></div>'
                        )
                    if _trk_roi not in ("", "—", "0", "0.0"):
                        _roi_html += (
                            '<div style="flex:1;background:#FFF7ED;border-radius:10px;'
                            'padding:12px 14px;border-top:3px solid #F97316;text-align:center">'
                            '<div style="font-size:18px;font-weight:900;color:#C2410C">'
                            + _trk_roi + '%</div>'
                            '<div style="font-size:10px;color:#9A3412;margin-top:3px">ROI</div></div>'
                        )
                    if _roi_html:
                        st.markdown(
                            '<div style="display:flex;gap:10px;margin-bottom:4px">'
                            + _roi_html + '</div>',
                            unsafe_allow_html=True
                        )

                    # ── Comments (persisted to DB) ────────────────────────────
                    _trk_pid = int(_trk_row["id"]) if _trk_row.get("id") else 0

                    st.markdown(
                        '<div style="margin-top:16px">'
                        '<div style="font-size:9px;color:#94A3B8;font-weight:700;'
                        'text-transform:uppercase;letter-spacing:.9px;margin-bottom:8px">'
                        'Comments</div>'
                        '</div>',
                        unsafe_allow_html=True
                    )
                    _existing = auth.get_project_comments(_trk_pid) if _trk_pid else []
                    if _existing:
                        for _ci, _cm in enumerate(_existing):
                            _cm_bg = "#F8FAFC" if _ci % 2 == 0 else "#FFFFFF"
                            _cm_ts = str(_cm["created_at"])
                            _cm_disp = (fmt_date(_cm_ts[:10]) + " " + _cm_ts[11:16]
                                        if len(_cm_ts) >= 16 else _cm_ts)
                            st.markdown(
                                f'<div style="background:{_cm_bg};border:1px solid #E2E8F0;'
                                f'border-radius:8px;padding:9px 12px;margin-bottom:5px">'
                                f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px">'
                                f'<span style="font-size:11px;font-weight:700;color:#3F8E91">{esc(_cm["user_name"])}</span>'
                                f'<span style="font-size:10px;color:#94A3B8">{esc(_cm_disp)}</span>'
                                f'</div>'
                                f'<div style="font-size:12px;color:#374151;line-height:1.5">{esc(_cm["comment"])}</div>'
                                f'</div>',
                                unsafe_allow_html=True
                            )
                    else:
                        st.markdown(
                            '<div style="font-size:11px;color:#CBD5E1;font-style:italic;margin-bottom:6px">'
                            'No comments yet.</div>',
                            unsafe_allow_html=True
                        )
                    _new_cmt = st.text_area(
                        "Add a comment",
                        key=f"cmt_input_{_trk_sel}",
                        placeholder="Write a comment…",
                        height=68,
                        label_visibility="collapsed"
                    )
                    _cmt_col1, _cmt_col2 = st.columns([1, 4])
                    if _cmt_col1.button("Post", key=f"cmt_post_{_trk_sel}",
                                        use_container_width=True, type="primary"):
                        _txt = (_new_cmt or "").strip()
                        if _txt and _trk_pid:
                            try:
                                auth.add_project_comment(
                                    _trk_pid, _trk_sel,
                                    cu["id"], cu["name"],
                                    _txt,
                                )
                                st.rerun()
                            except Exception as _cmt_err:
                                st.error(f"Failed to save comment: {_cmt_err}")
                        elif _txt and not _trk_pid:
                            st.error("Cannot save comment: project has no ID.")

            # ── Bot Metrics (RPA projects) ─────────────────────────────────────
            if str(_trk_row.get("proj_type", "")).strip() == "RPA" and _trk_pid:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                        'text-transform:uppercase;letter-spacing:.8px;margin-bottom:12px">'
                        '⏱️ Bot Metrics</div>',
                        unsafe_allow_html=True
                    )
                    # ensure columns
                    for _trbmc in ["num_bots", "manual_run_mins", "bot_run_mins", "num_persons"]:
                        if _trbmc not in st.session_state.projects.columns:
                            st.session_state.projects[_trbmc] = 0
                    _trb_np = int(float(_trk_row.get("num_persons",    0) or 0))
                    _trb_nb = int(float(_trk_row.get("num_bots",       0) or 0))
                    _trb_mr = float(_trk_row.get("manual_run_mins", 0) or 0)
                    _trb_br = float(_trk_row.get("bot_run_mins",    0) or 0)
                    _trb_month_start = date.today().replace(day=1).isoformat()
                    _trb_month_end   = date.today().isoformat()
                    _trb_logs_all = auth.get_bot_metric_logs(project_id=_trk_pid)
                    _trb_cur_logs = auth.get_bot_metric_logs(
                        project_id=_trk_pid,
                        start_date=_trb_month_start, end_date=_trb_month_end
                    )
                    _trb_month_qty = sum(int(l.get("qty", 0) or 0) for l in _trb_cur_logs)
                    _trb_saved = max(
                        float(_trb_mr) * float(_trb_np) - float(_trb_br) * float(_trb_nb), 0
                    ) * _trb_month_qty / 60
                    # KPI row
                    _trb_k1, _trb_k2, _trb_k3, _trb_k4 = st.columns(4)
                    _trb_k1.metric("Bots",           _trb_nb)
                    _trb_k2.metric("Persons (Manual)", _trb_np)
                    _trb_k3.metric("Month Qty",       _trb_month_qty)
                    _trb_k4.metric("Hrs Saved (Mo.)", f"{_trb_saved:.1f}")
                    if role in ("admin", "lead", "manager"):
                        st.markdown(
                            '<div style="font-size:10px;font-weight:700;color:#1F3B4D;margin:10px 0 4px">'
                            '⚙️ Settings</div>',
                            unsafe_allow_html=True
                        )
                        _trbs1, _trbs2, _trbs3, _trbs4 = st.columns(4)
                        _trb_inp_np = _trbs1.number_input("Persons",       min_value=0,   value=_trb_np,  step=1,   key=f"trb_np_{_trk_pid}")
                        _trb_inp_nb = _trbs2.number_input("Bots",          min_value=0,   value=_trb_nb,  step=1,   key=f"trb_nb_{_trk_pid}")
                        _trb_inp_mr = _trbs3.number_input("Manual (mins)", min_value=0.0, value=_trb_mr,  step=1.0, key=f"trb_mr_{_trk_pid}")
                        _trb_inp_br = _trbs4.number_input("Bot (mins)",    min_value=0.0, value=_trb_br,  step=1.0, key=f"trb_br_{_trk_pid}")
                        if st.button("💾 Save Bot Settings", key=f"trb_save_{_trk_pid}", type="primary"):
                            _trb_pi = st.session_state.projects[
                                st.session_state.projects["name"] == _trk_sel
                            ].index
                            if not _trb_pi.empty:
                                st.session_state.projects.loc[_trb_pi, "num_persons"]     = str(_trb_inp_np)
                                st.session_state.projects.loc[_trb_pi, "num_bots"]        = str(_trb_inp_nb)
                                st.session_state.projects.loc[_trb_pi, "manual_run_mins"] = str(_trb_inp_mr)
                                st.session_state.projects.loc[_trb_pi, "bot_run_mins"]    = str(_trb_inp_br)
                                save_projects_async(st.session_state.projects)
                                auth.log_audit(cu["id"], cu["name"], "UPDATE", "projects",
                                               str(_trk_pid),
                                               f'Bot settings updated for "{_trk_sel}"')
                                st.session_state.toast = {"msg": "Bot settings saved!", "type": "success"}
                                st.rerun()
                        st.markdown(
                            '<div style="font-size:10px;font-weight:700;color:#1F3B4D;margin:10px 0 4px">'
                            '📅 Log Daily Quantity</div>',
                            unsafe_allow_html=True
                        )
                        _trbl1, _trbl2, _trbl3 = st.columns([2, 2, 1])
                        _trb_log_date = _trbl1.date_input("Date", value=date.today(),
                                                          format="DD/MM/YYYY", key=f"trb_logdate_{_trk_pid}")
                        _trb_log_qty  = _trbl2.number_input("Quantity", min_value=0, value=0,
                                                             step=1, key=f"trb_logqty_{_trk_pid}")
                        _trbl3.markdown('<div style="height:28px"></div>', unsafe_allow_html=True)
                        if _trbl3.button("📥 Log", key=f"trb_log_{_trk_pid}", use_container_width=True):
                            auth.upsert_bot_metric_log(_trk_pid, _trk_sel,
                                                       str(_trb_log_date), _trb_log_qty)
                            auth.log_audit(cu["id"], cu["name"], "CREATE", "bot_metric_logs",
                                           str(_trk_pid),
                                           f'Qty {_trb_log_qty} logged for "{_trk_sel}" on {_trb_log_date}')
                            st.session_state.toast = {"msg": f"Logged {_trb_log_qty} for {_trb_log_date.strftime('%d/%m/%Y')}!", "type": "success"}
                            st.rerun()
                    if _trb_logs_all:
                        _trb_log_df = pd.DataFrame(_trb_logs_all[:15])[["log_date", "qty"]].copy()
                        _trb_log_df.columns = ["Date", "Qty"]
                        st.dataframe(_trb_log_df, use_container_width=True, hide_index=True)
                    elif _trb_nb == 0:
                        st.info("Configure bot settings above to start tracking.")

            # ── Checkpoint stepper (RPA only) ─────────────────────────────────
            if str(_trk_row.get("proj_type", "")).strip() != "RPA":
                st.stop()
            _ckpt_phases = [
                ("PDD / SDD",   "pdd_sdd"),
                ("Development", "development"),
                ("UAT",         "uat"),
                ("Deployment",  "deployment"),
            ]
            _ck_today = date.today()

            def _ck_parse(s):
                if not s or str(s).strip() in ("", "nan", "—"): return None
                for _f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                    try: return datetime.strptime(str(s).strip(), _f).date()
                    except (ValueError, TypeError): pass
                return None

            _ck_cfg = {
                "done":    {"bg":"#10B981","fg":"#fff","icon":"✓","anim":"",
                            "lbl":"Done",    "lbl_c":"#059669","ring":"yes"},
                "overdue": {"bg":"#EF4444","fg":"#fff","icon":"!","anim":"",
                            "lbl":"Overdue", "lbl_c":"#DC2626","ring":"yes"},
                "working": {"bg":"#F59E0B","fg":"#fff","icon":"▶",
                            "anim":"animation:ckpt-pulse 1.4s ease-in-out infinite",
                            "lbl":"Working", "lbl_c":"#D97706","ring":"yes"},
                "pending": {"bg":"#E2E8F0","fg":"#94A3B8","icon":"",  "anim":"",
                            "lbl":"Pending", "lbl_c":"#94A3B8","ring":"none"},
            }
            _ck_line = {
                "done":    "background:#10B981",
                "overdue": "background:#EF4444",
                "working": ("background:linear-gradient(90deg,#F59E0B 0%,#FDE68A 50%,#F59E0B 100%);"
                            "background-size:200% 100%;animation:ckpt-shimmer 1.6s linear infinite"),
                "pending": "background:#E2E8F0",
            }

            _ph_states = []
            for _phl, _phk in _ckpt_phases:
                _sk = "ckpt_" + _trk_sel + "_" + _phk + "_done"
                if _sk not in st.session_state:
                    st.session_state[_sk] = False
                _ps = _ck_parse(_trk_v("ckpt_" + _phk + "_start", ""))
                _pe = _ck_parse(_trk_v("ckpt_" + _phk + "_end",   ""))
                _pd = st.session_state[_sk]
                if _pd:                                                         _pst = "done"
                elif _pe and _ck_today > _pe:                                  _pst = "overdue"
                elif _ps and _ck_today >= _ps and (_pe is None or _ck_today <= _pe): _pst = "working"
                else:                                                           _pst = "pending"
                _ph_states.append({"label":_phl,"key":_phk,
                                   "start":_ps,"end":_pe,"done":_pd,"state":_pst,"sk":_sk})

            _bubbles = ""
            for _bi, _ph in enumerate(_ph_states):
                _cs   = _ck_cfg[_ph["state"]]
                _icon = _cs["icon"] if _ph["state"] != "pending" else str(_bi + 1)
                _dr_s = _ph["start"].strftime("%d %b %Y") if _ph["start"] else "Not set"
                _dr_e = _ph["end"].strftime("%d %b %Y")   if _ph["end"]   else "Not set"
                _ring = ("box-shadow:0 0 0 5px " + _cs["bg"] + "33;") if _cs["ring"] != "none" else ""
                _anim = (_cs["anim"] + ";") if _cs["anim"] else ""
                _bubbles += (
                    '<div style="display:flex;flex-direction:column;align-items:center;'
                    'width:80px;flex-shrink:0;flex-grow:0">'
                    '<div style="width:48px;height:48px;border-radius:50%;'
                    'background:' + _cs["bg"] + ';color:' + _cs["fg"] + ';'
                    'display:flex;align-items:center;justify-content:center;'
                    'font-size:17px;font-weight:900;' + _ring + _anim + '">'
                    + _icon + '</div>'
                    '<div style="font-size:10px;font-weight:700;color:#1F3B4D;'
                    'margin-top:7px;text-align:center;line-height:1.3;width:78px;word-break:break-word">'
                    + _ph["label"] + '</div>'
                    '<div style="font-size:8.5px;font-weight:600;color:' + _cs["lbl_c"] + ';'
                    'margin-top:3px;text-align:center">' + _cs["lbl"] + '</div>'
                    '<div style="font-size:8px;color:#94A3B8;text-align:center;'
                    'margin-top:3px;line-height:1.6;width:78px">'
                    + _dr_s + '<br>' + _dr_e + '</div>'
                    '</div>'
                )
                if _bi < len(_ph_states) - 1:
                    _bubbles += (
                        '<div style="flex:1;min-width:14px;height:3px;'
                        + _ck_line[_ph["state"]] + ';border-radius:3px;'
                        'margin-top:24px;align-self:flex-start"></div>'
                    )

            st.markdown(
                '<style>'
                '@keyframes ckpt-pulse{0%,100%{box-shadow:0 0 0 0 rgba(245,158,11,.55)}'
                '50%{box-shadow:0 0 0 10px rgba(245,158,11,0)}}'
                '@keyframes ckpt-shimmer{0%{background-position:200% center}'
                '100%{background-position:-200% center}}'
                '.ckpt-wrap{display:flex;align-items:flex-start;padding:18px 16px 16px;'
                'background:#F8FAFC;border-radius:12px;border:1px solid #E2E8F0;'
                'overflow-x:auto;gap:0;min-width:0;width:100%;box-sizing:border-box}'
                '</style>'
                '<div style="margin-top:14px">'
                '<div style="font-size:9px;color:#94A3B8;font-weight:700;'
                'text-transform:uppercase;letter-spacing:.9px;margin-bottom:10px">'
                'Project Checkpoints</div>'
                '<div class="ckpt-wrap">' + _bubbles + '</div>'
                '</div>',
                unsafe_allow_html=True
            )

            st.markdown(
                '<div style="font-size:9px;color:#94A3B8;font-weight:600;'
                'text-transform:uppercase;letter-spacing:.7px;margin:12px 0 4px">'
                'Mark Phase as Done</div>',
                unsafe_allow_html=True
            )
            _ck_cols = st.columns(4)
            _ck_changed = False
            _ck_new_vals = {}
            for _ci, _ph in enumerate(_ph_states):
                with _ck_cols[_ci]:
                    _nd = st.checkbox(_ph["label"], value=_ph["done"],
                                      key="_ck_" + _trk_sel + "_" + _ph["key"] + "_done")
                _ck_new_vals[_ph["sk"]] = _nd
                if _nd != _ph["done"]:
                    _ck_changed = True
            if _ck_changed:
                for _k, _v in _ck_new_vals.items():
                    st.session_state[_k] = _v
                st.rerun()

        st.stop()

    # ── Timesheet chart page (full-screen) ────────────────────────────────────
    if st.session_state.get("proj_timesheet_open"):
        _ts_info = st.session_state["proj_timesheet_open"]
        _ts_pid  = int(_ts_info.get("id", 0) or 0)
        _ts_name = str(_ts_info.get("name", ""))
        _ts_type = str(_ts_info.get("proj_type", ""))

        _tsb_col, _tst_col = st.columns([1, 9])
        with _tsb_col:
            if st.button("← Back", key="ts_back", use_container_width=True):
                st.session_state["proj_timesheet_open"] = None
                st.rerun()
        with _tst_col:
            _ts_type_html = (
                '<span style="font-size:11px;font-weight:700;background:#D9ECEC;color:#3F8E91;'
                'padding:2px 8px;border-radius:4px;margin-left:8px">RPA</span>'
                if _ts_type == "RPA" else
                '<span style="font-size:11px;font-weight:700;background:#E0F2FE;color:#0369A1;'
                'padding:2px 8px;border-radius:4px;margin-left:8px">Worksoft</span>'
                if _ts_type == "Worksoft" else ""
            )
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:6px;margin-bottom:2px">'
                f'<span style="font-size:16px;font-weight:800;color:#1F3B4D">Timesheet</span>'
                f'<span style="font-size:11px;font-weight:600;padding:3px 12px;border-radius:20px;'
                f'background:#0EA5E920;color:#0369A1;border:1px solid #0EA5E940">{esc(_ts_name)}</span>'
                f'{_ts_type_html}</div>',
                unsafe_allow_html=True
            )

        if _ts_type == "Worksoft":
            # ══ WORKSOFT: manual time entry timeline ═════════════════════════════
            from collections import defaultdict as _dd
            _ws_entries   = auth.get_project_punches(_ts_pid) if _ts_pid else []
            _ws_total_hrs = sum(p["hours_worked"] for p in _ws_entries)
            _ws_emps      = sorted({p["user_name"] for p in _ws_entries if p["user_name"]})

            # Metrics
            _tsm1, _tsm2, _tsm3 = st.columns(3)
            _tsm1.metric("Total Hours", f"{_ws_total_hrs:.1f}")
            _tsm2.metric("Entries", len(_ws_entries))
            _tsm3.metric("Team Members", len(_ws_emps))

            # Charts row
            _wscol_l, _wscol_r = st.columns([1.3, 1])

            with _wscol_l:
                if _ws_entries:
                    _ws_emp_hrs = _dd(float)
                    for _p in _ws_entries:
                        _ws_emp_hrs[_p["user_name"]] += _p["hours_worked"]
                    _ws_sorted = sorted(_ws_emp_hrs.keys())
                    _ws_bar = go.Figure(go.Bar(
                        x=_ws_sorted,
                        y=[_ws_emp_hrs[e] for e in _ws_sorted],
                        marker_color="#0EA5E9",
                        text=[f"{_ws_emp_hrs[e]:.1f}h" for e in _ws_sorted],
                        textposition="auto",
                    ))
                    _ws_bar.update_layout(
                        title=dict(text="Total Hours by Employee", font=dict(size=13, color="#1F3B4D")),
                        xaxis_title=None, yaxis_title="Hours",
                        height=300, margin=dict(l=0, r=0, t=40, b=0),
                        plot_bgcolor="#FAFAFA", paper_bgcolor="#FFFFFF",
                        font=dict(family="Manrope, sans-serif", size=11),
                    )
                    st.plotly_chart(_ws_bar, use_container_width=True)
                else:
                    st.info("No time entries yet — employees can log hours from the Tasks tab.")

            with _wscol_r:
                with st.container(border=True):
                    st.markdown('<div style="font-size:12px;font-weight:700;color:#1F3B4D;margin-bottom:8px">Hours per Day per Employee</div>', unsafe_allow_html=True)
                    if _ws_entries:
                        _daily = _dd(lambda: _dd(float))
                        for _p in _ws_entries:
                            _daily[_p["user_name"]][_p["work_date"]] += _p["hours_worked"]
                        for _emp in sorted(_daily.keys()):
                            _emp_total = sum(_daily[_emp].values())
                            st.markdown(
                                f'<div style="font-size:11px;font-weight:700;color:#0369A1;margin-top:8px">'
                                f'{esc(_emp)} &nbsp;<span style="font-weight:400;color:#64748B">({_emp_total:.1f}h total)</span></div>',
                                unsafe_allow_html=True,
                            )
                            for _d in sorted(_daily[_emp].keys(), reverse=True):
                                st.markdown(
                                    f'<div style="font-size:11px;color:#475569;padding-left:10px">'
                                    f'{fmt_date(_d) or _d}: <b>{_daily[_emp][_d]:.2f}h</b></div>',
                                    unsafe_allow_html=True,
                                )
                    else:
                        st.markdown('<span style="font-size:11px;color:#94A3B8">No data yet.</span>', unsafe_allow_html=True)

            # ── Date-range filter ─────────────────────────────────────────────
            st.markdown('<hr style="margin:14px 0 10px;border:none;border-top:1px solid #E2E8F0">', unsafe_allow_html=True)
            st.markdown('<div style="font-size:13px;font-weight:700;color:#1F3B4D;margin-bottom:10px">🗓 Filter by Date Range</div>', unsafe_allow_html=True)
            _tsf1, _tsf2, _tsf3 = st.columns([1.2, 1.2, 0.8])
            _ts_from = _tsf1.date_input("From", value=date.today().replace(day=1), format="DD/MM/YYYY", min_value=date(2000, 1, 1), key="ts_filter_from")
            _ts_to   = _tsf2.date_input("To",   value=date.today(),                format="DD/MM/YYYY", min_value=date(2000, 1, 1), key="ts_filter_to")
            _tsf3.markdown('<div style="height:28px"></div>', unsafe_allow_html=True)
            _tsf3.button("Apply Filter", key="ts_apply_filter", type="primary", use_container_width=True)

            if _ts_from > _ts_to:
                st.warning("'From' date must be on or before 'To' date.")
            elif _ws_entries:
                _tsf_from_s = _ts_from.strftime("%Y-%m-%d")
                _tsf_to_s   = _ts_to.strftime("%Y-%m-%d")
                _tsf_label  = f"{_ts_from.strftime('%d %b %Y')} → {_ts_to.strftime('%d %b %Y')}"
                _tsf_data   = [p for p in _ws_entries if _tsf_from_s <= (p.get("work_date") or "") <= _tsf_to_s]
                _tsf_total  = sum(p["hours_worked"] for p in _tsf_data)

                with st.container(border=True):
                    st.markdown(
                        f'<div style="font-size:12px;font-weight:700;color:#1F3B4D;margin-bottom:6px">'
                        f'Results for {esc(_tsf_label)}</div>',
                        unsafe_allow_html=True,
                    )
                    st.markdown(
                        f'<div style="font-size:22px;font-weight:800;color:#0369A1;margin-bottom:10px">'
                        f'{_tsf_total:.1f}h <span style="font-size:12px;font-weight:400;color:#64748B">total in period</span></div>',
                        unsafe_allow_html=True,
                    )
                    if _tsf_data:
                        from collections import defaultdict as _tsf_dd
                        _tsf_emp_hrs = _tsf_dd(float)
                        for _fp in _tsf_data:
                            _tsf_emp_hrs[_fp["user_name"]] += _fp["hours_worked"]
                        _tsf_hdr = st.columns([2.5, 1.0, 0.7])
                        for _fh, _fl in zip(_tsf_hdr, ["Employee", "Hours", "Entries"]):
                            _fh.markdown(f'<div style="{_HDR_STYLE}">{_fl}</div>', unsafe_allow_html=True)
                        for _fe in sorted(_tsf_emp_hrs.keys()):
                            _fe_count = sum(1 for p in _tsf_data if p["user_name"] == _fe)
                            _fr = st.columns([2.5, 1.0, 0.7])
                            _fr[0].markdown(f'<span style="font-size:12px;font-weight:600;color:#0369A1">{esc(_fe)}</span>', unsafe_allow_html=True)
                            _fr[1].markdown(f'<span style="font-size:12px;font-weight:700;color:#1F3B4D">{_tsf_emp_hrs[_fe]:.1f}h</span>', unsafe_allow_html=True)
                            _fr[2].markdown(f'<span style="font-size:11px;color:#64748B">{_fe_count}</span>', unsafe_allow_html=True)
                    else:
                        st.info("No entries found for the selected date range.")

            # ── Time entries table ────────────────────────────────────────────
            if _ws_entries:
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:700;'
                    'text-transform:uppercase;letter-spacing:.8px;margin:16px 0 8px">'
                    'Time Entries</div>',
                    unsafe_allow_html=True,
                )
                _ws_phdr = st.columns([1.5, 1.0, 0.7, 3.0, 0.5])
                for _ph, _pl in zip(_ws_phdr, ["Employee", "Date", "Hours", "Description", ""]):
                    _ph.markdown(f'<div style="{_HDR_STYLE}">{_pl}</div>', unsafe_allow_html=True)
                for _pp in _ws_entries:
                    _ws_pr = st.columns([1.5, 1.0, 0.7, 3.0, 0.5], vertical_alignment="center")
                    _ws_pr[0].markdown(f'<span style="font-size:11px;font-weight:600;color:#0369A1">{esc(_pp["user_name"])}</span>', unsafe_allow_html=True)
                    _ws_pr[1].markdown(f'<span style="font-size:11px;color:#64748B">{fmt_date(_pp["work_date"]) or _pp["work_date"]}</span>', unsafe_allow_html=True)
                    _ws_pr[2].markdown(f'<span style="font-size:12px;font-weight:700;color:#1F3B4D">{_pp["hours_worked"]:.2f}h</span>', unsafe_allow_html=True)
                    _ws_pr[3].markdown(f'<span style="font-size:11px;color:#475569">{esc(_pp.get("description", ""))}</span>', unsafe_allow_html=True)
                    if _ws_pr[4].button("🗑", key=f"del_punch_{_pp['id']}", help="Delete entry"):
                        auth.delete_worksoft_punch(_pp["id"])
                        st.session_state.toast = {"msg": "Entry deleted.", "type": "info"}
                        st.rerun()

            # ── Project Notes (all employees' comments) ───────────────────────
            _ts_ws_comments = auth.get_worksoft_project_comments(_ts_pid) if _ts_pid else []
            st.markdown(
                '<div style="font-size:9px;color:#94A3B8;font-weight:700;'
                'text-transform:uppercase;letter-spacing:.8px;margin:16px 0 8px">'
                f'Project Notes ({len(_ts_ws_comments)})</div>',
                unsafe_allow_html=True,
            )
            if _ts_ws_comments:
                for _twc in _ts_ws_comments:
                    _twc_ts = str(_twc["created_at"])
                    _twc_disp = (
                        fmt_date(_twc_ts[:10]) + " " + _twc_ts[11:16]
                        if len(_twc_ts) >= 16 else _twc_ts
                    )
                    _twc_cols = st.columns([0.4, 3.5, 0.3])
                    _twc_cols[0].markdown(
                        f'<div style="font-size:11px;font-weight:700;color:#0369A1">'
                        f'{esc(_twc["user_name"])}</div>'
                        f'<div style="font-size:10px;color:#94A3B8">{esc(_twc_disp)}</div>',
                        unsafe_allow_html=True,
                    )
                    _twc_cols[1].markdown(
                        f'<div style="font-size:12px;color:#1F3B4D;white-space:pre-wrap;'
                        f'background:#F8FAFC;border:1px solid #E2E8F0;border-radius:6px;'
                        f'padding:6px 10px">{esc(_twc["comment"])}</div>',
                        unsafe_allow_html=True,
                    )
                    if _twc_cols[2].button("🗑", key=f"del_wsc_{_twc['id']}", help="Delete note"):
                        auth.delete_worksoft_comment(_twc["id"])
                        st.session_state.toast = {"msg": "Note deleted.", "type": "info"}
                        st.rerun()
            else:
                st.markdown(
                    '<span style="font-size:11px;color:#94A3B8">No notes from employees yet.</span>',
                    unsafe_allow_html=True,
                )

        else:
            # ══ Non-Worksoft: manual timesheet ═══════════════════════════════════
            _ts_entries = auth.get_project_timesheets(_ts_pid) if _ts_pid else []
            _ts_total   = sum(e["hours"] for e in _ts_entries)
            _ts_emp_set = sorted({e["employee_name"] for e in _ts_entries if e["employee_name"]})

            _tsm1, _tsm2, _tsm3 = st.columns(3)
            _tsm1.metric("Total Hours", f"{_ts_total:.1f}")
            _tsm2.metric("Entries", len(_ts_entries))
            _tsm3.metric("Team Members", len(_ts_emp_set))

            _tscol_l, _tscol_r = st.columns([1.2, 1])

            with _tscol_l:
                if _ts_entries:
                    from collections import defaultdict as _dd
                    _emp_hours   = _dd(float)
                    for _e in _ts_entries:
                        _emp_hours[_e["employee_name"]] += _e["hours"]
                    _sorted_emps = sorted(_emp_hours.keys())
                    _ts_bar = go.Figure(go.Bar(
                        x=_sorted_emps,
                        y=[_emp_hours[e] for e in _sorted_emps],
                        marker_color="#3F8E91",
                        text=[f"{_emp_hours[e]:.1f}h" for e in _sorted_emps],
                        textposition="auto",
                    ))
                    _ts_bar.update_layout(
                        title=dict(text="Hours by Employee", font=dict(size=13, color="#1F3B4D")),
                        xaxis_title=None, yaxis_title="Hours",
                        height=300, margin=dict(l=0, r=0, t=40, b=0),
                        plot_bgcolor="#FAFAFA", paper_bgcolor="#FFFFFF",
                        font=dict(family="Manrope, sans-serif", size=11),
                    )
                    st.plotly_chart(_ts_bar, use_container_width=True)
                else:
                    st.info("No timesheet entries yet. Log the first entry on the right.")

            with _tscol_r:
                with st.container(border=True):
                    st.markdown(
                        '<div style="font-size:12px;font-weight:700;color:#1F3B4D;margin-bottom:10px">'
                        'Log Timesheet Entry</div>',
                        unsafe_allow_html=True,
                    )
                    _ts_f1, _ts_f2 = st.columns(2)
                    _ts_emp_input  = _ts_f1.text_input("Employee Name *", key="ts_emp", placeholder="e.g. Ravi Kumar")
                    _ts_date_input = _ts_f2.date_input("Work Date *", key="ts_date", value=date.today(), format="DD/MM/YYYY")
                    _ts_f3, _ts_f4 = st.columns(2)
                    _ts_hours_input = _ts_f3.number_input("Hours *", key="ts_hours", min_value=0.5, max_value=24.0, value=8.0, step=0.5)
                    _ts_desc_input  = _ts_f4.text_input("Description", key="ts_desc", placeholder="What was done?")
                    if st.button("Log Entry", type="primary", key="ts_log_btn", use_container_width=True):
                        if not _ts_emp_input.strip():
                            st.error("Employee name is required.")
                        elif not _ts_pid:
                            st.error("Project ID not found.")
                        else:
                            try:
                                auth.create_timesheet_entry(
                                    _ts_pid, _ts_name,
                                    _ts_emp_input.strip(),
                                    _ts_date_input.strftime("%Y-%m-%d"),
                                    float(_ts_hours_input),
                                    _ts_desc_input.strip(),
                                    cu["id"],
                                )
                                st.session_state.toast = {"msg": "Timesheet entry logged!", "type": "success"}
                                st.rerun()
                            except Exception as _ts_err:
                                st.error(f"Failed to log entry: {_ts_err}")

            if _ts_entries:
                st.markdown(
                    '<div style="font-size:9px;color:#94A3B8;font-weight:700;'
                    'text-transform:uppercase;letter-spacing:.8px;margin:14px 0 8px">'
                    'All Entries</div>',
                    unsafe_allow_html=True,
                )
                _ts_hdr = st.columns([1.5, 0.9, 0.7, 2.5, 0.5])
                for _th, _tl in zip(_ts_hdr, ["Employee", "Date", "Hours", "Description", ""]):
                    _th.markdown(f'<div style="{_HDR_STYLE}">{_tl}</div>', unsafe_allow_html=True)
                for _te in _ts_entries:
                    _ts_row = st.columns([1.5, 0.9, 0.7, 2.5, 0.5])
                    _ts_row[0].markdown(f'<span style="font-size:11px;font-weight:600;color:#3F8E91">{esc(_te["employee_name"])}</span>', unsafe_allow_html=True)
                    _ts_row[1].markdown(f'<span style="font-size:11px;color:#64748B">{esc(fmt_date(_te["work_date"]) or _te["work_date"])}</span>', unsafe_allow_html=True)
                    _ts_row[2].markdown(f'<span style="font-size:12px;font-weight:700;color:#1F3B4D">{_te["hours"]:.1f}h</span>', unsafe_allow_html=True)
                    _ts_row[3].markdown(f'<span style="font-size:11px;color:#475569">{esc(_te["description"])}</span>', unsafe_allow_html=True)
                    if _ts_row[4].button("🗑", key=f"ts_del_{_te['id']}", help="Delete entry"):
                        auth.delete_timesheet_entry(_te["id"])
                        st.session_state.toast = {"msg": "Entry deleted.", "type": "info"}
                        st.rerun()

        st.stop()

    # ── Shared filter bar ────────────────────────────────────────────────────────
    f1, f2, f3, f4 = st.columns([2, 1.5, 1.2, 1.5])
    search_q      = f1.text_input("Search", placeholder="Search projects…",
                                  label_visibility="collapsed")
    client_filter = f2.selectbox("Client", ["All"] + sorted(df["client"].dropna().unique().tolist()),
                                 label_visibility="collapsed")
    active_filter = f3.selectbox("Active", ["All", "Active", "Inactive"],
                                 label_visibility="collapsed")
    sort_by       = f4.selectbox(
        "Sort", ["Default","Name A→Z","Name Z→A","Due Date ↑","Due Date ↓","ROI % ↓","Client A→Z"],
        key="proj_sort_by", label_visibility="collapsed"
    )

    # ── Deadline Risk Alert Banner ────────────────────────────────────────────
    _today_proj = date.today()
    _at_risk = []
    for _, _pr in df.iterrows():
        _pr_due = _parse_dmy(str(_pr.get("due_date", "") or ""))
        if _pr_due and 0 <= (_pr_due - _today_proj).days <= 7 and str(_pr.get("status","")) not in ("Completed","Discontinued"):
            _at_risk.append(str(_pr.get("name","")))
    _overdue = []
    for _, _pr in df.iterrows():
        _pr_due = _parse_dmy(str(_pr.get("due_date", "") or ""))
        if _pr_due and (_pr_due - _today_proj).days < 0 and str(_pr.get("status","")) not in ("Completed","Discontinued"):
            _overdue.append(str(_pr.get("name","")))
    if _overdue:
        st.markdown(
            f'<div style="background:#FEF2F2;border:1px solid #FECACA;border-radius:10px;padding:10px 16px;margin-bottom:10px;display:flex;align-items:center;gap:10px">'
            f'<span style="font-size:18px">🚨</span>'
            f'<div><span style="font-size:12px;font-weight:700;color:#991B1B">Overdue: {len(_overdue)} project(s) — </span>'
            f'<span style="font-size:11px;color:#B91C1C">{", ".join(_overdue[:5])}{"..." if len(_overdue)>5 else ""}</span></div>'
            f'</div>',
            unsafe_allow_html=True
        )
    if _at_risk:
        st.markdown(
            f'<div style="background:#FFFBEB;border:1px solid #FDE68A;border-radius:10px;padding:10px 16px;margin-bottom:10px;display:flex;align-items:center;gap:10px">'
            f'<span style="font-size:18px">⚠️</span>'
            f'<div><span style="font-size:12px;font-weight:700;color:#92400E">Due this week: {len(_at_risk)} project(s) — </span>'
            f'<span style="font-size:11px;color:#B45309">{", ".join(_at_risk[:5])}{"..." if len(_at_risk)>5 else ""}</span></div>'
            f'</div>',
            unsafe_allow_html=True
        )


    # ── Helper: apply shared filters + sort to a sub-set of projects ─────────
    def _apply_filters(subset):
        out = subset.copy()
        if search_q:
            q = search_q.lower()
            sc = [c for c in ["name","employee","lead","client","desc"] if c in out.columns]
            mask = (out[sc].fillna("").astype(str)
                    .apply(lambda col: col.str.lower().str.contains(q, regex=False))
                    .any(axis=1))
            out = out[mask]
        if client_filter != "All":
            out = out[out["client"] == client_filter]
        if active_filter != "All" and "is_active" in out.columns:
            inactive_vals = {"false", "0", "no"}
            raw_col = out["is_active"].astype(str).str.strip().str.lower()
            if active_filter == "Active":
                out = out[~raw_col.isin(inactive_vals)]
            else:
                out = out[raw_col.isin(inactive_vals)]
        # ── Sorting ──────────────────────────────────────────────────────────
        _sb = sort_by
        if _sb == "Name A→Z" and "name" in out.columns:
            out = out.sort_values("name", key=lambda s: s.str.lower(), na_position="last")
        elif _sb == "Name Z→A" and "name" in out.columns:
            out = out.sort_values("name", key=lambda s: s.str.lower(), ascending=False, na_position="last")
        elif _sb in ("Due Date ↑", "Due Date ↓") and "due_date" in out.columns:
            _asc = (_sb == "Due Date ↑")
            _fill = date.max if _asc else date.min
            out = out.copy()
            out["_s_due"] = out["due_date"].apply(lambda x: _parse_dmy(str(x)) or _fill)
            out = out.sort_values("_s_due", ascending=_asc).drop(columns=["_s_due"])
        elif _sb == "ROI % ↓" and "roi_pct" in out.columns:
            out = out.copy()
            out["_s_roi"] = pd.to_numeric(out["roi_pct"], errors="coerce").fillna(0)
            out = out.sort_values("_s_roi", ascending=False).drop(columns=["_s_roi"])
        elif _sb == "Client A→Z" and "client" in out.columns:
            out = out.sort_values("client", key=lambda s: s.str.lower(), na_position="last")
        return out

    # ── Helper: render active/inactive pill ───────────────────────────────────
    def _active_pill(row):
        raw = str(row.get("is_active","True")).strip().lower()
        if raw in ["false","0","no"]:
            return '<span style="display:inline-flex;align-items:center;gap:3px;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;background:#FEF2F2;color:#991B1B"><span style="width:5px;height:5px;border-radius:50%;background:#EF4444;display:inline-block"></span>Inactive</span>'
        return '<span style="display:inline-flex;align-items:center;gap:3px;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;background:#ECFDF5;color:#065F46"><span style="width:5px;height:5px;border-radius:50%;background:#10B981;display:inline-block"></span>Active</span>'

    # ── Helper: render a project table for a given filtered DataFrame ──────────
    _ROW_BG = {
        "Important": "#FCEAEA",
        "Completed": "#E5F2EC",
        "R&M":       "#EFF7F7",
    }

    def _render_project_table(filtered, tab_key="", show_timesheet=True):
        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 10px"><b>{len(filtered)}</b> projects</p>',
                    unsafe_allow_html=True)
        if filtered.empty:
            st.info("No projects match the current filters.")
            return

        is_admin = (role in ("admin", "lead", "manager"))
        can_edit = role in ("admin", "lead", "manager")

        # ── Flat column layout — every column is its own Streamlit column ──────
        # [ID, Name, Status, Client, Lead, Employee, Type, Start, End, Due, PO, (Remaining*), ✏, 🗑, 📎]
        _is_ws_tab = (tab_key == "ws" or tab_key == "ws_mine")
        if _is_ws_tab:
            _CW_BASE = [0.38, 1.75, 0.95, 1.1, 0.85, 1.3, 0.55, 0.72]
            _HDR_LABELS = ["ID", "Project", "Status", "Client", "Lead", "Employee",
                           "Type", "Remaining"]
            _ws_all_totals = auth.get_all_worksoft_total_hours()
        else:
            _CW_BASE = [0.38, 1.75, 0.95, 1.1, 0.85, 1.3, 0.55, 0.75, 0.75, 0.85, 0.5]
            _HDR_LABELS = ["ID", "Project", "Status", "Client", "Lead", "Employee",
                           "Type", "Start", "End", "Due", "PO"]
            _ws_all_totals = {}
        if show_timesheet:
            if _is_ws_tab:
                _CW_ACT = ([0.38, 0.38, 0.38] if is_admin else
                           ([0.38, 0.38]       if can_edit  else [0.38]))
                _ACT_LABELS = (["", "", "📊"] if is_admin else
                               (["", "📊"]    if can_edit  else ["📊"]))
            else:
                _CW_ACT = ([0.38, 0.38, 0.38, 0.38] if is_admin else
                           ([0.38, 0.38, 0.38]       if can_edit  else [0.38, 0.38]))
                _ACT_LABELS = (["", "", "📎", "📊"] if is_admin else
                               (["", "📎", "📊"]    if can_edit  else ["📎", "📊"]))
        else:
            _CW_ACT = ([0.38, 0.38, 0.38] if is_admin else
                       ([0.38, 0.38]       if can_edit  else [0.38]))
            _ACT_LABELS = (["", "", "📎"] if is_admin else
                           (["", "📎"]    if can_edit  else ["📎"]))
        col_widths = _CW_BASE + _CW_ACT

        def _type_badge(pt):
            if pt == "RPA":
                return '<span style="font-size:9px;font-weight:700;background:#D9ECEC;color:#3F8E91;padding:1px 6px;border-radius:4px">RPA</span>'
            if pt == "Worksoft":
                return '<span style="font-size:9px;font-weight:700;background:#E0F2FE;color:#0369A1;padding:1px 6px;border-radius:4px">WS</span>'
            if pt == "AI Agent":
                return '<span style="font-size:9px;font-weight:700;background:#F3E8FF;color:#7C3AED;padding:1px 6px;border-radius:4px">AI</span>'
            return '<span style="font-size:10px;color:#CBD5E1">—</span>'

        # Header row
        hcols = st.columns(col_widths)
        for _hc, _hl in zip(hcols, _HDR_LABELS + _ACT_LABELS):
            _hc.markdown(f'<div style="{_HDR_STYLE}">{_hl}</div>', unsafe_allow_html=True)

        # Pre-compute per-row name-button CSS in one batch
        _css_parts = []
        _btn_css = (
            "background:transparent!important;border:none!important;box-shadow:none!important;"
            "color:#3F8E91!important;font-size:12px!important;font-weight:600!important;"
            "text-align:left!important;padding:4px 0!important;"
            "word-break:break-word!important;line-height:1.4!important;"
            "cursor:pointer!important;border-radius:0!important;width:100%!important"
        )
        for _prow in filtered.to_dict("records"):
            _rc = f"pr-{tab_key}-{str(_prow.get('id',''))}".replace(" ","_").replace(".","_")
            _bg = next((_ROW_BG[s] for s in _ROW_BG if s in str(_prow.get("status",""))), "#FFFFFF")
            _sel = f'[data-testid="stHorizontalBlock"]>[data-testid="stVerticalBlock"]:has(.{_rc})'
            _css_parts.append(
                f'{_sel}{{background:{_bg}!important}}'
                f'{_sel} [data-testid="stButton"]>button{{{_btn_css}}}'
                f'{_sel} [data-testid="stButton"]>button:hover{{color:#256B6E!important}}'
            )
        if _css_parts:
            st.markdown(f"<style>{''.join(_css_parts)}</style>", unsafe_allow_html=True)

        _file_counts = _get_file_counts_cached()

        # Build employee + client pools for inline edit forms
        _all_emps = sorted(set(
            n.strip()
            for raw in st.session_state.projects.get("employee", pd.Series(dtype=str)).dropna()
            for n in str(raw).replace("&", ",").split(",") if n.strip()
        ) | set(
            str(l).strip()
            for l in st.session_state.projects.get("lead", pd.Series(dtype=str)).dropna()
            if str(l).strip()
        ))
        _all_clients = sorted(set(
            str(c).strip()
            for c in st.session_state.projects.get("client", pd.Series(dtype=str)).dropna()
            if str(c).strip()
        ))
        _EMP_NEW    = "── Type new name ──"
        _CLIENT_NEW = "── Type new client ──"

        # ── Data rows ──────────────────────────────────────────────────────────
        for row in filtered.to_dict("records"):
            row_status = str(row.get("status",""))
            rid        = str(row.get("id",""))
            pname      = str(row.get("name",""))
            _row_cls   = f"pr-{tab_key}-{rid}".replace(" ","_").replace(".", "_")
            _edit_key  = f"proj_inline_edit_{tab_key}_{rid}"
            _inline_on = st.session_state.get(_edit_key, False)

            _ss = STATUS_STYLES.get(row_status, {"bg":"#F1F5F9","text":"#64748B","dot":"#94A3B8"})
            status_badge = (
                f'<span style="display:inline-flex;align-items:center;gap:3px;'
                f'background:{_ss["bg"]};color:{_ss["text"]};font-size:9px;font-weight:700;'
                f'padding:2px 6px;border-radius:20px;white-space:nowrap">'
                f'<span style="width:5px;height:5px;border-radius:50%;background:{_ss["dot"]};'
                f'display:inline-block;flex-shrink:0"></span>{esc(row_status)}</span>'
            ) if row_status else cell("—","10px","#CBD5E1")

            lead_val  = str(row.get("lead","")).strip()
            lead_disp = (f'<span style="font-size:11px;font-weight:600;color:#3F8E91">{esc(lead_val)}</span>'
                         if lead_val else '<span style="font-size:11px;color:#CBD5E1">—</span>')
            new_tag = (' <span style="font-size:9px;font-weight:700;background:#D9ECEC;'
                       'color:#3F8E91;padding:1px 5px;border-radius:4px">NEW</span>') if is_new(row) else ""

            rcols = st.columns(col_widths, vertical_alignment="center")
            ci = 0  # column index tracker

            # ID
            rcols[ci].markdown(f'<span style="font-size:10px;color:#94A3B8;font-weight:600">{esc(rid)}</span>', unsafe_allow_html=True); ci += 1
            # Project Name (styled link button)
            with rcols[ci]:
                st.markdown(f'<span class="{_row_cls}"></span>', unsafe_allow_html=True)
                _pname_display = pname + new_tag
                if st.button(pname, key=f"pname_{tab_key}_{rid}", use_container_width=True, help="Open tracker"):
                    st.session_state["proj_tracker_open"] = pname
                    st.rerun()
            ci += 1
            # Status + Health Score — quick-change selectbox for editors
            _hs = compute_health_score(row)
            _hs_html = (
                f'<span style="font-size:9px;font-weight:700;color:{_hs["color"]};'
                f'background:{_hs["color"]}12;padding:1px 5px;border-radius:8px;display:block;margin-top:2px">'
                f'{_hs["label"]} {_hs["score"]}%</span>'
            ) if _hs["label"] != "N/A" else ""
            if can_edit:
                _qs_opts = WS_STATUSES if _is_ws_tab else ALL_STATUSES
                _qs_idx  = _qs_opts.index(row_status) if row_status in _qs_opts else 0
                _qs_val  = rcols[ci].selectbox(
                    "Status", _qs_opts, index=_qs_idx,
                    key=f"qs_{tab_key}_{rid}",
                    label_visibility="collapsed",
                )
                if _hs_html:
                    rcols[ci].markdown(_hs_html, unsafe_allow_html=True)
                if _qs_val != row_status:
                    _updated_recs = [
                        dict(r, status=_qs_val) if str(r.get("id","")) == rid else r
                        for r in st.session_state.projects.to_dict("records")
                    ]
                    st.session_state.projects = pd.DataFrame(_updated_recs)
                    save_projects_async(st.session_state.projects)
                    auth.log_audit(cu["id"], cu["name"], "UPDATE", "projects", rid,
                                   f'Status: "{row_status}" → "{_qs_val}" on "{pname}"')
                    st.session_state.toast = {"msg": f'Status → {_qs_val}', "type": "success"}
                    if _qs_val == "Completed":
                        st.balloons()
                    st.rerun()
            else:
                rcols[ci].markdown(status_badge + _hs_html, unsafe_allow_html=True)
            ci += 1
            # Client
            rcols[ci].markdown(f'<span style="font-size:11px;color:#374151">{esc(str(row.get("client","")))}</span>', unsafe_allow_html=True); ci += 1
            # Lead
            rcols[ci].markdown(lead_disp, unsafe_allow_html=True); ci += 1
            # Employee
            rcols[ci].markdown(f'<span style="font-size:11px;color:#374151">{esc(str(row.get("employee","")))}</span>', unsafe_allow_html=True); ci += 1
            # Type
            rcols[ci].markdown(_type_badge(str(row.get("proj_type","")).strip()), unsafe_allow_html=True); ci += 1
            if not _is_ws_tab:
                # Start
                rcols[ci].markdown(f'<span style="font-size:11px;color:#64748B">{esc(fmt_date(str(row.get("start",""))))}</span>', unsafe_allow_html=True); ci += 1
                # End
                rcols[ci].markdown(f'<span style="font-size:11px;color:#64748B">{esc(fmt_date(str(row.get("end",""))))}</span>', unsafe_allow_html=True); ci += 1
                # Due
                rcols[ci].markdown(_due_cell(str(row.get("due_date",""))), unsafe_allow_html=True); ci += 1
                # PO
                rcols[ci].markdown(f'<span style="font-size:11px;color:#94A3B8">{esc(str(row.get("po","")))}</span>', unsafe_allow_html=True); ci += 1

            # Remaining (Worksoft tab only)
            if _is_ws_tab:
                try:
                    _p_id_int = int(float(str(row.get("id", 0) or 0)))
                    _p_alloc  = float(str(row.get("allocated_hours") or 0) or 0)
                    _p_total  = _ws_all_totals.get(_p_id_int, 0.0)
                except Exception:
                    _p_alloc, _p_total = 0.0, 0.0
                if _p_alloc > 0:
                    _p_rem = _p_alloc - _p_total
                    _p_rc  = "#DC2626" if _p_rem <= 0 else "#16A34A"
                    _p_rtx = f'{abs(_p_rem):.1f}h {"over" if _p_rem < 0 else "left"}'
                    rcols[ci].markdown(
                        f'<span style="font-size:11px;font-weight:700;color:{_p_rc}">{_p_rtx}</span>'
                        f'<br><span style="font-size:10px;color:#94A3B8">{_p_total:.1f}/{_p_alloc:.0f}h</span>',
                        unsafe_allow_html=True,
                    )
                else:
                    rcols[ci].markdown('<span style="font-size:10px;color:#CBD5E1">—</span>', unsafe_allow_html=True)
                ci += 1


            # ✏ Edit (inline toggle)
            if can_edit:
                with rcols[ci]:
                    _em = "act-warn-marker" if _inline_on else "act-edit-marker"
                    st.markdown(f'<span class="proj-act-marker {_em}"></span>', unsafe_allow_html=True)
                    if st.button("✕" if _inline_on else "✏",
                                 key=f"edit_{tab_key}_{rid}",
                                 help="Cancel edit" if _inline_on else "Edit project",
                                 use_container_width=True):
                        st.session_state[_edit_key] = not _inline_on
                        st.rerun()
                ci += 1

            # 🗑 Delete
            if is_admin:
                with rcols[ci]:
                    st.markdown('<span class="proj-act-marker proj-del-marker"></span>', unsafe_allow_html=True)
                    if st.button("🗑", key=f"del_{tab_key}_{rid}", help="Delete project", use_container_width=True):
                        st.session_state.confirm_delete = {"id": rid, "name": pname}
                        st.rerun()
                ci += 1

            # 📎 Files (not shown in Worksoft tab)
            if not _is_ws_tab:
                with rcols[ci]:
                    st.markdown('<span class="proj-act-marker proj-files-marker"></span>', unsafe_allow_html=True)
                    _fc = _file_counts.get(_safe_folder(pname), 0)
                    if st.button(f"📎{_fc}" if _fc else "📎",
                                 key=f"files_{tab_key}_{rid}",
                                 help="Upload / view project files",
                                 use_container_width=True):
                        st.session_state.file_panel_proj = rid
                        st.session_state.file_panel_name = pname
                        st.rerun()
            if show_timesheet:
                if not _is_ws_tab:
                    ci += 1
                with rcols[ci]:
                    st.markdown('<span class="proj-ts-marker"></span>', unsafe_allow_html=True)
                    if st.button("📊", key=f"ts_{tab_key}_{rid}",
                                 help="View timesheet chart",
                                 use_container_width=True):
                        st.session_state["proj_timesheet_open"] = {
                            "id": rid, "name": pname,
                            "proj_type": str(row.get("proj_type", "")).strip(),
                        }
                        st.rerun()

            # ── Inline edit form (merged details + edit) ───────────────────────
            if _inline_on and can_edit:
                with st.container(border=True):
                    st.markdown(
                        f'<div style="font-size:13px;font-weight:700;color:#1F3B4D;margin-bottom:10px">'
                        f'Edit — {esc(pname)}</div>',
                        unsafe_allow_html=True
                    )

                    if _is_ws_tab:
                        # ── Worksoft edit form (mirrors the create form) ──────────
                        _ie_all_users   = [u for u in auth.get_all_users() if u.get("is_active")]
                        _ie_id_to_label = {u["id"]: f"{u['name']} ({u['email']})" for u in _ie_all_users}
                        _ie_lead_users  = [u for u in _ie_all_users if u.get("role") in ("lead", "admin", "manager")]
                        _ie_cur_assigns = auth.get_worksoft_project_assignments(int(float(rid)))
                        _ie_cur_uid_set = {a["user_id"] for a in _ie_cur_assigns}
                        _ie_saved_dh    = {a["user_id"]: a["daily_hours"] for a in _ie_cur_assigns}

                        _iea, _ieb = st.columns(2)
                        _ie_ws_name = _iea.text_input("Project Name *", value=row.get("name", ""), key=f"ie_ws_name_{rid}")
                        _ie_cl_opts = _all_clients + [_CLIENT_NEW]
                        _ie_cl_cur  = row.get("client", "")
                        _ie_cl_idx  = _ie_cl_opts.index(_ie_cl_cur) if _ie_cl_cur in _ie_cl_opts else len(_ie_cl_opts) - 1
                        _ie_cl_sel  = _ieb.selectbox("Client *", _ie_cl_opts, index=_ie_cl_idx, key=f"ie_ws_client_{rid}")
                        _ie_ws_client = _ieb.text_input("New client name *", key=f"ie_ws_newclient_{rid}") if _ie_cl_sel == _CLIENT_NEW else _ie_cl_sel

                        # Lead selectbox — match by email to find current lead user
                        _ie_lead_email_cur = str(row.get("project_lead_email", "") or "")
                        _ie_lead_cur_id    = next((u["id"] for u in _ie_lead_users if u["email"] == _ie_lead_email_cur), None)
                        _ie_lead_id_opts   = [None] + [u["id"] for u in _ie_lead_users]
                        _ie_lead_sel_idx   = _ie_lead_id_opts.index(_ie_lead_cur_id) if _ie_lead_cur_id in _ie_lead_id_opts else 0

                        _iec, _ied, _iee = st.columns(3)
                        _ie_ws_lead_id = _iec.selectbox(
                            "Project Lead",
                            options=_ie_lead_id_opts,
                            index=_ie_lead_sel_idx,
                            format_func=lambda uid: "— Select Lead —" if uid is None else next(
                                (u["name"] for u in _ie_lead_users if u["id"] == uid), "—"
                            ),
                            key=f"ie_ws_lead_{rid}",
                        )
                        _ie_ws_lead_name  = next((u["name"]  for u in _ie_lead_users if u["id"] == _ie_ws_lead_id), "") if _ie_ws_lead_id else ""
                        _ie_ws_lead_email = next((u["email"] for u in _ie_lead_users if u["id"] == _ie_ws_lead_id), "") if _ie_ws_lead_id else ""

                        _ie_ws_st_cur = row.get("status", "")
                        _ie_ws_st_cur = _ie_ws_st_cur if _ie_ws_st_cur in WS_STATUSES else WS_STATUSES[0]
                        _ie_ws_status = _ied.selectbox("Status", WS_STATUSES, index=WS_STATUSES.index(_ie_ws_st_cur), key=f"ie_ws_status_{rid}")
                        _ie_ws_active = _iee.checkbox(
                            "Active",
                            value=(str(row.get("is_active", "True")).strip().lower() not in ["false", "0", "no"]),
                            key=f"ie_ws_active_{rid}",
                        )

                        _ief, _ieg = st.columns(2)
                        _ie_ws_alloc = _ief.number_input(
                            "Budget Hours (total project)", min_value=0.0, step=0.5,
                            value=float(row.get("allocated_hours") or 0),
                            key=f"ie_ws_alloc_{rid}",
                        )
                        _ie_ws_start_dt = _ieg.date_input(
                            "Start Date", value=_parse_dmy(row.get("start", "")),
                            key=f"ie_ws_start_{rid}", format="DD/MM/YYYY",
                        )
                        _ieh, _iei = st.columns(2)
                        _ie_ws_ongoing = _ieh.checkbox(
                            "Ongoing (no end date)",
                            value=not bool(str(row.get("end", "")).strip()),
                            key=f"ie_ws_ongoing_{rid}",
                        )
                        _ie_ws_end_dt = None if _ie_ws_ongoing else _iei.date_input(
                            "End Date",
                            value=_parse_dmy(str(row.get("end", "")).strip()) or date.today(),
                            key=f"ie_ws_end_{rid}", format="DD/MM/YYYY",
                        )

                        # Employee assignment with per-employee daily hours
                        _ie_ws_emp_sel = st.multiselect(
                            "Assign Employees",
                            options=[u["id"] for u in _ie_all_users],
                            default=list(_ie_cur_uid_set),
                            format_func=lambda uid: _ie_id_to_label.get(uid, str(uid)),
                            key=f"ie_ws_emps_{rid}",
                        )
                        _ie_ws_dh_map = {}
                        if _ie_ws_emp_sel:
                            st.markdown(
                                '<div style="font-size:12px;font-weight:600;color:#475569;margin:6px 0 4px">'
                                'Daily Hours Allocation per Employee</div>',
                                unsafe_allow_html=True,
                            )
                            _ie_dh_cols = st.columns(min(len(_ie_ws_emp_sel), 3))
                            for _ie_dhi, _ie_dhu in enumerate(_ie_ws_emp_sel):
                                _ie_dh_label = _ie_id_to_label.get(_ie_dhu, str(_ie_dhu)).split("(")[0].strip()
                                _ie_dh_cur   = _ie_saved_dh.get(_ie_dhu, 8.0)
                                _ie_ws_dh_map[_ie_dhu] = _ie_dh_cols[_ie_dhi % 3].number_input(
                                    _ie_dh_label, min_value=0.0, max_value=24.0, step=0.5,
                                    value=float(_ie_dh_cur or 8.0),
                                    key=f"ie_ws_dh_{rid}_{_ie_dhu}",
                                )
                        _ie_ws_emp_names = ", ".join(
                            u["name"] for u in _ie_all_users if u["id"] in _ie_ws_emp_sel
                        )

                        _ie_sb1, _ie_sb2, _ = st.columns([1, 1, 4])
                        if _ie_sb1.button("💾 Save", type="primary", use_container_width=True, key=f"ie_save_{tab_key}_{rid}"):
                            _ie_errs = []
                            if not _ie_ws_name.strip() or len(_ie_ws_name.strip()) < 3:
                                _ie_errs.append("Project name must be at least 3 characters.")
                            if not _ie_ws_client.strip():
                                _ie_errs.append("Client is required.")
                            if _ie_errs:
                                for _ie_er in _ie_errs: st.error(_ie_er)
                            else:
                                _ie_start = _ie_ws_start_dt.strftime("%d/%m/%Y") if _ie_ws_start_dt else ""
                                _ie_end   = "" if _ie_ws_ongoing else (_ie_ws_end_dt.strftime("%d/%m/%Y") if _ie_ws_end_dt else "")
                                _ie_recs  = []
                                for _ie_r in st.session_state.projects.to_dict("records"):
                                    if str(_ie_r.get("id", "")) == rid:
                                        _ie_r.update({
                                            "name": _ie_ws_name.strip(),
                                            "client": _ie_ws_client.strip(),
                                            "lead": _ie_ws_lead_name,
                                            "employee": _ie_ws_emp_names,
                                            "status": _ie_ws_status,
                                            "proj_type": "Worksoft",
                                            "start": _ie_start,
                                            "end": _ie_end,
                                            "is_active": _ie_ws_active,
                                            "allocated_hours": _ie_ws_alloc,
                                            "project_lead_email": _ie_ws_lead_email,
                                        })
                                    _ie_recs.append(_ie_r)
                                st.session_state.projects = pd.DataFrame(_ie_recs)
                                save_projects_async(st.session_state.projects)
                                auth.assign_worksoft_employees(
                                    int(float(rid)),
                                    _ie_ws_emp_sel,
                                    {u["id"]: u["name"] for u in _ie_all_users},
                                    _ie_ws_dh_map,
                                )
                                auth.log_audit(cu["id"], cu["name"], "UPDATE", "projects", rid,
                                               f'Updated Worksoft project "{_ie_ws_name.strip()}"')
                                st.session_state[_edit_key] = False
                                st.session_state.toast = {"msg": f'"{_ie_ws_name.strip()}" updated!', "type": "success"}
                                st.rerun()
                        if _ie_sb2.button("Cancel", use_container_width=True, key=f"ie_cancel_{tab_key}_{rid}"):
                            st.session_state[_edit_key] = False
                            st.rerun()

                    else:
                        # ── RPA / standard edit form ──────────────────────────────
                        _ea, _eb = st.columns(2)
                        _e_name   = _ea.text_input("Project Name *", value=row.get("name",""), key=f"ie_name_{tab_key}_{rid}")
                        _cl_opts  = _all_clients + [_CLIENT_NEW]
                        _cl_cur   = row.get("client","")
                        _cl_idx   = _cl_opts.index(_cl_cur) if _cl_cur in _cl_opts else len(_cl_opts)-1
                        _cl_sel   = _eb.selectbox("Client *", _cl_opts, index=_cl_idx, key=f"ie_client_{tab_key}_{rid}")
                        _e_client = _eb.text_input("New client name *", key=f"ie_newclient_{tab_key}_{rid}") if _cl_sel == _CLIENT_NEW else _cl_sel

                        _ec, _ed  = st.columns(2)
                        _ld_opts  = [""] + _all_emps + [_EMP_NEW]
                        _ld_cur   = row.get("lead","")
                        _ld_idx   = _ld_opts.index(_ld_cur) if _ld_cur in _ld_opts else 0
                        _ld_sel   = _ec.selectbox("Lead", _ld_opts, index=_ld_idx, key=f"ie_lead_{tab_key}_{rid}")
                        _e_lead   = _ec.text_input("New lead name", key=f"ie_newlead_{tab_key}_{rid}") if _ld_sel == _EMP_NEW else _ld_sel

                        _st_idx   = ALL_STATUSES.index(row.get("status","")) if row.get("status","") in ALL_STATUSES else 0
                        _e_status = _ed.selectbox("Status", ALL_STATUSES, index=_st_idx, key=f"ie_status_{tab_key}_{rid}")

                        _ee, _ef  = st.columns(2)
                        _PROJ_TYPES = ["", "RPA", "Worksoft", "AI Agent", "Presales"]
                        _pt_cur   = row.get("proj_type","")
                        _pt_idx   = _PROJ_TYPES.index(_pt_cur) if _pt_cur in _PROJ_TYPES else 0
                        _e_type   = _ee.selectbox("Type", _PROJ_TYPES, index=_pt_idx,
                                                  format_func=lambda x: "— Select type —" if x=="" else x,
                                                  key=f"ie_type_{tab_key}_{rid}")
                        _e_active = _ef.checkbox("Active", value=(str(row.get("is_active","True")).strip().lower() not in ["false","0","no"]),
                                                key=f"ie_active_{tab_key}_{rid}")

                        _ws_ie_alloc_val      = float(row.get("allocated_hours") or 0)
                        _ws_ie_lead_email_val = str(row.get("project_lead_email") or "")

                        _emp_raw  = str(row.get("employee",""))
                        _emp_list = [n.strip() for n in _emp_raw.replace("&",",").split(",") if n.strip()]
                        _emp_defs = [e for e in _emp_list if e in _all_emps]
                        _sel_emps = st.multiselect("Employees *", options=_all_emps, default=_emp_defs, key=f"ie_emps_{tab_key}_{rid}")
                        _new_emp  = st.text_input("Add new employee (optional)", key=f"ie_newemp_{tab_key}_{rid}", placeholder="Leave blank if not needed")
                        _e_emp    = ", ".join(_sel_emps + ([_new_emp.strip()] if _new_emp.strip() else []))

                        _eg, _eh, _ei = st.columns(3)
                        _e_start_dt = _eg.date_input("Start Date", value=_parse_dmy(row.get("start","")), key=f"ie_start_{tab_key}_{rid}", format="DD/MM/YYYY")
                        _e_ongoing  = _eh.checkbox("Ongoing", value=not bool(str(row.get("end","")).strip()), key=f"ie_ongoing_{tab_key}_{rid}")
                        _e_end_dt   = None if _e_ongoing else _eh.date_input("End Date", value=_parse_dmy(str(row.get("end","")).strip()) or date.today(), key=f"ie_end_{tab_key}_{rid}", format="DD/MM/YYYY")
                        _e_due_raw  = str(row.get("due_date","")).strip()
                        _e_due_dt   = _ei.date_input("Due Date", value=_parse_dmy(_e_due_raw) if _e_due_raw else None, key=f"ie_due_{tab_key}_{rid}", format="DD/MM/YYYY")

                        _ej, _ek    = st.columns(2)
                        _e_po       = _ej.text_input("PO Number", value=row.get("po",""), key=f"ie_po_{tab_key}_{rid}")
                        _e_desc     = _ek.text_input("Description", value=row.get("desc",""), key=f"ie_desc_{tab_key}_{rid}")

                        st.markdown('<div style="font-size:11px;font-weight:700;color:#64748B;margin:6px 0 2px">ROI Calculator <span style="font-weight:400">(optional)</span></div>', unsafe_allow_html=True)
                        _el, _em2, _en = st.columns(3)
                        _e_mhrs  = _el.text_input("Manual Hrs",  value=row.get("manual_hrs",""),  key=f"ie_mhrs_{tab_key}_{rid}")
                        _e_ahrs  = _em2.text_input("Auto Hrs",    value=row.get("auto_hrs",""),    key=f"ie_ahrs_{tab_key}_{rid}")
                        _e_cph   = _en.text_input("Cost/Hr (₹)", value=row.get("cost_per_hr",""), key=f"ie_cph_{tab_key}_{rid}")
                        _e_roi   = compute_roi(_e_mhrs, _e_ahrs, _e_cph)
                        if _e_roi:
                            st.success(f"ROI: **{_e_roi['pct']}%** | Hrs Saved: **{_e_roi['saved']}** | Cost Saved: **₹{_e_roi['cost']:,.0f}**")

                        with st.expander("Checkpoint Dates (optional)", expanded=False):
                            _ckpt_phases = [("PDD / SDD","pdd_sdd"),("Development","development"),("UAT","uat"),("Deployment","deployment")]
                            _ckpt_vals   = {}
                            for _cfl, _cfk in _ckpt_phases:
                                _cfa2, _cfb2 = st.columns(2)
                                _cf_ds = _parse_dmy(str(row.get(f"ckpt_{_cfk}_start","")).strip())
                                _cf_de = _parse_dmy(str(row.get(f"ckpt_{_cfk}_end",  "")).strip())
                                _cf_sd = _cfa2.date_input(f"{_cfl} Start", value=_cf_ds, key=f"ie_ck_{_cfk}_s_{tab_key}_{rid}", format="DD/MM/YYYY")
                                _cf_ed = _cfb2.date_input(f"{_cfl} End",   value=_cf_de, key=f"ie_ck_{_cfk}_e_{tab_key}_{rid}", format="DD/MM/YYYY")
                                _ckpt_vals[f"ckpt_{_cfk}_start"] = _cf_sd.strftime("%d/%m/%Y") if _cf_sd else ""
                                _ckpt_vals[f"ckpt_{_cfk}_end"]   = _cf_ed.strftime("%d/%m/%Y") if _cf_ed else ""

                        _sb1, _sb2, _sb3, _ = st.columns([1, 1, 1, 3])
                        if _sb1.button("💾 Save", type="primary", use_container_width=True, key=f"ie_save_{tab_key}_{rid}"):
                            _errs = []
                            if not _e_name.strip() or len(_e_name.strip()) < 3:
                                _errs.append("Project name must be at least 3 characters.")
                            if not _e_client.strip():
                                _errs.append("Client is required.")
                            if not _e_emp.strip():
                                _errs.append("At least one employee is required.")
                            if _errs:
                                for _er in _errs: st.error(_er)
                            else:
                                _e_start = _e_start_dt.strftime("%d/%m/%Y") if _e_start_dt else ""
                                _e_end   = "" if _e_ongoing else (_e_end_dt.strftime("%d/%m/%Y") if _e_end_dt else "")
                                _e_due   = _e_due_dt.strftime("%d/%m/%Y") if _e_due_dt else ""
                                _records = []
                                for _r in st.session_state.projects.to_dict("records"):
                                    if str(_r.get("id","")) == rid:
                                        _r.update({
                                            "name": _e_name.strip(), "client": _e_client.strip(),
                                            "lead": _e_lead.strip(), "employee": _e_emp.strip(),
                                            "status": _e_status, "proj_type": _e_type,
                                            "start": _e_start, "end": _e_end, "due_date": _e_due,
                                            "po": _e_po, "desc": _e_desc.strip(),
                                            "manual_hrs": _e_mhrs, "auto_hrs": _e_ahrs, "cost_per_hr": _e_cph,
                                            "hours_saved": str(_e_roi["saved"]) if _e_roi else _r.get("hours_saved",""),
                                            "cost_saved":  str(_e_roi["cost"])  if _e_roi else _r.get("cost_saved",""),
                                            "roi_pct":     str(_e_roi["pct"])   if _e_roi else _r.get("roi_pct",""),
                                            "is_active": _e_active,
                                            "allocated_hours": _ws_ie_alloc_val,
                                            "project_lead_email": _ws_ie_lead_email_val,
                                            **_ckpt_vals,
                                        })
                                    _records.append(_r)
                                st.session_state.projects = pd.DataFrame(_records)
                                save_projects_async(st.session_state.projects)
                                auth.log_audit(cu["id"], cu["name"], "UPDATE", "projects",
                                               rid, f'Updated project "{_e_name.strip()}"')
                                st.session_state[_edit_key] = False
                                st.session_state.toast = {"msg": f'"{_e_name.strip()}" updated!', "type": "success"}
                                if _e_status == "Completed":
                                    st.balloons()
                                st.rerun()
                        if _sb2.button("Cancel", use_container_width=True, key=f"ie_cancel_{tab_key}_{rid}"):
                            st.session_state[_edit_key] = False
                            st.rerun()
                        if _sb3.button("📋 Duplicate", use_container_width=True, key=f"ie_dup_{tab_key}_{rid}"):
                            _dup_all = st.session_state.projects.to_dict("records")
                            _max_id  = max(
                                (int(float(str(r.get("id",0) or 0)))
                                 for r in _dup_all
                                 if str(r.get("id","")).replace(".","",1).isdigit()),
                                default=0
                            )
                            _dup_entry = {str(k): str(v) for k, v in row.items()}
                            _dup_entry["id"]   = _max_id + 1
                            _dup_entry["name"] = pname + " (Copy)"
                            _dup_entry["is_new"] = True
                            _dup_all.append(_dup_entry)
                            st.session_state.projects = pd.DataFrame(_dup_all)
                            save_projects_async(st.session_state.projects)
                            auth.log_audit(cu["id"], cu["name"], "CREATE", "projects",
                                           str(_max_id + 1), f'Duplicated "{pname}"')
                            st.session_state[_edit_key] = False
                            st.session_state.toast = {"msg": f'"{pname}" duplicated!', "type": "success"}
                            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        _exp_c1, _exp_c2, _exp_c3 = st.columns([1, 1, 3])
        csv = filtered.to_csv(index=False)
        _exp_c1.download_button("Export CSV", csv,
                           file_name=f"qualesce_{tab_key}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                           mime="text/csv", key=f"csv_{tab_key}")
        _pdf_bytes = generate_pdf_report(filtered, get_stats(filtered))
        if _pdf_bytes:
            _exp_c2.download_button(
                "Export PDF",
                _pdf_bytes,
                file_name=f"qualesce_{tab_key}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                key=f"pdf_{tab_key}",
            )

    # ── Status sets ────────────────────────────────────────────────────────────
    _DEV_STATUSES        = {"In Progress", "PDD", "Important"}
    _RM_STATUSES           = {"R&M"}
    _COMPLETED_STATUSES    = {"Completed"}
    _UAT_STATUSES          = {"UAT"}
    _DISCONTINUED_STATUSES = {"Discontinued"}

    # ── Top-level department tabs: RPA | Worksoft ──────────────────────────────
    # Worksoft-department users only see the Worksoft tab
    _cu_dept = str(cu.get("department", "") or "")
    _ws_only = (_cu_dept == "Worksoft") and (role not in ("admin",))
    if _ws_only:
        _dept_tab_rpa  = None
        _dept_tab_ws   = st.container()
    else:
        _dept_tab_rpa, _dept_tab_ws = st.tabs(["🔧 RPA", "⚙️ Worksoft"])

    if _dept_tab_rpa is not None:
      with _dept_tab_rpa:
        _rpa_base = (df[df["proj_type"].fillna("").str.strip() == "RPA"]
                     if "proj_type" in df.columns else df)

        if role in ("admin", "lead", "manager"):
            _rpa_add_col, _ = st.columns([1, 6])
            if _rpa_add_col.button("➕ Add RPA Project", key="rpa_add_btn", type="primary", use_container_width=True):
                st.session_state["proj_add_type"] = "RPA"
                st.session_state.show_modal = "add"
                st.rerun()

        if role in ("lead", "manager"):
            tab_mine_rpa, tab_dev, tab_rm, tab_completed, tab_uat, tab_disc = st.tabs([
                "My Projects", "Development", "R&M", "Completed", "UAT", "Discontinued",
            ])
        else:
            tab_mine_rpa = None
            tab_dev, tab_rm, tab_completed, tab_uat, tab_disc = st.tabs([
                "Development", "R&M", "Completed", "UAT", "Discontinued",
            ])

        if tab_mine_rpa is not None:
            with tab_mine_rpa:
                _rpa_cu_name  = str(cu.get("name", "")).strip()
                _rpa_cu_email = str(cu.get("email", "")).strip()
                if "project_lead_email" in _rpa_base.columns:
                    _rpa_mine = _rpa_base[
                        (_rpa_base["lead"].fillna("").str.strip() == _rpa_cu_name) |
                        (_rpa_base["project_lead_email"].fillna("").str.strip() == _rpa_cu_email)
                    ]
                else:
                    _rpa_mine = _rpa_base[
                        _rpa_base["lead"].fillna("").str.strip() == _rpa_cu_name
                    ]
                if _rpa_mine.empty:
                    st.info("No RPA projects are assigned to you as lead.")
                else:
                    with st.container(border=True):
                        _render_project_table(_apply_filters(_rpa_mine), tab_key="rpa_mine", show_timesheet=False)

        with tab_dev:
            with st.container(border=True):
                _render_project_table(
                    _apply_filters(_rpa_base[_rpa_base["status"].isin(_DEV_STATUSES)]),
                    tab_key="rpa_dev", show_timesheet=False)

        with tab_rm:
            with st.container(border=True):
                _render_project_table(
                    _apply_filters(_rpa_base[_rpa_base["status"].isin(_RM_STATUSES)]),
                    tab_key="rpa_rm", show_timesheet=False)

        with tab_completed:
            with st.container(border=True):
                _render_project_table(
                    _apply_filters(_rpa_base[_rpa_base["status"].isin(_COMPLETED_STATUSES)]),
                    tab_key="rpa_comp", show_timesheet=False)

        with tab_uat:
            with st.container(border=True):
                _render_project_table(
                    _apply_filters(_rpa_base[_rpa_base["status"].isin(_UAT_STATUSES)]),
                    tab_key="rpa_uat", show_timesheet=False)

        with tab_disc:
            with st.container(border=True):
                _render_project_table(
                    _apply_filters(_rpa_base[_rpa_base["status"].isin(_DISCONTINUED_STATUSES)]),
                    tab_key="rpa_disc", show_timesheet=False)

    # ── Worksoft tab ────────────────────────────────────────────────────────────
    with _dept_tab_ws:
        _ws_base = (df[df["proj_type"].fillna("").str.strip() == "Worksoft"]
                    if "proj_type" in df.columns
                    else pd.DataFrame(columns=df.columns if not df.empty else []))

        # Add Worksoft Project form
        if role in ("admin", "lead", "manager"):
            with st.expander("➕ Add Worksoft Project", expanded=_ws_base.empty):
                _wsa, _wsb = st.columns(2)
                _ws_client  = _wsa.text_input("Client Name *", key="ws_add_client", placeholder="e.g. TEPL")
                _ws_heading = _wsb.text_input("Project Heading *", key="ws_add_heading", placeholder="e.g. Worksoft Certification")
                _ws_all_active = [u for u in auth.get_all_users() if u.get("is_active")]
                _ws_id_to_name = {u["id"]: f"{u['name']} ({u['email']})" for u in _ws_all_active}
                _ws_emp_sel = st.multiselect(
                    "Assign Employees (multi-select)",
                    options=[u["id"] for u in _ws_all_active],
                    format_func=lambda uid: _ws_id_to_name.get(uid, str(uid)),
                    key="ws_add_emps",
                )
                _ws_daily_hours_map = {}
                if _ws_emp_sel:
                    st.markdown(
                        '<div style="font-size:12px;font-weight:600;color:#475569;margin:6px 0 4px">'
                        'Daily Hours Allocation per Employee</div>',
                        unsafe_allow_html=True,
                    )
                    _ws_dh_cols = st.columns(min(len(_ws_emp_sel), 3))
                    for _dh_i, _dh_uid in enumerate(_ws_emp_sel):
                        _dh_label = _ws_id_to_name.get(_dh_uid, str(_dh_uid)).split("(")[0].strip()
                        _ws_daily_hours_map[_dh_uid] = _ws_dh_cols[_dh_i % 3].number_input(
                            _dh_label, min_value=0.0, max_value=24.0, step=0.5, value=8.0,
                            key=f"ws_add_dh_{_dh_uid}",
                        )
                _ws_lead_users = [u for u in _ws_all_active if u.get("role") in ("lead", "admin", "manager")]
                _wsc, _wsd, _wse = st.columns(3)
                _ws_lead_sel_id = _wsc.selectbox(
                    "Project Lead",
                    options=[u["id"] for u in _ws_lead_users] if _ws_lead_users else [0],
                    format_func=lambda uid: next(
                        (f"{u['name']} ({u['email']})" for u in _ws_lead_users if u["id"] == uid), "—"
                    ),
                    key="ws_add_lead",
                ) if _ws_lead_users else None
                _ws_alloc_hrs = _wsd.number_input(
                    "Allocated Hours", min_value=0.0, step=0.5, value=0.0, key="ws_add_hours"
                )
                _ws_add_status = _wse.selectbox("Status", WS_STATUSES, key="ws_add_status")
                _wsf, _wsg = st.columns(2)
                _ws_start_date = _wsf.date_input(
                    "Start Date", value=None, format="DD/MM/YYYY", key="ws_add_start"
                )
                _ws_end_date = _wsg.date_input(
                    "End Date", value=None, format="DD/MM/YYYY", key="ws_add_end"
                )
                if st.button("Add Worksoft Project", type="primary", key="ws_add_submit"):
                    if not _ws_client.strip():
                        st.error("Client name is required.")
                    elif not _ws_heading.strip():
                        st.error("Project heading is required.")
                    else:
                        _ws_new_id = (int(st.session_state.projects["id"].max()) + 1
                                      if not st.session_state.projects.empty else 1)
                        _ws_emp_names_str = ", ".join(
                            u["name"] for u in _ws_all_active if u["id"] in _ws_emp_sel
                        )
                        _ws_sel_lead_email = next(
                            (u["email"] for u in _ws_lead_users if u["id"] == _ws_lead_sel_id), ""
                        ) if _ws_lead_sel_id else ""
                        _ws_sel_lead_name = next(
                            (u["name"] for u in _ws_lead_users if u["id"] == _ws_lead_sel_id), ""
                        ) if _ws_lead_sel_id else ""
                        _ws_base_rec = {col: "" for col in st.session_state.projects.columns}
                        _ws_base_rec.update({
                            "id": _ws_new_id,
                            "name": _ws_heading.strip(),
                            "client": _ws_client.strip(),
                            "lead": _ws_sel_lead_name,
                            "employee": _ws_emp_names_str,
                            "status": _ws_add_status,
                            "proj_type": "Worksoft",
                            "is_active": True,
                            "is_new": False,
                            "allocated_hours": _ws_alloc_hrs,
                            "project_lead_email": _ws_sel_lead_email,
                            "start": _ws_start_date.strftime("%Y-%m-%d") if _ws_start_date else "",
                            "end":   _ws_end_date.strftime("%Y-%m-%d")   if _ws_end_date   else "",
                        })
                        st.session_state.projects = pd.concat(
                            [st.session_state.projects, pd.DataFrame([_ws_base_rec])],
                            ignore_index=True,
                        )
                        save_projects_async(st.session_state.projects)
                        if _ws_emp_sel:
                            auth.assign_worksoft_employees(
                                _ws_new_id, _ws_emp_sel,
                                {u["id"]: u["name"] for u in _ws_all_active},
                                _ws_daily_hours_map,
                            )
                        auth.log_audit(cu["id"], cu["name"], "ADD", "projects", str(_ws_new_id),
                                       f'Added Worksoft project "{_ws_heading.strip()}"')
                        st.session_state.toast = {"msg": f'"{_ws_heading.strip()}" added!', "type": "success"}
                        st.rerun()

        # Worksoft project list — sub-tabs
        _ws_tab_all, _ws_tab_mine = st.tabs(["All Projects", "My Projects"])

        with _ws_tab_all:
            with st.container(border=True):
                _render_project_table(_apply_filters(_ws_base), tab_key="ws")

        with _ws_tab_mine:
            if not _ws_base.empty:
                _cu_name  = str(cu.get("name", ""))
                _cu_email = str(cu.get("email", ""))
                _ws_mine  = _ws_base[
                    (_ws_base["lead"].fillna("").str.strip() == _cu_name) |
                    (_ws_base["project_lead_email"].fillna("").str.strip() == _cu_email)
                ]
            else:
                _ws_mine = _ws_base
            if _ws_mine.empty:
                st.info("No Worksoft projects are assigned to you as lead.")
            else:
                _wm_badge = (
                    '<span style="font-size:10px;font-weight:700;background:#E0F2FE;'
                    'color:#0369A1;padding:2px 8px;border-radius:4px">Worksoft</span>'
                )
                for _, _wmp_row in _ws_mine.iterrows():
                    _wmp    = _wmp_row.to_dict()
                    _wmp_id = int(float(str(_wmp.get("id", 0) or 0)))
                    with st.container(border=True):
                        _wmp_alloc  = float(_wmp.get("allocated_hours") or 0)
                        _wmp_total  = auth.get_project_total_hours(_wmp_id)
                        _wmp_remain = _wmp_alloc - _wmp_total if _wmp_alloc > 0 else None

                        _wph1, _wph2 = st.columns([4, 1])
                        _wph1.markdown(
                            f'<div style="font-size:15px;font-weight:700;color:#1F3B4D">{esc(str(_wmp.get("name", "")))}</div>'
                            f'<div style="font-size:11px;color:#64748B;margin-top:2px">'
                            f'Client: {esc(str(_wmp.get("client", "") or "—"))}'
                            f'&nbsp;&nbsp;·&nbsp;&nbsp;Status: {esc(str(_wmp.get("status", "") or "—"))}</div>',
                            unsafe_allow_html=True,
                        )
                        _wph2.markdown(_wm_badge, unsafe_allow_html=True)

                        if _wmp_alloc > 0:
                            _wmp_pct    = min((_wmp_total / _wmp_alloc) * 100, 100)
                            _wmp_rem_c  = "#DC2626" if (_wmp_remain or 0) <= 0 else "#16A34A"
                            _wmp_rem_tx = f'{abs(_wmp_remain):.1f}h over' if (_wmp_remain or 0) < 0 else f'{_wmp_remain:.1f}h remaining'
                            _wmp_bar_c  = "#DC2626" if _wmp_pct >= 100 else ("#F59E0B" if _wmp_pct >= 50 else "#3B82F6")
                            st.markdown(
                                f'<div style="margin:6px 0 2px">'
                                f'<div style="display:flex;justify-content:space-between;margin-bottom:3px">'
                                f'<span style="font-size:11px;font-weight:700;color:{_wmp_rem_c}">⏰ {_wmp_rem_tx}</span>'
                                f'<span style="font-size:10px;color:#64748B">Logged <b>{_wmp_total:.2f}h</b> / Budget <b>{_wmp_alloc:.1f}h</b></span>'
                                f'</div>'
                                f'<div style="background:#E2E8F0;border-radius:4px;height:7px">'
                                f'<div style="width:{_wmp_pct:.0f}%;background:{_wmp_bar_c};height:7px;border-radius:4px"></div>'
                                f'</div>'
                                f'<div style="font-size:9px;color:#94A3B8;margin-top:2px">{_wmp_pct:.0f}% of project budget used</div>'
                                f'</div>',
                                unsafe_allow_html=True,
                            )

                        st.markdown('<div style="height:4px"></div>', unsafe_allow_html=True)
                        st.markdown('<span style="font-size:11px;font-weight:700;color:#0369A1">📅 Log Work Hours</span>', unsafe_allow_html=True)
                        _wmp_lf1, _wmp_lf2 = st.columns([1.2, 1])
                        _wmp_date_in = _wmp_lf1.date_input(
                            "Work Date", value=date.today(), format="DD/MM/YYYY",
                            min_value=date(2000, 1, 1), key=f"wmp_date_{_wmp_id}",
                        )
                        _wmp_lf3, _wmp_lf4 = st.columns(2)
                        _wmp_from = _wmp_lf3.time_input("From", value=datetime(2000, 1, 1, 9, 0).time(), key=f"wmp_from_{_wmp_id}")
                        _wmp_to   = _wmp_lf4.time_input("To",   value=datetime(2000, 1, 1, 17, 0).time(), key=f"wmp_to_{_wmp_id}")
                        _wmp_from_dt  = datetime.combine(_wmp_date_in, _wmp_from)
                        _wmp_to_dt    = datetime.combine(_wmp_date_in, _wmp_to)
                        _wmp_calc_hrs = (_wmp_to_dt - _wmp_from_dt).total_seconds() / 3600
                        if _wmp_calc_hrs > 0:
                            st.markdown(
                                f'<div style="font-size:11px;color:#3B82F6;font-weight:600;margin:2px 0">'
                                f'⏱ {_wmp_calc_hrs:.2f}h &nbsp;({_wmp_from.strftime("%H:%M")} → {_wmp_to.strftime("%H:%M")})</div>',
                                unsafe_allow_html=True,
                            )
                        else:
                            st.markdown('<div style="font-size:11px;color:#DC2626;margin:2px 0">⚠ End time must be after start time</div>', unsafe_allow_html=True)
                        _wmp_desc_in = st.text_input(
                            "Description", placeholder="What did you work on?",
                            key=f"wmp_desc_{_wmp_id}", label_visibility="collapsed",
                        )
                        if st.button("📥 Log Hours", key=f"wmp_log_{_wmp_id}", type="primary",
                                     use_container_width=True, disabled=(_wmp_calc_hrs <= 0)):
                            auth.add_worksoft_manual_entry(
                                _wmp_id, str(_wmp.get("name", "")), cu["id"], cu["name"],
                                _wmp_date_in.strftime("%Y-%m-%d"), _wmp_calc_hrs, _wmp_desc_in,
                            )
                            _wmp_new_total = auth.get_project_total_hours(_wmp_id)
                            _wmp_al = float(_wmp.get("allocated_hours") or 0)
                            _wmp_le = str(_wmp.get("project_lead_email", ""))
                            _wmp_thresholds = auth.check_and_log_hours_alert(_wmp_id, _wmp_new_total, _wmp_al)
                            if 50 in _wmp_thresholds and _wmp_le:
                                threading.Thread(target=email_utils.send_worksoft_50pct_alert,
                                    args=(_wmp_le, str(_wmp.get("lead", "")), str(_wmp.get("name", "")), _wmp_new_total, _wmp_al),
                                    daemon=True).start()
                            if 100 in _wmp_thresholds:
                                if _wmp_le:
                                    threading.Thread(target=email_utils.send_worksoft_hours_alert,
                                        args=(_wmp_le, str(_wmp.get("lead", "")), str(_wmp.get("name", "")), _wmp_new_total, _wmp_al),
                                        daemon=True).start()
                                for _wmp_assign in auth.get_worksoft_project_assignments(_wmp_id):
                                    if _wmp_assign.get("email"):
                                        threading.Thread(target=email_utils.send_worksoft_100pct_employee_alert,
                                            args=(_wmp_assign["email"], _wmp_assign["name"], str(_wmp.get("name", "")), _wmp_new_total, _wmp_al),
                                            daemon=True).start()
                            st.session_state.toast = {"msg": f"{_wmp_calc_hrs:.2f}h logged!", "type": "success"}
                            st.rerun()
                        _wmp_entries = auth.get_user_worksoft_entries(_wmp_id, cu["id"])
                        if _wmp_entries:
                            with st.expander(f"My Recent Entries ({len(_wmp_entries)})", expanded=False):
                                for _wme in _wmp_entries[:10]:
                                    _wme_c1, _wme_c2, _wme_c3 = st.columns([1.2, 0.6, 2.5])
                                    _wme_c1.markdown(f'<span style="font-size:11px;color:#64748B">{fmt_date(_wme["work_date"]) or _wme["work_date"]}</span>', unsafe_allow_html=True)
                                    _wme_c2.markdown(f'<span style="font-size:12px;font-weight:700;color:#1F3B4D">{_wme["hours_worked"]:.1f}h</span>', unsafe_allow_html=True)
                                    _wme_c3.markdown(f'<span style="font-size:11px;color:#475569">{esc(_wme["description"])}</span>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB: PRESALES
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "presales" and role not in ("employee",):
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px;font-family:Manrope,sans-serif">Presales / POC</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Presales pipeline and proof-of-concept projects</p>', unsafe_allow_html=True)

    _POC_DEFAULT = {"Presales", "Internal POC", "External POC"}
    _POC_CLIENTS = {"Internal POC", "External POC"}

    PS_ROW_BG = {
        "Important":    "#FCEAEA",
        "Presales":     "#EFF7F7",
        "Internal POC": "#EFF7F7",
        "External POC": "#F7F8F9",
        "Completed":    "#E5F2EC",
        "In Progress":  "#EFF7F7",
        "Discontinued": "#FCEAEA",
    }

    def _ps_type_badge(pt):
        if pt == "RPA":
            return '<span style="font-size:9px;font-weight:700;background:#D9ECEC;color:#3F8E91;padding:1px 6px;border-radius:4px">RPA</span>'
        if pt == "AI Agent":
            return '<span style="font-size:9px;font-weight:700;background:#F3E8FF;color:#7C3AED;padding:1px 6px;border-radius:4px">AI</span>'
        if pt == "Presales":
            return '<span style="font-size:9px;font-weight:700;background:#FEF9C3;color:#854D0E;padding:1px 6px;border-radius:4px">Pre</span>'
        return '<span style="font-size:10px;color:#CBD5E1">—</span>'

    def _render_poc_table(data, tab_key):
        _is_adm = (role in ("admin", "lead", "manager"))
        _can_ed = role in ("admin", "lead", "manager")
        _cw = [10, 0.4, 0.4] if _is_adm else ([10, 0.4] if _can_ed else [10])
        if data.empty:
            st.info("No projects found.")
            return
        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 12px"><b>{len(data)}</b> project(s)</p>',
                    unsafe_allow_html=True)
        _hc = st.columns(_cw)
        _hc[0].markdown(
            f'<div style="display:flex;gap:0;align-items:center">'
            f'<div style="width:3%;{_HDR_STYLE}">ID</div>'
            f'<div style="width:17%;{_HDR_STYLE}">Project Name</div>'
            f'<div style="width:10%;{_HDR_STYLE}">Client</div>'
            f'<div style="width:9%;{_HDR_STYLE}">Lead</div>'
            f'<div style="width:12%;{_HDR_STYLE}">Employee</div>'
            f'<div style="width:6%;{_HDR_STYLE}">Type</div>'
            f'<div style="width:10%;{_HDR_STYLE}">Status</div>'
            f'<div style="width:7%;{_HDR_STYLE}">Start</div>'
            f'<div style="width:7%;{_HDR_STYLE}">End</div>'
            f'<div style="width:8%;{_HDR_STYLE}">Due Date</div>'
            f'<div style="width:11%;{_HDR_STYLE}">Notes</div>'
            f'</div>',
            unsafe_allow_html=True
        )
        if _can_ed: _hc[1].markdown(f'<div style="{_HDR_STYLE}"></div>', unsafe_allow_html=True)
        if _is_adm: _hc[2].markdown(f'<div style="{_HDR_STYLE}"></div>', unsafe_allow_html=True)
        for _, _row in data.iterrows():
            _rstat = str(_row.get("status",""))
            _bg = next((PS_ROW_BG[s] for s in PS_ROW_BG if s in _rstat), "#FFFFFF")
            _new_tag = (' <span style="font-size:9px;font-weight:700;background:#D9ECEC;color:#3F8E91;'
                        'padding:1px 5px;border-radius:4px">NEW</span>') if is_new(_row) else ""
            _lv = str(_row.get("lead","")).strip()
            _lead_html = (f'<span style="font-size:11px;font-weight:600;color:#3F8E91">{esc(_lv)}</span>'
                          if _lv else '<span style="font-size:11px;color:#CBD5E1">—</span>')
            _rid = str(_row.get("id",""))
            _inline_active = (st.session_state.poc_row_edit == _rid)
            _notes_val = str(_row.get("desc","")).strip()
            _notes_disp = (f'<span style="font-size:11px;color:#374151">{esc(_notes_val)}</span>'
                           if _notes_val else '<span style="font-size:11px;color:#CBD5E1">—</span>')
            _rc = st.columns(_cw, vertical_alignment="center")
            _rc[0].markdown(
                f'<div style="display:flex;gap:0;align-items:center;background:{_bg};padding:7px 0;border-bottom:1px solid #F1F5F9">'
                f'<div style="width:3%;font-size:10px;color:#94A3B8">{esc(str(_row.get("id","")))}</div>'
                f'<div style="width:17%;font-size:12px;font-weight:600;color:#111827">{esc(str(_row.get("name","")))}{_new_tag}</div>'
                f'<div style="width:10%;font-size:12px;color:#374151">{esc(str(_row.get("client","")))}</div>'
                f'<div style="width:9%">{_lead_html}</div>'
                f'<div style="width:12%;font-size:11px;color:#374151">{esc(str(_row.get("employee","")))}</div>'
                f'<div style="width:6%">{_ps_type_badge(str(_row.get("proj_type","")).strip())}</div>'
                f'<div style="width:10%">{badge_html(str(_row.get("status","")))}</div>'
                f'<div style="width:7%;font-size:11px;color:#64748B">{esc(fmt_date(str(_row.get("start",""))))}</div>'
                f'<div style="width:7%;font-size:11px;color:#64748B">{esc(fmt_date(str(_row.get("end",""))))}</div>'
                f'<div style="width:8%">{_due_cell(str(_row.get("due_date","")))}</div>'
                f'<div style="width:11%">{_notes_disp}</div>'
                f'</div>',
                unsafe_allow_html=True
            )
            if _can_ed:
                with _rc[1]:
                    _marker = "act-warn-marker" if _inline_active else "act-edit-marker"
                    st.markdown(f'<span class="{_marker}"></span>', unsafe_allow_html=True)
                    if st.button("✕" if _inline_active else "✏", key=f"{tab_key}_edit_{_rid}",
                                 help="Cancel" if _inline_active else "Edit notes",
                                 use_container_width=True):
                        if _inline_active:
                            st.session_state.poc_row_edit = None
                        else:
                            st.session_state.poc_row_edit = _rid
                        st.rerun()
            if _is_adm:
                with _rc[2]:
                    st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
                    if st.button("🗑", key=f"{tab_key}_del_{_rid}", help="Delete", use_container_width=True):
                        st.session_state.confirm_delete = {"id": _rid, "name": str(_row.get("name",""))}
                        st.rerun()
            if _inline_active and _can_ed:
                with st.container():
                    _ic1, _ic2 = st.columns([3, 1])
                    _new_comment = _ic1.text_area(
                        "Notes / Comment", value=_notes_val,
                        key=f"{tab_key}_inline_comment_{_rid}", height=72,
                        label_visibility="collapsed", placeholder="Add notes or comment…"
                    )
                    _b1, _b2, _b3 = _ic2.columns(3)
                    if _b1.button("💾", key=f"{tab_key}_save_cmt_{_rid}", help="Save comment"):
                        _proj_idx = st.session_state.projects.index[
                            st.session_state.projects["id"].astype(str) == _rid
                        ]
                        if len(_proj_idx) > 0:
                            st.session_state.projects.at[_proj_idx[0], "desc"] = _new_comment.strip()
                            save_projects_async(st.session_state.projects)
                        st.session_state.poc_row_edit = None
                        st.session_state.toast = {"msg": "Comment saved!", "type": "success"}
                        st.rerun()
                    if _b2.button("✏", key=f"{tab_key}_full_edit_{_rid}", help="Full edit"):
                        st.session_state.poc_row_edit = None
                        st.session_state.show_modal = {"edit": _row.to_dict()}
                        st.rerun()
                    if _b3.button("✕", key=f"{tab_key}_cancel_cmt_{_rid}", help="Cancel"):
                        st.session_state.poc_row_edit = None
                        st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button("Export CSV", data.to_csv(index=False),
                           file_name=f"qualesce_{tab_key}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                           mime="text/csv", key=f"csv_{tab_key}")

    # ── Create New form (admin / lead / manager only) ─────────────────────────
    if role in ("admin", "lead", "manager"):
        with st.expander("+ Create New Presales / POC Entry", expanded=False):
            with st.container():
                _pc1, _pc2 = st.columns(2)
                _ps_new_name   = _pc1.text_input("Project Name *", key="ps_new_name")
                _ps_new_client = _pc2.text_input("Client Name *",  key="ps_new_client")

                _pc3, _pc4 = st.columns(2)
                _ps_all_emp = sorted(set(
                    n.strip()
                    for raw in df.get("employee", pd.Series(dtype=str)).dropna()
                    for n in str(raw).replace("&", ",").split(",")
                    if n.strip()
                ))
                _ps_all_leads = sorted(set(
                    str(l).strip() for l in df.get("lead", pd.Series(dtype=str)).dropna() if str(l).strip()
                )) if "lead" in df.columns else []

                _ps_new_lead = _pc3.selectbox(
                    "Lead", [""] + _ps_all_leads + ["── Type new ──"],
                    key="ps_new_lead"
                )
                if _ps_new_lead == "── Type new ──":
                    _ps_new_lead = _pc3.text_input("Enter lead name", key="ps_new_lead_txt")

                _ps_new_emp_sel = _pc4.multiselect(
                    "Employee(s)", options=_ps_all_emp, key="ps_new_emp_sel"
                )
                _ps_new_emp_txt = _pc4.text_input(
                    "Add new employee (optional)", key="ps_new_emp_txt",
                    placeholder="leave blank if not needed"
                )
                _ps_new_emp = ", ".join(_ps_new_emp_sel + ([_ps_new_emp_txt.strip()] if _ps_new_emp_txt.strip() else []))

                _pc5, _pc6 = st.columns(2)
                _PS_NEW_TYPES    = ["", "RPA", "AI Agent", "Presales"]
                _ps_new_type     = _pc5.selectbox("Type", _PS_NEW_TYPES, key="ps_new_type",
                                                   format_func=lambda x: "— Select type —" if x == "" else x)
                _PS_NEW_STATUSES = ["Internal POC", "External POC",
                                    "In Progress", "Completed", "Discontinued"]
                _ps_new_status   = _pc6.selectbox("Status", _PS_NEW_STATUSES, key="ps_new_status")

                _pc7, _pc8, _pc9 = st.columns([1.5, 1.5, 2])
                _ps_new_start_dt = _pc7.date_input("Start Date (optional)", value=None,
                                                    key="ps_new_start", format="DD/MM/YYYY")
                _ps_new_end_dt   = _pc8.date_input("End Date (optional)", value=None,
                                                    key="ps_new_end", format="DD/MM/YYYY")
                _ps_new_comment  = _pc9.text_area("Notes / Comment", key="ps_new_comment", height=68)
                _ps_new_start = _ps_new_start_dt.strftime("%d/%m/%Y") if _ps_new_start_dt else ""
                _ps_new_end   = _ps_new_end_dt.strftime("%d/%m/%Y")   if _ps_new_end_dt   else ""

                if st.button("Save New Entry", type="primary", key="ps_new_save"):
                    _ps_errs = []
                    if not _ps_new_name.strip():   _ps_errs.append("Project name is required.")
                    if not _ps_new_client.strip(): _ps_errs.append("Client name is required.")
                    if _ps_errs:
                        for _e in _ps_errs: st.error(_e)
                    else:
                        _ps_row = {
                            "id": st.session_state.next_id,
                            "name": _ps_new_name.strip(),
                            "client": _ps_new_client.strip(),
                            "lead": _ps_new_lead.strip() if _ps_new_lead != "── Type new ──" else "",
                            "employee": _ps_new_emp,
                            "status": _ps_new_status,
                            "proj_type": _ps_new_type,
                            "start": _ps_new_start,
                            "end": _ps_new_end,
                            "po": "", "desc": _ps_new_comment.strip(),
                            "manual_hrs": "", "auto_hrs": "", "cost_per_hr": "",
                            "hours_saved": "", "cost_saved": "", "roi_pct": "",
                            "is_new": True, "is_active": True,
                        }
                        st.session_state.projects = pd.concat(
                            [st.session_state.projects, pd.DataFrame([_ps_row])], ignore_index=True
                        )
                        st.session_state.next_id += 1
                        save_projects_async(st.session_state.projects)
                        st.session_state.toast = {"msg": f'"{_ps_new_name.strip()}" added!', "type": "success"}
                        st.rerun()

    # ── Sub-tabs ──────────────────────────────────────────────────────────────
    _ps_t2, _ps_t3, _ps_t4 = st.tabs(["In Progress", "Completed", "Discontinued"])

    _POC_MASK = df["client"].isin(_POC_CLIENTS) | (df["proj_type"].fillna("").str.strip() == "Presales")
    _ACTIVE_STATUSES = {"In Progress", "Presales", "Internal POC", "External POC"}

    with _ps_t2:
        _ip_df = df[_POC_MASK & df["status"].isin(_ACTIVE_STATUSES)].copy()
        st.markdown('<p style="color:#64748B;font-size:12px;margin:0 0 10px">'
                    'POC / Presales projects currently in development</p>', unsafe_allow_html=True)
        with st.container(border=True):
            _render_poc_table(_ip_df, "poc_ip")

    with _ps_t3:
        _done_df = df[_POC_MASK & (df["status"] == "Completed")].copy()
        st.markdown('<p style="color:#64748B;font-size:12px;margin:0 0 10px">'
                    'Successfully completed POC / Presales projects</p>', unsafe_allow_html=True)
        with st.container(border=True):
            _render_poc_table(_done_df, "poc_done")

    with _ps_t4:
        _disc_df = df[_POC_MASK & (df["status"] == "Discontinued")].copy()
        st.markdown('<p style="color:#64748B;font-size:12px;margin:0 0 10px">'
                    'Discontinued POC / Presales projects</p>', unsafe_allow_html=True)
        with st.container(border=True):
            _render_poc_table(_disc_df, "poc_disc")

# ══════════════════════════════════════════════════════════════════════════════
# TAB: LICENSE
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "license" and role != "employee":
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px;font-family:Manrope,sans-serif">License Management</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Track purchased and sold licenses</p>', unsafe_allow_html=True)

    def _lc_expiry_badge(end_date: str) -> str:
        if not end_date:
            return '<span style="font-size:10px;color:#94A3B8">—</span>'
        try:
            exp = datetime.strptime(end_date, "%Y-%m-%d").date()
            today = datetime.now().date()
            diff  = (exp - today).days
            if diff < 0:
                return (f'<span style="background:#FEF2F2;color:#991B1B;font-size:10px;font-weight:700;'
                        f'padding:2px 8px;border-radius:10px">Expired</span>')
            elif diff <= 30:
                return (f'<span style="background:#FEF2F2;color:#DC2626;font-size:10px;font-weight:700;'
                        f'padding:2px 8px;border-radius:10px">30d: {diff}d left</span>')
            elif diff <= 60:
                return (f'<span style="background:#FFFBEB;color:#B45309;font-size:10px;font-weight:700;'
                        f'padding:2px 8px;border-radius:10px">60d: {diff}d left</span>')
            elif diff <= 90:
                return (f'<span style="background:#FEF3C7;color:#92400E;font-size:10px;font-weight:700;'
                        f'padding:2px 8px;border-radius:10px">90d: {diff}d left</span>')
            else:
                return (f'<span style="background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;'
                        f'padding:2px 8px;border-radius:10px">Active</span>')
        except ValueError:
            return f'<span style="font-size:11px;color:#64748B">{esc(end_date)}</span>'

    def _notif_threshold(days_left) -> str | None:
        if days_left is None:
            return None
        if days_left < 0:
            return "expired"
        elif days_left <= 30:
            return "30d"
        elif days_left <= 60:
            return "60d"
        elif days_left <= 90:
            return "90d"
        return None

    _licenses_all     = auth.get_all_licenses()
    _sold_licenses_all = auth.get_all_sold_licenses()

    # Tool names from purchased licenses (for Sold License dropdown)
    _purchased_tool_names = sorted({l["tool_name"].strip() for l in _licenses_all if l["tool_name"].strip()})

    # ── Auto-send expiry notifications (once per day per session) ────────────
    _today_str = datetime.now().strftime("%Y-%m-%d")
    if st.session_state.lc_last_notif_check != _today_str:
        st.session_state.lc_last_notif_check = _today_str
        _auto_results = []
        for _al in _licenses_all:
            if not _al.get("client_email") or not _al.get("end_date"):
                continue
            try:
                _dl = (datetime.strptime(_al["end_date"], "%Y-%m-%d").date() - datetime.now().date()).days
            except ValueError:
                continue
            _thr = _notif_threshold(_dl)
            if _thr and not auth.has_notification_been_sent(_al["id"], "purchased", _thr):
                _ok, _ = email_utils.send_license_expiry_email(
                    _al["client_email"], _al["tool_name"], _al["tool_name"], _al["end_date"], _dl)
                if _ok:
                    auth.mark_notification_sent(_al["id"], "purchased", _thr)
                    _auto_results.append(f'{_al["tool_name"]} ({_thr})')
        for _sl in _sold_licenses_all:
            if not _sl.get("client_email") or not _sl.get("end_date"):
                continue
            try:
                _dl = (datetime.strptime(_sl["end_date"], "%Y-%m-%d").date() - datetime.now().date()).days
            except ValueError:
                continue
            _thr = _notif_threshold(_dl)
            if _thr and not auth.has_notification_been_sent(_sl["id"], "sold", _thr):
                _ok, _ = email_utils.send_license_expiry_email(
                    _sl["client_email"], _sl["client"], _sl["tool_name"], _sl["end_date"], _dl)
                if _ok:
                    auth.mark_notification_sent(_sl["id"], "sold", _thr)
                    _auto_results.append(f'{_sl["tool_name"]} → {_sl["client"]} ({_thr})')
        if _auto_results:
            st.session_state.toast = {
                "msg": f"Auto-sent {len(_auto_results)} expiry notification(s): {', '.join(_auto_results[:3])}{'…' if len(_auto_results) > 3 else ''}",
                "type": "info"
            }

    _lc_tab1, _lc_tab2 = st.tabs(["Purchased License", "Sold License"])

    # ══════════════════════════════════════════════════════════════════════════
    # SUB-TAB 1 — PURCHASED LICENSE
    # ══════════════════════════════════════════════════════════════════════════
    with _lc_tab1:
        # ── Mail compose form (purchased) ─────────────────────────────────────
        if st.session_state.lc_mail_id is not None:
            _lm_rec = next((x for x in _licenses_all if x["id"] == st.session_state.lc_mail_id), None)
            if _lm_rec:
                try:
                    _lm_dl = (datetime.strptime(_lm_rec["end_date"], "%Y-%m-%d").date() - datetime.now().date()).days if _lm_rec.get("end_date") else None
                except ValueError:
                    _lm_dl = None
                _lm_status = ("Expired" if _lm_dl is not None and _lm_dl < 0 else
                               f"Expiring in {_lm_dl}d" if _lm_dl is not None else "No expiry date")
                with st.container(border=True):
                    st.markdown(
                        f'<div style="font-size:13px;font-weight:700;color:#1F3B4D;margin-bottom:4px">'
                        f'Manage Notification Emails — {esc(_lm_rec["tool_name"])}</div>'
                        f'<div style="font-size:11px;color:#64748B;margin-bottom:12px">'
                        f'Expiry: <b>{_lm_rec["end_date"] or "—"}</b> &nbsp;|&nbsp; Status: <b>{_lm_status}</b></div>',
                        unsafe_allow_html=True
                    )
                    _lm_emails_raw = st.text_area(
                        "Notification Email Addresses",
                        value=_lm_rec.get("client_email", ""),
                        height=90, key="lc_mail_recipients",
                        placeholder="Enter one or more addresses, separated by commas or new lines"
                    )
                    _lmb1, _lmb2, _lmb3, _ = st.columns([1, 1.4, 1, 5])
                    if _lmb1.button("Save", type="primary", key="lc_mail_save"):
                        auth.update_license(
                            _lm_rec["id"], _lm_rec["tool_name"],
                            int(_lm_rec["no_of_licenses"]),
                            _lm_rec["start_date"], _lm_rec["end_date"],
                            _lm_emails_raw.strip()
                        )
                        st.session_state.lc_mail_id = None
                        st.session_state.toast = {"msg": "Email address(es) saved.", "type": "success"}
                        st.rerun()
                    if _lmb2.button("Send", key="lc_mail_send"):
                        import re as _re
                        _lm_addrs = [a.strip() for a in _re.split(r"[,\n;]+", _lm_emails_raw) if a.strip() and "@" in a.strip()]
                        if not _lm_addrs:
                            st.error("Please enter at least one valid email address.")
                        elif _lm_dl is None:
                            st.error("No expiry date set — cannot send notification.")
                        else:
                            _lm_ok, _lm_fail = 0, 0
                            for _addr in _lm_addrs:
                                _ok, _ = email_utils.send_license_expiry_email(
                                    _addr, _lm_rec["tool_name"], _lm_rec["tool_name"], _lm_rec["end_date"], _lm_dl)
                                if _ok: _lm_ok += 1
                                else:   _lm_fail += 1
                            st.session_state.lc_mail_id = None
                            st.session_state.toast = {
                                "msg": f"Sent to {_lm_ok} recipient(s)." + (f" {_lm_fail} failed." if _lm_fail else ""),
                                "type": "success" if _lm_fail == 0 else "warning"
                            }
                            st.rerun()
                    if _lmb3.button("Cancel", key="lc_mail_cancel"):
                        st.session_state.lc_mail_id = None
                        st.rerun()

        # ── Edit form ────────────────────────────────────────────────────────
        if st.session_state.lc_edit_id is not None:
            _lc_rec = next((x for x in _licenses_all if x["id"] == st.session_state.lc_edit_id), None)
            if _lc_rec:
                with st.container(border=True):
                    st.markdown('<div style="font-size:13px;font-weight:700;color:#1F3B4D;margin-bottom:10px">Edit Purchased License</div>', unsafe_allow_html=True)
                    _ec1, _ec2 = st.columns(2)
                    _e_tool  = _ec1.text_input("Tool Name *", value=_lc_rec["tool_name"], key="lc_e_tool")
                    _e_seats = _ec2.number_input("No. of Licenses *", min_value=1, value=int(_lc_rec["no_of_licenses"]), step=1, key="lc_e_seats")
                    _ec3, _ec4 = st.columns(2)
                    _e_start_dt = _ec3.date_input("Start Date", value=_parse_ymd(_lc_rec["start_date"]), key="lc_e_start", format="DD/MM/YYYY")
                    _e_end_dt   = _ec4.date_input("End Date", value=_parse_ymd(_lc_rec["end_date"]), key="lc_e_end", format="DD/MM/YYYY")
                    _e_start = _e_start_dt.strftime("%Y-%m-%d") if _e_start_dt else ""
                    _e_end   = _e_end_dt.strftime("%Y-%m-%d") if _e_end_dt else ""
                    _ec5, _ec6 = st.columns(2)
                    _lc_plan_opts = ["", "Monthly", "Quarterly", "Yearly", "Lifetime"]
                    _e_plan_cur = _lc_rec.get("license_plan", "") or ""
                    _e_plan_idx = _lc_plan_opts.index(_e_plan_cur) if _e_plan_cur in _lc_plan_opts else 0
                    _e_plan  = _ec5.selectbox("License Plan", _lc_plan_opts, index=_e_plan_idx, key="lc_e_plan")
                    _e_email = _ec6.text_input("Notification Email(s)", value=_lc_rec.get("client_email", ""), key="lc_e_email", placeholder="email1@company.com, email2@company.com")
                    _eb1, _eb2 = st.columns([1, 4])
                    if _eb1.button("Save Changes", type="primary", key="lc_save_edit"):
                        if not _e_tool.strip():
                            st.error("Tool name is required.")
                        else:
                            auth.update_license(st.session_state.lc_edit_id, _e_tool, int(_e_seats), _e_start, _e_end, _e_email, _e_plan)
                            save_projects_async(st.session_state.projects)
                            st.session_state.lc_edit_id = None
                            st.session_state.toast = {"msg": "License updated!", "type": "success"}
                            st.rerun()
                    if _eb2.button("Cancel", key="lc_cancel_edit"):
                        st.session_state.lc_edit_id = None
                        st.rerun()

        # ── Add License form ─────────────────────────────────────────────────
        with st.expander("Add Purchased License", expanded=False):
            _lc1, _lc2 = st.columns(2)
            _n_tool  = _lc1.text_input("Tool Name *", key="lc_n_tool")
            _n_seats = _lc2.number_input("No. of Licenses *", min_value=1, value=1, step=1, key="lc_n_seats")
            _lc3, _lc4 = st.columns(2)
            _n_start_dt = _lc3.date_input("Start Date (optional)", value=None, key="lc_n_start", format="DD/MM/YYYY")
            _n_end_dt   = _lc4.date_input("End Date (optional)", value=None, key="lc_n_end", format="DD/MM/YYYY")
            _n_start = _n_start_dt.strftime("%Y-%m-%d") if _n_start_dt else ""
            _n_end   = _n_end_dt.strftime("%Y-%m-%d") if _n_end_dt else ""
            _lc5, _lc6 = st.columns(2)
            _n_plan  = _lc5.selectbox("License Plan", ["", "Monthly", "Quarterly", "Yearly", "Lifetime"], key="lc_n_plan")
            _n_email = _lc6.text_input("Notification Email(s) (for auto expiry alerts)", key="lc_n_email", placeholder="email1@company.com, email2@company.com")
            if st.button("Add License", type="primary", key="lc_add_btn"):
                if not _n_tool.strip():
                    st.error("Tool name is required.")
                else:
                    auth.create_license(_n_tool, int(_n_seats), _n_start, _n_end, _n_email, _n_plan)
                    save_projects_async(st.session_state.projects)
                    st.session_state.toast = {"msg": f'License "{_n_tool}" added!', "type": "success"}
                    st.rerun()

        # ── Purchased License table ──────────────────────────────────────────
        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 12px"><b>{len(_licenses_all)}</b> license(s) tracked</p>', unsafe_allow_html=True)
        if not _licenses_all:
            st.info("No licenses added yet. Use the form above to add one.")
        else:
            with st.container(border=True):
                _lhdr = st.columns([0.3, 2.0, 0.9, 1.1, 1.1, 1.1, 1.3, 0.5, 0.5, 0.5])
                for _lc, _ll in zip(_lhdr, ["#", "Tool Name", "Licenses", "Start", "End", "Plan", "Status", "", "", ""]):
                    _lc.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;'
                                 f'letter-spacing:.5px;padding:8px 4px;border-bottom:2px solid #DFE3E7;white-space:nowrap;background:#F8FAFC">{_ll}</div>',
                                 unsafe_allow_html=True)
                for _lic in _licenses_all:
                    _lr = st.columns([0.3, 2.0, 0.9, 1.1, 1.1, 1.1, 1.3, 0.5, 0.5, 0.5], vertical_alignment="center")
                    _lr[0].markdown(cell(_lic["id"], size="11px", color="#94A3B8"), unsafe_allow_html=True)
                    _lr[1].markdown(f'<span style="font-size:13px;font-weight:700;color:#111827">{esc(_lic["tool_name"])}</span>', unsafe_allow_html=True)
                    _lr[2].markdown(f'<span style="font-size:13px;font-weight:600;color:#3F8E91">{_lic["no_of_licenses"]}</span>', unsafe_allow_html=True)
                    _lr[3].markdown(cell(_lic["start_date"] or "—", size="12px", color="#64748B"), unsafe_allow_html=True)
                    _lr[4].markdown(cell(_lic["end_date"] or "—", size="12px", color="#64748B"), unsafe_allow_html=True)
                    _lp_val = _lic.get("license_plan", "") or "—"
                    _lp_colors = {"Monthly": "#3B82F6", "Quarterly": "#8B5CF6", "Yearly": "#059669", "Lifetime": "#D97706"}
                    _lp_c = _lp_colors.get(_lp_val, "#94A3B8")
                    _lr[5].markdown(f'<span style="font-size:10px;font-weight:700;color:{_lp_c}">{esc(_lp_val)}</span>', unsafe_allow_html=True)
                    _lr[6].markdown(_lc_expiry_badge(_lic["end_date"]), unsafe_allow_html=True)
                    with _lr[7]:
                        st.markdown('<span class="act-warn-marker"></span>', unsafe_allow_html=True)
                        if st.button("✉", key=f"lc_mail_{_lic['id']}", use_container_width=True):
                            st.session_state.lc_mail_id = _lic["id"]
                            st.session_state.lc_edit_id = None
                            st.rerun()
                    if role in ("admin", "lead", "manager"):
                        with _lr[8]:
                            st.markdown('<span class="act-edit-marker"></span>', unsafe_allow_html=True)
                            if st.button("✏", key=f"lc_e_{_lic['id']}", help="Edit license", use_container_width=True):
                                st.session_state.lc_edit_id = _lic["id"]
                                st.session_state.lc_mail_id = None
                                st.session_state.sl_edit_id = None
                                st.rerun()
                        with _lr[9]:
                            st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
                            if st.button("🗑", key=f"lc_d_{_lic['id']}", help="Delete license", use_container_width=True):
                                auth.delete_license(_lic["id"])
                                save_projects_async(st.session_state.projects)
                                st.session_state.toast = {"msg": f'License "{_lic["tool_name"]}" deleted.', "type": "info"}
                                st.rerun()

    # ══════════════════════════════════════════════════════════════════════════
    # SUB-TAB 2 — SOLD LICENSE
    # ══════════════════════════════════════════════════════════════════════════
    with _lc_tab2:
        # ── Edit form ────────────────────────────────────────────────────────
        if st.session_state.sl_edit_id is not None:
            _sl_rec = next((x for x in _sold_licenses_all if x["id"] == st.session_state.sl_edit_id), None)
            if _sl_rec:
                with st.container(border=True):
                    st.markdown('<div style="font-size:13px;font-weight:700;color:#1F3B4D;margin-bottom:10px">Edit Sold License</div>', unsafe_allow_html=True)
                    _se1, _se2 = st.columns(2)
                    _sl_tool_opts = _purchased_tool_names or [""]
                    _sl_e_tool_idx = _sl_tool_opts.index(_sl_rec["tool_name"]) if _sl_rec["tool_name"] in _sl_tool_opts else 0
                    _sl_e_tool   = _se1.selectbox("Tool Name *", _sl_tool_opts, index=_sl_e_tool_idx, key="sl_e_tool")
                    _sl_e_client = _se2.text_input("Client *", value=_sl_rec["client"], key="sl_e_client")
                    _se3, _se4 = st.columns(2)
                    _sl_e_seats  = _se3.number_input("No. of Licenses *", min_value=1, value=int(_sl_rec["no_of_licenses"]), step=1, key="sl_e_seats")
                    _sl_e_notes  = _se4.text_input("Notes", value=_sl_rec["notes"], key="sl_e_notes")
                    _se5, _se6 = st.columns(2)
                    _sl_e_start_dt = _se5.date_input("Start Date", value=_parse_ymd(_sl_rec["start_date"]), key="sl_e_start", format="DD/MM/YYYY")
                    _sl_e_end_dt   = _se6.date_input("End Date", value=_parse_ymd(_sl_rec["end_date"]), key="sl_e_end", format="DD/MM/YYYY")
                    _sl_e_start = _sl_e_start_dt.strftime("%Y-%m-%d") if _sl_e_start_dt else ""
                    _sl_e_end   = _sl_e_end_dt.strftime("%Y-%m-%d") if _sl_e_end_dt else ""
                    _se7, _se8 = st.columns(2)
                    _sl_plan_opts = ["", "Monthly", "Quarterly", "Yearly", "Lifetime"]
                    _sl_e_plan_cur = _sl_rec.get("license_plan", "") or ""
                    _sl_e_plan_idx = _sl_plan_opts.index(_sl_e_plan_cur) if _sl_e_plan_cur in _sl_plan_opts else 0
                    _sl_e_plan  = _se7.selectbox("License Plan", _sl_plan_opts, index=_sl_e_plan_idx, key="sl_e_plan")
                    _sl_e_email = _se8.text_input("Client Email (for expiry notifications)", value=_sl_rec.get("client_email", ""), key="sl_e_email", placeholder="client@company.com")
                    _sb1, _sb2 = st.columns([1, 4])
                    if _sb1.button("Save Changes", type="primary", key="sl_save_edit"):
                        if not _sl_e_tool or not _sl_e_client.strip():
                            st.error("Tool name and client are required.")
                        else:
                            auth.update_sold_license(st.session_state.sl_edit_id, _sl_e_tool,
                                                     _sl_e_client, int(_sl_e_seats),
                                                     _sl_e_start, _sl_e_end, _sl_e_notes,
                                                     _sl_e_email, _sl_e_plan)
                            save_projects_async(st.session_state.projects)
                            st.session_state.sl_edit_id = None
                            st.session_state.toast = {"msg": "Sold license updated!", "type": "success"}
                            st.rerun()
                    if _sb2.button("Cancel", key="sl_cancel_edit"):
                        st.session_state.sl_edit_id = None
                        st.rerun()

        # ── Add Sold License form ────────────────────────────────────────────
        with st.expander("Add Sold License", expanded=False):
            if not _purchased_tool_names:
                st.info("Add at least one purchased license first — the tool name list comes from there.")
            else:
                _sa1, _sa2 = st.columns(2)
                _sl_n_tool   = _sa1.selectbox("Tool Name *", _purchased_tool_names, key="sl_n_tool")
                _sl_n_client = _sa2.text_input("Client *", key="sl_n_client")
                _sa3, _sa4 = st.columns(2)
                _sl_n_seats  = _sa3.number_input("No. of Licenses *", min_value=1, value=1, step=1, key="sl_n_seats")
                _sl_n_notes  = _sa4.text_input("Notes (optional)", key="sl_n_notes")
                _sa5, _sa6 = st.columns(2)
                _sl_n_start_dt = _sa5.date_input("Start Date (optional)", value=None, key="sl_n_start", format="DD/MM/YYYY")
                _sl_n_end_dt   = _sa6.date_input("End Date (optional)", value=None, key="sl_n_end", format="DD/MM/YYYY")
                _sl_n_start = _sl_n_start_dt.strftime("%Y-%m-%d") if _sl_n_start_dt else ""
                _sl_n_end   = _sl_n_end_dt.strftime("%Y-%m-%d") if _sl_n_end_dt else ""
                _sa7, _sa8 = st.columns(2)
                _sl_n_plan  = _sa7.selectbox("License Plan", ["", "Monthly", "Quarterly", "Yearly", "Lifetime"], key="sl_n_plan")
                _sl_n_email = _sa8.text_input("Client Email (for expiry notifications)", key="sl_n_email", placeholder="client@company.com")
                if st.button("Add Sold License", type="primary", key="sl_add_btn"):
                    if not _sl_n_client.strip():
                        st.error("Client is required.")
                    else:
                        auth.create_sold_license(_sl_n_tool, _sl_n_client, int(_sl_n_seats),
                                                 _sl_n_start, _sl_n_end, _sl_n_notes,
                                                 _sl_n_email, _sl_n_plan)
                        save_projects_async(st.session_state.projects)
                        st.session_state.toast = {"msg": f'Sold license "{_sl_n_tool}" added!', "type": "success"}
                        st.rerun()

        # ── Sold License table ───────────────────────────────────────────────
        # Helper: compute days until expiry
        def _days_left(end_date_str: str):
            if not end_date_str:
                return None
            try:
                return (datetime.strptime(end_date_str, "%Y-%m-%d").date() - datetime.now().date()).days
            except ValueError:
                return None

        # ── Mail compose form (shown when a row's Mail button is clicked) ─────
        if st.session_state.sl_mail_id is not None:
            _ml_rec = next((x for x in _sold_licenses_all if x["id"] == st.session_state.sl_mail_id), None)
            if _ml_rec:
                _ml_dl = _days_left(_ml_rec.get("end_date", ""))
                if _ml_dl is None:
                    _ml_status_txt = "No expiry date set"
                elif _ml_dl < 0:
                    _ml_status_txt = f"Expired {abs(_ml_dl)} day(s) ago"
                else:
                    _ml_status_txt = f"Expiring in {_ml_dl} day(s)"
                with st.container(border=True):
                    st.markdown(
                        f'<div style="font-size:13px;font-weight:700;color:#1F3B4D;margin-bottom:4px">'
                        f'Manage Notification Emails — {esc(_ml_rec["tool_name"])}</div>'
                        f'<div style="font-size:11px;color:#64748B;margin-bottom:12px">'
                        f'Client: <b>{esc(_ml_rec["client"])}</b> &nbsp;|&nbsp; '
                        f'Expiry: <b>{_ml_rec["end_date"] or "—"}</b> &nbsp;|&nbsp; '
                        f'Status: <b>{_ml_status_txt}</b></div>',
                        unsafe_allow_html=True
                    )
                    _ml_emails_raw = st.text_area(
                        "Notification Email Addresses",
                        value=_ml_rec.get("client_email", ""),
                        height=90,
                        key="sl_mail_recipients",
                        placeholder="Enter one or more addresses, separated by commas or new lines"
                    )
                    _mb1, _mb2, _mb3, _ = st.columns([1, 1.4, 1, 5])
                    if _mb1.button("Save", type="primary", key="sl_mail_save"):
                        auth.update_sold_license(
                            _ml_rec["id"], _ml_rec["tool_name"], _ml_rec["client"],
                            int(_ml_rec["no_of_licenses"]),
                            _ml_rec["start_date"], _ml_rec["end_date"],
                            _ml_rec["notes"], _ml_emails_raw.strip()
                        )
                        st.session_state.sl_mail_id = None
                        st.session_state.toast = {"msg": "Email address(es) saved.", "type": "success"}
                        st.rerun()
                    if _mb2.button("Send", key="sl_mail_send"):
                        import re as _re
                        _valid_addrs = [a.strip() for a in _re.split(r"[,\n;]+", _ml_emails_raw) if a.strip() and "@" in a.strip()]
                        if not _valid_addrs:
                            st.error("Please enter at least one valid email address.")
                        elif _ml_dl is None:
                            st.error("This license has no expiry date set — cannot send notification.")
                        else:
                            _m_ok, _m_fail, _m_errs = 0, 0, []
                            for _addr in _valid_addrs:
                                _ok, _err = email_utils.send_license_expiry_email(
                                    _addr, _ml_rec["client"],
                                    _ml_rec["tool_name"], _ml_rec["end_date"], _ml_dl
                                )
                                if _ok:   _m_ok += 1
                                else:     _m_fail += 1; _m_errs.append(f"{_addr}: {_err}")
                            st.session_state.sl_mail_id = None
                            st.session_state.toast = {
                                "msg": f"Sent to {_m_ok} recipient(s)." + (f" {_m_fail} failed — {'; '.join(_m_errs[:2])}" if _m_fail else ""),
                                "type": "success" if _m_fail == 0 else "error"
                            }
                            st.rerun()
                    if _mb3.button("Cancel", key="sl_mail_cancel"):
                        st.session_state.sl_mail_id = None
                        st.rerun()

        # ── Send All Notifications button (90d / 30d / expired) ──────────────
        _notifiable = [
            s for s in _sold_licenses_all
            if s.get("end_date") and _days_left(s["end_date"]) is not None
            and _days_left(s["end_date"]) <= 90
        ]
        if _notifiable and role in ("admin", "lead", "manager"):
            _snb_col, _ = st.columns([3, 5])
            if _snb_col.button(f"Send All Expiry Notifications ({len(_notifiable)})", key="sl_send_all", type="primary"):
                st.session_state.sl_mail_id = None
                # Open a combined compose for all expiring licenses
                st.session_state.sl_send_all_trigger = True
                st.rerun()

        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 12px"><b>{len(_sold_licenses_all)}</b> sold license record(s)</p>', unsafe_allow_html=True)
        if not _sold_licenses_all:
            st.info("No sold licenses recorded yet. Use the form above to add one.")
        else:
            with st.container(border=True):
                _slhdr = st.columns([0.3, 1.6, 1.6, 0.8, 1.0, 1.0, 1.0, 1.1, 1.3, 0.5, 0.5, 0.5])
                for _slc, _sll in zip(_slhdr, ["#", "Tool Name", "Client", "Qty", "Start", "End", "Plan", "Status", "Notes", "", "", ""]):
                    _slc.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;'
                                  f'letter-spacing:.5px;padding:8px 4px;border-bottom:2px solid #DFE3E7;white-space:nowrap;background:#F8FAFC">{_sll}</div>',
                                  unsafe_allow_html=True)
                for _sl in _sold_licenses_all:
                    _slr = st.columns([0.3, 1.6, 1.6, 0.8, 1.0, 1.0, 1.0, 1.1, 1.3, 0.5, 0.5, 0.5], vertical_alignment="center")
                    _slr[0].markdown(cell(_sl["id"], size="11px", color="#94A3B8"), unsafe_allow_html=True)
                    _slr[1].markdown(f'<span style="font-size:12px;font-weight:700;color:#111827">{esc(_sl["tool_name"])}</span>', unsafe_allow_html=True)
                    _slr[2].markdown(f'<span style="font-size:12px;color:#374151">{esc(_sl["client"])}</span>', unsafe_allow_html=True)
                    _slr[3].markdown(f'<span style="font-size:13px;font-weight:600;color:#3F8E91">{_sl["no_of_licenses"]}</span>', unsafe_allow_html=True)
                    _slr[4].markdown(cell(_sl["start_date"] or "—", size="12px", color="#64748B"), unsafe_allow_html=True)
                    _slr[5].markdown(cell(_sl["end_date"] or "—", size="12px", color="#64748B"), unsafe_allow_html=True)
                    _slp_val = _sl.get("license_plan", "") or "—"
                    _slp_colors = {"Monthly": "#3B82F6", "Quarterly": "#8B5CF6", "Yearly": "#059669", "Lifetime": "#D97706"}
                    _slp_c = _slp_colors.get(_slp_val, "#94A3B8")
                    _slr[6].markdown(f'<span style="font-size:10px;font-weight:700;color:{_slp_c}">{esc(_slp_val)}</span>', unsafe_allow_html=True)
                    _slr[7].markdown(_lc_expiry_badge(_sl["end_date"]), unsafe_allow_html=True)
                    _slr[8].markdown(cell(_sl["notes"] or "—", size="11px", color="#64748B"), unsafe_allow_html=True)
                    with _slr[9]:
                        st.markdown('<span class="act-warn-marker"></span>', unsafe_allow_html=True)
                        if st.button("✉", key=f"sl_mail_{_sl['id']}", use_container_width=True):
                            st.session_state.sl_mail_id = _sl["id"]
                            st.session_state.sl_edit_id = None
                            st.rerun()
                    if role in ("admin", "lead", "manager"):
                        with _slr[10]:
                            st.markdown('<span class="act-edit-marker"></span>', unsafe_allow_html=True)
                            if st.button("✏", key=f"sl_e_{_sl['id']}", help="Edit sold license", use_container_width=True):
                                st.session_state.sl_edit_id = _sl["id"]
                                st.session_state.sl_mail_id = None
                                st.session_state.lc_edit_id = None
                                st.rerun()
                        with _slr[11]:
                            st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
                            if st.button("🗑", key=f"sl_d_{_sl['id']}", help="Delete sold license", use_container_width=True):
                                auth.delete_sold_license(_sl["id"])
                                save_projects_async(st.session_state.projects)
                                st.session_state.toast = {"msg": f'Sold license deleted.', "type": "info"}
                                st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB: AI AGENT
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "agent" and role in ("admin", "lead", "manager"):
    api_key = get_api_key()

    with st.expander("API Key Settings", expanded=not bool(api_key)):
        _new_key_input = st.text_input(
            "Anthropic API Key",
            value=api_key,
            type="password",
            placeholder="sk-ant-...",
            help="Get your key at console.anthropic.com",
            key="agent_api_key_input",
        )
        _ak1, _ak2 = st.columns([1, 4])
        if _ak1.button("Save API Key", type="primary", key="save_api_key_btn"):
            if _new_key_input.strip():
                auth.save_anthropic_api_key(_new_key_input.strip())
                get_api_key.clear()
                st.session_state.toast = {"msg": "API key saved!", "type": "success"}
                st.rerun()
            else:
                st.error("API key cannot be empty.")
        if _new_key_input.strip() and not api_key:
            api_key = _new_key_input.strip()

    if not api_key:
        st.info("Enter and save your Anthropic API Key in the settings above to use the AI Agent.")
    else:
        # Chat history
        for msg in st.session_state.messages:
            content = md_to_html(msg["content"])
            if msg["role"] == "user":
                st.markdown(_TMPL_USER_MSG.render(content=content), unsafe_allow_html=True)
            else:
                st.markdown(_TMPL_BOT_MSG.render(content=content), unsafe_allow_html=True)

        # Suggested prompts (expanded set)
        st.markdown('<div style="margin:12px 0 6px;font-size:10px;font-weight:600;color:#94A3B8;text-transform:uppercase;letter-spacing:.6px">Suggested Prompts</div>', unsafe_allow_html=True)
        _all_quick_qs = [
            "Which projects are In Progress?",
            "Show team workload summary",
            "How many UAT projects?",
            "List all TEPL projects",
            "What is the ROI formula?",
            "Which projects are overdue?",
            "Show ROI for all completed projects",
            "Who has the most projects?",
            "What are the presales opportunities?",
            "Summarize this month's progress",
            "Which projects have the highest cost savings?",
            "List projects by client",
        ]
        _qrow1, _qrow2, _qrow3 = st.columns(4), st.columns(4), st.columns(4)
        for _qi, (_qcol, _qq) in enumerate(zip(list(_qrow1)+list(_qrow2)+list(_qrow3), _all_quick_qs)):
            if _qcol.button(_qq, key=f"qq_{_qi}", use_container_width=True):
                st.session_state.messages.append({"role":"user","content":_qq})
                st.markdown(_TMPL_TYPING.render(), unsafe_allow_html=True)
                try:
                    reply = call_claude(api_key, st.session_state.messages, df)
                    st.session_state.messages.append({"role":"assistant","content":reply})
                except Exception as e:
                    st.session_state.messages.append({"role":"assistant","content":f"Error: {e}"})
                st.rerun()

        # Chat input
        user_input = st.chat_input("Ask anything about projects, team, ROI…")
        if user_input:
            st.session_state.messages.append({"role":"user","content":user_input})
            st.markdown(_TMPL_TYPING.render(), unsafe_allow_html=True)
            try:
                reply = call_claude(api_key, st.session_state.messages, df)
                st.session_state.messages.append({"role":"assistant","content":reply})
            except Exception as e:
                st.session_state.messages.append({"role":"assistant","content":f"Error: {e}"})
            st.rerun()

        _chat_act1, _chat_act2, _ = st.columns([1, 1, 4])
        if _chat_act1.button("Clear Chat", key="clear_chat"):
            st.session_state.messages = [st.session_state.messages[0]]
            st.rerun()

        # Export chat as text
        if len(st.session_state.messages) > 1:
            _chat_export_lines = []
            for _cm in st.session_state.messages:
                _role_label = "You" if _cm["role"] == "user" else "AI Agent"
                _chat_export_lines.append(f"[{_role_label}]\n{_cm['content']}\n")
            _chat_txt = "\n".join(_chat_export_lines)
            _chat_act2.download_button(
                "Export Chat",
                _chat_txt,
                file_name=f"agent_chat_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain",
                key="export_chat_btn",
            )

# ══════════════════════════════════════════════════════════════════════════════
# TAB: USER MANAGEMENT  (admin only)
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "users" and role == "admin":
    _um_h1, _um_h2 = st.columns([5, 1])
    _um_h1.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px;font-family:Manrope,sans-serif">User Management</h2>', unsafe_allow_html=True)
    _um_h1.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Create accounts, assign roles, and manage password resets</p>', unsafe_allow_html=True)
    if _um_h2.button("Refresh Users", use_container_width=True, key="sync_users_btn"):
        st.session_state.toast = {"msg": "Users loaded from database.", "type": "success"}
        st.rerun()

    _users_cache = auth.get_all_users()

    # ── Export Database ───────────────────────────────────────────────────────
    with st.expander("📦 Export Full Database", expanded=False):
        st.markdown(
            '<p style="color:#64748B;font-size:12px;margin-bottom:12px">'
            'Download the entire database as an Excel workbook (one sheet per table) '
            'or as the raw SQLite <code>.db</code> file.</p>',
            unsafe_allow_html=True,
        )
        _exp_c1, _exp_c2 = st.columns(2)

        # ── Excel export ──────────────────────────────────────────────────────
        with _exp_c1:
            if st.button("📊 Generate Excel Export", key="gen_excel_export", use_container_width=True, type="primary"):
                import io as _exp_io
                _exp_conn = auth.get_conn()
                _exp_buf  = _exp_io.BytesIO()
                try:
                    # Discover every table that actually exists in the DB
                    _exp_cur = _exp_conn.cursor()
                    _exp_cur.execute(
                        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
                    )
                    _exp_tables = [r[0] for r in _exp_cur.fetchall()]
                    with pd.ExcelWriter(_exp_buf, engine="openpyxl") as _xw:
                        for _tbl in _exp_tables:
                            try:
                                _tdf = pd.read_sql_query(f"SELECT * FROM [{_tbl}]", _exp_conn)
                                _tdf.to_excel(_xw, sheet_name=_tbl[:31], index=False)
                            except Exception:
                                pass
                        # Projects in-memory dataframe (not stored as a plain table)
                        try:
                            _proj_sheet = "projects"
                            if _proj_sheet not in _exp_tables and not st.session_state.projects.empty:
                                st.session_state.projects.to_excel(_xw, sheet_name=_proj_sheet, index=False)
                        except Exception:
                            pass
                    _exp_conn.close()
                    st.session_state["_db_excel_bytes"] = _exp_buf.getvalue()
                    _sheet_count = len(_exp_tables) + (1 if "projects" not in _exp_tables else 0)
                    st.session_state.toast = {"msg": f"Excel export ready — {_sheet_count} sheet(s). Click Download.", "type": "success"}
                    st.rerun()
                except Exception as _exp_err:
                    st.error(f"Export failed: {_exp_err}")

            if st.session_state.get("_db_excel_bytes"):
                _exp_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    "⬇️ Download Excel",
                    data=st.session_state["_db_excel_bytes"],
                    file_name=f"qualesce_db_{_exp_ts}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_excel_export",
                    use_container_width=True,
                )

        # ── SQLite .db export ─────────────────────────────────────────────────
        with _exp_c2:
            if st.button("🗄️ Download .db File", key="dl_db_file", use_container_width=True):
                import shutil as _shutil
                import io as _db_io
                try:
                    _db_buf = _db_io.BytesIO()
                    with open(auth.DB_PATH, "rb") as _dbf:
                        _db_buf.write(_dbf.read())
                    _db_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.session_state["_db_sqlite_bytes"] = _db_buf.getvalue()
                    st.session_state["_db_sqlite_name"]  = f"qualesce_{_db_ts}.db"
                    st.rerun()
                except Exception as _db_err:
                    st.error(f"Could not read database: {_db_err}")

            if st.session_state.get("_db_sqlite_bytes"):
                st.download_button(
                    "⬇️ Download SQLite DB",
                    data=st.session_state["_db_sqlite_bytes"],
                    file_name=st.session_state.get("_db_sqlite_name", "qualesce.db"),
                    mime="application/octet-stream",
                    key="dl_sqlite_export",
                    use_container_width=True,
                )

    # ── Outlook Email Settings ────────────────────────────────────────────────
    with st.expander("Outlook Email Settings", expanded=False):
        _cfg = auth.get_email_settings()
        st.markdown(
            '<p style="color:#64748B;font-size:12px;margin-bottom:10px">'
            'Configure the Outlook / Office 365 account used to send password reset codes '
            'and task notifications. These credentials are stored in the local database.</p>',
            unsafe_allow_html=True)
        if _cfg["outlook_email"]:
            st.markdown(
                f'<div style="background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;'
                f'padding:10px 14px;font-size:12px;color:#16A34A;margin-bottom:12px">'
                f'Currently configured: <b>{_cfg["outlook_email"]}</b>'
                f'&nbsp;&nbsp;·&nbsp;&nbsp;Last updated: {_cfg["updated_at"] or "—"}'
                f'</div>',
                unsafe_allow_html=True)
        else:
            st.markdown(
                '<div style="background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;'
                'padding:10px 14px;font-size:12px;color:#92400E;margin-bottom:12px">'
                'Not configured — email notifications are disabled until you save credentials below.'
                '</div>',
                unsafe_allow_html=True)
        _oc1, _oc2 = st.columns(2)
        _new_outlook_email = _oc1.text_input(
            "Outlook / Office 365 Email",
            value=_cfg["outlook_email"],
            placeholder="sender@yourcompany.com",
            key="outlook_email_input")
        _new_outlook_pwd = _oc2.text_input(
            "Password / App Password",
            value=_cfg["outlook_password"],
            type="password",
            placeholder="Enter password or App Password",
            key="outlook_pwd_input",
            help="If MFA is enabled, generate an App Password in your Microsoft account settings.")
        _oc3, _oc4 = st.columns([1, 3])
        if _oc3.button("Save Settings", type="primary", key="save_outlook_cfg"):
            if not _new_outlook_email.strip() or not _new_outlook_pwd.strip():
                st.error("Both email and password are required.")
            else:
                auth.save_email_settings(_new_outlook_email.strip(), _new_outlook_pwd.strip())
                st.session_state.toast = {"msg": "Outlook settings saved!", "type": "success"}
                st.rerun()
        if _oc4.button("Clear / Disable Email", key="clear_outlook_cfg"):
            auth.save_email_settings("", "")
            st.session_state.toast = {"msg": "Email settings cleared.", "type": "info"}
            st.rerun()

    # ── Import users from Excel ───────────────────────────────────────────────
    with st.expander("Import Users from Excel", expanded=False):
        st.markdown(
            '<p style="color:#64748B;font-size:12px;margin-bottom:10px">'
            'Upload an Excel file with columns: <b>Name</b>, <b>Email</b>, <b>Password</b>, '
            '<b>Role</b> (optional — defaults to <i>employee</i>), '
            '<b>Department</b> (optional — <i>RPA</i> or <i>Worksoft</i>). '
            'Existing users (same email) will be skipped.</p>',
            unsafe_allow_html=True
        )
        _tmpl_df = pd.DataFrame([
            {"Name": "Alice Smith", "Email": "alice@example.com", "Password": "Alice@123", "Role": "employee", "Department": "RPA"},
            {"Name": "Bob Lead",    "Email": "bob@example.com",   "Password": "Bob@456",   "Role": "lead",     "Department": "Worksoft"},
        ])
        import io as _io
        _tmpl_buf = _io.BytesIO()
        _tmpl_df.to_excel(_tmpl_buf, index=False, engine="openpyxl")
        st.download_button(
            "Download Template", data=_tmpl_buf.getvalue(),
            file_name="users_import_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_user_tmpl"
        )
        _uploaded = st.file_uploader("Upload Excel", type=["xlsx", "xls"], key="import_users_file")
        if _uploaded:
            try:
                _imp_df = pd.read_excel(_uploaded, dtype=str, engine="openpyxl").fillna("")
                _imp_df.columns = [c.strip() for c in _imp_df.columns]
                _required = {"Name", "Email", "Password"}
                if not _required.issubset(set(_imp_df.columns)):
                    st.error(f"Missing columns. Required: {', '.join(_required)}")
                else:
                    _existing_emails = {u["email"].strip().lower() for u in _users_cache}
                    _created, _skipped, _errors = [], [], []
                    _pw_map = {}
                    _VALID_DEPTS = {"RPA", "Worksoft"}
                    for _, _ir in _imp_df.iterrows():
                        _iname  = str(_ir["Name"]).strip()
                        _iemail = str(_ir["Email"]).strip().lower()
                        _ipass  = str(_ir["Password"]).strip()
                        _irole  = str(_ir.get("Role", "employee")).strip().lower()
                        if _irole not in auth.ROLES:
                            _irole = "employee"
                        _idept_raw = str(_ir.get("Department", "")).strip()
                        _idept = _idept_raw if _idept_raw in _VALID_DEPTS else ""
                        if not _iname or not _iemail or "@" not in _iemail:
                            _errors.append(f"Invalid row: {_iname} / {_iemail}")
                            continue
                        if len(_ipass) < 6:
                            _errors.append(f"Password too short for {_iemail} (min 6 chars)")
                            continue
                        if _iemail in _existing_emails:
                            _skipped.append(_iemail)
                            continue
                        try:
                            auth.create_user(_iname, _iemail, _ipass, _irole, _idept)
                            _pw_map[_iemail] = _ipass
                            _created.append(_iname)
                            _existing_emails.add(_iemail)
                        except Exception as _ex:
                            _errors.append(f"{_iemail}: {_ex}")
                    if _pw_map:
                        pass  # users already saved to DB via auth.create_user()
                    if _created:
                        st.success(f"Created {len(_created)} user(s): {', '.join(_created)}")
                    if _skipped:
                        st.info(f"Skipped {len(_skipped)} existing email(s): {', '.join(_skipped)}")
                    for _e in _errors:
                        st.error(_e)
                    if _created:
                        st.rerun()
            except Exception as _ex:
                st.error(f"Could not read file: {_ex}")

    # ── Department tabs ───────────────────────────────────────────────────────
    def _render_user_dept_tab(dept_label: str, dept_key: str):
        _dept_users = [u for u in _users_cache if u.get("department", "") == dept_label]
        with st.expander(f"➕ Add {dept_label} User", expanded=False):
            with st.container():
                _ua, _ub = st.columns(2)
                _nu_name  = _ua.text_input("Full Name *",     key=f"nu_name_{dept_key}")
                _nu_email = _ub.text_input("Email Address *", key=f"nu_email_{dept_key}")
                _uc2, _ud = st.columns(2)
                _nu_pass  = _uc2.text_input("Password *",     type="password", key=f"nu_pass_{dept_key}")
                _nu_role  = _ud.selectbox("Role",             auth.ROLES,      key=f"nu_role_{dept_key}")
                if st.button(f"Create {dept_label} User", type="primary", key=f"create_user_btn_{dept_key}"):
                    _errs = []
                    if not _nu_name.strip():                          _errs.append("Name is required.")
                    if not _nu_email.strip() or "@" not in _nu_email: _errs.append("Valid email is required.")
                    if not _nu_pass or len(_nu_pass) < 6:             _errs.append("Password must be at least 6 characters.")
                    if _errs:
                        for _e in _errs: st.error(_e)
                    else:
                        try:
                            auth.create_user(_nu_name.strip(), _nu_email.strip(), _nu_pass, _nu_role, dept_label)
                            auth.log_audit(cu["id"], cu["name"], "CREATE", "users", "",
                                           f'Created {dept_label} user "{_nu_name.strip()}" ({_nu_email.strip()}) role={_nu_role}')
                            st.session_state.toast = {"msg": f'User "{_nu_name.strip()}" created!', "type": "success"}
                            st.rerun()
                        except Exception as _ex:
                            st.error(f"Could not create user: {_ex}")

        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 10px"><b>{len(_dept_users)}</b> {dept_label} users</p>', unsafe_allow_html=True)
        with st.container(border=True):
            _uhdr = st.columns([0.3, 1.6, 2.2, 1.0, 0.7, 0.45, 0.45, 0.45, 0.45])
            for _col, _lbl in zip(_uhdr, ["ID", "Name", "Email", "Role", "Active", "", "", "", ""]):
                _col.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;letter-spacing:.5px;padding:8px 4px;border-bottom:2px solid #DFE3E7;white-space:nowrap;background:#F8FAFC">{_lbl}</div>', unsafe_allow_html=True)
            _role_colors = {"admin": "#3F8E91", "lead": "#2E7D5B", "manager": "#966D17", "employee": "#4E5860", "sales": "#5FA9AB"}
            if not _dept_users:
                st.markdown('<div style="padding:16px;color:#94A3B8;font-size:12px;text-align:center">No users in this department yet.</div>', unsafe_allow_html=True)
            for _u in _dept_users:
                _uc = st.columns([0.3, 1.6, 2.2, 1.0, 0.7, 0.45, 0.45, 0.45, 0.45], vertical_alignment="center")
                _uc[0].markdown(cell(_u["id"], size="10px", color="#94A3B8"), unsafe_allow_html=True)
                _uc[1].markdown(f'<span style="font-size:12px;font-weight:600;color:#111827">{esc(_u["name"])}</span>', unsafe_allow_html=True)
                _uc[2].markdown(cell(_u["email"]), unsafe_allow_html=True)
                _rc = _role_colors.get(_u["role"], "#374151")
                _uc[3].markdown(f'<span style="font-size:11px;font-weight:700;color:{_rc}">{_u["role"].upper()}</span>', unsafe_allow_html=True)
                _uc[4].markdown(f'<span style="font-size:11px;font-weight:700;color:{"#10B981" if _u["is_active"] else "#EF4444"}">{"Yes" if _u["is_active"] else "No"}</span>', unsafe_allow_html=True)
                with _uc[5]:
                    st.markdown('<span class="act-edit-marker"></span>', unsafe_allow_html=True)
                    if st.button("✏", key=f"eu_{dept_key}_{_u['id']}", help="Edit user", use_container_width=True):
                        st.session_state.user_edit_id = _u["id"]
                        st.session_state.reset_pwd_uid = None
                        st.rerun()
                with _uc[6]:
                    st.markdown('<span class="act-warn-marker"></span>', unsafe_allow_html=True)
                    if st.button("🔑", key=f"rp_{dept_key}_{_u['id']}", help="Reset password", use_container_width=True):
                        st.session_state.reset_pwd_uid = _u["id"]
                        st.session_state.user_edit_id = None
                        st.rerun()
                _tog_lbl = "🔒" if _u["is_active"] else "🔓"
                _tog_tip = "Deactivate" if _u["is_active"] else "Activate"
                with _uc[7]:
                    st.markdown('<span class="act-warn-marker"></span>', unsafe_allow_html=True)
                    if st.button(_tog_lbl, key=f"tog_{dept_key}_{_u['id']}", help=_tog_tip, use_container_width=True):
                        if _u["id"] != cu["id"]:
                            auth.set_active(_u["id"], not _u["is_active"])
                            _tog_action = "DEACTIVATE" if _u["is_active"] else "ACTIVATE"
                            auth.log_audit(cu["id"], cu["name"], _tog_action, "users",
                                           str(_u["id"]), f'{"Deactivated" if _u["is_active"] else "Activated"} user "{_u["name"]}"')
                            st.session_state.toast = {"msg": f'User {"deactivated" if _u["is_active"] else "activated"}.', "type": "info"}
                            st.rerun()
                        else:
                            st.warning("You cannot deactivate your own account.")
                with _uc[8]:
                    st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
                    if st.button("🗑", key=f"du_{dept_key}_{_u['id']}", help="Delete user", use_container_width=True):
                        if _u["id"] != cu["id"]:
                            auth.delete_user(_u["id"])
                            auth.log_audit(cu["id"], cu["name"], "DELETE", "users",
                                           str(_u["id"]), f'Deleted user "{_u["name"]}"')
                            st.session_state.toast = {"msg": f'User "{_u["name"]}" deleted.', "type": "info"}
                            st.rerun()
                        else:
                            st.warning("You cannot delete your own account.")

    # ── Edit user form (shown when a row's Edit button is clicked) ───────────
    if st.session_state.user_edit_id is not None:
        _eu_all  = _users_cache
        _eu_rec  = next((u for u in _eu_all if u["id"] == st.session_state.user_edit_id), None)
        if _eu_rec:
            with st.container(border=True):
                st.markdown(f'<div style="font-size:13px;font-weight:700;color:#1F3B4D;margin-bottom:10px">Edit User — <span style="color:#3F8E91">{esc(_eu_rec["name"])}</span></div>', unsafe_allow_html=True)
                _ea, _eb = st.columns(2)
                _eu_name  = _ea.text_input("Full Name *",    value=_eu_rec["name"],  key="eu_name")
                _eu_email = _eb.text_input("Email *",        value=_eu_rec["email"], key="eu_email")
                _ec, _ed  = st.columns(2)
                _eu_role  = _ec.selectbox("Role", auth.ROLES,
                                          index=auth.ROLES.index(_eu_rec["role"]) if _eu_rec["role"] in auth.ROLES else 0,
                                          key="eu_role")
                _eu_dept_opts = ["", "RPA", "Worksoft"]
                _eu_dept_idx  = _eu_dept_opts.index(_eu_rec.get("department", "")) if _eu_rec.get("department", "") in _eu_dept_opts else 0
                _eu_dept = _ed.selectbox("Department", _eu_dept_opts, index=_eu_dept_idx, key="eu_dept")
                _es1, _es2 = st.columns([1, 4])
                if _es1.button("Save", type="primary", key="eu_save"):
                    _errs = []
                    if not _eu_name.strip():                           _errs.append("Name is required.")
                    if not _eu_email.strip() or "@" not in _eu_email:  _errs.append("Valid email is required.")
                    if _errs:
                        for _e in _errs: st.error(_e)
                    else:
                        try:
                            auth.update_user(st.session_state.user_edit_id, _eu_name, _eu_email, _eu_role, _eu_dept)
                            auth.log_audit(cu["id"], cu["name"], "UPDATE", "users",
                                           str(st.session_state.user_edit_id),
                                           f'Updated user "{_eu_name.strip()}" role={_eu_role} dept={_eu_dept}')
                            st.session_state.user_edit_id = None
                            st.session_state.toast = {"msg": f'User "{_eu_name.strip()}" updated!', "type": "success"}
                            st.rerun()
                        except Exception as _ex:
                            st.error(f"Could not update user: {_ex}")
                if _es2.button("Cancel", key="eu_cancel"):
                    st.session_state.user_edit_id = None
                    st.rerun()
            st.markdown("---")

    # ── Password reset form (shown when a row's Reset button is clicked) ──────
    _rp_uid = st.session_state.get("reset_pwd_uid")
    if _rp_uid:
        _rp_users = _users_cache
        _rp_user  = next((u for u in _rp_users if u["id"] == _rp_uid), None)
        if _rp_user:
            with st.container(border=True):
                st.markdown(f'<div style="font-size:13px;font-weight:700;color:#1F3B4D;margin-bottom:8px">Reset password for <span style="color:#3F8E91">{esc(_rp_user["name"])}</span></div>', unsafe_allow_html=True)
                rpa, rpb = st.columns([2, 1])
                _new_pwd = rpa.text_input("New Password (min 6 chars)", type="password", key="rp_new_pwd")
                rpb.write("")
                rpc, rpd = st.columns(2)
                if rpc.button("Save Password", type="primary", key="rp_save"):
                    if _new_pwd and len(_new_pwd) >= 6:
                        auth.reset_password(_rp_uid, _new_pwd)
                        st.session_state.reset_pwd_uid = None
                        st.session_state.toast = {"msg": "Password reset successfully!", "type": "success"}
                        st.rerun()
                    else:
                        st.error("Password must be at least 6 characters.")
                if rpd.button("Cancel", key="rp_cancel"):
                    st.session_state.reset_pwd_uid = None
                    st.rerun()
            st.markdown("---")

    _udept_rpa, _udept_ws = st.tabs(["🔧 RPA Users", "⚙️ Worksoft Users"])
    with _udept_rpa:
        _render_user_dept_tab("RPA", "rpa")
    with _udept_ws:
        _render_user_dept_tab("Worksoft", "ws")

# ══════════════════════════════════════════════════════════════════════════════
# TAB: TASKS  (all roles — employees see only their own tasks)
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "settings" and role in ("admin", "lead", "manager"):
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px;font-family:Manrope,sans-serif">Settings</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Configure integrations and notifications</p>', unsafe_allow_html=True)

    with st.expander("Outlook Email Settings", expanded=True):
        _cfg = auth.get_email_settings()
        st.markdown(
            '<p style="color:#64748B;font-size:12px;margin-bottom:10px">'
            'Configure the Outlook / Office 365 account used to send task assignment emails '
            'and notifications. Credentials are stored in the local database.</p>',
            unsafe_allow_html=True)
        if _cfg["outlook_email"]:
            st.markdown(
                f'<div style="background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;'
                f'padding:10px 14px;font-size:12px;color:#16A34A;margin-bottom:12px">'
                f'Currently configured: <b>{_cfg["outlook_email"]}</b>'
                f'&nbsp;&nbsp;·&nbsp;&nbsp;Last updated: {_cfg["updated_at"] or "—"}'
                f'</div>',
                unsafe_allow_html=True)
        else:
            st.markdown(
                '<div style="background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;'
                'padding:10px 14px;font-size:12px;color:#92400E;margin-bottom:12px">'
                'Not configured — task notification emails will not be sent until credentials are saved below.'
                '</div>',
                unsafe_allow_html=True)
        _sc1, _sc2 = st.columns(2)
        _s_email = _sc1.text_input(
            "Outlook / Office 365 Email",
            value=_cfg["outlook_email"],
            placeholder="sender@yourcompany.com",
            key="settings_outlook_email")
        _s_pwd = _sc2.text_input(
            "Password / App Password",
            value=_cfg["outlook_password"],
            type="password",
            placeholder="Enter password or App Password",
            key="settings_outlook_pwd",
            help="If MFA is enabled, generate an App Password in your Microsoft account settings.")
        _sc3, _sc4, _sc5 = st.columns([1, 1, 2])
        if _sc3.button("Save Settings", type="primary", key="settings_save_outlook"):
            if not _s_email.strip() or not _s_pwd.strip():
                st.error("Both email and password are required.")
            else:
                auth.save_email_settings(_s_email.strip(), _s_pwd.strip())
                st.session_state.toast = {"msg": "Outlook settings saved!", "type": "success"}
                st.rerun()
        if _sc4.button("Test Connection", key="settings_test_outlook"):
            if not _s_email.strip() or not _s_pwd.strip():
                st.error("Enter email and password first.")
            else:
                import email_utils as _eu
                _test_ok, _test_err = _eu._smtp_send(
                    _s_email.strip(), _s_pwd.strip(),
                    _s_email.strip(),
                    "Qualesce – SMTP Test",
                    "<p>Your Outlook connection is working correctly.</p>")
                if _test_ok:
                    st.success("Connection successful! Test email sent to your inbox.")
                else:
                    st.error(f"Connection failed: {_test_err}")
        if _sc5.button("Clear / Disable Email", key="settings_clear_outlook"):
            auth.save_email_settings("", "")
            st.session_state.toast = {"msg": "Email settings cleared.", "type": "info"}
            st.rerun()

    if role == "admin":
        with st.expander("Project Types", expanded=False):
            st.markdown(
                '<p style="color:#64748B;font-size:12px;margin-bottom:12px">'
                'Manage the list of project types available when creating or editing projects. '
                'Changes apply immediately.</p>',
                unsafe_allow_html=True)

            _pt_list = auth.get_project_types()

            # ── Existing types table ──────────────────────────────────────────
            if _pt_list:
                _pt_hdr = st.columns([0.6, 2.5, 1.2, 0.7, 0.7])
                for _lbl, _col in zip(["#", "Name", "Color", "Edit", "Delete"], _pt_hdr):
                    _col.markdown(
                        f'<div style="font-size:11px;font-weight:700;color:#64748B;'
                        f'padding-bottom:4px;border-bottom:1px solid #E2E8F0">{_lbl}</div>',
                        unsafe_allow_html=True)

                for _idx, _pt in enumerate(_pt_list):
                    _edit_key = f"pt_edit_open_{_pt['id']}"
                    _pc1, _pc2, _pc3, _pc4, _pc5 = st.columns([0.6, 2.5, 1.2, 0.7, 0.7],
                                                                 vertical_alignment="center")
                    _pc1.markdown(
                        f'<div style="font-size:12px;color:#94A3B8">{_idx + 1}</div>',
                        unsafe_allow_html=True)
                    _pc2.markdown(
                        f'<div style="font-size:13px;font-weight:600;color:#1F3B4D">{esc(_pt["name"])}</div>',
                        unsafe_allow_html=True)
                    _pc3.markdown(
                        f'<div style="display:flex;align-items:center;gap:6px">'
                        f'<div style="width:16px;height:16px;border-radius:4px;background:{esc(_pt["color"])};'
                        f'border:1px solid #E2E8F0;flex-shrink:0"></div>'
                        f'<span style="font-size:11px;color:#64748B;font-family:monospace">{esc(_pt["color"])}</span>'
                        f'</div>',
                        unsafe_allow_html=True)
                    if _pc4.button("Edit", key=f"pt_edit_btn_{_pt['id']}", use_container_width=True):
                        st.session_state[_edit_key] = not st.session_state.get(_edit_key, False)
                        st.rerun()
                    if _pc5.button("Delete", key=f"pt_del_btn_{_pt['id']}", use_container_width=True):
                        auth.delete_project_type(_pt["id"])
                        load_project_types.clear()
                        st.session_state.toast = {"msg": f'Type "{_pt["name"]}" deleted.', "type": "info"}
                        st.rerun()

                    if st.session_state.get(_edit_key, False):
                        with st.container(border=True):
                            st.markdown(
                                f'<div style="font-size:12px;font-weight:600;color:#1F3B4D;margin-bottom:6px">'
                                f'Edit type: {esc(_pt["name"])}</div>',
                                unsafe_allow_html=True)
                            _ec1, _ec2, _ec3 = st.columns([2, 1, 1])
                            _pt_new_name = _ec1.text_input(
                                "Name", value=_pt["name"],
                                key=f"pt_name_inp_{_pt['id']}")
                            _pt_new_color = _ec2.color_picker(
                                "Color", value=_pt["color"],
                                key=f"pt_color_inp_{_pt['id']}")
                            _ec3.markdown('<div style="height:26px"></div>', unsafe_allow_html=True)
                            _save_col, _cancel_col = st.columns([1, 1])
                            if _save_col.button("Save", type="primary",
                                                key=f"pt_save_{_pt['id']}", use_container_width=True):
                                if not _pt_new_name.strip():
                                    st.error("Name cannot be empty.")
                                else:
                                    ok = auth.update_project_type(_pt["id"], _pt_new_name.strip(), _pt_new_color)
                                    if ok:
                                        load_project_types.clear()
                                        st.session_state[_edit_key] = False
                                        st.session_state.toast = {
                                            "msg": f'Type updated to "{_pt_new_name.strip()}".',
                                            "type": "success"}
                                        st.rerun()
                                    else:
                                        st.error(f'A type named "{_pt_new_name.strip()}" already exists.')
                            if _cancel_col.button("Cancel", key=f"pt_cancel_{_pt['id']}", use_container_width=True):
                                st.session_state[_edit_key] = False
                                st.rerun()
            else:
                st.info("No project types defined yet. Add one below.")

            st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)
            st.markdown(
                '<div style="font-size:12px;font-weight:700;color:#1F3B4D;margin-bottom:6px">'
                'Add New Type</div>',
                unsafe_allow_html=True)
            _na1, _na2, _na3 = st.columns([2.5, 1, 1])
            _new_pt_name = _na1.text_input(
                "Type Name", placeholder="e.g. Consulting",
                key="settings_new_pt_name")
            _new_pt_color = _na2.color_picker(
                "Color", value="#3F8E91",
                key="settings_new_pt_color")
            _na3.markdown('<div style="height:26px"></div>', unsafe_allow_html=True)
            if st.button("Add Type", type="primary", key="settings_add_pt"):
                if not _new_pt_name.strip():
                    st.error("Type name cannot be empty.")
                else:
                    ok = auth.add_project_type(_new_pt_name.strip(), _new_pt_color)
                    if ok:
                        load_project_types.clear()
                        st.session_state.toast = {
                            "msg": f'Type "{_new_pt_name.strip()}" added.',
                            "type": "success"}
                        st.rerun()
                    else:
                        st.error(f'A type named "{_new_pt_name.strip()}" already exists.')

    # ── Slack / Teams Webhook ────────────────────────────────────────────────
    with st.expander("Slack / Teams Webhook (optional)", expanded=False):
        st.markdown(
            '<p style="color:#64748B;font-size:12px;margin-bottom:10px">'
            'Add an Incoming Webhook URL to receive project notifications in Slack or Teams. '
            'Leave blank to disable.</p>',
            unsafe_allow_html=True
        )
        _sw_url = auth.get_slack_webhook_url()
        _sw_new = st.text_input(
            "Webhook URL",
            value=_sw_url,
            placeholder="https://hooks.slack.com/services/...",
            key="slack_webhook_input",
            type="password"
        )
        _swc1, _swc2, _ = st.columns([1, 1, 4])
        if _swc1.button("Save Webhook", type="primary", key="save_slack_webhook"):
            auth.save_slack_webhook_url(_sw_new.strip())
            st.session_state.toast = {"msg": "Webhook saved!", "type": "success"}
            st.rerun()
        if _swc2.button("Test Webhook", key="test_slack_webhook"):
            _wh = _sw_new.strip() or _sw_url
            if _wh:
                _ok = auth.send_slack_notification(_wh, f"*Qualesce Dashboard* — test notification from {cu['name']} ✅")
                if _ok:
                    st.success("Test message sent successfully!")
                else:
                    st.error("Failed to send test message. Check the webhook URL.")
            else:
                st.warning("Enter a webhook URL first.")

    # ── Audit Log Viewer ─────────────────────────────────────────────────────
    if role == "admin":
        with st.expander("Audit Log", expanded=False):
            st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:10px">All user actions recorded in the system</p>', unsafe_allow_html=True)
            _al_c1, _al_c2, _al_c3 = st.columns(3)
            _al_table  = _al_c1.selectbox("Table", ["", "projects","tasks","users","licenses","sold_licenses"], key="al_table_filter")
            _al_action = _al_c2.selectbox("Action", ["","CREATE","UPDATE","DELETE","LOGIN"], key="al_action_filter")
            _al_user   = _al_c3.text_input("User name", placeholder="Filter by user…", key="al_user_filter")
            _al_logs = auth.get_audit_logs(table_filter=_al_table, action_filter=_al_action, user_filter=_al_user)
            st.markdown(f'<p style="font-size:11px;color:#94A3B8;margin-bottom:8px"><b>{len(_al_logs)}</b> entries</p>', unsafe_allow_html=True)
            if not _al_logs:
                st.info("No audit entries match the current filters.")
            else:
                _al_hdr = st.columns([0.4, 1.2, 1.5, 1.0, 3.0, 1.2])
                for _h, _l in zip(_al_hdr, ["ID", "User", "Action", "Table", "Description", "Time"]):
                    _h.markdown(f'<div style="{_HDR_STYLE}">{_l}</div>', unsafe_allow_html=True)
                for _ale in _al_logs[:100]:
                    _alc = {"CREATE":"#10B981","UPDATE":"#3B82F6","DELETE":"#EF4444"}.get(_ale["action"],"#94A3B8")
                    _alr = st.columns([0.4, 1.2, 1.5, 1.0, 3.0, 1.2])
                    _alr[0].markdown(f'<span style="font-size:10px;color:#94A3B8">{_ale["id"]}</span>', unsafe_allow_html=True)
                    _alr[1].markdown(f'<span style="font-size:11px;color:#374151">{esc(_ale["user_name"])}</span>', unsafe_allow_html=True)
                    _alr[2].markdown(f'<span style="font-size:10px;font-weight:700;color:{_alc};background:{_alc}18;padding:1px 6px;border-radius:6px">{esc(_ale["action"])}</span>', unsafe_allow_html=True)
                    _alr[3].markdown(f'<span style="font-size:10px;color:#64748B">{esc(_ale["table_name"])}</span>', unsafe_allow_html=True)
                    _alr[4].markdown(f'<span style="font-size:11px;color:#374151">{esc(_ale["description"][:80])}</span>', unsafe_allow_html=True)
                    _alr[5].markdown(f'<span style="font-size:10px;color:#94A3B8">{format_relative_time(_ale["created_at"])}</span>', unsafe_allow_html=True)

    # ── PDF Report Download ───────────────────────────────────────────────────
    with st.expander("Portfolio Reports", expanded=False):
        st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:10px">Download a PDF summary of all projects</p>', unsafe_allow_html=True)
        _full_stats = get_stats(df)
        _rpt_pdf = generate_pdf_report(df, _full_stats)
        if _rpt_pdf:
            st.download_button(
                "Download Portfolio PDF",
                _rpt_pdf,
                file_name=f"qualesce_portfolio_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                key="settings_pdf_dl",
                type="primary",
            )
        else:
            st.info("Install fpdf2 (pip install fpdf2) to enable PDF reports.")

elif st.session_state.active_tab == "tasks":
    _STAT_COLORS = {
        "Not Started": "#94A3B8", "In Progress": "#3B82F6",
        "Completed": "#10B981",   "On Hold": "#F59E0B",
    }

    if role == "employee":
        # ── Employee view: own tasks + progress update ────────────────────────
        _tp = st.session_state.get("task_popup")
        if _tp == "success":
            st.markdown(_POPUP_SUCCESS, unsafe_allow_html=True)
            st.session_state.task_popup = None
        elif _tp == "error":
            st.markdown(_POPUP_ERROR, unsafe_allow_html=True)
            st.session_state.task_popup = None

        _sp = st.session_state.get("save_popup")
        if _sp:
            st.markdown(_POPUP_SAVED, unsafe_allow_html=True)
            st.session_state.save_popup = None

        st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px;font-family:Manrope,sans-serif">My Tasks</h2>', unsafe_allow_html=True)
        st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:12px">Tasks assigned to you — update your progress here</p>', unsafe_allow_html=True)

        _my_tasks     = auth.get_user_tasks(cu["id"])
        _my_active_t  = [t for t in _my_tasks if t.get("status", "") != "Completed"]
        _my_done_t    = [t for t in _my_tasks if t.get("status", "") == "Completed"]

        _emp_ttab1, _emp_ttab2 = st.tabs([
            f"Active ({len(_my_active_t)})",
            f"Completed ({len(_my_done_t)})",
        ])

        with _emp_ttab2:
            if not _my_done_t:
                st.info("No completed tasks yet.")
            else:
                for _dt in _my_done_t:
                    with st.container(border=True):
                        _dtl, _dtr = st.columns([4, 1])
                        with _dtl:
                            st.markdown(
                                f'<div style="font-size:13px;font-weight:700;color:#111827">'
                                f'{esc(_dt["title"])}</div>',
                                unsafe_allow_html=True
                            )
                            _dt_meta = f'Assigned by: <b>{esc(_dt["assigned_by"])}</b>'
                            if _dt.get("due_date"):
                                _dt_meta += f' &nbsp;·&nbsp; Due: <b>{esc(fmt_date(_dt["due_date"]))}</b>'
                            st.markdown(
                                f'<div style="font-size:11px;color:#64748B;margin-top:3px">{_dt_meta}</div>',
                                unsafe_allow_html=True
                            )
                        with _dtr:
                            st.markdown(
                                '<span style="font-size:10px;font-weight:700;background:#D1FAE5;'
                                'color:#065F46;padding:3px 10px;border-radius:12px">✓ Completed</span>',
                                unsafe_allow_html=True
                            )

        with _emp_ttab1:
            if not _my_active_t:
                st.info("No active tasks assigned to you.")
            else:
                st.markdown(f'<p style="color:#64748B;font-size:12px;margin-bottom:12px"><b>{len(_my_active_t)}</b> active task(s)</p>', unsafe_allow_html=True)
            for _t in _my_active_t:
                with st.container(border=True):
                    _tl, _tr = st.columns([3, 1.2])
                    _pct = int(_t["progress"])
                    _bar_c = "#10B981" if _pct == 100 else "#3B82F6"
                    with _tl:
                        st.markdown(f'<div style="font-size:14px;font-weight:700;color:#111827;margin-bottom:4px">{esc(_t["title"])}</div>', unsafe_allow_html=True)
                        if _t["description"]:
                            st.markdown(f'<div style="font-size:12px;color:#64748B;margin-bottom:6px;font-style:italic">{esc(_t["description"])}</div>', unsafe_allow_html=True)
                        _date_meta = f'Assigned by: <b>{esc(_t["assigned_by"])}</b>'
                        if _t.get("start_date"):
                            _date_meta += f' &nbsp;·&nbsp; Start: <b>{esc(fmt_date(_t["start_date"]))}</b>'
                        if _t.get("due_date"):
                            _date_meta += f' &nbsp;·&nbsp; Due: <b>{esc(fmt_date(_t["due_date"]))}</b>'
                        st.markdown(f'<div style="font-size:11px;color:#64748B;margin-bottom:6px">{_date_meta}</div>', unsafe_allow_html=True)
                        st.markdown(f'<div class="progress-bar-outer"><div class="progress-bar-inner" style="width:{_pct}%;background:{_bar_c}"></div></div>'
                                    f'<div style="font-size:10px;color:#64748B;margin-top:2px">{_pct}% complete</div>', unsafe_allow_html=True)
                    with _tr:
                        _new_prog = st.slider("Progress %", 0, 100, _pct, step=5, key=f"prog_{_t['id']}")
                        _stat_idx = auth.TASK_STATUSES.index(_t["status"]) if _t["status"] in auth.TASK_STATUSES else 0
                        _new_stat = st.selectbox("Status", auth.TASK_STATUSES, index=_stat_idx, key=f"stat_{_t['id']}")

                    # ── Weekly comment (inline, no separate button) ────────────
                    _wk_start = auth.get_week_start()
                    _wk_end_dt = date.fromisoformat(_wk_start) + timedelta(days=6)
                    _wk_label = (f"{date.fromisoformat(_wk_start).strftime('%d %b')} – "
                                 f"{_wk_end_dt.strftime('%d %b %Y')}")
                    _existing_wc = auth.get_user_week_comment(_t["id"], cu["id"], _wk_start)
                    if _existing_wc:
                        _wc_ca = str(_existing_wc["created_at"])
                        _wc_disp = fmt_date(_wc_ca[:10]) + " " + _wc_ca[11:16] if len(_wc_ca) >= 16 else _wc_ca
                        st.markdown(
                            f'<div style="font-size:10px;color:#64748B;margin-top:6px;margin-bottom:2px">'
                            f'Weekly note ({_wk_label}) — saved {esc(_wc_disp)}</div>',
                            unsafe_allow_html=True)
                        _wc_input = st.text_area(
                            "Weekly note", value=_existing_wc["comment"], height=60,
                            key=f"wc_{_t['id']}",
                            label_visibility="collapsed")
                    else:
                        st.markdown(
                            f'<div style="font-size:10px;color:#64748B;margin-top:6px;margin-bottom:2px">'
                            f'Weekly note ({_wk_label}) — optional</div>', unsafe_allow_html=True)
                        _wc_input = st.text_area(
                            "Weekly note", height=60, key=f"wc_{_t['id']}",
                            placeholder="Describe your progress this week…",
                            label_visibility="collapsed")

                    _sb1, _sb2 = st.columns([1, 3])
                    if _sb1.button("Save", type="primary", key=f"save_p_{_t['id']}", use_container_width=True):
                        auth.update_task_progress(_t["id"], _new_prog, _new_stat, _t.get("comment", ""))
                        if _wc_input.strip():
                            try:
                                auth.add_task_comment(_t["id"], cu["id"], _wc_input, _wk_start)
                            except Exception as _wc_err:
                                st.session_state.toast = {"msg": f"Progress saved but weekly note failed: {_wc_err}", "type": "error"}
                        if _t.get("assigned_by_email"):
                            _upd_ok, _upd_err = email_utils.send_task_updated_email(
                                _t["assigned_by_email"], _t["assigned_by"],
                                cu["name"], _t["title"], _new_stat, _new_prog, "")
                            if not _upd_ok and _upd_err:
                                st.session_state.toast = {"msg": f"Task saved, but email failed: {_upd_err}", "type": "error"}
                        save_projects(st.session_state.projects)
                        st.session_state.save_popup = "success"
                        st.rerun()

    else:
        # ── Admin / Lead / Manager: create + view all tasks ───────────────────
        st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px;font-family:Manrope,sans-serif">Task Management</h2>', unsafe_allow_html=True)
        st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Assign and track tasks for your team</p>', unsafe_allow_html=True)

        def _render_my_tasks_panel(key_prefix):
            _tp = st.session_state.get("task_popup")
            if _tp == "success":
                st.markdown(_POPUP_SUCCESS, unsafe_allow_html=True)
                st.session_state.task_popup = None
            elif _tp == "error":
                st.markdown(_POPUP_ERROR, unsafe_allow_html=True)
                st.session_state.task_popup = None
            _sp2 = st.session_state.get("save_popup")
            if _sp2:
                st.markdown(_POPUP_SAVED, unsafe_allow_html=True)
                st.session_state.save_popup = None
            # ── Worksoft — Log Hours (leads/managers assigned to Worksoft projects)
            _lead_ws_projs = auth.get_user_worksoft_projects(cu["id"])
            if _lead_ws_projs:
                st.markdown(
                    '<div style="font-size:13px;font-weight:700;color:#0369A1;margin-bottom:8px">'
                    '⏱ Worksoft — Log Hours</div>',
                    unsafe_allow_html=True,
                )
                for _lwp in _lead_ws_projs:
                    with st.container(border=True):
                        _lwp_alloc  = float(_lwp.get("allocated_hours") or 0)
                        _lwp_daily  = float(_lwp.get("daily_hours") or 0)
                        _lwp_total  = auth.get_project_total_hours(_lwp["id"])
                        _lwp_remain = _lwp_alloc - _lwp_total if _lwp_alloc > 0 else None

                        _lwp_daily_tag = (
                            f'&nbsp;·&nbsp;<span style="color:#3B82F6">{_lwp_daily:.1f}h/day</span>'
                            if _lwp_daily > 0 else ""
                        )
                        st.markdown(
                            f'<span style="font-size:13px;font-weight:600;color:#1F3B4D">{esc(_lwp["name"])}</span>'
                            f'&nbsp;&nbsp;<span style="font-size:11px;color:#64748B">{esc(_lwp["client"] or "")}{_lwp_daily_tag}</span>',
                            unsafe_allow_html=True,
                        )
                        _lwp_f1, _lwp_f2 = st.columns(2)
                        _lwp_date_in = _lwp_f1.date_input(
                            "Work Date", value=date.today(), format="DD/MM/YYYY",
                            min_value=date(2000, 1, 1), key=f"lwp_date_{_lwp['id']}",
                        )
                        _lwp_hrs_in  = _lwp_f2.number_input(
                            "Hours", min_value=0.5, max_value=24.0,
                            value=float(_lwp_daily) if _lwp_daily > 0 else 8.0,
                            step=0.5, key=f"lwp_hrs_{_lwp['id']}",
                        )
                        _lwp_desc_in = st.text_input(
                            "Note / Description", placeholder="What did you work on?",
                            key=f"lwp_desc_{_lwp['id']}",
                            label_visibility="collapsed",
                        )
                        if st.button("📥 Log Hours", key=f"{key_prefix}lwp_log_{_lwp['id']}", type="primary", use_container_width=True):
                            auth.add_worksoft_manual_entry(
                                _lwp["id"], _lwp["name"], cu["id"], cu["name"],
                                _lwp_date_in.strftime("%Y-%m-%d"),
                                float(_lwp_hrs_in), _lwp_desc_in,
                            )
                            _lwp_new_total = auth.get_project_total_hours(_lwp["id"])
                            _lwp_al = float(_lwp.get("allocated_hours") or 0)
                            _lwp_lead_email = _lwp.get("project_lead_email", "")
                            _lwp_new_thresholds = auth.check_and_log_hours_alert(_lwp["id"], _lwp_new_total, _lwp_al)
                            if 50 in _lwp_new_thresholds and _lwp_lead_email:
                                threading.Thread(
                                    target=email_utils.send_worksoft_50pct_alert,
                                    args=(_lwp_lead_email, _lwp.get("lead", ""), _lwp["name"], _lwp_new_total, _lwp_al),
                                    daemon=True,
                                ).start()
                            if 100 in _lwp_new_thresholds:
                                if _lwp_lead_email:
                                    threading.Thread(
                                        target=email_utils.send_worksoft_hours_alert,
                                        args=(_lwp_lead_email, _lwp.get("lead", ""), _lwp["name"], _lwp_new_total, _lwp_al),
                                        daemon=True,
                                    ).start()
                                for _lwp_assign in auth.get_worksoft_project_assignments(_lwp["id"]):
                                    if _lwp_assign.get("email"):
                                        threading.Thread(
                                            target=email_utils.send_worksoft_100pct_employee_alert,
                                            args=(_lwp_assign["email"], _lwp_assign["name"], _lwp["name"], _lwp_new_total, _lwp_al),
                                            daemon=True,
                                        ).start()
                            st.session_state.toast = {"msg": f"{_lwp_hrs_in:.1f}h logged on {_lwp_date_in.strftime('%d/%m/%Y')}!", "type": "success"}
                            st.rerun()
                        # Recent entries for this user
                        _lwp_my_entries = auth.get_user_worksoft_entries(_lwp["id"], cu["id"])
                        if _lwp_my_entries:
                            with st.expander(f"My Recent Entries ({len(_lwp_my_entries)})", expanded=False):
                                for _lwe in _lwp_my_entries[:10]:
                                    _lwe_c1, _lwe_c2, _lwe_c3 = st.columns([1.2, 0.6, 2.5])
                                    _lwe_c1.markdown(f'<span style="font-size:11px;color:#64748B">{fmt_date(_lwe["work_date"]) or _lwe["work_date"]}</span>', unsafe_allow_html=True)
                                    _lwe_c2.markdown(f'<span style="font-size:12px;font-weight:700;color:#1F3B4D">{_lwe["hours_worked"]:.1f}h</span>', unsafe_allow_html=True)
                                    _lwe_c3.markdown(f'<span style="font-size:11px;color:#475569">{esc(_lwe["description"])}</span>', unsafe_allow_html=True)
                        if _lwp_remain is not None:
                            _lwp_total2 = auth.get_project_total_hours(_lwp["id"])
                            _lwp_rc    = "#DC2626" if _lwp_remain <= 0 else "#16A34A"
                            _lwp_rtx   = f'{abs(_lwp_remain):.1f}h over' if _lwp_remain < 0 else f'{_lwp_remain:.1f}h left'
                            _lwp_pct   = min((_lwp_total2 / _lwp_alloc) * 100, 100) if _lwp_alloc > 0 else 0
                            _lwp_bar_c = "#DC2626" if _lwp_pct >= 100 else ("#F59E0B" if _lwp_pct >= 50 else "#3B82F6")
                            st.markdown(
                                f'<div style="margin-top:6px;font-size:10px;font-weight:700;color:{_lwp_rc}">'
                                f'⏰ {_lwp_rtx} &nbsp;·&nbsp; '
                                f'<span style="color:#94A3B8;font-weight:400">'
                                f'Logged {_lwp_total2:.2f}h / Budget {_lwp_alloc:.1f}h</span></div>'
                                f'<div style="margin-top:3px;background:#E2E8F0;border-radius:4px;height:5px">'
                                f'<div style="width:{_lwp_pct:.0f}%;background:{_lwp_bar_c};height:5px;border-radius:4px"></div>'
                                f'</div><div style="font-size:9px;color:#94A3B8;margin-top:1px">'
                                f'{_lwp_pct:.0f}% of budget</div>',
                                unsafe_allow_html=True,
                            )
                st.markdown('<hr style="margin:10px 0 14px">', unsafe_allow_html=True)
            _my_tasks = auth.get_user_tasks(cu["id"])
            if not _my_tasks:
                st.info("No tasks assigned to you yet.")
            else:
                st.markdown(f'<p style="color:#64748B;font-size:12px;margin-bottom:12px"><b>{len(_my_tasks)}</b> task(s) assigned to you</p>', unsafe_allow_html=True)
                for _t in _my_tasks:
                    with st.container(border=True):
                        _tl, _tr = st.columns([3, 1.2])
                        _pct   = int(_t["progress"])
                        _bar_c = "#10B981" if _pct == 100 else "#3B82F6"
                        with _tl:
                            st.markdown(f'<div style="font-size:14px;font-weight:700;color:#111827;margin-bottom:4px">{esc(_t["title"])}</div>', unsafe_allow_html=True)
                            if _t["description"]:
                                st.markdown(f'<div style="font-size:12px;color:#64748B;margin-bottom:6px;font-style:italic">{esc(_t["description"])}</div>', unsafe_allow_html=True)
                            _lmeta = f'Assigned by: <b>{esc(_t["assigned_by"])}</b>'
                            if _t.get("start_date"):
                                _lmeta += f' &nbsp;·&nbsp; Start: <b>{esc(fmt_date(_t["start_date"]))}</b>'
                            if _t.get("due_date"):
                                _lmeta += f' &nbsp;·&nbsp; Due: <b>{esc(fmt_date(_t["due_date"]))}</b>'
                            st.markdown(f'<div style="font-size:11px;color:#64748B;margin-bottom:6px">{_lmeta}</div>', unsafe_allow_html=True)
                            st.markdown(
                                f'<div class="progress-bar-outer"><div class="progress-bar-inner" style="width:{_pct}%;background:{_bar_c}"></div></div>'
                                f'<div style="font-size:10px;color:#64748B;margin-top:2px">{_pct}% complete</div>',
                                unsafe_allow_html=True)
                        with _tr:
                            _new_prog = st.slider("Progress %", 0, 100, _pct, step=5, key=f"{key_prefix}prog_{_t['id']}")
                            _stat_idx = auth.TASK_STATUSES.index(_t["status"]) if _t["status"] in auth.TASK_STATUSES else 0
                            _new_stat = st.selectbox("Status", auth.TASK_STATUSES, index=_stat_idx, key=f"{key_prefix}stat_{_t['id']}")

                    # ── Weekly comment (inline) ────────────────────────────────
                    _wk_start = auth.get_week_start()
                    _wk_end_dt = date.fromisoformat(_wk_start) + timedelta(days=6)
                    _wk_label = (f"{date.fromisoformat(_wk_start).strftime('%d %b')} – "
                                 f"{_wk_end_dt.strftime('%d %b %Y')}")
                    _existing_wc = auth.get_user_week_comment(_t["id"], cu["id"], _wk_start)
                    if _existing_wc:
                        _wc_ca2 = str(_existing_wc["created_at"])
                        _wc_disp2 = fmt_date(_wc_ca2[:10]) + " " + _wc_ca2[11:16] if len(_wc_ca2) >= 16 else _wc_ca2
                        st.markdown(
                            f'<div style="font-size:10px;color:#64748B;margin-top:6px;margin-bottom:2px">'
                            f'Weekly note ({_wk_label}) — saved {esc(_wc_disp2)}</div>',
                            unsafe_allow_html=True)
                        _wc_input2 = st.text_area(
                            "Weekly note", value=_existing_wc["comment"], height=60,
                            key=f"{key_prefix}wc_{_t['id']}",
                            label_visibility="collapsed")
                    else:
                        st.markdown(
                            f'<div style="font-size:10px;color:#64748B;margin-top:6px;margin-bottom:2px">'
                            f'Weekly note ({_wk_label}) — optional</div>', unsafe_allow_html=True)
                        _wc_input2 = st.text_area(
                            "Weekly note", height=60, key=f"{key_prefix}wc_{_t['id']}",
                            placeholder="Describe your progress this week…",
                            label_visibility="collapsed")

                    _sb1, _sb2 = st.columns([1, 3])
                    if _sb1.button("Save", type="primary", key=f"{key_prefix}save_{_t['id']}", use_container_width=True):
                        auth.update_task_progress(_t["id"], _new_prog, _new_stat, _t.get("comment", ""))
                        if _wc_input2.strip():
                            try:
                                auth.add_task_comment(_t["id"], cu["id"], _wc_input2, _wk_start)
                            except Exception as _wc_err2:
                                st.session_state.toast = {"msg": f"Progress saved but weekly note failed: {_wc_err2}", "type": "error"}
                        save_projects(st.session_state.projects)
                        st.session_state.save_popup = "success"
                        st.rerun()

        def _render_all_tasks_panel(dept=""):
            _dept_sfx = dept.lower() or "all"
            # ── Popup overlay (success / error from previous action) ───────────
            _tp = st.session_state.get("task_popup")
            if _tp == "success":
                st.markdown(_POPUP_SUCCESS, unsafe_allow_html=True)
                st.session_state.task_popup = None
            elif _tp == "error":
                st.markdown(_POPUP_ERROR, unsafe_allow_html=True)
                st.session_state.task_popup = None

            with st.expander(f"Assign New {dept} Task" if dept else "Assign New Task", expanded=False):
                # Show which email will be used to send the task notification
                _preview_cfg = auth.get_email_settings()
                if _preview_cfg["outlook_email"]:
                    st.markdown(
                        f'<div style="background:#EFF7F7;border:1px solid #B6DADB;border-radius:8px;'
                        f'padding:8px 14px;font-size:11px;color:#3F8E91;margin-bottom:10px">'
                        f'Notification email will be sent from: <b>{_preview_cfg["outlook_email"]}</b></div>',
                        unsafe_allow_html=True)
                else:
                    st.markdown(
                        '<div style="background:#FFFBEB;border:1px solid #FDE68A;border-radius:8px;'
                        'padding:8px 14px;font-size:11px;color:#92400E;margin-bottom:10px">'
                        'No email configured — task will be assigned but no notification email will be sent. '
                        'Go to Settings tab to add Outlook credentials.</div>',
                        unsafe_allow_html=True)
                _assignable = auth.get_employees_and_leads()
                if not _assignable:
                    st.warning("No employee or lead accounts found. Create users under the Users tab first.")
                else:
                    _ta1, _ta2 = st.columns(2)
                    _nt_title = _ta1.text_input("Task Title *", key=f"nt_title_{_dept_sfx}")
                    _emp_opts  = [f"{_e['name']}  [{_e['role'].upper()}]  ({_e['email']})" for _e in _assignable]
                    _emp_sel   = _ta2.selectbox("Assign To *", _emp_opts, key=f"nt_emp_{_dept_sfx}",
                                                help="Employees and Leads are listed here.")
                    _nt_desc   = st.text_area("Description (optional)", key=f"nt_desc_{_dept_sfx}", height=80)
                    _ta3, _ta4, _ta5 = st.columns(3)
                    _nt_start_dt = _ta3.date_input("Start Date (optional)", value=None, key=f"nt_start_dt_{_dept_sfx}", format="DD/MM/YYYY")
                    _nt_due_dt   = _ta4.date_input("Due Date (optional)", value=None, key=f"nt_due_dt_{_dept_sfx}", format="DD/MM/YYYY")
                    _ta5.text_input("Assigned By", value=cu["name"], disabled=True, key=f"nt_assigned_by_{_dept_sfx}")
                    _nt_start = _nt_start_dt.strftime("%Y-%m-%d") if _nt_start_dt else ""
                    _nt_due   = _nt_due_dt.strftime("%Y-%m-%d") if _nt_due_dt else ""
                    if st.button("Assign Task", type="primary", key=f"assign_task_btn_{_dept_sfx}"):
                        if not _nt_title.strip():
                            st.error("Task title is required.")
                        else:
                            _sel_idx = _emp_opts.index(_emp_sel)
                            _sel_emp = _assignable[_sel_idx]
                            # Show loading overlay immediately (streams to browser before work starts)
                            _loading_ph = st.empty()
                            _loading_ph.markdown(_POPUP_LOADING, unsafe_allow_html=True)
                            try:
                                auth.create_task(_nt_title, _nt_desc or "", _sel_emp["id"], cu["id"], _nt_due.strip(), _nt_start.strip(), dept)
                                auth.log_audit(cu["id"], cu["name"], "CREATE", "tasks", "",
                                               f'Created {dept} task "{_nt_title}" assigned to {_sel_emp["name"]}')
                                save_projects_async(st.session_state.projects)
                                # Email sends in background — does not block task creation
                                _emp_email = _sel_emp["email"]
                                _emp_name  = _sel_emp["name"]
                                _by_name   = cu["name"]
                                _due_str   = _nt_due.strip()
                                _title_str = _nt_title
                                # Pre-fetch credentials in main thread (st.secrets is
                                # unreliable from background threads)
                                _s_cfg = auth.get_email_settings()
                                _s_e   = _s_cfg["outlook_email"]
                                _s_p   = _s_cfg["outlook_password"]
                                if not _s_e:
                                    _s_e = os.environ.get("OUTLOOK_EMAIL", "")
                                    _s_p = os.environ.get("OUTLOOK_PASSWORD", "")
                                def _send_mail(_se=_s_e, _sp=_s_p):
                                    ok, err = email_utils.send_task_assigned_email(
                                        _emp_email, _emp_name, _title_str, _by_name,
                                        _due_str, _se, _sp)
                                    if not ok and err:
                                        email_utils._error_queue.put(err)
                                threading.Thread(target=_send_mail, daemon=True).start()
                                st.session_state.task_popup = "success"
                            except Exception:
                                st.session_state.task_popup = "error"
                            st.rerun()

            _raw_tasks = auth.get_all_tasks()
            _all_tasks = [t for t in _raw_tasks if t.get("department", "") == dept] if dept else _raw_tasks
            st.markdown(
                f'<p style="color:#64748B;font-size:12px;margin:6px 0 12px">'
                f'<b>{len(_all_tasks)}</b> {dept + " " if dept else ""}tasks</p>',
                unsafe_allow_html=True)

            # ── Comment date-range filter ──────────────────────────────────────
            with st.container(border=True):
                st.markdown('<div style="font-size:11px;font-weight:700;color:#475569;margin-bottom:8px">Weekly Comment Filter</div>', unsafe_allow_html=True)
                _cf1, _cf2 = st.columns(2)
                _cm_from_dt = _cf1.date_input(
                    "From (week start)", key=f"cm_from_{_dept_sfx}",
                    value=date.today() - timedelta(weeks=4),
                    format="DD/MM/YYYY")
                _cm_to_dt = _cf2.date_input(
                    "To (week start)", key=f"cm_to_{_dept_sfx}",
                    value=date.today(),
                    format="DD/MM/YYYY")
                _cm_from_str = _cm_from_dt.strftime("%Y-%m-%d") if _cm_from_dt else None
                _cm_to_str   = _cm_to_dt.strftime("%Y-%m-%d") if _cm_to_dt else None

            if not _all_tasks:
                st.info(f"No {dept + ' ' if dept else ''}tasks yet. Use the form above to assign tasks to employees.")
            else:
                def _render_task_rows(tasks, tab_sfx):
                    if not tasks:
                        st.info("No tasks in this category.")
                        return
                    _thdr = st.columns([2.0, 1.6, 1.5, 1.4, 0.9, 1.0, 1.0, 0.4, 0.4])
                    for _col, _lbl in zip(_thdr, ["Task", "Assigned To", "Assigned By", "Status", "Progress", "Start Date", "Due Date", "", ""]):
                        _col.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;letter-spacing:.5px;padding:8px 4px;border-bottom:2px solid #DFE3E7;white-space:nowrap;background:#F8FAFC">{_lbl}</div>', unsafe_allow_html=True)

                    for _t in tasks:
                        _tc = st.columns([2.0, 1.6, 1.5, 1.4, 0.9, 1.0, 1.0, 0.4, 0.4], vertical_alignment="center")
                        _tdesc = _t["description"]
                        _tdesc_short = (_tdesc[:50] + "…") if len(_tdesc) > 50 else _tdesc
                        _tc[0].markdown(
                            f'<span style="font-size:12px;font-weight:600;color:#111827">{esc(_t["title"])}</span>'
                            + (f'<br><span style="font-size:10px;color:#64748B">{esc(_tdesc_short)}</span>' if _tdesc_short else ""),
                            unsafe_allow_html=True)
                        _tc[1].markdown(
                            f'<span style="font-size:12px">{esc(_t["assigned_to"])}</span>'
                            f'<br><span style="font-size:10px;color:#64748B">{esc(_t["assigned_to_email"])}</span>',
                            unsafe_allow_html=True)
                        _tc[2].markdown(
                            f'<span style="font-size:12px;color:#374151">{esc(_t["assigned_by"])}</span>',
                            unsafe_allow_html=True)
                        _tsc = _STAT_COLORS.get(_t["status"], "#94A3B8")
                        _tc[3].markdown(f'<span style="font-size:11px;font-weight:700;color:{_tsc}">{esc(_t["status"])}</span>', unsafe_allow_html=True)
                        _tpct = int(_t["progress"])
                        _tbar = "#10B981" if _tpct == 100 else "#3B82F6"
                        _tc[4].markdown(
                            f'<div class="progress-bar-outer"><div class="progress-bar-inner" style="width:{_tpct}%;background:{_tbar}"></div></div>'
                            f'<div style="font-size:10px;color:#64748B">{_tpct}%</div>',
                            unsafe_allow_html=True)
                        _tc[5].markdown(cell(fmt_date(_t.get("start_date") or "") or "—", size="11px", color="#64748B"), unsafe_allow_html=True)
                        _tc[6].markdown(cell(fmt_date(_t["due_date"]) or "—", size="11px", color="#64748B"), unsafe_allow_html=True)
                        _edit_key = f"editing_task_{tab_sfx}_{_t['id']}"
                        with _tc[7]:
                            st.markdown('<span class="act-edit-marker"></span>', unsafe_allow_html=True)
                            if st.button("✏", key=f"edit_btn_{tab_sfx}_{_t['id']}", help="Edit task", use_container_width=True):
                                st.session_state[_edit_key] = not st.session_state.get(_edit_key, False)
                                st.rerun()
                        with _tc[8]:
                            st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
                            if st.button("🗑", key=f"dt_{tab_sfx}_{_t['id']}", help="Delete task", use_container_width=True):
                                auth.log_audit(cu["id"], cu["name"], "DELETE", "tasks",
                                               str(_t["id"]), f'Deleted task "{_t["title"]}"')
                                auth.delete_task(_t["id"])
                                save_projects_async(st.session_state.projects)
                                st.session_state.toast = {"msg": "Task deleted.", "type": "info"}
                                st.rerun()

                        # ── Inline edit form ───────────────────────────────────
                        if st.session_state.get(_edit_key, False):
                            with st.container(border=True):
                                st.markdown('<div style="font-size:12px;font-weight:700;color:#3F8E91;margin-bottom:8px">Edit Task</div>', unsafe_allow_html=True)
                                _ea, _eb = st.columns(2)
                                _e_title = _ea.text_input("Title *", value=_t["title"], key=f"etitle_{tab_sfx}_{_t['id']}")
                                _e_desc  = st.text_area("Description", value=_t.get("description", ""), key=f"edesc_{tab_sfx}_{_t['id']}", height=70)
                                _ec, _ed = st.columns(2)
                                _e_start_val = date.fromisoformat(_t["start_date"]) if _t.get("start_date") else None
                                _e_due_val   = date.fromisoformat(_t["due_date"])   if _t.get("due_date")   else None
                                _e_start_dt  = _ec.date_input("Start Date", value=_e_start_val, key=f"estart_{tab_sfx}_{_t['id']}", format="DD/MM/YYYY")
                                _e_due_dt    = _ed.date_input("Due Date",   value=_e_due_val,   key=f"edue_{tab_sfx}_{_t['id']}",   format="DD/MM/YYYY")
                                _es1, _es2 = st.columns([1, 5])
                                if _es1.button("Save", type="primary", key=f"esave_{tab_sfx}_{_t['id']}"):
                                    if not _e_title.strip():
                                        st.error("Title cannot be empty.")
                                    else:
                                        _e_start_str = _e_start_dt.strftime("%Y-%m-%d") if _e_start_dt else ""
                                        _e_due_str   = _e_due_dt.strftime("%Y-%m-%d")   if _e_due_dt   else ""
                                        auth.update_task_meta(_t["id"], _e_title, _e_desc or "", _e_start_str, _e_due_str)
                                        auth.log_audit(cu["id"], cu["name"], "UPDATE", "tasks",
                                                       str(_t["id"]), f'Updated task "{_e_title}"')
                                        save_projects_async(st.session_state.projects)
                                        st.session_state[_edit_key] = False
                                        st.session_state.toast = {"msg": "Task updated!", "type": "success"}
                                        st.rerun()
                                if _es2.button("Cancel", key=f"ecancel_{tab_sfx}_{_t['id']}"):
                                    st.session_state[_edit_key] = False
                                    st.rerun()

                        # ── Per-task weekly comments expander ──────────────────
                        _wc_list = auth.get_task_comments_with_users(
                            task_id=_t["id"], from_date=_cm_from_str, to_date=_cm_to_str)
                        with st.expander(f"Weekly Comments ({len(_wc_list)})", expanded=False):
                            if not _wc_list:
                                st.info("No weekly comments in the selected period.")
                            else:
                                _wch = st.columns([1.2, 3.5, 2.0, 1.8])
                                for _c, _l in zip(_wch, ["Week", "Comment", "Employee", "Submitted"]):
                                    _c.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;border-bottom:2px solid #DFE3E7;padding:6px 2px;background:#F8FAFC">{_l}</div>', unsafe_allow_html=True)
                                for _wc in _wc_list:
                                    _wr = st.columns([1.2, 3.5, 2.0, 1.8])
                                    _wk_d = date.fromisoformat(_wc["week_start"])
                                    _wk_end = _wk_d + timedelta(days=6)
                                    _wr[0].markdown(f'<span style="font-size:10px;color:#475569">{_wk_d.strftime("%d %b")}–{_wk_end.strftime("%d %b")}</span>', unsafe_allow_html=True)
                                    _wr[1].markdown(f'<span style="font-size:11px;color:#111827">{esc(_wc["comment"])}</span>', unsafe_allow_html=True)
                                    _wr[2].markdown(f'<span style="font-size:11px;color:#374151">{esc(_wc["user_name"])}</span>', unsafe_allow_html=True)
                                    _wc_ts = str(_wc["created_at"]); _wc_ts_disp = fmt_date(_wc_ts[:10]) + " " + _wc_ts[11:16] if len(_wc_ts) >= 16 else _wc_ts
                                    _wr[3].markdown(f'<span style="font-size:10px;color:#94A3B8">{esc(_wc_ts_disp)}</span>', unsafe_allow_html=True)

                def _render_tab_with_filters(base_tasks, tab_sfx):
                    _emp_names = sorted({t["assigned_to"] for t in base_tasks})
                    _ff1, _ff2 = st.columns([1.5, 2.5])
                    _emp_f  = _ff1.selectbox("Employee", ["All"] + _emp_names, key=f"emp_f_{tab_sfx}")
                    _name_f = _ff2.text_input("Task name", placeholder="Filter by task title…", key=f"name_f_{tab_sfx}")
                    _visible = list(base_tasks)
                    if _emp_f != "All":
                        _visible = [t for t in _visible if t["assigned_to"] == _emp_f]
                    if _name_f.strip():
                        _nq = _name_f.strip().lower()
                        _visible = [t for t in _visible if _nq in t["title"].lower()]
                    st.markdown(
                        f'<p style="color:#64748B;font-size:11px;margin:4px 0 8px">'
                        f'<b>{len(_visible)}</b> task(s)</p>',
                        unsafe_allow_html=True)
                    _render_task_rows(_visible, tab_sfx)

                # ── Status sub-tabs ────────────────────────────────────────────
                _ip_tasks   = [t for t in _all_tasks if t["status"] == "In Progress"]
                _comp_tasks = [t for t in _all_tasks if t["status"] == "Completed"]
                _hold_tasks = [t for t in _all_tasks if t["status"] == "On Hold"]

                _stab_all, _stab_ip, _stab_comp, _stab_hold = st.tabs([
                    f"All ({len(_all_tasks)})",
                    f"In Progress ({len(_ip_tasks)})",
                    f"Completed ({len(_comp_tasks)})",
                    f"On Hold ({len(_hold_tasks)})",
                ])
                with _stab_all:
                    _render_tab_with_filters(_all_tasks, f"{_dept_sfx}_all")
                with _stab_ip:
                    _render_tab_with_filters(_ip_tasks, f"{_dept_sfx}_ip")
                with _stab_comp:
                    _render_tab_with_filters(_comp_tasks, f"{_dept_sfx}_comp")
                with _stab_hold:
                    _render_tab_with_filters(_hold_tasks, f"{_dept_sfx}_hold")

                # ── All comments summary (collapsible) ────────────────────────
                _all_comments = auth.get_task_comments_with_users(from_date=_cm_from_str, to_date=_cm_to_str)
                with st.expander(f"All Weekly Comments Summary ({len(_all_comments)} entries)", expanded=False):
                    if not _all_comments:
                        st.info("No comments in the selected period.")
                    else:
                        _ach = st.columns([1.5, 2.5, 2.5, 1.8, 1.8])
                        for _c, _l in zip(_ach, ["Week", "Task", "Employee", "Comment", "Submitted"]):
                            _c.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;border-bottom:2px solid #DFE3E7;padding:6px 2px;background:#F8FAFC">{_l}</div>', unsafe_allow_html=True)
                        for _ac in _all_comments:
                            _ar = st.columns([1.5, 2.5, 2.5, 1.8, 1.8])
                            _wk_d = date.fromisoformat(_ac["week_start"])
                            _wk_end = _wk_d + timedelta(days=6)
                            _ar[0].markdown(f'<span style="font-size:10px;color:#475569">{_wk_d.strftime("%d %b")}–{_wk_end.strftime("%d %b")}</span>', unsafe_allow_html=True)
                            _ar[1].markdown(f'<span style="font-size:11px;font-weight:600;color:#111827">{esc(_ac["task_title"])}</span>', unsafe_allow_html=True)
                            _ar[2].markdown(f'<span style="font-size:11px;color:#374151">{esc(_ac["user_name"])}</span>', unsafe_allow_html=True)
                            _ar[3].markdown(f'<span style="font-size:11px;color:#64748B">{esc(_ac["comment"][:80])}{"…" if len(_ac["comment"])>80 else ""}</span>', unsafe_allow_html=True)
                            _ac_ts = str(_ac["created_at"]); _ac_ts_disp = fmt_date(_ac_ts[:10]) + " " + _ac_ts[11:16] if len(_ac_ts) >= 16 else _ac_ts
                            _ar[4].markdown(f'<span style="font-size:10px;color:#94A3B8">{esc(_ac_ts_disp)}</span>', unsafe_allow_html=True)

        def _render_all_completed_cross_portal():
            _cross_done = [t for t in auth.get_all_tasks() if t.get("status") == "Completed"]
            if not _cross_done:
                st.info("No completed tasks found across any portal.")
                return
            st.markdown(
                f'<p style="color:#64748B;font-size:12px;margin-bottom:12px">'
                f'<b>{len(_cross_done)}</b> completed task(s) across all portals</p>',
                unsafe_allow_html=True,
            )
            _ff_c1, _ff_c2, _ff_c3 = st.columns([1.5, 1.5, 2])
            _cf_dept = _ff_c1.selectbox("Portal", ["All", "RPA", "Worksoft"], key="cf_dept_filter")
            _cf_emp  = _ff_c2.selectbox(
                "Employee",
                ["All"] + sorted({t.get("assigned_to", "") for t in _cross_done}),
                key="cf_emp_filter",
            )
            _cf_name = _ff_c3.text_input("Task name", placeholder="Filter by title…", key="cf_name_filter")
            _vis = list(_cross_done)
            if _cf_dept != "All":
                _vis = [t for t in _vis if t.get("department", "") == _cf_dept]
            if _cf_emp != "All":
                _vis = [t for t in _vis if t.get("assigned_to", "") == _cf_emp]
            if _cf_name.strip():
                _vis = [t for t in _vis if _cf_name.strip().lower() in t.get("title", "").lower()]
            st.markdown(
                f'<p style="color:#64748B;font-size:11px;margin:4px 0 8px"><b>{len(_vis)}</b> task(s)</p>',
                unsafe_allow_html=True,
            )
            with st.container(border=True):
                _ch = st.columns([2.5, 1.5, 1.0, 1.0, 0.8])
                for _c, _l in zip(_ch, ["Task", "Assigned To", "Portal", "Due Date", "Progress"]):
                    _c.markdown(
                        f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;'
                        f'color:#64748B;letter-spacing:.5px;padding:6px 4px;border-bottom:2px solid #DFE3E7;'
                        f'background:#F8FAFC">{_l}</div>',
                        unsafe_allow_html=True,
                    )
                for _ct in _vis:
                    _cr = st.columns([2.5, 1.5, 1.0, 1.0, 0.8], vertical_alignment="center")
                    _cr[0].markdown(
                        f'<div style="font-size:12px;font-weight:700;color:#111827">{esc(_ct["title"])}</div>'
                        + (f'<div style="font-size:10px;color:#64748B;font-style:italic">{esc(_ct.get("description","")[:60])}</div>' if _ct.get("description") else ""),
                        unsafe_allow_html=True,
                    )
                    _cr[1].markdown(f'<span style="font-size:11px;color:#374151">{esc(_ct.get("assigned_to","—"))}</span>', unsafe_allow_html=True)
                    _dept_v = _ct.get("department", "—") or "—"
                    _dept_c = "#3F8E91" if _dept_v == "RPA" else ("#7C3AED" if _dept_v == "Worksoft" else "#94A3B8")
                    _cr[2].markdown(f'<span style="font-size:10px;font-weight:700;color:{_dept_c}">{esc(_dept_v)}</span>', unsafe_allow_html=True)
                    _cr[3].markdown(f'<span style="font-size:11px;color:#64748B">{esc(fmt_date(_ct.get("due_date","")) or "—")}</span>', unsafe_allow_html=True)
                    _ct_pct = int(_ct.get("progress", 100))
                    _cr[4].markdown(
                        f'<div class="progress-bar-outer"><div class="progress-bar-inner" '
                        f'style="width:{_ct_pct}%;background:#10B981"></div></div>'
                        f'<div style="font-size:9px;color:#10B981;font-weight:700">{_ct_pct}%</div>',
                        unsafe_allow_html=True,
                    )

        def _render_task_dept_tabs():
            _tdept_rpa, _tdept_ws, _tdept_done = st.tabs(["🔧 RPA", "⚙️ Worksoft", "✅ All Completed"])
            with _tdept_rpa:
                _render_all_tasks_panel("RPA")
            with _tdept_ws:
                _render_all_tasks_panel("Worksoft")
            with _tdept_done:
                _render_all_completed_cross_portal()

        def _render_assigned_tasks_panel():
            _at_tasks = auth.get_tasks_assigned_by(cu["id"])
            if not _at_tasks:
                st.info("You haven't assigned any tasks yet.")
                return
            st.markdown(
                f'<p style="color:#64748B;font-size:12px;margin-bottom:12px">'
                f'<b>{len(_at_tasks)}</b> task(s) assigned by you</p>',
                unsafe_allow_html=True,
            )
            for _at in _at_tasks:
                with st.container(border=True):
                    _at_l, _at_r = st.columns([3, 1.2])
                    _at_pct   = int(_at["progress"])
                    _at_bar_c = "#10B981" if _at_pct == 100 else "#3B82F6"
                    with _at_l:
                        st.markdown(
                            f'<div style="font-size:14px;font-weight:700;color:#111827;margin-bottom:2px">'
                            f'{esc(_at["title"])}</div>',
                            unsafe_allow_html=True,
                        )
                        if _at.get("description"):
                            st.markdown(
                                f'<div style="font-size:12px;color:#64748B;font-style:italic;margin-bottom:4px">'
                                f'{esc(_at["description"])}</div>',
                                unsafe_allow_html=True,
                            )
                        _at_meta = f'Assigned to: <b>{esc(_at["assigned_to"])}</b>'
                        if _at.get("start_date"):
                            _at_meta += f' &nbsp;·&nbsp; Start: <b>{esc(fmt_date(_at["start_date"]))}</b>'
                        if _at.get("due_date"):
                            _at_meta += f' &nbsp;·&nbsp; Due: <b>{esc(fmt_date(_at["due_date"]))}</b>'
                        if _at.get("department"):
                            _at_meta += f' &nbsp;·&nbsp; Dept: <b>{esc(_at["department"])}</b>'
                        st.markdown(
                            f'<div style="font-size:11px;color:#64748B;margin-bottom:6px">{_at_meta}</div>',
                            unsafe_allow_html=True,
                        )
                        st.markdown(
                            f'<div class="progress-bar-outer"><div class="progress-bar-inner" '
                            f'style="width:{_at_pct}%;background:{_at_bar_c}"></div></div>'
                            f'<div style="font-size:10px;color:#64748B;margin-top:2px">{_at_pct}% complete</div>',
                            unsafe_allow_html=True,
                        )
                    with _at_r:
                        st.markdown(
                            f'<div style="text-align:right">{badge_html(_at["status"])}</div>',
                            unsafe_allow_html=True,
                        )
                        if _at.get("comment"):
                            st.markdown(
                                f'<div style="font-size:11px;color:#64748B;margin-top:6px;'
                                f'font-style:italic">{esc(_at["comment"])}</div>',
                                unsafe_allow_html=True,
                            )

        if role == "lead":
            _ltab_mine, _ltab_assigned, _ltab_all = st.tabs(["My Tasks", "Assigned Tasks", "All Tasks"])
            with _ltab_mine:
                _render_my_tasks_panel(key_prefix="lt_")
            with _ltab_assigned:
                _render_assigned_tasks_panel()
            with _ltab_all:
                _render_task_dept_tabs()
        elif role in ("manager", "admin"):
            _ltab_assigned, _ltab_all = st.tabs(["Assigned Tasks", "All Tasks"])
            with _ltab_assigned:
                _render_assigned_tasks_panel()
            with _ltab_all:
                _render_task_dept_tabs()
        else:
            _render_task_dept_tabs()

