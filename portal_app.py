"""
portal_app.py — Qualesce Dual-Portal Dashboard
================================================
Two fully-separated portals:
  • RPA Portal      — for users with department = RPA
  • Worksoft Portal — for users with department = Worksoft

Role hierarchy inside each portal
  Lead     → can CREATE projects / tasks; in Worksoft sees ONLY their own projects
  Manager  → full visibility of all projects in their portal
  User     → sees only tasks/projects directly assigned to them

Worksoft isolation rule (enforced on every view):
  A Lead's projects are visible to that Lead + any Manager/Admin.
  A Lead cannot see another Lead's Worksoft projects.

Run with:  streamlit run portal_app.py --server.port 8502
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import os, html, re, threading, io
from datetime import datetime, date, timedelta
import auth
import email_utils

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
_DIR  = os.path.dirname(os.path.abspath(__file__))
_LOGO = os.path.join(_DIR, "qualesce_logo.png")
st.set_page_config(
    page_title="Qualesce Portal",
    page_icon=_LOGO if os.path.exists(_LOGO) else "🏢",
    layout="wide",
    initial_sidebar_state="collapsed",
)

if "db_initialized" not in st.session_state:
    auth.init_db()
    st.session_state.db_initialized = True

if hasattr(email_utils, "start_license_notification_scheduler"):
    email_utils.start_license_notification_scheduler()

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
PROJECT_COLS = [
    "id","name","client","lead","employee","status","proj_type",
    "start","end","due_date","po","desc","manual_hrs","auto_hrs",
    "cost_per_hr","hours_saved","cost_saved","roi_pct","is_new","is_active",
    "ckpt_pdd_sdd_start","ckpt_pdd_sdd_end","ckpt_development_start","ckpt_development_end",
    "ckpt_uat_start","ckpt_uat_end","ckpt_deployment_start","ckpt_deployment_end",
    "num_bots","manual_run_mins","bot_run_mins","monthly_runs","num_persons",
    "allocated_hours","project_lead_email",
]
ALL_STATUSES = ["R&M","UAT","In Progress","Completed","PDD","Discontinued",
                "Internal POC","External POC","Important","Presales"]
WS_STATUSES  = ["In Progress","Completed","On Hold","Discontinued"]

STATUS_STYLES = {
    "R&M":         {"bg":"#EFF7F7","text":"#3F8E91","dot":"#5FA9AB"},
    "UAT":         {"bg":"#FBF6E7","text":"#966D17","dot":"#D4A02C"},
    "Completed":   {"bg":"#E5F2EC","text":"#2E7D5B","dot":"#2E7D5B"},
    "In Progress": {"bg":"#EFF7F7","text":"#2F6F72","dot":"#5FA9AB"},
    "PDD":         {"bg":"#FBF6E7","text":"#966D17","dot":"#D4A02C"},
    "Discontinued":{"bg":"#FCEAEA","text":"#B23A3A","dot":"#B23A3A"},
    "On Hold":     {"bg":"#FEF3C7","text":"#92400E","dot":"#F59E0B"},
    "Internal POC":{"bg":"#EFF7F7","text":"#2F6F72","dot":"#4A989B"},
    "External POC":{"bg":"#F7F8F9","text":"#4E5860","dot":"#9BA5AE"},
}

_TASK_STAT_COLORS = {
    "Not Started":"#94A3B8","In Progress":"#3B82F6",
    "Completed":"#10B981","On Hold":"#F59E0B",
}

# Portal themes
_PORTAL_THEME = {
    "RPA":      {"color":"#3F8E91","bg":"#EFF7F7","dark":"#162C3B","icon":"🔧","label":"RPA Portal"},
    "Worksoft": {"color":"#7C3AED","bg":"#F5F3FF","dark":"#1E1B4B","icon":"⚙️","label":"Worksoft Portal"},
}

# Role display labels (employee shown as "USER" in the UI)
_ROLE_LABEL = {
    "admin":    "ADMIN",
    "lead":     "LEAD",
    "manager":  "MANAGER",
    "employee": "USER",
    "sales":    "SALES",
}

# ── HELPERS ───────────────────────────────────────────────────────────────────
esc = html.escape

def fmt_date(d: str) -> str:
    if not d or not str(d).strip(): return ""
    s = str(d).strip()
    for fmt in ("%Y-%m-%d","%d/%m/%Y","%d-%m-%Y","%Y/%m/%d"):
        try: return datetime.strptime(s, fmt).strftime("%d-%m-%Y")
        except ValueError: pass
    return s

def _parse_dmy(s: str):
    v = str(s).strip()
    for fmt in ("%d-%m-%Y","%d/%m/%Y","%Y-%m-%d"):
        try: return datetime.strptime(v, fmt).date()
        except ValueError: pass
    return None

def cell(val, size="11px", color="#374151"):
    return f'<span style="font-size:{size};color:{color}">{esc(str(val))}</span>'

def is_new(row) -> bool:
    return str(row.get("is_new","")).lower() in ["true","1","yes"]

@st.cache_data(ttl=30, show_spinner=False)
def load_projects() -> pd.DataFrame:
    records = auth.get_all_projects()
    if records:
        df = pd.DataFrame(records)
        for col in PROJECT_COLS:
            if col not in df.columns: df[col] = ""
        return df
    return pd.DataFrame(columns=PROJECT_COLS)

def save_projects_async(df: pd.DataFrame):
    _c = df.copy()
    def _w():
        try: auth.upsert_projects(_c.to_dict("records")); load_projects.clear()
        except Exception: pass
    threading.Thread(target=_w, daemon=True).start()

def get_stats(d: pd.DataFrame) -> dict:
    if d is None or d.empty:
        return {"total":0,"in_progress":0,"completed":0,"uat":0,"rm":0,
                "new_added":0,"total_hrs":0.0,"total_cost":0.0}
    return {
        "total":       len(d),
        "in_progress": int((d["status"]=="In Progress").sum()),
        "completed":   int((d["status"]=="Completed").sum()),
        "uat":         int((d["status"]=="UAT").sum()),
        "rm":          int((d["status"]=="R&M").sum()),
        "new_added":   int(d["is_new"].astype(str).str.lower().isin(["true","1","yes"]).sum()),
        "total_hrs":   pd.to_numeric(d.get("hours_saved", pd.Series(dtype=float)), errors="coerce").fillna(0).sum(),
        "total_cost":  pd.to_numeric(d.get("cost_saved",  pd.Series(dtype=float)), errors="coerce").fillna(0).sum(),
    }

# ── SESSION STATE INIT ────────────────────────────────────────────────────────
_SS_DEFAULTS = {
    "current_user": None, "active_tab": "dashboard", "toast": None,
    "active_portal": None, "login_attempts": 0, "projects": None,
    "proj_tracker_open": None, "user_edit_id": None,
    "lc_edit_id": None, "sl_edit_id": None,
    "_pex_excel": None, "_pex_db": None, "_pex_db_name": None,
    "portal_home": True,  # show portal selection screen after login
}
for _k, _v in _SS_DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ── GLOBAL CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.block-container,[data-testid="stMainBlockContainer"]{padding-top:0!important}
.stTabs [data-baseweb="tab"]{font-size:12px;font-weight:600}
.progress-bar-outer{background:#E2E8F0;border-radius:6px;height:8px;overflow:hidden;margin:2px 0}
.progress-bar-inner{height:8px;border-radius:6px;transition:width .3s ease}
.isolation-notice{
  background:#F5F3FF;border:1px solid #C4B5FD;border-radius:8px;
  padding:8px 14px;font-size:11px;color:#5B21B6;margin-bottom:10px
}
</style>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# LOGIN GATE
# ══════════════════════════════════════════════════════════════════════════════
def _render_login():
    st.markdown("""<style>
    [data-testid="stAppViewContainer"]>.main{
      background:linear-gradient(135deg,#0F2233 0%,#1A3347 50%,#162C3B 100%)!important}
    </style>""", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1.1, 1])
    with col:
        st.markdown('<div style="height:80px"></div>', unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(
                '<div style="text-align:center;padding:16px 0 8px">'
                '<div style="font-size:26px;font-weight:900;color:#1F3B4D;letter-spacing:-1px">Qualesce</div>'
                '<div style="font-size:11px;color:#64748B;margin-top:2px;font-weight:600;'
                'text-transform:uppercase;letter-spacing:1px">Portal Dashboard</div>'
                '</div>',
                unsafe_allow_html=True,
            )
            st.markdown("---")
            _em = st.text_input("Email", placeholder="your@email.com", key="pl_email")
            _pw = st.text_input("Password", type="password", placeholder="Password", key="pl_pwd")
            if st.button("Sign In", type="primary", use_container_width=True, key="pl_signin"):
                _u = auth.authenticate(_em.strip().lower(), _pw)
                if _u:
                    st.session_state.current_user   = _u
                    st.session_state.login_attempts = 0
                    st.session_state.active_tab     = "tasks" if _u["role"] == "employee" else "dashboard"
                    auth.log_audit(_u["id"], _u["name"], "LOGIN", "users",
                                   str(_u["id"]), f'Portal login: {_u["name"]}')
                    st.rerun()
                else:
                    st.session_state.login_attempts += 1
                    remaining = max(0, 3 - st.session_state.login_attempts)
                    msg = (f"Invalid credentials. ({remaining} attempt{'s' if remaining!=1 else ''} left)"
                           if remaining > 0 else "Invalid credentials or account inactive.")
                    st.error(msg)

if st.session_state.current_user is None:
    _render_login()
    st.stop()

# ── POST-LOGIN SETUP ──────────────────────────────────────────────────────────
cu   = st.session_state.current_user
role = cu["role"]

if st.session_state.projects is None or not isinstance(st.session_state.projects, pd.DataFrame):
    st.session_state.projects = load_projects()
df = st.session_state.projects

# ── PORTAL DETECTION ─────────────────────────────────────────────────────────
_cu_dept = (cu.get("department") or "").strip()
if _cu_dept == "Worksoft":
    _locked_portal = "Worksoft"
elif _cu_dept == "RPA":
    _locked_portal = "RPA"
else:
    _locked_portal = None  # admin / no dept → can choose either

# ── PORTAL HOME SCREEN ────────────────────────────────────────────────────────
# Shown after login so the user can see and choose between both portals.
if st.session_state.get("portal_home", True):
    st.markdown("""
    <style>
    [data-testid="stAppViewContainer"]>.main{
      background:linear-gradient(135deg,#0F2233 0%,#1A3347 50%,#162C3B 100%)!important}
    </style>""", unsafe_allow_html=True)

    st.markdown('<div style="height:40px"></div>', unsafe_allow_html=True)
    st.markdown(
        f'<div style="text-align:center;margin-bottom:8px">'
        f'<div style="font-size:30px;font-weight:900;color:#fff;letter-spacing:-1px">Qualesce</div>'
        f'<div style="font-size:12px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin-top:4px">'
        f'Welcome, {esc(cu["name"])} &nbsp;·&nbsp; '
        f'<span style="color:#5FA9AB">{_ROLE_LABEL.get(role,role.upper())}</span></div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div style="text-align:center;color:#64748B;font-size:13px;margin-bottom:32px">Select a portal to continue</div>', unsafe_allow_html=True)

    _can_rpa      = _locked_portal in (None, "RPA")
    _can_worksoft = _locked_portal in (None, "Worksoft")

    _hc1, _hc2, _hc3 = st.columns([1, 2, 1])
    with _hc2:
        _pc1, _pc2 = st.columns(2)

        # RPA Portal card
        with _pc1:
            _rpa_opacity = "1" if _can_rpa else "0.4"
            st.markdown(
                f'<div style="background:linear-gradient(135deg,#162C3B,#1F4A50);border:2px solid {"#3F8E91" if _can_rpa else "#2d3748"};'
                f'border-radius:16px;padding:32px 24px;text-align:center;opacity:{_rpa_opacity};cursor:{"pointer" if _can_rpa else "not-allowed"};'
                f'margin-bottom:12px">'
                f'<div style="font-size:48px">🔧</div>'
                f'<div style="font-size:18px;font-weight:900;color:#fff;margin:12px 0 6px">RPA Portal</div>'
                f'<div style="font-size:11px;color:#94A3B8">Projects · Tasks · License</div>'
                f'{"<div style=\'font-size:10px;color:#3F8E91;margin-top:8px;font-weight:700\'>Your Portal</div>" if _locked_portal=="RPA" else ""}'
                f'</div>',
                unsafe_allow_html=True,
            )
            if _can_rpa:
                if st.button("Enter RPA Portal", key="home_rpa", use_container_width=True, type="primary"):
                    st.session_state.active_portal = "RPA"
                    st.session_state.portal_home   = False
                    st.session_state.active_tab    = "tasks" if role == "employee" else "dashboard"
                    st.rerun()
            else:
                st.button("No Access", key="home_rpa_locked", use_container_width=True, disabled=True)

        # Worksoft Portal card
        with _pc2:
            _ws_opacity = "1" if _can_worksoft else "0.4"
            st.markdown(
                f'<div style="background:linear-gradient(135deg,#1E1B4B,#2D1B69);border:2px solid {"#7C3AED" if _can_worksoft else "#2d3748"};'
                f'border-radius:16px;padding:32px 24px;text-align:center;opacity:{_ws_opacity};cursor:{"pointer" if _can_worksoft else "not-allowed"};'
                f'margin-bottom:12px">'
                f'<div style="font-size:48px">⚙️</div>'
                f'<div style="font-size:18px;font-weight:900;color:#fff;margin:12px 0 6px">Worksoft Portal</div>'
                f'<div style="font-size:11px;color:#94A3B8">Projects · Tasks · Hours</div>'
                f'{"<div style=\'font-size:10px;color:#7C3AED;margin-top:8px;font-weight:700\'>Your Portal</div>" if _locked_portal=="Worksoft" else ""}'
                f'</div>',
                unsafe_allow_html=True,
            )
            if _can_worksoft:
                if st.button("Enter Worksoft Portal", key="home_ws", use_container_width=True, type="primary"):
                    st.session_state.active_portal = "Worksoft"
                    st.session_state.portal_home   = False
                    st.session_state.active_tab    = "tasks" if role == "employee" else "dashboard"
                    st.rerun()
            else:
                st.button("No Access", key="home_ws_locked", use_container_width=True, disabled=True)

        st.markdown('<div style="text-align:center;margin-top:16px">', unsafe_allow_html=True)
        if st.button("Logout", key="home_logout", use_container_width=False):
            st.session_state.current_user = None
            st.session_state.portal_home  = True
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    st.stop()

portal = st.session_state.get("active_portal", "RPA")
_pc    = _PORTAL_THEME[portal]

# ── TOAST ─────────────────────────────────────────────────────────────────────
if st.session_state.toast:
    _t   = st.session_state.toast
    _tc  = {"success":("#064E3B","#10B981","✓"),"error":("#7F1D1D","#EF4444","✗"),"info":("#244E51","#5FA9AB","ℹ")}
    _tbg, _tbd, _tico = _tc.get(_t.get("type","success"),("#064E3B","#10B981","✓"))
    st.markdown(
        f'<div style="background:{_tbg};border:1px solid {_tbd};border-radius:10px;'
        f'padding:10px 18px;color:#fff;font-size:13px;font-weight:600;margin-bottom:10px;'
        f'display:flex;align-items:center;gap:10px">'
        f'<span>{_tico}</span><span>{html.escape(_t["msg"])}</span></div>',
        unsafe_allow_html=True)
    st.session_state.toast = None

# ── PORTAL HEADER ─────────────────────────────────────────────────────────────
_role_colors = {"admin":"#3F8E91","lead":"#2E7D5B","manager":"#966D17","employee":"#4E5860","sales":"#5FA9AB"}
_role_bg     = {"admin":"#EFF7F7","lead":"#E5F2EC","manager":"#FBF6E7","employee":"#F1F5F9","sales":"#EFF7F7"}

_hc1, _hc2 = st.columns([3, 1])
with _hc1:
    st.markdown(
        f'<div style="background:linear-gradient(90deg,{_pc["dark"]},#1F3B4D);'
        f'border-radius:12px;padding:14px 22px;margin-bottom:12px;'
        f'display:flex;align-items:center;gap:14px">'
        f'<span style="font-size:28px">{_pc["icon"]}</span>'
        f'<div>'
        f'<div style="font-size:18px;font-weight:900;color:#fff;letter-spacing:-.4px">{_pc["label"]}</div>'
        f'<div style="font-size:11px;color:#94A3B8;margin-top:3px">'
        f'Logged in as <b style="color:#E2E8F0">{esc(cu["name"])}</b>'
        f' &nbsp;·&nbsp; '
        f'<span style="background:{_role_bg.get(role,"#F1F5F9")};color:{_role_colors.get(role,"#374151")};'
        f'padding:2px 8px;border-radius:12px;font-weight:700;font-size:10px">'
        f'{_ROLE_LABEL.get(role, role.upper())}</span>'
        f'</div></div></div>',
        unsafe_allow_html=True,
    )

with _hc2:
    _btn_col, _logout_col = st.columns(2)
    if _btn_col.button("⬅ Portals", use_container_width=True, key="back_to_home"):
        st.session_state.portal_home        = True
        st.session_state.active_portal      = None
        st.session_state.proj_tracker_open  = None
        st.rerun()
    if _logout_col.button("Logout", use_container_width=True, key="portal_logout"):
        st.session_state.current_user = None
        st.session_state.portal_home  = True
        st.rerun()

# ── NAVIGATION TABS ───────────────────────────────────────────────────────────
# Tab visibility per portal × role:
#   User (employee) : My Tasks, My Projects
#   Worksoft Lead   : Dashboard, Projects (own only), Tasks (own projects only)
#   Worksoft Manager: Dashboard, Projects (all), Tasks (all)
#   RPA Lead        : Dashboard, Projects (all), Tasks, License
#   RPA Manager     : Dashboard, Projects (all), Tasks, License
#   Admin           : all tabs + Users + Settings
if role == "employee":
    _tab_defs = [("tasks","My Tasks"), ("projects","My Projects")]
elif portal == "RPA":
    if role == "admin":
        _tab_defs = [("dashboard","Dashboard"), ("projects","Projects"),
                     ("tasks","Tasks"), ("license","License"),
                     ("users","Users"), ("settings","Settings")]
    else:  # lead or manager
        _tab_defs = [("dashboard","Dashboard"), ("projects","Projects"),
                     ("tasks","Tasks"), ("license","License")]
else:  # Worksoft portal
    if role == "admin":
        _tab_defs = [("dashboard","Dashboard"), ("projects","Projects"),
                     ("tasks","Tasks"), ("users","Users"), ("settings","Settings")]
    else:  # lead or manager
        _tab_defs = [("dashboard","Dashboard"), ("projects","Projects"), ("tasks","Tasks")]

_valid_tabs = [t[0] for t in _tab_defs]
if st.session_state.active_tab not in _valid_tabs:
    st.session_state.active_tab = _tab_defs[0][0]

_nav_cols = st.columns(len(_tab_defs))
for _ni, (_tid, _tlbl) in enumerate(_tab_defs):
    _is_active = st.session_state.active_tab == _tid
    if _nav_cols[_ni].button(_tlbl, key=f"nav_{_tid}", use_container_width=True,
                              type="primary" if _is_active else "secondary"):
        st.session_state.active_tab = _tid
        st.rerun()

st.markdown(f'<hr style="margin:6px 0 16px;border:none;border-top:2px solid {_pc["color"]}40">',
            unsafe_allow_html=True)

active = st.session_state.active_tab

# ══════════════════════════════════════════════════════════════════════════════
# PROJECT FILTER FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def _portal_projects() -> pd.DataFrame:
    """All active projects belonging to the current portal (matched by proj_type)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=PROJECT_COLS)
    if "proj_type" in df.columns:
        return df[df["proj_type"].fillna("").str.strip() == portal].copy()
    return df.copy()


def _apply_role_filter(base: pd.DataFrame) -> pd.DataFrame:
    """
    Apply role-based project visibility rules:

    Employee / User:
      → Can only see projects where their name appears in the employee field.

    Lead (isolation rule — all portals):
      → Can only see projects where they are the assigned Lead.
      → Cannot see projects belonging to other Leads.
      → Managers and Admins can see ALL projects.

    Manager / Admin:
      → No project-level isolation; all portal projects are visible.
    """
    if base.empty:
        return base

    # User (employee): only projects assigned to them
    if role == "employee":
        _name = str(cu.get("name", "")).strip()
        if not _name:
            return base.iloc[0:0]
        _mask = base["employee"].fillna("").str.contains(_name, case=False, regex=False)
        return base[_mask].copy()

    # Lead isolation rule — applies to all portals
    if role == "lead":
        _name  = str(cu.get("name", "")).strip()
        _email = str(cu.get("email", "")).strip()
        _mask  = base["lead"].fillna("").str.strip() == _name
        if "project_lead_email" in base.columns:
            _mask = _mask | (base["project_lead_email"].fillna("").str.strip() == _email)
        return base[_mask].copy()

    # Manager / Admin: see everything in this portal
    return base


def _visible_projects() -> pd.DataFrame:
    """Portal-scoped + role-filtered projects (single convenience call)."""
    return _apply_role_filter(_portal_projects())


def _show_isolation_notice():
    """Display the Lead isolation notice where relevant."""
    if role == "lead":
        st.markdown(
            '<div class="isolation-notice">'
            '🔒 <b>Lead view:</b> Showing only projects assigned to you as Lead. '
            'Managers can see all projects.'
            '</div>',
            unsafe_allow_html=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
# TAB: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if active == "dashboard":
    _pdf = _visible_projects()
    _s   = get_stats(_pdf)

    st.markdown(
        f'<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:8px">'
        f'{_pc["icon"]} {portal} Dashboard</h2>',
        unsafe_allow_html=True,
    )
    _show_isolation_notice()

    # KPI row
    _k1, _k2, _k3, _k4, _k5 = st.columns(5)
    def _kpi(col, label, value, color):
        col.markdown(
            f'<div style="background:#fff;border:1px solid #E2E8F0;border-radius:10px;'
            f'padding:16px;text-align:center;border-top:3px solid {color}">'
            f'<div style="font-size:28px;font-weight:900;color:{color}">{value}</div>'
            f'<div style="font-size:10px;font-weight:700;color:#94A3B8;text-transform:uppercase;'
            f'letter-spacing:.6px;margin-top:4px">{label}</div></div>',
            unsafe_allow_html=True,
        )
    _kpi(_k1, "Total Projects",  _s["total"],               _pc["color"])
    _kpi(_k2, "In Progress",     _s["in_progress"],         "#3B82F6")
    _kpi(_k3, "Completed",       _s["completed"],           "#10B981")
    _kpi(_k4, "Hrs Saved",       f'{_s["total_hrs"]:,.0f}', "#F59E0B")
    _kpi(_k5, "Cost Saved",      f'₹{_s["total_cost"]:,.0f}', "#8B5CF6")

    st.markdown("<br>", unsafe_allow_html=True)

    if _pdf.empty:
        st.info(f"No {portal} projects visible for your role.")
    else:
        _ch1, _ch2 = st.columns([1.5, 1])
        with _ch1:
            with st.container(border=True):
                st.markdown('<div style="font-size:11px;font-weight:700;color:#64748B;margin-bottom:8px">Projects by Status</div>', unsafe_allow_html=True)
                _sc = _pdf["status"].value_counts().reset_index()
                _sc.columns = ["Status","Count"]
                _fig = px.bar(_sc, x="Status", y="Count", color="Status",
                              color_discrete_sequence=["#3F8E91","#D4A02C","#10B981","#3B82F6",
                                                       "#F59E0B","#EF4444","#8B5CF6","#64748B"],
                              height=260)
                _fig.update_layout(showlegend=False, margin=dict(t=10,b=10,l=0,r=0),
                                   plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(_fig, use_container_width=True)
        with _ch2:
            with st.container(border=True):
                st.markdown('<div style="font-size:11px;font-weight:700;color:#64748B;margin-bottom:8px">Status Breakdown</div>', unsafe_allow_html=True)
                _pie_df = _pdf["status"].value_counts().reset_index()
                _pie_df.columns = ["Status","Count"]
                _pfig = px.pie(_pie_df, names="Status", values="Count", height=260,
                               color_discrete_sequence=["#3F8E91","#D4A02C","#10B981","#3B82F6",
                                                        "#F59E0B","#EF4444","#8B5CF6","#64748B"])
                _pfig.update_layout(margin=dict(t=10,b=10,l=0,r=0), showlegend=True,
                                    legend=dict(font=dict(size=10)))
                st.plotly_chart(_pfig, use_container_width=True)

        with st.container(border=True):
            st.markdown('<div style="font-size:11px;font-weight:700;color:#64748B;margin-bottom:8px">Recent Projects</div>', unsafe_allow_html=True)
            for _, _rr in _pdf.head(10).iterrows():
                _rs  = str(_rr.get("status",""))
                _rss = STATUS_STYLES.get(_rs, {"dot":"#94A3B8"})
                st.markdown(
                    f'<div style="display:flex;align-items:center;gap:10px;padding:6px 0;'
                    f'border-bottom:1px solid #F8FAFC">'
                    f'<div style="width:8px;height:8px;border-radius:50%;background:{_rss["dot"]};flex-shrink:0"></div>'
                    f'<div style="flex:1;font-size:12px;font-weight:600;color:#111827">{esc(str(_rr.get("name","")))}</div>'
                    f'<div style="font-size:11px;color:#64748B">{esc(str(_rr.get("client","")))}</div>'
                    f'<div style="font-size:10px;font-weight:700;color:{_rss["dot"]}">{esc(_rs)}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )


# ══════════════════════════════════════════════════════════════════════════════
# TAB: PROJECTS  (also "My Projects" for employee)
# ══════════════════════════════════════════════════════════════════════════════
elif active == "projects":
    _title = "My Projects" if role == "employee" else f"{_pc['icon']} {portal} Projects"
    st.markdown(
        f'<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px">{_title}</h2>',
        unsafe_allow_html=True,
    )
    _show_isolation_notice()

    # ── Project tracker detail view ───────────────────────────────────────────
    if st.session_state.get("proj_tracker_open"):
        _psel = st.session_state["proj_tracker_open"]
        _bk, _ttl = st.columns([1, 9])
        if _bk.button("← Back", key="ptr_back", use_container_width=True):
            st.session_state["proj_tracker_open"] = None
            st.rerun()
        _ttl.markdown(
            f'<div style="display:flex;align-items:center;gap:10px">'
            f'<span style="font-size:15px;font-weight:800;color:#1F3B4D">Project Tracker</span>'
            f'<span style="font-size:11px;font-weight:600;padding:3px 12px;border-radius:20px;'
            f'background:{_pc["bg"]};color:{_pc["color"]};border:1px solid {_pc["color"]}40">'
            f'{esc(_psel)}</span></div>',
            unsafe_allow_html=True,
        )

        _vis = _visible_projects()
        _pmatch = _vis[_vis["name"] == _psel]
        if _pmatch.empty:
            st.warning("Project not found or you do not have access to this project.")
        else:
            _pr = _pmatch.iloc[0]
            def _pv(k, fb="—"):
                v = _pr.get(k, fb); s = str(v).strip()
                return fb if s in ("","nan","None","NaN") else s

            _pstatus = _pv("status","")
            _pcolor  = STATUS_STYLES.get(_pstatus, {"dot":"#94A3B8"})["dot"]
            _tc1, _tc2 = st.columns([1, 1.3])

            with _tc1:
                with st.container(border=True):
                    st.markdown(
                        f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:12px">'
                        f'<div style="width:10px;height:10px;border-radius:50%;background:{_pcolor}"></div>'
                        f'<span style="font-size:13px;font-weight:800;color:#1F3B4D">{esc(_psel)}</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                    def _pfield(lbl, val, col="#374151"):
                        return (f'<div style="display:flex;gap:8px;padding:5px 0;border-bottom:1px solid #F8FAFC">'
                                f'<div style="font-size:10px;font-weight:700;color:#94A3B8;min-width:80px">{lbl}</div>'
                                f'<div style="font-size:11px;color:{col};font-weight:600">{esc(str(val))}</div>'
                                f'</div>')
                    st.markdown(
                        _pfield("Client",  _pv("client"))
                        + _pfield("Lead",  _pv("lead"))
                        + _pfield("Team",  _pv("employee"))
                        + _pfield("Status",_pstatus, _pcolor)
                        + _pfield("Start", fmt_date(_pv("start",""))    or "—")
                        + _pfield("End",   fmt_date(_pv("end",""))      or "Ongoing")
                        + _pfield("Due",   fmt_date(_pv("due_date","")) or "—"),
                        unsafe_allow_html=True,
                    )
                    # Timeline progress
                    _ts = _parse_dmy(_pv("start",""))
                    _te = _parse_dmy(_pv("end",""))
                    if _ts and _te and _te > _ts:
                        _today = date.today()
                        _tot   = (_te - _ts).days
                        _ela   = max(0, min(_tot, (_today - _ts).days))
                        _pct   = 100 if _pstatus=="Completed" else round((_ela/_tot)*100)
                        _tc    = "#10B981" if _pct<70 else ("#F59E0B" if _pct<90 else "#EF4444")
                        if _pstatus == "Completed": _tc = "#10B981"
                        _dl   = (_te - _today).days
                        _dlbl = (f"{_dl}d left" if _dl>0
                                 else ("Completed" if _pstatus=="Completed" else f"{abs(_dl)}d overdue"))
                        st.markdown(
                            f'<div style="margin-top:12px;padding-top:10px;border-top:1px solid #F1F5F9">'
                            f'<div style="display:flex;justify-content:space-between;margin-bottom:4px">'
                            f'<span style="font-size:10px;font-weight:700;color:#64748B">Timeline</span>'
                            f'<span style="font-size:10px;font-weight:700;color:{_tc}">{_dlbl}</span></div>'
                            f'<div style="background:#E2E8F0;border-radius:6px;height:8px;overflow:hidden">'
                            f'<div style="width:{_pct}%;background:{_tc};height:8px;border-radius:6px"></div></div>'
                            f'<div style="display:flex;justify-content:space-between;margin-top:3px">'
                            f'<span style="font-size:9px;color:#94A3B8">{fmt_date(_pv("start",""))}</span>'
                            f'<span style="font-size:10px;font-weight:700;color:{_tc}">{_pct}%</span>'
                            f'<span style="font-size:9px;color:#94A3B8">{fmt_date(_pv("end",""))}</span>'
                            f'</div></div>',
                            unsafe_allow_html=True,
                        )

            with _tc2:
                with st.container(border=True):
                    _pid = int(float(_pr.get("id",0) or 0))

                    # Metrics row
                    _mh = _pv("hours_saved",""); _mc = _pv("cost_saved","")
                    _mhtml = ""
                    if _mh not in ("","—","0","0.0"):
                        _mhtml += (f'<div style="flex:1;background:#F0FDF4;border-radius:10px;padding:12px;border-top:3px solid #10B981;text-align:center">'
                                   f'<div style="font-size:20px;font-weight:900;color:#059669">{_mh}</div>'
                                   f'<div style="font-size:10px;color:#14532D;margin-top:3px">Hours Saved</div></div>')
                    if _mc not in ("","—","0","0.0"):
                        _mhtml += (f'<div style="flex:1;background:#EFF7F7;border-radius:10px;padding:12px;border-top:3px solid #5FA9AB;text-align:center">'
                                   f'<div style="font-size:20px;font-weight:900;color:#3F8E91">₹{_mc}</div>'
                                   f'<div style="font-size:10px;color:#1E40AF;margin-top:3px">Cost Saved</div></div>')
                    if _mhtml:
                        st.markdown('<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:6px">Metrics</div>', unsafe_allow_html=True)
                        st.markdown(f'<div style="display:flex;gap:8px;margin-bottom:12px">{_mhtml}</div>', unsafe_allow_html=True)

                    # RPA-specific: bot metrics
                    if portal == "RPA" and _pid:
                        st.markdown('<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">Bot Metrics</div>', unsafe_allow_html=True)
                        _nb  = int(float(_pr.get("num_bots",0)       or 0))
                        _np2 = int(float(_pr.get("num_persons",0)    or 0))
                        _mr2 = float(_pr.get("manual_run_mins",0)    or 0)
                        _br2 = float(_pr.get("bot_run_mins",0)       or 0)
                        _ms  = date.today().replace(day=1).isoformat()
                        _me  = date.today().isoformat()
                        _cur = auth.get_bot_metric_logs(project_id=_pid, start_date=_ms, end_date=_me)
                        _mq   = sum(int(l.get("qty",0) or 0) for l in _cur)
                        _svd  = max(float(_mr2)*float(_np2) - float(_br2)*float(_nb), 0)*_mq/60 if (_mr2 or _br2) else 0
                        _riv  = float(_pr.get("run_interval_value", 0) or 0)
                        _riu  = str(_pr.get("run_interval_unit", "Minutes") or "Minutes")
                        _rif  = str(_pr.get("run_frequency", "Daily") or "Daily")
                        _rim  = _riv * 60 if _riu == "Hours" else _riv
                        # actual run time = logged qty × interval (runs every X mins)
                        _mo_run_time = _mq * _rim / 60
                        # projected monthly runs = active hrs × 60 / interval
                        _freq_hrs = {"Daily": 176, "Weekly": 160, "Monthly": 8}.get(_rif, 176)
                        _est_mo_runs = int(_freq_hrs * 60 / _rim) if _rim > 0 else 0
                        _bk1,_bk2,_bk3,_bk4,_bk5,_bk6 = st.columns(6)
                        _bk1.metric("Bots",_nb); _bk2.metric("Persons",_np2)
                        _bk3.metric("Month Qty",_mq); _bk4.metric("Est. Mo. Runs",_est_mo_runs)
                        _bk5.metric("Hrs Saved",f"{_svd:.1f}")
                        _bk6.metric("Mo. Run Time (hrs)", f"{_mo_run_time:.1f}")

                        # ── Log Daily Quantity ────────────────────────────────
                        st.markdown(
                            '<div style="font-size:10px;font-weight:700;color:#1F3B4D;margin:12px 0 6px">'
                            '📅 Log Daily Quantity</div>',
                            unsafe_allow_html=True,
                        )
                        _bl1, _bl2, _bl3 = st.columns([2, 2, 1])
                        _bm_log_date = _bl1.date_input(
                            "Date", value=date.today(), format="DD/MM/YYYY",
                            key=f"bm_logdate_{_pid}",
                        )
                        _bm_log_qty = _bl2.number_input(
                            "Quantity", min_value=0, value=0, step=1,
                            key=f"bm_logqty_{_pid}",
                        )
                        _bl3.markdown('<div style="height:28px"></div>', unsafe_allow_html=True)
                        if _bl3.button("📥 Log", key=f"bm_log_{_pid}", use_container_width=True):
                            auth.upsert_bot_metric_log(_pid, _psel, str(_bm_log_date), _bm_log_qty)
                            auth.log_audit(cu["id"], cu["name"], "CREATE", "bot_metric_logs",
                                           str(_pid),
                                           f'Qty {_bm_log_qty} logged for "{_psel}" on {_bm_log_date}')
                            st.session_state.toast = {
                                "msg": f"Logged {_bm_log_qty} for {_bm_log_date.strftime('%d/%m/%Y')}!",
                                "type": "success",
                            }
                            st.rerun()
                        # live calculation preview
                        if _rim > 0:
                            _prev_mins = _bm_log_qty * _rim
                            st.markdown(
                                f'<div style="background:#F0F9FF;border:1px solid #BAE6FD;border-radius:8px;'
                                f'padding:8px 12px;margin-top:6px;font-size:11px;color:#0369A1">'
                                f'<b>This entry:</b> {_bm_log_qty} runs × {_riv} {_riu.lower()} = '
                                f'<b>{_prev_mins:.0f} mins ({_prev_mins/60:.2f} hrs)</b> &nbsp;|&nbsp; '
                                f'<b>Est. Mo. Runs:</b> {_est_mo_runs:,} &nbsp;|&nbsp; '
                                f'<b>Freq:</b> {_rif}'
                                f'</div>',
                                unsafe_allow_html=True
                            )

                        # ── Recent log entries ────────────────────────────────
                        _bm_all_logs = auth.get_bot_metric_logs(project_id=_pid)
                        if _bm_all_logs:
                            with st.expander(f"📋 Log History ({len(_bm_all_logs)} entries)", expanded=False):
                                _bm_log_df = pd.DataFrame(_bm_all_logs[:15])[["log_date", "qty"]].copy()
                                _bm_log_df["run_time_mins"] = (_bm_log_df["qty"] * _rim).round(1)
                                _bm_log_df["est_mo_runs"]   = _est_mo_runs
                                _bm_log_df.columns = ["Date", "Qty", "Run Time (mins)", "Est. Mo. Runs"]
                                st.dataframe(_bm_log_df, use_container_width=True, hide_index=True)

                    # Worksoft-specific: hours budget
                    if portal == "Worksoft" and _pid:
                        _ah = 0.0
                        try: _ah = float(_pv("allocated_hours","0") or 0)
                        except (ValueError, TypeError): pass
                        if _ah:
                            _ws_punches = auth.get_worksoft_punches(project_id=_pid) if hasattr(auth,"get_worksoft_punches") else []
                            _used_hrs   = sum(float(p.get("hours_worked",0) or 0) for p in _ws_punches)
                            _pct_used   = min(100, round((_used_hrs/_ah)*100)) if _ah else 0
                            _hc         = "#10B981" if _pct_used<50 else ("#F59E0B" if _pct_used<90 else "#EF4444")
                            st.markdown(
                                f'<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:6px">Hours Budget</div>'
                                f'<div style="background:#E2E8F0;border-radius:6px;height:10px;overflow:hidden;margin-bottom:4px">'
                                f'<div style="width:{_pct_used}%;background:{_hc};height:10px;border-radius:6px"></div></div>'
                                f'<div style="font-size:11px;color:#374151">{_used_hrs:.1f} / {_ah:.0f} hrs ({_pct_used}%)</div>',
                                unsafe_allow_html=True,
                            )

                    # Comments
                    st.markdown('<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin:10px 0 6px">Comments</div>', unsafe_allow_html=True)
                    _cmts = auth.get_project_comments(_pid) if _pid else []
                    for _ci, _cm in enumerate(_cmts):
                        _cm_ts   = str(_cm["created_at"])
                        _cm_disp = fmt_date(_cm_ts[:10]) + " " + _cm_ts[11:16] if len(_cm_ts)>=16 else _cm_ts
                        st.markdown(
                            f'<div style="background:{"#F8FAFC" if _ci%2==0 else "#fff"};border:1px solid #E2E8F0;'
                            f'border-radius:8px;padding:8px 12px;margin-bottom:4px">'
                            f'<div style="display:flex;justify-content:space-between">'
                            f'<span style="font-size:11px;font-weight:700;color:{_pc["color"]}">{esc(_cm["user_name"])}</span>'
                            f'<span style="font-size:10px;color:#94A3B8">{esc(_cm_disp)}</span>'
                            f'</div><div style="font-size:12px;color:#374151;margin-top:3px">{esc(_cm["comment"])}</div>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                    if not _cmts:
                        st.markdown('<div style="font-size:11px;color:#CBD5E1;font-style:italic">No comments yet.</div>', unsafe_allow_html=True)
                    _new_cmt = st.text_area("Add comment", key=f"ptr_cmt_{_psel}", height=65,
                                            label_visibility="collapsed", placeholder="Write a comment…")
                    if st.button("Post", key=f"ptr_post_{_psel}", type="primary"):
                        if _new_cmt.strip() and _pid:
                            auth.add_project_comment(_pid, _psel, cu["id"], cu["name"], _new_cmt.strip())
                            st.rerun()

        st.stop()

    # ── Project listing ───────────────────────────────────────────────────────
    _base = _visible_projects()

    _fstatus_opts = WS_STATUSES if portal == "Worksoft" else ALL_STATUSES
    _f1, _f2, _f3 = st.columns([2, 2, 2])
    _fstatus = _f1.selectbox("Status", ["All"] + _fstatus_opts, key="p_fstat")
    _fclient = _f2.selectbox("Client", ["All"] + sorted({
        str(r).strip() for r in _base.get("client", pd.Series(dtype=str)).dropna() if str(r).strip()
    }), key="p_fcli")
    _fsearch = _f3.text_input("Search", placeholder="Project name…", key="p_fsrch")

    _filtered = _base.copy()
    if _fstatus != "All": _filtered = _filtered[_filtered["status"] == _fstatus]
    if _fclient != "All": _filtered = _filtered[_filtered["client"].fillna("") == _fclient]
    if _fsearch.strip():
        _q = _fsearch.strip().lower()
        _filtered = _filtered[_filtered["name"].fillna("").str.lower().str.contains(_q)]

    st.markdown(f'<p style="color:#64748B;font-size:12px;margin:4px 0 10px"><b>{len(_filtered)}</b> project(s)</p>',
                unsafe_allow_html=True)

    if _filtered.empty:
        st.info(f"No {portal} projects match the current filters.")
    else:
        with st.container(border=True):
            _hcols = st.columns([0.3,2.5,1.2,1.2,1.2,1.0,1.0,1.0])
            for _c, _l in zip(_hcols, ["#","Project","Client","Lead","Status","Start","End","Due"]):
                _c.markdown(
                    f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;'
                    f'letter-spacing:.5px;padding:6px 4px;border-bottom:2px solid #DFE3E7;background:#F8FAFC">{_l}</div>',
                    unsafe_allow_html=True)
            for _, _row in _filtered.iterrows():
                _rstat  = str(_row.get("status",""))
                _rid    = str(_row.get("id",""))
                _rpname = str(_row.get("name",""))
                _rss    = STATUS_STYLES.get(_rstat, {"bg":"#F1F5F9","text":"#64748B","dot":"#94A3B8"})
                _rcols  = st.columns([0.3,2.5,1.2,1.2,1.2,1.0,1.0,1.0], vertical_alignment="center")
                _rcols[0].markdown(cell(_rid, size="10px", color="#94A3B8"), unsafe_allow_html=True)
                with _rcols[1]:
                    if st.button(_rpname, key=f"ptr_open_{_rid}", use_container_width=True, help="Open tracker"):
                        st.session_state["proj_tracker_open"] = _rpname
                        st.rerun()
                _rcols[2].markdown(cell(str(_row.get("client","")), size="11px"), unsafe_allow_html=True)
                _rcols[3].markdown(
                    f'<span style="font-size:11px;font-weight:600;color:{_pc["color"]}">'
                    f'{esc(str(_row.get("lead","")))}</span>', unsafe_allow_html=True)
                _rcols[4].markdown(
                    f'<span style="background:{_rss["bg"]};color:{_rss["text"]};font-size:10px;font-weight:700;'
                    f'padding:2px 8px;border-radius:12px;display:inline-flex;align-items:center;gap:4px">'
                    f'<span style="width:5px;height:5px;border-radius:50%;background:{_rss["dot"]};display:inline-block"></span>'
                    f'{esc(_rstat)}</span>', unsafe_allow_html=True)
                _rcols[5].markdown(cell(fmt_date(str(_row.get("start",""))),    size="11px", color="#64748B"), unsafe_allow_html=True)
                _rcols[6].markdown(cell(fmt_date(str(_row.get("end",""))),      size="11px", color="#64748B"), unsafe_allow_html=True)
                _rcols[7].markdown(cell(fmt_date(str(_row.get("due_date",""))), size="11px", color="#64748B"), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB: TASKS  (also "My Tasks" for employee)
# ══════════════════════════════════════════════════════════════════════════════
elif active == "tasks":
    _title = "My Tasks" if role == "employee" else f"{_pc['icon']} {portal} Tasks"
    st.markdown(
        f'<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px">{_title}</h2>',
        unsafe_allow_html=True,
    )

    if role == "employee":
        # User role: own tasks only, scoped to this portal's department
        _my = auth.get_user_tasks(cu["id"])
        _my = [t for t in _my
               if (t.get("department","") or "") == portal or not (t.get("department","") or "")]
        _active_t = [t for t in _my if t.get("status","") != "Completed"]
        _done_t   = [t for t in _my if t.get("status","") == "Completed"]

        _et1, _et2 = st.tabs([f"Active ({len(_active_t)})", f"Completed ({len(_done_t)})"])
        with _et1:
            if not _active_t:
                st.info("No active tasks assigned to you.")
            for _t in _active_t:
                with st.container(border=True):
                    _tl, _tr = st.columns([3, 1.2])
                    _pct = int(_t.get("progress", 0))
                    with _tl:
                        st.markdown(f'<div style="font-size:13px;font-weight:700;color:#111827">{esc(_t["title"])}</div>', unsafe_allow_html=True)
                        if _t.get("description"):
                            st.markdown(f'<div style="font-size:11px;color:#64748B;font-style:italic">{esc(_t["description"])}</div>', unsafe_allow_html=True)
                        st.markdown(
                            f'<div class="progress-bar-outer"><div class="progress-bar-inner" style="width:{_pct}%;background:#3B82F6"></div></div>'
                            f'<div style="font-size:10px;color:#64748B">{_pct}% complete</div>',
                            unsafe_allow_html=True)
                    with _tr:
                        _np = st.slider("Progress", 0, 100, _pct, step=5, key=f"et_prog_{_t['id']}")
                        _ns = st.selectbox("Status", auth.TASK_STATUSES,
                                           index=auth.TASK_STATUSES.index(_t["status"]) if _t["status"] in auth.TASK_STATUSES else 0,
                                           key=f"et_stat_{_t['id']}")
                    if st.button("Save", type="primary", key=f"et_save_{_t['id']}", use_container_width=True):
                        auth.update_task_progress(_t["id"], _np, _ns, _t.get("comment",""))
                        st.session_state.toast = {"msg":"Progress saved!", "type":"success"}
                        st.rerun()
        with _et2:
            if not _done_t:
                st.info("No completed tasks yet.")
            for _dt in _done_t:
                with st.container(border=True):
                    st.markdown(
                        f'<div style="display:flex;justify-content:space-between;align-items:center">'
                        f'<span style="font-size:13px;font-weight:700;color:#111827">{esc(_dt["title"])}</span>'
                        f'<span style="font-size:10px;font-weight:700;background:#D1FAE5;color:#065F46;'
                        f'padding:2px 10px;border-radius:12px">✓ Completed</span></div>',
                        unsafe_allow_html=True,
                    )

    else:
        # Lead / Manager / Admin: full task board for their portal
        _all_t = auth.get_all_tasks()
        _all_t = [t for t in _all_t if (t.get("department","") or "") == portal]

        # Worksoft lead sees only tasks tied to their own projects
        if portal == "Worksoft" and role == "lead":
            _show_isolation_notice()
            _my_proj_names = {str(r.get("name","")).strip() for _, r in _visible_projects().iterrows()}
            _all_t = [t for t in _all_t
                      if str(t.get("project_name","")).strip() in _my_proj_names
                      or t.get("assigned_by") == cu["name"]
                      or t.get("assigned_to") == cu["name"]]

        # Assign task form (lead/manager/admin only)
        if role in ("lead","manager","admin"):
            with st.expander("Assign New Task", expanded=False):
                _assignable = auth.get_employees_and_leads()
                _assignable = [e for e in _assignable
                               if (e.get("department","") or "") == portal
                               or not (e.get("department","") or "")]
                if not _assignable:
                    st.warning(f"No {portal} team members found.")
                else:
                    _ta1, _ta2 = st.columns(2)
                    _nt_title = _ta1.text_input("Task Title *", key="nt_title_portal")
                    _emp_opts = [
                        f"{_e['name']}  [{_ROLE_LABEL.get(_e['role'],_e['role'].upper())}]  ({_e['email']})"
                        for _e in _assignable
                    ]
                    _emp_sel  = _ta2.selectbox("Assign To *", _emp_opts, key="nt_emp_portal")
                    _tb1, _tb2 = st.columns(2)
                    _nt_start_dt = _tb1.date_input("Start Date", value=None, key="nt_start_portal", format="DD/MM/YYYY")
                    _nt_due_dt   = _tb2.date_input("Due Date",   value=None, key="nt_due_portal",   format="DD/MM/YYYY")
                    _nt_desc = st.text_area("Description (optional)", key="nt_desc_portal", height=60)
                    if st.button("Assign Task", type="primary", key="nt_submit_portal"):
                        if not _nt_title.strip():
                            st.error("Title is required.")
                        else:
                            _sel_user = _assignable[_emp_opts.index(_emp_sel)]
                            _st2 = _nt_start_dt.strftime("%Y-%m-%d") if _nt_start_dt else ""
                            _du2 = _nt_due_dt.strftime("%Y-%m-%d")   if _nt_due_dt   else ""
                            auth.create_task(_nt_title.strip(), _nt_desc.strip(),
                                             _sel_user["id"], cu["id"], _du2, _st2, portal)
                            st.session_state.toast = {"msg":f'Task "{_nt_title.strip()}" assigned!', "type":"success"}
                            st.rerun()

        _ip   = [t for t in _all_t if t["status"] == "In Progress"]
        _comp = [t for t in _all_t if t["status"] == "Completed"]
        _hold = [t for t in _all_t if t["status"] == "On Hold"]

        _tab_all, _tab_ip, _tab_comp, _tab_hold = st.tabs([
            f"All ({len(_all_t)})", f"In Progress ({len(_ip)})",
            f"Completed ({len(_comp)})", f"On Hold ({len(_hold)})",
        ])

        def _render_task_list(tlist, sfx):
            _emp_names = sorted({t["assigned_to"] for t in tlist})
            _ff1, _ff2 = st.columns([1.5, 2.5])
            _ef = _ff1.selectbox("Team member", ["All"] + _emp_names, key=f"emp_f_{sfx}")
            _nf = _ff2.text_input("Filter title", placeholder="Search…", key=f"nm_f_{sfx}")
            _vis = [t for t in tlist
                    if (_ef == "All" or t["assigned_to"] == _ef)
                    and (_nf.strip().lower() in t["title"].lower() if _nf.strip() else True)]
            st.markdown(f'<p style="font-size:11px;color:#64748B;margin:4px 0 8px"><b>{len(_vis)}</b> task(s)</p>',
                        unsafe_allow_html=True)
            for _t in _vis:
                with st.container(border=True):
                    _tl, _tr = st.columns([3.5, 1])
                    _pct = int(_t.get("progress", 0))
                    _sc  = _TASK_STAT_COLORS.get(_t["status"], "#94A3B8")
                    with _tl:
                        st.markdown(f'<div style="font-size:13px;font-weight:700;color:#111827">{esc(_t["title"])}</div>', unsafe_allow_html=True)
                        _tmeta = f'Assigned to: <b>{esc(_t["assigned_to"])}</b>'
                        if _t.get("due_date"):
                            _tmeta += f' &nbsp;·&nbsp; Due: <b>{esc(fmt_date(_t["due_date"]))}</b>'
                        st.markdown(f'<div style="font-size:11px;color:#64748B;margin-bottom:4px">{_tmeta}</div>', unsafe_allow_html=True)
                        st.markdown(
                            f'<div class="progress-bar-outer"><div class="progress-bar-inner" style="width:{_pct}%;background:{_sc}"></div></div>'
                            f'<div style="font-size:10px;color:#64748B">{_pct}% · {_t["status"]}</div>',
                            unsafe_allow_html=True)
                    with _tr:
                        if role in ("admin","lead","manager"):
                            if st.button("🗑", key=f"dt_{sfx}_{_t['id']}", help="Delete task", use_container_width=True):
                                auth.delete_task(_t["id"])
                                st.session_state.toast = {"msg":"Task deleted.", "type":"info"}
                                st.rerun()

        with _tab_all:  _render_task_list(_all_t, f"{portal.lower()}_all")
        with _tab_ip:   _render_task_list(_ip,    f"{portal.lower()}_ip")
        with _tab_comp: _render_task_list(_comp,  f"{portal.lower()}_comp")
        with _tab_hold: _render_task_list(_hold,  f"{portal.lower()}_hold")


# ══════════════════════════════════════════════════════════════════════════════
# TAB: LICENSE  (RPA portal — lead / manager / admin only)
# ══════════════════════════════════════════════════════════════════════════════
elif active == "license":
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px">License Management</h2>', unsafe_allow_html=True)

    def _lc_badge(end_date: str) -> str:
        if not end_date: return '<span style="font-size:10px;color:#94A3B8">—</span>'
        try:
            diff = (datetime.strptime(end_date,"%Y-%m-%d").date() - datetime.now().date()).days
            if diff<0:   return f'<span style="background:#FEF2F2;color:#991B1B;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px">Expired</span>'
            if diff<=30: return f'<span style="background:#FEF2F2;color:#DC2626;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px">30d: {diff}d left</span>'
            if diff<=60: return f'<span style="background:#FFFBEB;color:#B45309;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px">60d: {diff}d left</span>'
            if diff<=90: return f'<span style="background:#FEF3C7;color:#92400E;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px">90d: {diff}d left</span>'
            return f'<span style="background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px">Active</span>'
        except ValueError:
            return f'<span style="font-size:11px;color:#64748B">{esc(end_date)}</span>'

    _lplan_opts   = ["","Monthly","Quarterly","Yearly","Lifetime"]
    _lplan_colors = {"Monthly":"#3B82F6","Quarterly":"#8B5CF6","Yearly":"#059669","Lifetime":"#D97706"}
    _lic_all    = auth.get_all_licenses()
    _sl_all     = auth.get_all_sold_licenses()
    _tool_names = sorted({l["tool_name"].strip() for l in _lic_all if l["tool_name"].strip()})

    _lc_t1, _lc_t2 = st.tabs(["Purchased License","Sold License"])

    with _lc_t1:
        if role in ("admin","lead","manager"):
            with st.expander("Add Purchased License", expanded=False):
                _la1,_la2 = st.columns(2)
                _n_tool  = _la1.text_input("Tool Name *", key="plc_n_tool")
                _n_seats = _la2.number_input("No. of Licenses *", min_value=1, value=1, step=1, key="plc_n_seats")
                _lb1,_lb2 = st.columns(2)
                _n_start_dt = _lb1.date_input("Start Date", value=None, key="plc_n_start", format="DD/MM/YYYY")
                _n_end_dt   = _lb2.date_input("End Date",   value=None, key="plc_n_end",   format="DD/MM/YYYY")
                _lc1,_lc2  = st.columns(2)
                _n_plan  = _lc1.selectbox("License Plan", _lplan_opts, key="plc_n_plan")
                _n_email = _lc2.text_input("Notification Email(s)", key="plc_n_email")
                _n_start = _n_start_dt.strftime("%Y-%m-%d") if _n_start_dt else ""
                _n_end   = _n_end_dt.strftime("%Y-%m-%d")   if _n_end_dt   else ""
                if st.button("Add License", type="primary", key="plc_add_btn"):
                    if not _n_tool.strip(): st.error("Tool name required.")
                    else:
                        auth.create_license(_n_tool, int(_n_seats), _n_start, _n_end, _n_email, _n_plan)
                        st.session_state.toast = {"msg":f'License "{_n_tool}" added!', "type":"success"}
                        st.rerun()

        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 12px"><b>{len(_lic_all)}</b> license(s)</p>', unsafe_allow_html=True)
        if not _lic_all:
            st.info("No licenses yet.")
        else:
            with st.container(border=True):
                _lhdr = st.columns([0.3,2.0,0.9,1.1,1.1,1.1,1.3,0.6,0.6])
                for _lc,_ll in zip(_lhdr,["#","Tool Name","Qty","Start","End","Plan","Status","",""]):
                    _lc.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;padding:6px 4px;border-bottom:2px solid #DFE3E7;background:#F8FAFC">{_ll}</div>', unsafe_allow_html=True)
                for _lic in _lic_all:
                    _lr = st.columns([0.3,2.0,0.9,1.1,1.1,1.1,1.3,0.6,0.6], vertical_alignment="center")
                    _lr[0].markdown(cell(_lic["id"],size="10px",color="#94A3B8"), unsafe_allow_html=True)
                    _lr[1].markdown(f'<span style="font-size:13px;font-weight:700;color:#111827">{esc(_lic["tool_name"])}</span>', unsafe_allow_html=True)
                    _lr[2].markdown(f'<span style="font-size:13px;font-weight:600;color:#3F8E91">{_lic["no_of_licenses"]}</span>', unsafe_allow_html=True)
                    _lr[3].markdown(cell(_lic["start_date"] or "—",size="11px",color="#64748B"), unsafe_allow_html=True)
                    _lr[4].markdown(cell(_lic["end_date"]   or "—",size="11px",color="#64748B"), unsafe_allow_html=True)
                    _lpv = _lic.get("license_plan","") or "—"
                    _lr[5].markdown(f'<span style="font-size:10px;font-weight:700;color:{_lplan_colors.get(_lpv,"#94A3B8")}">{esc(_lpv)}</span>', unsafe_allow_html=True)
                    _lr[6].markdown(_lc_badge(_lic["end_date"]), unsafe_allow_html=True)
                    if role in ("admin","lead","manager"):
                        with _lr[7]:
                            if st.button("✏", key=f"plc_e_{_lic['id']}", use_container_width=True):
                                st.session_state.lc_edit_id = _lic["id"]
                                st.rerun()
                        with _lr[8]:
                            if st.button("🗑", key=f"plc_d_{_lic['id']}", use_container_width=True):
                                auth.delete_license(_lic["id"])
                                st.session_state.toast = {"msg":f'Deleted "{_lic["tool_name"]}"', "type":"info"}
                                st.rerun()

    with _lc_t2:
        if role in ("admin","lead","manager"):
            with st.expander("Add Sold License", expanded=False):
                if not _tool_names:
                    st.info("Add a purchased license first.")
                else:
                    _sa1,_sa2 = st.columns(2)
                    _sn_tool   = _sa1.selectbox("Tool *", _tool_names, key="psl_n_tool")
                    _sn_client = _sa2.text_input("Client *", key="psl_n_client")
                    _sb1,_sb2  = st.columns(2)
                    _sn_seats  = _sb1.number_input("Qty *", min_value=1, value=1, step=1, key="psl_n_seats")
                    _sn_notes  = _sb2.text_input("Notes", key="psl_n_notes")
                    _sc1,_sc2  = st.columns(2)
                    _sn_start_dt = _sc1.date_input("Start", value=None, key="psl_n_start", format="DD/MM/YYYY")
                    _sn_end_dt   = _sc2.date_input("End",   value=None, key="psl_n_end",   format="DD/MM/YYYY")
                    _sd1,_sd2  = st.columns(2)
                    _sn_plan   = _sd1.selectbox("License Plan", _lplan_opts, key="psl_n_plan")
                    _sn_email  = _sd2.text_input("Client Email", key="psl_n_email")
                    _sn_start  = _sn_start_dt.strftime("%Y-%m-%d") if _sn_start_dt else ""
                    _sn_end    = _sn_end_dt.strftime("%Y-%m-%d")   if _sn_end_dt   else ""
                    if st.button("Add Sold License", type="primary", key="psl_add_btn"):
                        if not _sn_client.strip(): st.error("Client required.")
                        else:
                            auth.create_sold_license(_sn_tool, _sn_client, int(_sn_seats),
                                                     _sn_start, _sn_end, _sn_notes, _sn_email, _sn_plan)
                            st.session_state.toast = {"msg":"Sold license added!", "type":"success"}
                            st.rerun()

        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 12px"><b>{len(_sl_all)}</b> sold license record(s)</p>', unsafe_allow_html=True)
        if not _sl_all:
            st.info("No sold licenses yet.")
        else:
            with st.container(border=True):
                _slhdr = st.columns([0.3,1.6,1.6,0.8,1.0,1.0,1.0,1.1,0.6,0.6])
                for _sc,_sl in zip(_slhdr,["#","Tool","Client","Qty","Start","End","Plan","Status","",""]):
                    _sc.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;padding:6px 4px;border-bottom:2px solid #DFE3E7;background:#F8FAFC">{_sl}</div>', unsafe_allow_html=True)
                for _sl in _sl_all:
                    _slr = st.columns([0.3,1.6,1.6,0.8,1.0,1.0,1.0,1.1,0.6,0.6], vertical_alignment="center")
                    _slr[0].markdown(cell(_sl["id"],size="10px",color="#94A3B8"), unsafe_allow_html=True)
                    _slr[1].markdown(f'<span style="font-size:12px;font-weight:700;color:#111827">{esc(_sl["tool_name"])}</span>', unsafe_allow_html=True)
                    _slr[2].markdown(f'<span style="font-size:12px;color:#374151">{esc(_sl["client"])}</span>', unsafe_allow_html=True)
                    _slr[3].markdown(f'<span style="font-size:12px;font-weight:600;color:#3F8E91">{_sl["no_of_licenses"]}</span>', unsafe_allow_html=True)
                    _slr[4].markdown(cell(_sl["start_date"] or "—",size="11px",color="#64748B"), unsafe_allow_html=True)
                    _slr[5].markdown(cell(_sl["end_date"]   or "—",size="11px",color="#64748B"), unsafe_allow_html=True)
                    _slpv = _sl.get("license_plan","") or "—"
                    _slr[6].markdown(f'<span style="font-size:10px;font-weight:700;color:{_lplan_colors.get(_slpv,"#94A3B8")}">{esc(_slpv)}</span>', unsafe_allow_html=True)
                    _slr[7].markdown(_lc_badge(_sl["end_date"]), unsafe_allow_html=True)
                    if role in ("admin","lead","manager"):
                        with _slr[8]:
                            if st.button("✏", key=f"psl_e_{_sl['id']}", use_container_width=True):
                                st.session_state.sl_edit_id = _sl["id"]
                                st.rerun()
                        with _slr[9]:
                            if st.button("🗑", key=f"psl_d_{_sl['id']}", use_container_width=True):
                                auth.delete_sold_license(_sl["id"])
                                st.session_state.toast = {"msg":"Sold license deleted.", "type":"info"}
                                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB: USERS  (admin only)
# ══════════════════════════════════════════════════════════════════════════════
elif active == "users" and role == "admin":
    st.markdown(
        f'<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px">'
        f'User Management — {portal} Portal</h2>',
        unsafe_allow_html=True,
    )
    st.markdown(f'<p style="color:#64748B;font-size:12px;margin-bottom:16px">'
                f'Showing {portal} department users + admins</p>', unsafe_allow_html=True)

    _uc = auth.get_all_users()
    _portal_users = [u for u in _uc if u.get("department","") == portal or u.get("role") == "admin"]

    with st.expander(f"➕ Add {portal} User", expanded=False):
        _ua, _ub = st.columns(2)
        _nu_name  = _ua.text_input("Full Name *", key="pnu_name")
        _nu_email = _ub.text_input("Email *", key="pnu_email")
        _uc2, _ud = st.columns(2)
        _nu_pass = _uc2.text_input("Password *", type="password", key="pnu_pass")
        _nu_role = _ud.selectbox("Role", auth.ROLES, key="pnu_role")
        if st.button(f"Create {portal} User", type="primary", key="pnu_create"):
            _errs = []
            if not _nu_name.strip():                           _errs.append("Name required.")
            if not _nu_email.strip() or "@" not in _nu_email: _errs.append("Valid email required.")
            if not _nu_pass or len(_nu_pass) < 6:             _errs.append("Password min 6 chars.")
            if _errs:
                for e in _errs: st.error(e)
            else:
                try:
                    auth.create_user(_nu_name.strip(), _nu_email.strip(), _nu_pass, _nu_role, portal)
                    st.session_state.toast = {"msg":f'User "{_nu_name.strip()}" created!', "type":"success"}
                    st.rerun()
                except Exception as _ex:
                    st.error(f"Error: {_ex}")

    st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 10px"><b>{len(_portal_users)}</b> user(s)</p>',
                unsafe_allow_html=True)
    with st.container(border=True):
        _uhdr = st.columns([0.3,1.8,2.5,1.0,0.8,1.2,0.5,0.5,0.5])
        for _c,_l in zip(_uhdr,["ID","Name","Email","Role","Active","Dept","","",""]):
            _c.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;padding:6px 4px;border-bottom:2px solid #DFE3E7;background:#F8FAFC">{_l}</div>', unsafe_allow_html=True)
        _rc2 = {"admin":"#3F8E91","lead":"#2E7D5B","manager":"#966D17","employee":"#4E5860","sales":"#5FA9AB"}
        for _u in _portal_users:
            _ur = st.columns([0.3,1.8,2.5,1.0,0.8,1.2,0.5,0.5,0.5], vertical_alignment="center")
            _ur[0].markdown(cell(_u["id"],size="10px",color="#94A3B8"), unsafe_allow_html=True)
            _ur[1].markdown(f'<span style="font-size:12px;font-weight:600;color:#111827">{esc(_u["name"])}</span>', unsafe_allow_html=True)
            _ur[2].markdown(cell(_u["email"]), unsafe_allow_html=True)
            _rlbl = _ROLE_LABEL.get(_u["role"], _u["role"].upper())
            _ur[3].markdown(f'<span style="font-size:11px;font-weight:700;color:{_rc2.get(_u["role"],"#374151")}">{_rlbl}</span>', unsafe_allow_html=True)
            _ur[4].markdown(f'<span style="font-size:11px;font-weight:700;color:{"#10B981" if _u["is_active"] else "#EF4444"}">{"Yes" if _u["is_active"] else "No"}</span>', unsafe_allow_html=True)
            _udept   = _u.get("department","") or "—"
            _udept_c = "#3F8E91" if _udept=="RPA" else ("#7C3AED" if _udept=="Worksoft" else "#94A3B8")
            _ur[5].markdown(f'<span style="font-size:11px;font-weight:700;color:{_udept_c}">{esc(_udept)}</span>', unsafe_allow_html=True)
            with _ur[6]:
                if st.button("✏", key=f"peu_{_u['id']}", use_container_width=True):
                    st.session_state.user_edit_id = _u["id"]
                    st.rerun()
            with _ur[7]:
                _tl = "🔒" if _u["is_active"] else "🔓"
                if st.button(_tl, key=f"ptog_{_u['id']}", use_container_width=True):
                    if _u["id"] != cu["id"]:
                        auth.set_active(_u["id"], not _u["is_active"])
                        st.session_state.toast = {"msg":"User status updated.", "type":"info"}
                        st.rerun()
            with _ur[8]:
                if st.button("🗑", key=f"pdu_{_u['id']}", use_container_width=True):
                    if _u["id"] != cu["id"]:
                        auth.delete_user(_u["id"])
                        st.session_state.toast = {"msg":"User deleted.", "type":"info"}
                        st.rerun()

    if st.session_state.user_edit_id:
        _eu_rec = next((u for u in _uc if u["id"] == st.session_state.user_edit_id), None)
        if _eu_rec:
            with st.container(border=True):
                st.markdown(f'<div style="font-size:13px;font-weight:700;color:#1F3B4D;margin-bottom:8px">Edit — {esc(_eu_rec["name"])}</div>', unsafe_allow_html=True)
                _ea, _eb = st.columns(2)
                _eu_name  = _ea.text_input("Full Name *", value=_eu_rec["name"],  key="peu_name")
                _eu_email = _eb.text_input("Email *",     value=_eu_rec["email"], key="peu_email")
                _ec, _ed  = st.columns(2)
                _eu_role  = _ec.selectbox("Role", auth.ROLES,
                                          index=auth.ROLES.index(_eu_rec["role"]) if _eu_rec["role"] in auth.ROLES else 0,
                                          key="peu_role")
                _eu_dept_opts = ["","RPA","Worksoft"]
                _eu_dept  = _ed.selectbox("Department", _eu_dept_opts,
                                          index=_eu_dept_opts.index(_eu_rec.get("department","")) if _eu_rec.get("department","") in _eu_dept_opts else 0,
                                          key="peu_dept")
                _es1, _es2 = st.columns([1, 4])
                if _es1.button("Save", type="primary", key="peu_save"):
                    if _eu_name.strip() and _eu_email.strip():
                        auth.update_user(st.session_state.user_edit_id, _eu_name, _eu_email, _eu_role, _eu_dept)
                        st.session_state.user_edit_id = None
                        st.session_state.toast = {"msg":f'"{_eu_name}" updated!', "type":"success"}
                        st.rerun()
                if _es2.button("Cancel", key="peu_cancel"):
                    st.session_state.user_edit_id = None
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB: SETTINGS  (admin only)
# ══════════════════════════════════════════════════════════════════════════════
elif active == "settings" and role == "admin":
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:16px">Settings</h2>', unsafe_allow_html=True)

    with st.expander("Email (Outlook) Settings", expanded=True):
        _cfg = auth.get_email_settings()
        if _cfg.get("outlook_email"):
            st.success(f'Configured: {_cfg["outlook_email"]}')
        else:
            st.warning("Email not configured — notifications disabled.")
        _oc1, _oc2 = st.columns(2)
        _ne = _oc1.text_input("Outlook Email",  value=_cfg.get("outlook_email",""),    key="ps_email")
        _np = _oc2.text_input("Password",        value=_cfg.get("outlook_password",""), type="password", key="ps_pwd")
        _os1, _os2 = st.columns([1, 3])
        if _os1.button("Save", type="primary", key="ps_save"):
            if _ne.strip() and _np.strip():
                auth.save_email_settings(_ne.strip(), _np.strip())
                st.session_state.toast = {"msg":"Email settings saved!", "type":"success"}
                st.rerun()
        if _os2.button("Clear", key="ps_clear"):
            auth.save_email_settings("","")
            st.session_state.toast = {"msg":"Email settings cleared.", "type":"info"}
            st.rerun()

    with st.expander("Export Database", expanded=False):
        _ec1, _ec2 = st.columns(2)
        with _ec1:
            if st.button("Export as Excel", type="primary", key="pex_excel", use_container_width=True):
                _econn = auth.get_conn()
                _ebuf  = io.BytesIO()
                try:
                    _etables = [r[0] for r in _econn.execute(
                        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").fetchall()]
                    with pd.ExcelWriter(_ebuf, engine="openpyxl") as _xw:
                        for _tbl in _etables:
                            try: pd.read_sql_query(f"SELECT * FROM [{_tbl}]", _econn).to_excel(_xw, sheet_name=_tbl[:31], index=False)
                            except Exception: pass
                    _econn.close()
                    st.session_state["_pex_excel"] = _ebuf.getvalue()
                    st.session_state.toast = {"msg":"Excel ready — click Download.", "type":"success"}
                    st.rerun()
                except Exception as _ee:
                    st.error(f"Export failed: {_ee}")
            if st.session_state.get("_pex_excel"):
                st.download_button("Download Excel", data=st.session_state["_pex_excel"],
                                   file_name=f"qualesce_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="pex_dl_excel", use_container_width=True)
        with _ec2:
            if st.button("Download .db", key="pex_db", use_container_width=True):
                try:
                    _db_buf = io.BytesIO()
                    with open(auth.DB_PATH,"rb") as _dbf: _db_buf.write(_dbf.read())
                    st.session_state["_pex_db"]      = _db_buf.getvalue()
                    st.session_state["_pex_db_name"] = f"qualesce_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
                    st.rerun()
                except Exception as _de:
                    st.error(f"Error: {_de}")
            if st.session_state.get("_pex_db"):
                st.download_button("Download SQLite DB", data=st.session_state["_pex_db"],
                                   file_name=st.session_state.get("_pex_db_name","qualesce.db"),
                                   mime="application/octet-stream",
                                   key="pex_dl_db", use_container_width=True)
