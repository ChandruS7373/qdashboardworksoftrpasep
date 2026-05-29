"""
portal_app.py — Qualesce Portal Dashboard
Separate portals for RPA and Worksoft.
  • Users with department=RPA  → RPA Portal only
  • Users with department=Worksoft → Worksoft Portal only
  • Admin → can switch portals via top bar
  • Worksoft Portal: leads see ONLY their own projects; managers/admins see all.
Run with:  streamlit run portal_app.py --server.port 8502
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os, html, re, threading, io
import base64 as _b64
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
_DEV_STATUSES       = ["In Progress","PDD"]
_RM_STATUSES        = ["R&M"]
_COMPLETED_STATUSES = ["Completed"]
_UAT_STATUSES       = ["UAT"]
_DISC_STATUSES      = ["Discontinued"]

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

def _parse_ymd(s: str):
    v = str(s).strip()
    for fmt in ("%Y-%m-%d","%d-%m-%Y","%d/%m/%Y"):
        try: return datetime.strptime(v, fmt).date()
        except ValueError: pass
    return None

def cell(val, size="11px", color="#374151"):
    return f'<span style="font-size:{size};color:{color}">{esc(str(val))}</span>'

def badge_html(status: str) -> str:
    s = STATUS_STYLES.get(status, {"bg":"#F1F5F9","text":"#475569","dot":"#94A3B8"})
    return (f'<span style="display:inline-flex;align-items:center;gap:4px;padding:3px 10px;'
            f'border-radius:20px;font-size:11px;font-weight:700;background:{s["bg"]};color:{s["text"]}">'
            f'<span style="width:6px;height:6px;border-radius:50%;background:{s["dot"]};display:inline-block"></span>'
            f'{esc(status)}</span>')

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
        "total": len(d),
        "in_progress": int((d["status"]=="In Progress").sum()),
        "completed":   int((d["status"]=="Completed").sum()),
        "uat":         int((d["status"]=="UAT").sum()),
        "rm":          int((d["status"]=="R&M").sum()),
        "new_added":   int(d["is_new"].astype(str).str.lower().isin(["true","1","yes"]).sum()),
        "total_hrs":   pd.to_numeric(d.get("hours_saved",pd.Series(dtype=float)), errors="coerce").fillna(0).sum(),
        "total_cost":  pd.to_numeric(d.get("cost_saved", pd.Series(dtype=float)), errors="coerce").fillna(0).sum(),
    }

# ── SESSION STATE INIT ────────────────────────────────────────────────────────
_SS = {
    "current_user": None, "active_tab": "dashboard", "toast": None,
    "active_portal": "RPA", "login_attempts": 0, "projects": None,
    "proj_tracker_open": None, "user_edit_id": None, "reset_pwd_uid": None,
    "lc_edit_id": None, "sl_edit_id": None, "sl_mail_id": None,
    "lc_mail_id": None, "lc_last_notif_check": "", "save_popup": None,
    "task_popup": None,
}
for _k, _v in _SS.items():
    if _k not in st.session_state: st.session_state[_k] = _v

# ── GLOBAL CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.block-container,[data-testid="stMainBlockContainer"]{padding-top:0!important}
.stTabs [data-baseweb="tab"]{font-size:12px;font-weight:600}
.progress-bar-outer{background:#E2E8F0;border-radius:6px;height:8px;overflow:hidden;margin:2px 0}
.progress-bar-inner{height:8px;border-radius:6px;transition:width .3s ease}
.portal-header{
  background:linear-gradient(90deg,#162C3B,#1F3B4D);
  border-radius:12px;padding:16px 24px;margin-bottom:16px;
  display:flex;align-items:center;justify-content:space-between
}
</style>""", unsafe_allow_html=True)

# ── LOGIN GATE ────────────────────────────────────────────────────────────────
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
                    st.session_state.current_user  = _u
                    st.session_state.login_attempts = 0
                    st.session_state.active_tab    = "tasks" if _u["role"] == "employee" else "dashboard"
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
    _locked_portal = None   # admin / no dept → can switch

portal = _locked_portal or st.session_state.get("active_portal", "RPA")

# ── PORTAL VISUAL THEME ───────────────────────────────────────────────────────
_P = {
    "RPA":      {"color":"#3F8E91","bg":"#EFF7F7","dark":"#162C3B","icon":"🔧","label":"RPA Portal"},
    "Worksoft": {"color":"#7C3AED","bg":"#F5F3FF","dark":"#1E1B4B","icon":"⚙️","label":"Worksoft Portal"},
}
_pc = _P[portal]

# ── TOAST ─────────────────────────────────────────────────────────────────────
if st.session_state.toast:
    _t = st.session_state.toast
    _tc = {"success":("#064E3B","#10B981","✓"),"error":("#7F1D1D","#EF4444","✗"),"info":("#244E51","#5FA9AB","ℹ")}
    _tbg, _tbd, _tico = _tc.get(_t.get("type","success"),("#064E3B","#10B981","✓"))
    st.markdown(
        f'<div style="background:{_tbg};border:1px solid {_tbd};border-radius:10px;'
        f'padding:10px 18px;color:#fff;font-size:13px;font-weight:600;margin-bottom:10px;'
        f'display:flex;align-items:center;gap:10px">'
        f'<span>{_tico}</span><span>{html.escape(_t["msg"])}</span></div>',
        unsafe_allow_html=True)
    st.session_state.toast = None

# ── PORTAL HEADER ─────────────────────────────────────────────────────────────
_hc1, _hc2 = st.columns([3, 1])
with _hc1:
    st.markdown(
        f'<div style="background:linear-gradient(90deg,{_pc["dark"]},#1F3B4D);'
        f'border-radius:12px;padding:14px 22px;margin-bottom:12px;'
        f'display:flex;align-items:center;gap:14px">'
        f'<span style="font-size:26px">{_pc["icon"]}</span>'
        f'<div>'
        f'<div style="font-size:18px;font-weight:900;color:#fff;letter-spacing:-.4px">{_pc["label"]}</div>'
        f'<div style="font-size:11px;color:#94A3B8;margin-top:2px">'
        f'Logged in as <b style="color:#E2E8F0">{esc(cu["name"])}</b>'
        f' &nbsp;·&nbsp; <span style="color:{_pc["color"]};font-weight:700">{role.upper()}</span>'
        f'</div></div></div>',
        unsafe_allow_html=True,
    )
with _hc2:
    _btn_col, _logout_col = st.columns(2)
    # Portal switcher (admins and users without a fixed dept only)
    if _locked_portal is None:
        _sw = _btn_col.selectbox(
            "Portal", ["RPA","Worksoft"],
            index=0 if portal=="RPA" else 1,
            key="portal_switch_sel",
            label_visibility="collapsed",
        )
        if _sw != portal:
            st.session_state.active_portal = _sw
            st.session_state.active_tab = "dashboard"
            st.rerun()
    else:
        _btn_col.markdown(
            f'<div style="text-align:center;padding:8px 0;font-size:11px;'
            f'font-weight:700;color:{_pc["color"]};background:{_pc["bg"]};'
            f'border-radius:8px">{_pc["icon"]} {portal}</div>',
            unsafe_allow_html=True,
        )
    if _logout_col.button("Logout", use_container_width=True, key="portal_logout"):
        st.session_state.current_user = None
        st.rerun()

# ── NAVIGATION ────────────────────────────────────────────────────────────────
if role == "employee":
    _tab_defs = [("tasks","My Tasks"),("projects","Projects")]
elif portal == "RPA":
    if role == "admin":
        _tab_defs = [("dashboard","Dashboard"),("projects","Projects"),
                     ("tasks","Tasks"),("license","License"),("users","Users"),("settings","Settings")]
    else:
        _tab_defs = [("dashboard","Dashboard"),("projects","Projects"),
                     ("tasks","Tasks"),("license","License")]
else:  # Worksoft portal
    if role == "admin":
        _tab_defs = [("dashboard","Dashboard"),("projects","Projects"),
                     ("tasks","Tasks"),("users","Users"),("settings","Settings")]
    else:
        _tab_defs = [("dashboard","Dashboard"),("projects","Projects"),("tasks","Tasks")]

_valid = [t[0] for t in _tab_defs]
if st.session_state.active_tab not in _valid:
    st.session_state.active_tab = _tab_defs[0][0]

_nav_cols = st.columns(len(_tab_defs))
for _ni, (_tid, _tlbl) in enumerate(_tab_defs):
    _is_active = st.session_state.active_tab == _tid
    _btn_type = "primary" if _is_active else "secondary"
    if _nav_cols[_ni].button(_tlbl, key=f"nav_{_tid}", use_container_width=True, type=_btn_type):
        st.session_state.active_tab = _tid
        st.rerun()

st.markdown(f'<hr style="margin:6px 0 16px;border:none;border-top:2px solid {_pc["color"]}40">', unsafe_allow_html=True)

active = st.session_state.active_tab

# ── FILTER PROJECTS BY PORTAL ─────────────────────────────────────────────────
def _portal_projects() -> pd.DataFrame:
    """Return projects belonging to the current portal."""
    if df is None or df.empty: return pd.DataFrame(columns=PROJECT_COLS)
    _ptype = portal  # "RPA" or "Worksoft"
    if "proj_type" in df.columns:
        return df[df["proj_type"].fillna("").str.strip() == _ptype].copy()
    return df.copy()

def _apply_worksoft_lead_filter(base: pd.DataFrame) -> pd.DataFrame:
    """In Worksoft portal, leads can only see their own projects."""
    if portal != "Worksoft" or role not in ("lead",):
        return base
    _name  = str(cu.get("name","")).strip()
    _email = str(cu.get("email","")).strip()
    _mask = (
        base["lead"].fillna("").str.strip() == _name
    )
    if "project_lead_email" in base.columns:
        _mask = _mask | (base["project_lead_email"].fillna("").str.strip() == _email)
    return base[_mask].copy()

# ──────────────────────────────────────────────────────────────────────────────
# TAB: DASHBOARD
# ──────────────────────────────────────────────────────────────────────────────
if active == "dashboard":
    _pdf = _portal_projects()
    _pdf = _apply_worksoft_lead_filter(_pdf) if portal == "Worksoft" and role == "lead" else _pdf
    _s   = get_stats(_pdf)

    st.markdown(
        f'<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:16px">'
        f'{_pc["icon"]} {portal} Dashboard</h2>',
        unsafe_allow_html=True,
    )

    # ── KPI row ───────────────────────────────────────────────────────────────
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
    _kpi(_k1, "Total Projects",   _s["total"],       _pc["color"])
    _kpi(_k2, "In Progress",      _s["in_progress"], "#3B82F6")
    _kpi(_k3, "Completed",        _s["completed"],   "#10B981")
    _kpi(_k4, "Hrs Saved",        f'{_s["total_hrs"]:,.0f}', "#F59E0B")
    _kpi(_k5, "Cost Saved",       f'₹{_s["total_cost"]:,.0f}', "#8B5CF6")

    st.markdown("<br>", unsafe_allow_html=True)

    if _pdf.empty:
        st.info(f"No {portal} projects found.")
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

        # Recent activity
        with st.container(border=True):
            st.markdown('<div style="font-size:11px;font-weight:700;color:#64748B;margin-bottom:8px">Recent Projects</div>', unsafe_allow_html=True)
            _rec = _pdf.head(10)
            for _, _rr in _rec.iterrows():
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

# ──────────────────────────────────────────────────────────────────────────────
# TAB: PROJECTS
# ──────────────────────────────────────────────────────────────────────────────
elif active == "projects":
    st.markdown(
        f'<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px">'
        f'{_pc["icon"]} {portal} Projects</h2>',
        unsafe_allow_html=True,
    )

    if portal == "Worksoft" and role == "lead":
        st.markdown(
            f'<div style="background:#F5F3FF;border:1px solid #C4B5FD;border-radius:8px;'
            f'padding:8px 14px;font-size:11px;color:#5B21B6;margin-bottom:10px">'
            f'🔒 Showing only projects assigned to you as Lead.</div>',
            unsafe_allow_html=True,
        )

    # ── Project tracker detail page ───────────────────────────────────────────
    if st.session_state.get("proj_tracker_open"):
        _psel = st.session_state["proj_tracker_open"]
        _bk, _ttl = st.columns([1,9])
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
        _all_p = _portal_projects()
        _all_p = _apply_worksoft_lead_filter(_all_p)
        _pmatch = _all_p[_all_p["name"] == _psel]
        if _pmatch.empty:
            st.warning("Project not found.")
        else:
            _pr = _pmatch.iloc[0]
            def _pv(k, fb="—"):
                v = _pr.get(k, fb); s = str(v).strip()
                return fb if s in ("","nan","None","NaN") else s

            _pstatus = _pv("status","")
            _pcolor  = STATUS_STYLES.get(_pstatus,{"dot":"#94A3B8"})["dot"]
            _tc1, _tc2 = st.columns([1,1.3])
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
                    _details = (
                        _pfield("Client",   _pv("client"))
                      + _pfield("Lead",     _pv("lead"))
                      + _pfield("Employee", _pv("employee"))
                      + _pfield("Status",   _pstatus, _pcolor)
                      + _pfield("PO No.",   _pv("po"))
                      + _pfield("Start",    fmt_date(_pv("start","")) or "—")
                      + _pfield("End",      fmt_date(_pv("end",""))   or "Ongoing")
                      + _pfield("Due Date", fmt_date(_pv("due_date","")) or "—")
                    )
                    if _pv("desc","") not in ("","—"):
                        _details += _pfield("Description", _pv("desc"))
                    st.markdown(_details, unsafe_allow_html=True)

                    # Timeline progress bar
                    _ts = _parse_dmy(_pv("start",""))
                    _te = _parse_dmy(_pv("end",""))
                    if _ts and _te and _te > _ts:
                        _today = date.today()
                        _tot = (_te - _ts).days
                        _ela = max(0, min(_tot, (_today - _ts).days))
                        _pct = 100 if _pstatus=="Completed" else round((_ela/_tot)*100)
                        _tc = "#10B981" if _pct<70 else ("#F59E0B" if _pct<90 else "#EF4444")
                        if _pstatus=="Completed": _tc="#10B981"
                        _dl = (_te - _today).days
                        _dlbl = f"{_dl}d left" if _dl>0 else ("Completed" if _pstatus=="Completed" else f"{abs(_dl)}d overdue")
                        st.markdown(
                            f'<div style="margin-top:12px;padding-top:10px;border-top:1px solid #F1F5F9">'
                            f'<div style="display:flex;justify-content:space-between;margin-bottom:4px">'
                            f'<span style="font-size:10px;font-weight:700;color:#64748B">Timeline</span>'
                            f'<span style="font-size:10px;font-weight:700;color:{_tc}">{_dlbl}</span></div>'
                            f'<div style="background:#E2E8F0;border-radius:6px;height:8px;overflow:hidden">'
                            f'<div style="width:{_pct}%;background:{_tc};height:8px;border-radius:6px"></div>'
                            f'</div>'
                            f'<div style="display:flex;justify-content:space-between;margin-top:3px">'
                            f'<span style="font-size:9px;color:#94A3B8">{fmt_date(_pv("start",""))}</span>'
                            f'<span style="font-size:10px;font-weight:700;color:{_tc}">{_pct}%</span>'
                            f'<span style="font-size:9px;color:#94A3B8">{fmt_date(_pv("end",""))}</span>'
                            f'</div></div>',
                            unsafe_allow_html=True,
                        )

            with _tc2:
                with st.container(border=True):
                    st.markdown('<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:12px">Metrics</div>', unsafe_allow_html=True)
                    _mh = _pv("hours_saved","")
                    _mc = _pv("cost_saved","")
                    _mr = _pv("roi_pct","")
                    _mhtml = ""
                    if _mh not in ("","—","0","0.0"):
                        _mhtml += (f'<div style="flex:1;background:#F0FDF4;border-radius:10px;padding:12px;border-top:3px solid #10B981;text-align:center">'
                                   f'<div style="font-size:20px;font-weight:900;color:#059669">{_mh}</div>'
                                   f'<div style="font-size:10px;color:#14532D;margin-top:3px">Hours Saved</div></div>')
                    if _mc not in ("","—","0","0.0"):
                        _mhtml += (f'<div style="flex:1;background:#EFF7F7;border-radius:10px;padding:12px;border-top:3px solid #5FA9AB;text-align:center">'
                                   f'<div style="font-size:20px;font-weight:900;color:#3F8E91">₹{_mc}</div>'
                                   f'<div style="font-size:10px;color:#1E40AF;margin-top:3px">Cost Saved</div></div>')
                    if _mr not in ("","—","0","0.0"):
                        _mhtml += (f'<div style="flex:1;background:#FFF7ED;border-radius:10px;padding:12px;border-top:3px solid #F97316;text-align:center">'
                                   f'<div style="font-size:20px;font-weight:900;color:#C2410C">{_mr}%</div>'
                                   f'<div style="font-size:10px;color:#9A3412;margin-top:3px">ROI</div></div>')
                    if _mhtml:
                        st.markdown(f'<div style="display:flex;gap:8px;margin-bottom:12px">{_mhtml}</div>', unsafe_allow_html=True)

                    # Bot metrics (RPA only)
                    _pid = int(float(_pr.get("id",0) or 0))
                    if portal == "RPA" and _pid:
                        st.markdown('<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">Bot Metrics</div>', unsafe_allow_html=True)
                        _nb = int(float(_pr.get("num_bots",0) or 0))
                        _np = int(float(_pr.get("num_persons",0) or 0))
                        _mr2 = float(_pr.get("manual_run_mins",0) or 0)
                        _br2 = float(_pr.get("bot_run_mins",0) or 0)
                        _bm_logs = auth.get_bot_metric_logs(project_id=_pid)
                        _ms = date.today().replace(day=1).isoformat()
                        _me = date.today().isoformat()
                        _cur = auth.get_bot_metric_logs(project_id=_pid, start_date=_ms, end_date=_me)
                        _mq  = sum(int(l.get("qty",0) or 0) for l in _cur)
                        _svd = max(float(_mr2)*float(_np) - float(_br2)*float(_nb), 0)*_mq/60
                        _bk1,_bk2,_bk3,_bk4 = st.columns(4)
                        _bk1.metric("Bots",_nb); _bk2.metric("Persons",_np)
                        _bk3.metric("Month Qty",_mq); _bk4.metric("Hrs Saved",f"{_svd:.1f}")
                        if role in ("admin","lead","manager") and _bm_logs:
                            _ldf = pd.DataFrame(_bm_logs[:10])[["log_date","qty"]].copy()
                            _ldf.columns=["Date","Qty"]
                            st.dataframe(_ldf, use_container_width=True, hide_index=True)

                    # Comments
                    st.markdown('<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin:10px 0 6px">Comments</div>', unsafe_allow_html=True)
                    _cmts = auth.get_project_comments(_pid) if _pid else []
                    for _ci, _cm in enumerate(_cmts):
                        _cm_ts = str(_cm["created_at"])
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
                    _new_cmt = st.text_area("Add comment", key=f"ptr_cmt_{_psel}", height=65, label_visibility="collapsed", placeholder="Write a comment…")
                    if st.button("Post", key=f"ptr_post_{_psel}", type="primary"):
                        if _new_cmt.strip() and _pid:
                            auth.add_project_comment(_pid, _psel, cu["id"], cu["name"], _new_cmt.strip())
                            st.rerun()
        st.stop()

    # ── Project listing ───────────────────────────────────────────────────────
    _base = _portal_projects()
    _base = _apply_worksoft_lead_filter(_base)

    # Filters
    _f1, _f2, _f3 = st.columns([2, 2, 2])
    _fstatus  = _f1.selectbox("Status", ["All"] + (WS_STATUSES if portal=="Worksoft" else ALL_STATUSES), key="p_fstat")
    _fclient  = _f2.selectbox("Client", ["All"] + sorted({str(r).strip() for r in _base.get("client", pd.Series(dtype=str)).dropna() if str(r).strip()}), key="p_fcli")
    _fsearch  = _f3.text_input("Search", placeholder="Project name…", key="p_fsrch")

    _filtered = _base.copy()
    if _fstatus != "All": _filtered = _filtered[_filtered["status"]==_fstatus]
    if _fclient != "All": _filtered = _filtered[_filtered["client"].fillna("")==_fclient]
    if _fsearch.strip():
        _q = _fsearch.strip().lower()
        _filtered = _filtered[_filtered["name"].fillna("").str.lower().str.contains(_q)]

    st.markdown(f'<p style="color:#64748B;font-size:12px;margin:4px 0 10px"><b>{len(_filtered)}</b> project(s)</p>', unsafe_allow_html=True)

    if _filtered.empty:
        st.info(f"No {portal} projects match the current filters.")
    else:
        # Table header
        with st.container(border=True):
            _can_edit = role in ("admin","lead","manager")
            _hcols = st.columns([0.3,2.5,1.2,1.2,1.2,1.0,1.0,1.0])
            for _c,_l in zip(_hcols,["#","Project","Client","Lead","Status","Start","End","Due"]):
                _c.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;'
                            f'letter-spacing:.5px;padding:6px 4px;border-bottom:2px solid #DFE3E7;background:#F8FAFC">{_l}</div>', unsafe_allow_html=True)
            for _, _row in _filtered.iterrows():
                _rstat = str(_row.get("status",""))
                _rid   = str(_row.get("id",""))
                _rpname = str(_row.get("name",""))
                _rss   = STATUS_STYLES.get(_rstat,{"bg":"#F1F5F9","text":"#64748B","dot":"#94A3B8"})
                _rcols = st.columns([0.3,2.5,1.2,1.2,1.2,1.0,1.0,1.0], vertical_alignment="center")
                _rcols[0].markdown(cell(_rid,size="10px",color="#94A3B8"), unsafe_allow_html=True)
                with _rcols[1]:
                    _new_tag = (' <span style="font-size:9px;font-weight:700;background:#D9ECEC;color:#3F8E91;padding:1px 5px;border-radius:4px">NEW</span>') if is_new(_row.to_dict()) else ""
                    if st.button(_rpname, key=f"ptr_open_{_rid}", use_container_width=True, help="Open tracker"):
                        st.session_state["proj_tracker_open"] = _rpname
                        st.rerun()
                _rcols[2].markdown(cell(str(_row.get("client","")),size="11px"), unsafe_allow_html=True)
                _rcols[3].markdown(f'<span style="font-size:11px;font-weight:600;color:{_pc["color"]}">{esc(str(_row.get("lead","")))}</span>', unsafe_allow_html=True)
                _rcols[4].markdown(
                    f'<span style="background:{_rss["bg"]};color:{_rss["text"]};font-size:10px;font-weight:700;'
                    f'padding:2px 8px;border-radius:12px;display:inline-flex;align-items:center;gap:4px">'
                    f'<span style="width:5px;height:5px;border-radius:50%;background:{_rss["dot"]};display:inline-block"></span>'
                    f'{esc(_rstat)}</span>',
                    unsafe_allow_html=True,
                )
                _rcols[5].markdown(cell(fmt_date(str(_row.get("start",""))),size="11px",color="#64748B"), unsafe_allow_html=True)
                _rcols[6].markdown(cell(fmt_date(str(_row.get("end",""))),size="11px",color="#64748B"), unsafe_allow_html=True)
                _rcols[7].markdown(cell(fmt_date(str(_row.get("due_date",""))),size="11px",color="#64748B"), unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# TAB: TASKS
# ──────────────────────────────────────────────────────────────────────────────
elif active == "tasks":
    st.markdown(
        f'<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px">'
        f'{_pc["icon"]} {portal} Tasks</h2>',
        unsafe_allow_html=True,
    )

    if role == "employee":
        # Employee: own tasks only
        _my = auth.get_user_tasks(cu["id"])
        # filter by portal dept
        _my = [t for t in _my if (t.get("department","") or "") == portal or not (t.get("department","") or "")]
        _active_t = [t for t in _my if t.get("status","") != "Completed"]
        _done_t   = [t for t in _my if t.get("status","") == "Completed"]

        _et1, _et2 = st.tabs([f"Active ({len(_active_t)})", f"Completed ({len(_done_t)})"])
        with _et1:
            if not _active_t: st.info("No active tasks.")
            for _t in _active_t:
                with st.container(border=True):
                    _tl, _tr = st.columns([3,1.2])
                    _pct = int(_t.get("progress",0))
                    with _tl:
                        st.markdown(f'<div style="font-size:13px;font-weight:700;color:#111827">{esc(_t["title"])}</div>', unsafe_allow_html=True)
                        if _t.get("description"):
                            st.markdown(f'<div style="font-size:11px;color:#64748B;font-style:italic">{esc(_t["description"])}</div>', unsafe_allow_html=True)
                        st.markdown(f'<div class="progress-bar-outer"><div class="progress-bar-inner" style="width:{_pct}%;background:#3B82F6"></div></div>'
                                    f'<div style="font-size:10px;color:#64748B">{_pct}% complete</div>', unsafe_allow_html=True)
                    with _tr:
                        _np = st.slider("Progress", 0, 100, _pct, step=5, key=f"et_prog_{_t['id']}")
                        _ns = st.selectbox("Status", auth.TASK_STATUSES, index=auth.TASK_STATUSES.index(_t["status"]) if _t["status"] in auth.TASK_STATUSES else 0, key=f"et_stat_{_t['id']}")
                    if st.button("Save", type="primary", key=f"et_save_{_t['id']}", use_container_width=True):
                        auth.update_task_progress(_t["id"], _np, _ns, _t.get("comment",""))
                        st.session_state.toast = {"msg":"Progress saved!", "type":"success"}
                        st.rerun()
        with _et2:
            if not _done_t: st.info("No completed tasks yet.")
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
        # Admin/Lead/Manager
        def _render_portal_tasks(dept_filter=""):
            _all_t = auth.get_all_tasks()
            # Filter by portal dept
            _all_t = [t for t in _all_t if (t.get("department","") or "") == portal]
            if dept_filter:
                _all_t = [t for t in _all_t if t.get("department","") == dept_filter]

            # For Worksoft lead: only tasks for their projects
            if portal == "Worksoft" and role == "lead":
                _my_proj_names = set(
                    str(r.get("name","")).strip()
                    for _, r in _apply_worksoft_lead_filter(_portal_projects()).iterrows()
                )
                _all_t = [t for t in _all_t if str(t.get("project_name","")).strip() in _my_proj_names
                          or t.get("assigned_by") == cu["name"] or t.get("assigned_to") == cu["name"]]

            with st.expander("Assign New Task", expanded=False):
                _assignable = auth.get_employees_and_leads()
                if not _assignable:
                    st.warning("No employees found.")
                else:
                    _ta1, _ta2 = st.columns(2)
                    _nt_title = _ta1.text_input("Task Title *", key="nt_title_portal")
                    _emp_opts = [f"{_e['name']}  [{_e['role'].upper()}]  ({_e['email']})" for _e in _assignable]
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

            _ip   = [t for t in _all_t if t["status"]=="In Progress"]
            _comp = [t for t in _all_t if t["status"]=="Completed"]
            _hold = [t for t in _all_t if t["status"]=="On Hold"]

            _tab_all, _tab_ip, _tab_comp, _tab_hold = st.tabs([
                f"All ({len(_all_t)})", f"In Progress ({len(_ip)})",
                f"Completed ({len(_comp)})", f"On Hold ({len(_hold)})",
            ])

            def _render_task_list(tlist, sfx):
                _emp_names = sorted({t["assigned_to"] for t in tlist})
                _ff1, _ff2 = st.columns([1.5, 2.5])
                _ef = _ff1.selectbox("Employee", ["All"]+_emp_names, key=f"emp_f_{sfx}")
                _nf = _ff2.text_input("Filter title", placeholder="Search…", key=f"nm_f_{sfx}")
                _vis = [t for t in tlist if (_ef=="All" or t["assigned_to"]==_ef)
                        and (_nf.strip().lower() in t["title"].lower() if _nf.strip() else True)]
                st.markdown(f'<p style="font-size:11px;color:#64748B;margin:4px 0 8px"><b>{len(_vis)}</b> task(s)</p>', unsafe_allow_html=True)
                for _t in _vis:
                    with st.container(border=True):
                        _tl, _tr = st.columns([3.5, 1])
                        _pct = int(_t.get("progress",0))
                        _sc  = _TASK_STAT_COLORS.get(_t["status"],"#94A3B8")
                        with _tl:
                            st.markdown(f'<div style="font-size:13px;font-weight:700;color:#111827">{esc(_t["title"])}</div>', unsafe_allow_html=True)
                            _tmeta = f'Assigned to: <b>{esc(_t["assigned_to"])}</b>'
                            if _t.get("due_date"): _tmeta += f' &nbsp;·&nbsp; Due: <b>{esc(fmt_date(_t["due_date"]))}</b>'
                            st.markdown(f'<div style="font-size:11px;color:#64748B;margin-bottom:4px">{_tmeta}</div>', unsafe_allow_html=True)
                            st.markdown(f'<div class="progress-bar-outer"><div class="progress-bar-inner" style="width:{_pct}%;background:{_sc}"></div></div>'
                                        f'<div style="font-size:10px;color:#64748B">{_pct}% · {_t["status"]}</div>', unsafe_allow_html=True)
                        with _tr:
                            if st.button("🗑", key=f"dt_{sfx}_{_t['id']}", help="Delete", use_container_width=True):
                                auth.delete_task(_t["id"])
                                st.session_state.toast={"msg":"Task deleted.","type":"info"}
                                st.rerun()

            with _tab_all:  _render_task_list(_all_t,  f"{portal.lower()}_all")
            with _tab_ip:   _render_task_list(_ip,     f"{portal.lower()}_ip")
            with _tab_comp: _render_task_list(_comp,   f"{portal.lower()}_comp")
            with _tab_hold: _render_task_list(_hold,   f"{portal.lower()}_hold")

        _render_portal_tasks()

# ──────────────────────────────────────────────────────────────────────────────
# TAB: LICENSE  (RPA portal or admin)
# ──────────────────────────────────────────────────────────────────────────────
elif active == "license":
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px">License Management</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Track purchased and sold licenses</p>', unsafe_allow_html=True)

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

    _lplan_opts = ["","Monthly","Quarterly","Yearly","Lifetime"]
    _lplan_colors = {"Monthly":"#3B82F6","Quarterly":"#8B5CF6","Yearly":"#059669","Lifetime":"#D97706"}
    _lic_all  = auth.get_all_licenses()
    _sl_all   = auth.get_all_sold_licenses()
    _tool_names = sorted({l["tool_name"].strip() for l in _lic_all if l["tool_name"].strip()})

    _lc_t1, _lc_t2 = st.tabs(["Purchased License","Sold License"])

    with _lc_t1:
        with st.expander("Add Purchased License", expanded=False):
            _la1,_la2 = st.columns(2)
            _n_tool  = _la1.text_input("Tool Name *", key="plc_n_tool")
            _n_seats = _la2.number_input("No. of Licenses *", min_value=1, value=1, step=1, key="plc_n_seats")
            _lb1,_lb2 = st.columns(2)
            _n_start_dt = _lb1.date_input("Start Date", value=None, key="plc_n_start", format="DD/MM/YYYY")
            _n_end_dt   = _lb2.date_input("End Date",   value=None, key="plc_n_end",   format="DD/MM/YYYY")
            _lc1,_lc2 = st.columns(2)
            _n_plan  = _lc1.selectbox("License Plan", _lplan_opts, key="plc_n_plan")
            _n_email = _lc2.text_input("Notification Email(s)", key="plc_n_email", placeholder="email@company.com")
            _n_start = _n_start_dt.strftime("%Y-%m-%d") if _n_start_dt else ""
            _n_end   = _n_end_dt.strftime("%Y-%m-%d")   if _n_end_dt   else ""
            if st.button("Add License", type="primary", key="plc_add_btn"):
                if not _n_tool.strip(): st.error("Tool name required.")
                else:
                    auth.create_license(_n_tool, int(_n_seats), _n_start, _n_end, _n_email, _n_plan)
                    st.session_state.toast={"msg":f'License "{_n_tool}" added!',"type":"success"}
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
                                st.session_state.toast={"msg":f'Deleted "{_lic["tool_name"]}"',"type":"info"}
                                st.rerun()

    with _lc_t2:
        with st.expander("Add Sold License", expanded=False):
            if not _tool_names:
                st.info("Add a purchased license first.")
            else:
                _sa1,_sa2 = st.columns(2)
                _sn_tool   = _sa1.selectbox("Tool *", _tool_names, key="psl_n_tool")
                _sn_client = _sa2.text_input("Client *", key="psl_n_client")
                _sb1,_sb2 = st.columns(2)
                _sn_seats  = _sb1.number_input("Qty *", min_value=1, value=1, step=1, key="psl_n_seats")
                _sn_notes  = _sb2.text_input("Notes", key="psl_n_notes")
                _sc1,_sc2 = st.columns(2)
                _sn_start_dt = _sc1.date_input("Start", value=None, key="psl_n_start", format="DD/MM/YYYY")
                _sn_end_dt   = _sc2.date_input("End",   value=None, key="psl_n_end",   format="DD/MM/YYYY")
                _sd1,_sd2 = st.columns(2)
                _sn_plan  = _sd1.selectbox("License Plan", _lplan_opts, key="psl_n_plan")
                _sn_email = _sd2.text_input("Client Email", key="psl_n_email")
                _sn_start = _sn_start_dt.strftime("%Y-%m-%d") if _sn_start_dt else ""
                _sn_end   = _sn_end_dt.strftime("%Y-%m-%d")   if _sn_end_dt   else ""
                if st.button("Add Sold License", type="primary", key="psl_add_btn"):
                    if not _sn_client.strip(): st.error("Client required.")
                    else:
                        auth.create_sold_license(_sn_tool, _sn_client, int(_sn_seats),
                                                 _sn_start, _sn_end, _sn_notes, _sn_email, _sn_plan)
                        st.session_state.toast={"msg":"Sold license added!","type":"success"}
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
                                st.session_state.toast={"msg":"Sold license deleted.","type":"info"}
                                st.rerun()

# ──────────────────────────────────────────────────────────────────────────────
# TAB: USERS (admin only)
# ──────────────────────────────────────────────────────────────────────────────
elif active == "users" and role == "admin":
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:4px">User Management</h2>', unsafe_allow_html=True)
    st.markdown(f'<p style="color:#64748B;font-size:12px;margin-bottom:16px">Managing users — {portal} Portal view</p>', unsafe_allow_html=True)

    _uc = auth.get_all_users()
    # Show users of this portal's department + admins
    _portal_users = [u for u in _uc if u.get("department","") == portal or u.get("role") == "admin"]

    # Create user
    with st.expander(f"➕ Add {portal} User", expanded=False):
        _ua,_ub = st.columns(2)
        _nu_name  = _ua.text_input("Full Name *", key="pnu_name")
        _nu_email = _ub.text_input("Email *", key="pnu_email")
        _uc2,_ud = st.columns(2)
        _nu_pass = _uc2.text_input("Password *", type="password", key="pnu_pass")
        _nu_role = _ud.selectbox("Role", auth.ROLES, key="pnu_role")
        if st.button(f"Create {portal} User", type="primary", key="pnu_create"):
            _errs = []
            if not _nu_name.strip():                           _errs.append("Name required.")
            if not _nu_email.strip() or "@" not in _nu_email: _errs.append("Valid email required.")
            if not _nu_pass or len(_nu_pass)<6:               _errs.append("Password min 6 chars.")
            if _errs:
                for e in _errs: st.error(e)
            else:
                try:
                    auth.create_user(_nu_name.strip(), _nu_email.strip(), _nu_pass, _nu_role, portal)
                    st.session_state.toast={"msg":f'User "{_nu_name.strip()}" created!',"type":"success"}
                    st.rerun()
                except Exception as _ex:
                    st.error(f"Error: {_ex}")

    st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 10px"><b>{len(_portal_users)}</b> user(s) in {portal} portal</p>', unsafe_allow_html=True)
    with st.container(border=True):
        _uhdr = st.columns([0.3,1.8,2.5,1.0,0.8,1.2,0.5,0.5,0.5])
        for _c,_l in zip(_uhdr,["ID","Name","Email","Role","Active","Dept","","",""]):
            _c.markdown(f'<div style="font-size:10px;font-weight:700;text-transform:uppercase;color:#64748B;padding:6px 4px;border-bottom:2px solid #DFE3E7;background:#F8FAFC">{_l}</div>', unsafe_allow_html=True)
        _rc = {"admin":"#3F8E91","lead":"#2E7D5B","manager":"#966D17","employee":"#4E5860","sales":"#5FA9AB"}
        for _u in _portal_users:
            _ur = st.columns([0.3,1.8,2.5,1.0,0.8,1.2,0.5,0.5,0.5], vertical_alignment="center")
            _ur[0].markdown(cell(_u["id"],size="10px",color="#94A3B8"), unsafe_allow_html=True)
            _ur[1].markdown(f'<span style="font-size:12px;font-weight:600;color:#111827">{esc(_u["name"])}</span>', unsafe_allow_html=True)
            _ur[2].markdown(cell(_u["email"]), unsafe_allow_html=True)
            _ur[3].markdown(f'<span style="font-size:11px;font-weight:700;color:{_rc.get(_u["role"],"#374151")}">{_u["role"].upper()}</span>', unsafe_allow_html=True)
            _ur[4].markdown(f'<span style="font-size:11px;font-weight:700;color:{"#10B981" if _u["is_active"] else "#EF4444"}">{"Yes" if _u["is_active"] else "No"}</span>', unsafe_allow_html=True)
            _udept = _u.get("department","") or "—"
            _udept_c = "#3F8E91" if _udept=="RPA" else ("#7C3AED" if _udept=="Worksoft" else "#94A3B8")
            _ur[5].markdown(f'<span style="font-size:11px;font-weight:700;color:{_udept_c}">{esc(_udept)}</span>', unsafe_allow_html=True)
            with _ur[6]:
                if st.button("✏", key=f"peu_{_u['id']}", use_container_width=True):
                    st.session_state.user_edit_id = _u["id"]
                    st.rerun()
            with _ur[7]:
                _tl = "🔒" if _u["is_active"] else "🔓"
                if st.button(_tl, key=f"ptog_{_u['id']}", use_container_width=True, help="Toggle active"):
                    if _u["id"] != cu["id"]:
                        auth.set_active(_u["id"], not _u["is_active"])
                        st.session_state.toast={"msg":"User status updated.","type":"info"}
                        st.rerun()
            with _ur[8]:
                if st.button("🗑", key=f"pdu_{_u['id']}", use_container_width=True):
                    if _u["id"] != cu["id"]:
                        auth.delete_user(_u["id"])
                        st.session_state.toast={"msg":"User deleted.","type":"info"}
                        st.rerun()

    # Edit user form
    if st.session_state.user_edit_id:
        _eu_rec = next((u for u in _uc if u["id"]==st.session_state.user_edit_id), None)
        if _eu_rec:
            with st.container(border=True):
                st.markdown(f'<div style="font-size:13px;font-weight:700;color:#1F3B4D;margin-bottom:8px">Edit — {esc(_eu_rec["name"])}</div>', unsafe_allow_html=True)
                _ea,_eb = st.columns(2)
                _eu_name  = _ea.text_input("Full Name *", value=_eu_rec["name"],  key="peu_name")
                _eu_email = _eb.text_input("Email *",     value=_eu_rec["email"], key="peu_email")
                _ec,_ed = st.columns(2)
                _eu_role = _ec.selectbox("Role", auth.ROLES, index=auth.ROLES.index(_eu_rec["role"]) if _eu_rec["role"] in auth.ROLES else 0, key="peu_role")
                _eu_dept_opts = ["","RPA","Worksoft"]
                _eu_dept = _ed.selectbox("Department", _eu_dept_opts, index=_eu_dept_opts.index(_eu_rec.get("department","")) if _eu_rec.get("department","") in _eu_dept_opts else 0, key="peu_dept")
                _es1,_es2 = st.columns([1,4])
                if _es1.button("Save", type="primary", key="peu_save"):
                    if _eu_name.strip() and _eu_email.strip():
                        auth.update_user(st.session_state.user_edit_id, _eu_name, _eu_email, _eu_role, _eu_dept)
                        st.session_state.user_edit_id = None
                        st.session_state.toast={"msg":f'"{_eu_name}" updated!',"type":"success"}
                        st.rerun()
                if _es2.button("Cancel", key="peu_cancel"):
                    st.session_state.user_edit_id = None
                    st.rerun()

# ──────────────────────────────────────────────────────────────────────────────
# TAB: SETTINGS
# ──────────────────────────────────────────────────────────────────────────────
elif active == "settings":
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#1F3B4D;margin-bottom:16px">Settings</h2>', unsafe_allow_html=True)

    with st.expander("Email (Outlook) Settings", expanded=True):
        _cfg = auth.get_email_settings()
        if _cfg.get("outlook_email"):
            st.success(f'Configured: {_cfg["outlook_email"]}')
        else:
            st.warning("Email not configured — notifications disabled.")
        _oc1,_oc2 = st.columns(2)
        _ne = _oc1.text_input("Outlook Email", value=_cfg.get("outlook_email",""), key="ps_email")
        _np = _oc2.text_input("Password", value=_cfg.get("outlook_password",""), type="password", key="ps_pwd")
        _os1,_os2 = st.columns([1,3])
        if _os1.button("Save", type="primary", key="ps_save"):
            if _ne.strip() and _np.strip():
                auth.save_email_settings(_ne.strip(), _np.strip())
                st.session_state.toast={"msg":"Email settings saved!","type":"success"}
                st.rerun()
        if _os2.button("Clear", key="ps_clear"):
            auth.save_email_settings("","")
            st.session_state.toast={"msg":"Email settings cleared.","type":"info"}
            st.rerun()

    with st.expander("📦 Export Database", expanded=False):
        _ec1,_ec2 = st.columns(2)
        with _ec1:
            if st.button("📊 Export as Excel", type="primary", key="pex_excel", use_container_width=True):
                _econn = auth.get_conn()
                _ebuf  = io.BytesIO()
                try:
                    _etables = [r[0] for r in _econn.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").fetchall()]
                    with pd.ExcelWriter(_ebuf, engine="openpyxl") as _xw:
                        for _tbl in _etables:
                            try:
                                pd.read_sql_query(f"SELECT * FROM [{_tbl}]", _econn).to_excel(_xw, sheet_name=_tbl[:31], index=False)
                            except Exception: pass
                        try:
                            if "projects" not in _etables:
                                st.session_state.projects.to_excel(_xw, sheet_name="projects", index=False)
                        except Exception: pass
                    _econn.close()
                    st.session_state["_pex_excel"] = _ebuf.getvalue()
                    st.session_state.toast={"msg":"Excel ready — click Download.","type":"success"}
                    st.rerun()
                except Exception as _ee: st.error(f"Export failed: {_ee}")
            if st.session_state.get("_pex_excel"):
                st.download_button("⬇️ Download Excel", data=st.session_state["_pex_excel"],
                                   file_name=f"qualesce_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="pex_dl_excel", use_container_width=True)
        with _ec2:
            if st.button("🗄️ Download .db", key="pex_db", use_container_width=True):
                try:
                    _db_buf = io.BytesIO()
                    with open(auth.DB_PATH,"rb") as _dbf: _db_buf.write(_dbf.read())
                    st.session_state["_pex_db"] = _db_buf.getvalue()
                    st.session_state["_pex_db_name"] = f"qualesce_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
                    st.rerun()
                except Exception as _de: st.error(f"Error: {_de}")
            if st.session_state.get("_pex_db"):
                st.download_button("⬇️ Download SQLite DB", data=st.session_state["_pex_db"],
                                   file_name=st.session_state.get("_pex_db_name","qualesce.db"),
                                   mime="application/octet-stream",
                                   key="pex_dl_db", use_container_width=True)

# ── Init export state ─────────────────────────────────────────────────────────
for _sk in ("_pex_excel","_pex_db","_pex_db_name"):
    if _sk not in st.session_state: st.session_state[_sk] = None
