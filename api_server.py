"""Qualesce HTML Frontend REST API
Runs on port 8503 alongside the Streamlit app (port 8502).
Call start_background() once at Streamlit startup.
"""
import os, sys, threading
from datetime import datetime, date

_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _DIR)
import auth

try:
    from fastapi import FastAPI
    from fastapi.middleware.cors import CORSMiddleware
    from pydantic import BaseModel
    import uvicorn
    _OK = True
except ImportError:
    _OK = False

# ── CLIENT / STATUS MAPS ───────────────────────────────────────────────────────
_CLIENT_LOOKUP = {
    "raychem": "RAY",
    "swagekklok - california": "SWC", "swagelok - california": "SWC",
    "swagelok - alabama": "SWA",
    "tepl": "TEPL",
    "internal poc": "INT",
    "external poc": "EXT",
    "presales": "EXT",
}
CLIENT_INFO = {
    "RAY":  {"name": "Raychem",       "short": "RAY", "color": "#F0ABFC"},
    "SWC":  {"name": "Swagelok — CA", "short": "SWC", "color": "#67E8F9"},
    "SWA":  {"name": "Swagelok — AL", "short": "SWA", "color": "#6EE7B7"},
    "TEPL": {"name": "TEPL",          "short": "TPL", "color": "#A78BFA"},
    "INT":  {"name": "Internal POC",  "short": "INT", "color": "#94A3B8"},
    "EXT":  {"name": "External POC",  "short": "EXT", "color": "#FB923C"},
}
_STATUS_MAP = {
    "r&m": "rm", "uat": "uat", "completed": "completed", "in progress": "progress",
    "pdd": "pdd", "discontinued": "disc", "internal poc": "poc_int",
    "external poc": "poc_ext", "important": "progress", "presales": "pdd",
}
_STATUS_COLORS = {
    "rm": "#60A5FA", "uat": "#FBBF24", "completed": "#6EE7B7", "progress": "#67E8F9",
    "pdd": "#A78BFA", "disc": "#FB7185", "poc_int": "#94A3B8", "poc_ext": "#FB923C",
}
_EMP_PALETTE = [
    "#F0ABFC","#67E8F9","#A78BFA","#6EE7B7","#FB923C",
    "#F472B6","#60A5FA","#FBBF24","#94A3B8","#34D399",
    "#FB7185","#22D3EE","#FBBF24","#A78BFA","#6EE7B7","#F472B6",
]

# ── HELPERS ────────────────────────────────────────────────────────────────────
def _key(name: str) -> str:
    """Generate short unique-ish key from name."""
    p = [w for w in name.strip().split() if w]
    if len(p) >= 2:
        return (p[0][0] + p[-1][0]).upper()
    return p[0][:3].upper() if p and len(p[0]) >= 3 else (p[0][:2].upper() if p else "??")

def _norm_client(client: str) -> str:
    c = client.strip().lower()
    for k, v in _CLIENT_LOOKUP.items():
        if k in c:
            return v
    return client.strip()[:4].upper()

def _norm_status(status: str) -> str:
    return _STATUS_MAP.get(status.strip().lower(), "progress")

def _fmt_date(d: str) -> str:
    """'20/07/2025' → '20 Jul'. Returns '' for blank."""
    s = str(d).strip()
    if not s or s in ("nan", "None", "NaT"):
        return ""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            dt = datetime.strptime(s, fmt)
            return f"{dt.day} {dt.strftime('%b')}"
        except (ValueError, AttributeError):
            pass
    return s

def _days_until(end: str) -> int:
    s = str(end).strip()
    if not s or s in ("nan", "None", "NaT"):
        return 999
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return (datetime.strptime(s, fmt).date() - date.today()).days
        except (ValueError, AttributeError):
            pass
    return 999

def _date_month(d: str):
    """Return 0-indexed month (Jan=0) or None."""
    s = str(d).strip()
    if not s or s in ("nan", "None", "NaT"):
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).month - 1
        except (ValueError, AttributeError):
            pass
    return None

def _rel_time(ts: str) -> str:
    if not ts:
        return "recently"
    try:
        diff = datetime.now() - datetime.fromisoformat(ts)
        s = diff.total_seconds()
        if s < 60:   return "just now"
        if s < 3600: return f"{int(s/60)}m ago"
        if s < 86400:return f"{int(s/3600)}h ago"
        return f"{int(s/86400)}d ago"
    except Exception:
        return "recently"

def _load_projects() -> list:
    try:
        records = auth.get_all_projects()
        if records:
            return records
    except Exception:
        pass
    return []

def _build_emp_map(projects: list) -> dict:
    """name → {i: key, n: name, c: color}"""
    names: set = set()
    for p in projects:
        for part in str(p.get("employee", "")).split("&"):
            n = part.strip()
            if n:
                names.add(n)
        lead = str(p.get("lead", "")).strip()
        if lead:
            names.add(lead)
    try:
        for u in auth.get_all_users():
            names.add(u["name"])
    except Exception:
        pass

    # Deduplicate keys
    key_to_names: dict = {}
    result: dict = {}
    for nm in sorted(names):
        k = _key(nm)
        if k in key_to_names:
            k = nm[:3].upper()  # fallback: first 3 chars
        key_to_names[k] = nm
        result[nm] = k

    emp_map = {}
    for i, (nm, k) in enumerate(result.items()):
        emp_map[nm] = {"i": k, "n": nm, "c": _EMP_PALETTE[i % len(_EMP_PALETTE)]}
    return emp_map


# ── FastAPI ────────────────────────────────────────────────────────────────────
if _OK:
    api = FastAPI(title="Qualesce API", docs_url=None, redoc_url=None)
    api.add_middleware(
        CORSMiddleware,
        allow_origins=["*"], allow_methods=["*"], allow_headers=["*"],
    )

    class _ChatBody(BaseModel):
        message: str
        api_key: str
        history: list = []

    # ── /api/projects ─────────────────────────────────────────────────────────
    @api.get("/api/projects")
    def endpoint_projects():
        projs = _load_projects()
        emp_map = _build_emp_map(projs)
        rows = []
        for p in projs:
            active = str(p.get("is_active", "True")).strip().lower()
            if active in ("0", "false"):
                continue
            ck = _norm_client(str(p.get("client", "")))
            sk = _norm_status(str(p.get("status", "")))
            empl_first = str(p.get("employee", "")).split("&")[0].strip()
            lead_name = str(p.get("lead", "")).strip() or empl_first
            e_obj = emp_map.get(empl_first, {"i": _key(empl_first), "c": "#94A3B8"})
            l_obj = emp_map.get(lead_name,  {"i": _key(lead_name),  "c": "#94A3B8"})
            code = f"{ck[:3]}-{str(p.get('id', ''))}"
            end_raw = str(p.get("end", "")).strip()
            rows.append([
                code,
                str(p.get("name", ""))[:50],
                ck,
                l_obj["i"],
                e_obj["i"],
                sk,
                0,
                _fmt_date(str(p.get("start", ""))),
                _fmt_date(end_raw) if end_raw and end_raw not in ("nan","None","NaT") else None,
                0,
            ])

        employees = [{"i": v["i"], "n": v["n"], "c": v["c"]} for v in emp_map.values()]
        return {"rows": rows, "clients": CLIENT_INFO, "employees": employees}

    # ── /api/dashboard ────────────────────────────────────────────────────────
    @api.get("/api/dashboard")
    def endpoint_dashboard():
        projs = _load_projects()
        emp_map = _build_emp_map(projs)

        sc: dict = {}   # status → count
        cc: dict = {}   # client_key → {status_key: count}
        for p in projs:
            active = str(p.get("is_active", "True")).strip().lower()
            if active in ("0", "false"):
                continue
            s = str(p.get("status", "")).strip()
            ck = _norm_client(str(p.get("client", "")))
            sc[s] = sc.get(s, 0) + 1
            if ck not in cc: cc[ck] = {}
            sk = _norm_status(s)
            cc[ck][sk] = cc[ck].get(sk, 0) + 1

        total = sum(sc.values())

        # KPIs
        def _kpi(key, label, statuses, color, spark="0,12 6,11 12,10 18,9 24,8 30,6 36,5 44,4"):
            val = sum(sc.get(s, 0) for s in statuses)
            return {"key": key, "label": label, "value": val, "delta": "",
                    "deltaCls": "up" if val > 0 else "", "color": color, "sparkPts": spark}

        kpis = [
            {"key":"total","label":"Total","value":total,"delta":f"{total} projects",
             "deltaCls":"up","color":"#F0ABFC","sparkPts":"0,12 6,11 12,10 18,9 24,8 30,6 36,5 44,4"},
            _kpi("rm",       "R&M",      ["R&M"],                               "#60A5FA"),
            _kpi("uat",      "UAT",      ["UAT"],                               "#FBBF24"),
            _kpi("completed","Completed",["Completed"],                         "#6EE7B7"),
            _kpi("progress", "In prog",  ["In Progress"],                       "#67E8F9"),
            _kpi("poc",      "POC",      ["Internal POC","External POC"],        "#FB923C"),
            _kpi("disc",     "Disc.",    ["Discontinued"],                       "#FB7185"),
            _kpi("pdd",      "PDD",      ["PDD"],                               "#A78BFA"),
            _kpi("important","Flagged",  ["Important"],                          "#F472B6"),
            _kpi("new",      "Presales", ["Presales"],                           "#34D399"),
        ]

        # DONUT
        _color_for = {
            "In Progress":"#67E8F9","Completed":"#6EE7B7","R&M":"#60A5FA",
            "UAT":"#FBBF24","PDD":"#A78BFA","Discontinued":"#FB7185",
            "Internal POC":"#94A3B8","External POC":"#FB923C",
        }
        donut = [
            {"name": s, "n": n, "color": _color_for.get(s, "#94A3B8")}
            for s, n in sorted(sc.items(), key=lambda x: -x[1]) if n > 0
        ]

        # BARS
        bars = []
        for ck, sm in sorted(cc.items(), key=lambda x: -sum(x[1].values())):
            info = CLIENT_INFO.get(ck, {"name": ck, "color": "#94A3B8"})
            segs = [{"c": _STATUS_COLORS.get(sk, "#94A3B8"), "n": n}
                    for sk, n in sorted(sm.items(), key=lambda x: -x[1])]
            bars.append({"name": info["name"], "color": info["color"], "segs": segs})

        # ACTIVITY (from recent tasks or project list)
        activity = []
        try:
            tasks = auth.get_all_tasks()[:8]
            _amap = {"In Progress":("play","#67E8F9"),"Completed":("check","#6EE7B7"),
                     "Not Started":("task","#94A3B8"),"On Hold":("warn","#FBBF24")}
            for t in tasks:
                ico, col = _amap.get(t["status"], ("task","#94A3B8"))
                nm = t["assigned_to"] or "Team"
                parts = nm.split()
                short = f"{parts[0]} {parts[-1][0]}." if len(parts) > 1 else nm
                activity.append({
                    "type": t["status"].lower().replace(" ", "_"),
                    "icon": ico, "color": col, "who": short,
                    "msg": f"<b>{t['title']}</b> — <span class='tag'>{t['status']}</span>",
                    "time": _rel_time(t["updated_at"]),
                })
        except Exception:
            pass

        if not activity:
            act_colors = ["#67E8F9","#6EE7B7","#F0ABFC","#A78BFA","#FBBF24"]
            for i, p in enumerate(projs[:5]):
                nm = str(p.get("employee","Team")).split("&")[0].strip().split()[0]
                activity.append({
                    "type":"task","icon":"task","color":act_colors[i % len(act_colors)],
                    "who": nm,
                    "msg": f"<b>{p['name']}</b> · {p['status']}",
                    "time": "recently"
                })

        # TIMELINE
        today_m = date.today().month - 1
        timeline = []
        for p in projs:
            if len(timeline) >= 8: break
            active = str(p.get("is_active","True")).strip().lower()
            if active in ("0","false"): continue
            if str(p.get("status","")) in ("Discontinued","Completed"): continue
            ck = _norm_client(str(p.get("client","")))
            color = CLIENT_INFO.get(ck, {}).get("color", "#94A3B8")
            from_m = _date_month(str(p.get("start","")))
            to_m   = _date_month(str(p.get("end","")))
            if from_m is None: from_m = today_m
            if to_m   is None: to_m   = 11
            timeline.append({
                "name": str(p.get("name",""))[:32],
                "code": f"{ck[:3]}-{p.get('id','')}",
                "color": color,
                "from": max(0, min(from_m, 11)),
                "to":   max(0, min(to_m,   11)),
                "label": str(p.get("status","")),
            })

        # ROWS preview (dashboard slicer)
        rows_preview = []
        for p in projs[:14]:
            active = str(p.get("is_active","True")).strip().lower()
            if active in ("0","false"): continue
            ck = _norm_client(str(p.get("client","")))
            sk = _norm_status(str(p.get("status","")))
            empl = str(p.get("employee","")).split("&")[0].strip()
            e_obj = emp_map.get(empl, {"i": _key(empl), "c": "#94A3B8"})
            code  = f"{ck[:3]}-{p.get('id','')}"
            end_raw = str(p.get("end","")).strip()
            rows_preview.append([
                code, str(p.get("name",""))[:40], ck,
                e_obj["i"], e_obj["i"], sk, 0,
                _fmt_date(str(p.get("start",""))),
                _fmt_date(end_raw) if end_raw and end_raw not in ("nan","None","NaT") else None,
                0,
            ])

        employees = [{"i": v["i"], "n": v["n"], "c": v["c"]} for v in emp_map.values()]
        emp_names  = {v["i"]: v["n"] for v in emp_map.values()}
        emp_colors = {v["i"]: v["c"] for v in emp_map.values()}

        return {
            "kpis": kpis, "donut": donut, "bars": bars,
            "activity": activity, "timeline": timeline,
            "rows": rows_preview, "clients": CLIENT_INFO,
            "employees": employees,
            "emp_names": emp_names, "emp_colors": emp_colors,
            "stats": {"total": total, "status_counts": sc},
        }

    # ── /api/licenses ─────────────────────────────────────────────────────────
    @api.get("/api/licenses")
    def endpoint_licenses():
        lics = auth.get_all_licenses()
        return [
            [l["tool_name"], l["no_of_licenses"],
             l["start_date"], l["end_date"],
             _days_until(l["end_date"]), l["id"]]
            for l in lics
        ]

    # ── /api/users ────────────────────────────────────────────────────────────
    @api.get("/api/users")
    def endpoint_users():
        users = auth.get_all_users()
        rows, names, colors = [], {}, {}
        for i, u in enumerate(users):
            k = _key(u["name"])
            col = _EMP_PALETTE[i % len(_EMP_PALETTE)]
            names[k]  = u["name"]
            colors[k] = col
            rows.append([k, u["email"], u["role"], 0, u["is_active"], "—"])
        return {"rows": rows, "names": names, "colors": colors}

    # ── /api/tasks ────────────────────────────────────────────────────────────
    @api.get("/api/tasks")
    def endpoint_tasks():
        _sk = {"Not Started":"notstarted","In Progress":"progress",
               "Completed":"completed","On Hold":"hold"}
        tasks = auth.get_all_tasks()
        return [
            [t["title"], _key(t["assigned_to"] or "??"),
             (t.get("description","") or "")[:30] or "—",
             t["due_date"] or "—",
             t["progress"], _sk.get(t["status"], "notstarted"),
             "med", 0]
            for t in tasks
        ]

    # ── /api/chat ─────────────────────────────────────────────────────────────
    @api.post("/api/chat")
    def endpoint_chat(body: _ChatBody):
        if not body.api_key:
            return {"ok": False, "response": "No API key provided. Enter your Anthropic key in the AI Agent panel."}
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=body.api_key)
            system = (
                "You are an AI Project Manager Agent for Qualesce (RPA automation company). "
                "Portfolio: projects across Raychem, Swagelok-CA, Swagelok-AL, TEPL, Internal POC, External POC. "
                "Statuses: R&M, UAT, In Progress, Completed, PDD, Discontinued, Internal POC, External POC. "
                "Team: Akhila Kovuri, Avinash, Chethan B N, Faiyaz, Mathan, Nandukanth, Narendra, "
                "Nischal, Radhika, Sharan, Shiv Shankar, Sivin, Sushma, Vikas, Chandru S, Rubika AE.\n"
                "RULES: Use markdown tables for lists. Use bullet points for explanations. "
                "No long prose. Be concise and data-driven."
            )
            msgs = [{"role": m["role"], "content": m["content"]} for m in body.history[-10:]]
            msgs.append({"role": "user", "content": body.message})
            resp = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=1024,
                system=system,
                messages=msgs,
            )
            return {"ok": True, "response": resp.content[0].text}
        except Exception as e:
            return {"ok": False, "response": f"Error: {e}"}

    # ── background launcher ───────────────────────────────────────────────────
    _started = False
    _lock    = threading.Lock()

    def start_background(port: int = 8503, host: str = "127.0.0.1"):
        global _started
        with _lock:
            if _started:
                return
            _started = True
        t = threading.Thread(
            target=lambda: uvicorn.run(api, host=host, port=port, log_level="error"),
            daemon=True,
        )
        t.start()

else:
    def start_background(port=8503, host="127.0.0.1"):
        pass  # fastapi/uvicorn not installed
