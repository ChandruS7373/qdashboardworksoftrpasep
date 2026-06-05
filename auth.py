import sqlite3
import hashlib
import os
import secrets
from datetime import datetime, date, timedelta

DB_PATH = os.environ.get(
    "QDASH_DB_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "qualesce.db"),
)

ROLES = ["admin", "lead", "manager", "employee", "sales", "project_lead"]
RPA_ROLES = [r for r in ROLES if r != "project_lead"]
TASK_STATUSES = ["Not Started", "In Progress", "Completed", "On Hold"]

_license_extras_ready = False
_sold_licenses_ready = False
_crm_tables_ready = False
_audit_tables_ready = False
_timesheets_ready = False
_worksoft_tables_ready = False
_worksoft_comments_ready = False


def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=30)
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def init_db():
    conn = get_conn()
    conn.execute("PRAGMA journal_mode=WAL")
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'employee',
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            description TEXT DEFAULT '',
            assigned_to_id INTEGER NOT NULL,
            assigned_by_id INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'Not Started',
            progress INTEGER NOT NULL DEFAULT 0,
            due_date TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (assigned_to_id) REFERENCES users(id),
            FOREIGN KEY (assigned_by_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS licenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tool_name TEXT NOT NULL,
            no_of_licenses INTEGER NOT NULL DEFAULT 1,
            start_date TEXT DEFAULT '',
            end_date TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
    """)
    conn.commit()
    try:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS task_comments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id INTEGER NOT NULL,
                user_id INTEGER NOT NULL,
                comment TEXT NOT NULL,
                week_start TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (task_id) REFERENCES tasks(id),
                FOREIGN KEY (user_id) REFERENCES users(id),
                UNIQUE (task_id, user_id, week_start)
            );
        """)
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE tasks ADD COLUMN start_date TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE tasks ADD COLUMN comment TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS email_settings (
                id INTEGER PRIMARY KEY DEFAULT 1,
                outlook_email TEXT NOT NULL DEFAULT '',
                outlook_password TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT ''
            );
            INSERT OR IGNORE INTO email_settings (id, outlook_email, outlook_password, updated_at)
            VALUES (1, '', '', '');
        """)
        conn.commit()
    except Exception:
        pass
    try:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY,
                name TEXT DEFAULT '',
                client TEXT DEFAULT '',
                lead TEXT DEFAULT '',
                employee TEXT DEFAULT '',
                status TEXT DEFAULT '',
                proj_type TEXT DEFAULT '',
                start TEXT DEFAULT '',
                end TEXT DEFAULT '',
                due_date TEXT DEFAULT '',
                po TEXT DEFAULT '',
                desc TEXT DEFAULT '',
                manual_hrs TEXT DEFAULT '',
                auto_hrs TEXT DEFAULT '',
                cost_per_hr TEXT DEFAULT '',
                hours_saved TEXT DEFAULT '',
                cost_saved TEXT DEFAULT '',
                roi_pct TEXT DEFAULT '',
                is_new INTEGER DEFAULT 0,
                is_active INTEGER DEFAULT 1,
                ckpt_pdd_sdd_start TEXT DEFAULT '',
                ckpt_pdd_sdd_end TEXT DEFAULT '',
                ckpt_development_start TEXT DEFAULT '',
                ckpt_development_end TEXT DEFAULT '',
                ckpt_uat_start TEXT DEFAULT '',
                ckpt_uat_end TEXT DEFAULT '',
                ckpt_deployment_start TEXT DEFAULT '',
                ckpt_deployment_end TEXT DEFAULT '',
                created_at TEXT DEFAULT '',
                updated_at TEXT DEFAULT ''
            );
        """)
        conn.commit()
    except Exception:
        pass
    try:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS project_types (
                id   INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                color TEXT NOT NULL DEFAULT '#94A3B8',
                created_at TEXT NOT NULL
            );
            INSERT OR IGNORE INTO project_types (name, color, created_at) VALUES ('RPA',      '#3F8E91', datetime('now'));
            INSERT OR IGNORE INTO project_types (name, color, created_at) VALUES ('AI Agent', '#7C3AED', datetime('now'));
            INSERT OR IGNORE INTO project_types (name, color, created_at) VALUES ('Presales', '#854D0E', datetime('now'));
            INSERT OR IGNORE INTO project_types (name, color, created_at) VALUES ('Worksoft', '#0EA5E9', datetime('now'));
        """)
        conn.commit()
    except Exception:
        pass
    # Ensure audit_log and project_comments tables exist at startup
    try:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                user_name TEXT NOT NULL DEFAULT '',
                action TEXT NOT NULL,
                table_name TEXT NOT NULL DEFAULT '',
                record_id TEXT DEFAULT '',
                description TEXT DEFAULT '',
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS project_comments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                project_name TEXT NOT NULL DEFAULT '',
                user_id INTEGER NOT NULL,
                user_name TEXT NOT NULL DEFAULT '',
                comment TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_audit_log_created  ON audit_log(created_at);
            CREATE INDEX IF NOT EXISTS idx_audit_log_table    ON audit_log(table_name);
            CREATE INDEX IF NOT EXISTS idx_proj_comments_pid  ON project_comments(project_id);
        """)
        conn.commit()
    except Exception:
        pass
    # Add indexes for frequently filtered columns (safe if they already exist)
    try:
        c.executescript("""
            CREATE INDEX IF NOT EXISTS idx_projects_status   ON projects(status);
            CREATE INDEX IF NOT EXISTS idx_projects_client   ON projects(client);
            CREATE INDEX IF NOT EXISTS idx_projects_employee ON projects(employee);
            CREATE INDEX IF NOT EXISTS idx_projects_active   ON projects(is_active);
            CREATE INDEX IF NOT EXISTS idx_tasks_assigned_to ON tasks(assigned_to_id);
            CREATE INDEX IF NOT EXISTS idx_tasks_status      ON tasks(status);
            CREATE INDEX IF NOT EXISTS idx_task_comments_tid ON task_comments(task_id);
            CREATE INDEX IF NOT EXISTS idx_task_comments_uid ON task_comments(user_id);
        """)
        conn.commit()
    except Exception:
        pass
    for _col_def in [
        "ALTER TABLE projects ADD COLUMN num_bots INTEGER DEFAULT 0",
        "ALTER TABLE projects ADD COLUMN manual_run_mins REAL DEFAULT 0",
        "ALTER TABLE projects ADD COLUMN bot_run_mins REAL DEFAULT 0",
        "ALTER TABLE projects ADD COLUMN monthly_runs INTEGER DEFAULT 0",
        "ALTER TABLE projects ADD COLUMN num_persons INTEGER DEFAULT 0",
    ]:
        try:
            c.execute(_col_def)
            conn.commit()
        except Exception:
            pass
    try:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS app_settings (
                id INTEGER PRIMARY KEY DEFAULT 1,
                anthropic_api_key TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT ''
            );
            INSERT OR IGNORE INTO app_settings (id, anthropic_api_key, updated_at)
            VALUES (1, '', '');
        """)
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE users ADD COLUMN department TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE tasks ADD COLUMN department TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE projects ADD COLUMN allocated_hours REAL DEFAULT 0")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE projects ADD COLUMN project_lead_email TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE projects ADD COLUMN run_interval_value REAL DEFAULT 0")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE projects ADD COLUMN run_interval_unit TEXT DEFAULT 'Minutes'")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE projects ADD COLUMN run_frequency TEXT DEFAULT 'Daily'")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("""
            CREATE TABLE IF NOT EXISTS bot_metric_logs (
                id                INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id        INTEGER NOT NULL,
                project_name      TEXT    DEFAULT '',
                log_date          TEXT    NOT NULL,
                qty               INTEGER DEFAULT 0,
                run_interval_mins REAL    DEFAULT 0,
                created_at        TEXT    DEFAULT '',
                UNIQUE(project_id, log_date)
            )
        """)
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE bot_metric_logs ADD COLUMN run_interval_mins REAL DEFAULT 0")
        conn.commit()
    except Exception:
        pass
    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        _seed_admin(c)
        conn.commit()
    conn.close()


_PROJECT_COLS = [
    "id","name","client","lead","employee","status","proj_type","start","end","due_date","po","desc",
    "manual_hrs","auto_hrs","cost_per_hr","hours_saved","cost_saved","roi_pct","is_new","is_active",
    "ckpt_pdd_sdd_start","ckpt_pdd_sdd_end","ckpt_development_start","ckpt_development_end",
    "ckpt_uat_start","ckpt_uat_end","ckpt_deployment_start","ckpt_deployment_end",
    "allocated_hours","project_lead_email",
    "num_bots","num_persons","manual_run_mins","bot_run_mins","monthly_runs",
    "run_interval_value","run_interval_unit","run_frequency",
]


def upsert_bot_metric_log(project_id: int, project_name: str, log_date: str, qty: int):
    conn = get_conn()
    c = conn.cursor()
    try:
        c.execute("""
            INSERT INTO bot_metric_logs (project_id, project_name, log_date, qty, created_at)
            VALUES (?, ?, ?, ?, datetime('now'))
            ON CONFLICT(project_id, log_date) DO UPDATE SET
                qty          = excluded.qty,
                project_name = excluded.project_name,
                created_at   = datetime('now')
        """, (int(project_id), str(project_name), str(log_date), int(qty)))
        conn.commit()
    except Exception:
        pass
    conn.close()


def get_bot_metric_logs(project_id=None, start_date=None, end_date=None) -> list:
    conn = get_conn()
    c = conn.cursor()
    q = "SELECT * FROM bot_metric_logs WHERE 1=1"
    params: list = []
    if project_id is not None:
        q += " AND project_id=?"
        params.append(int(project_id))
    if start_date:
        q += " AND log_date>=?"
        params.append(str(start_date))
    if end_date:
        q += " AND log_date<=?"
        params.append(str(end_date))
    q += " ORDER BY log_date DESC"
    c.execute(q, params)
    cols = [d[0] for d in c.description]
    rows = [dict(zip(cols, r)) for r in c.fetchall()]
    conn.close()
    return rows


def get_all_projects() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(f"SELECT {','.join(_PROJECT_COLS)} FROM projects ORDER BY id")
    rows = c.fetchall()
    conn.close()
    return [dict(zip(_PROJECT_COLS, r)) for r in rows]


def upsert_projects(records: list):
    """Persist the full project list to SQLite (upsert by id, delete removed rows)."""
    if not records:
        return
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    incoming_ids = []
    rows_to_save = []
    for r in records:
        try:
            pid = int(float(str(r.get("id", "") or 0)))
        except Exception:
            pid = 0
        if not pid:
            continue
        incoming_ids.append(pid)
        vals = [pid] + [str(r.get(col, "") or "") for col in _PROJECT_COLS[1:]]
        vals[_PROJECT_COLS.index("is_new")] = 1 if r.get("is_new") else 0
        vals[_PROJECT_COLS.index("is_active")] = 1 if r.get("is_active", True) else 0
        # preserve numeric type for run_interval_value so SQLite stores it as REAL
        try:
            vals[_PROJECT_COLS.index("run_interval_value")] = float(r.get("run_interval_value") or 0)
        except (ValueError, TypeError):
            vals[_PROJECT_COLS.index("run_interval_value")] = 0.0
        rows_to_save.append(vals)
    if not rows_to_save:
        conn.close()
        return
    placeholders = ",".join("?" * len(incoming_ids))
    # Preserve original created_at for rows that already exist
    c.execute(f"SELECT id, created_at FROM projects WHERE id IN ({placeholders})", incoming_ids)
    existing_created = {row[0]: row[1] for row in c.fetchall()}
    c.execute(f"DELETE FROM projects WHERE id NOT IN ({placeholders})", incoming_ids)
    full_cols = _PROJECT_COLS + ["created_at", "updated_at"]
    ph = ",".join(["?"] * len(full_cols))
    c.executemany(
        f"INSERT OR REPLACE INTO projects ({','.join(full_cols)}) VALUES ({ph})",
        [vals + [existing_created.get(vals[0], now), now] for vals in rows_to_save],
    )
    conn.commit()
    conn.close()


def _seed_admin(cur):
    cur.execute(
        "INSERT INTO users (name, email, password_hash, role, is_active, created_at) VALUES (?,?,?,?,?,?)",
        ("Admin", "admin@qualesce.com", _hash("Admin@123"), "admin", 1, _now()),
    )


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _hash(password: str) -> str:
    salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 200_000)
    return f"{salt}:{dk.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        salt, dk_hex = stored.split(":", 1)
        dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 200_000)
        return secrets.compare_digest(dk.hex(), dk_hex)
    except Exception:
        # Run full derivation so missing/malformed hash doesn't leak via timing
        hashlib.pbkdf2_hmac("sha256", password.encode(), b"timingguard", 200_000)
        return False


def authenticate(email: str, password: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, name, email, password_hash, role, is_active, department FROM users WHERE email=?",
        (email.strip().lower(),),
    )
    row = c.fetchone()
    conn.close()
    if row and row[5] == 1 and verify_password(password, row[3]):
        return {"id": row[0], "name": row[1], "email": row[2], "role": row[4], "department": row[6] or ""}
    return None


def get_user_by_email(email: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, name, email, role, is_active FROM users WHERE email=?",
        (email.strip().lower(),),
    )
    row = c.fetchone()
    conn.close()
    if row and row[4] == 1:
        return {"id": row[0], "name": row[1], "email": row[2], "role": row[3]}
    return None


# ── USER CRUD ──────────────────────────────────────────────────────────────────

def create_user(name: str, email: str, password: str, role: str, department: str = "") -> int:
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "INSERT INTO users (name, email, password_hash, role, is_active, created_at, department) VALUES (?,?,?,?,?,?,?)",
        (name.strip(), email.strip().lower(), _hash(password), role, 1, _now(), department.strip()),
    )
    conn.commit()
    uid = c.lastrowid
    conn.close()
    return uid


def get_all_users() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, name, email, role, is_active, created_at, department FROM users ORDER BY id")
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "name": r[1], "email": r[2], "role": r[3],
         "is_active": bool(r[4]), "created_at": r[5], "department": r[6] or ""}
        for r in rows
    ]


def get_employees() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, name, email FROM users WHERE role='employee' AND is_active=1 ORDER BY name"
    )
    rows = c.fetchall()
    conn.close()
    return [{"id": r[0], "name": r[1], "email": r[2]} for r in rows]


def get_leads() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, name, email FROM users WHERE role='lead' AND is_active=1 ORDER BY name"
    )
    rows = c.fetchall()
    conn.close()
    return [{"id": r[0], "name": r[1], "email": r[2]} for r in rows]


def get_employees_and_leads() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, name, email, role FROM users WHERE role IN ('employee','lead') AND is_active=1 ORDER BY role, name"
    )
    rows = c.fetchall()
    conn.close()
    return [{"id": r[0], "name": r[1], "email": r[2], "role": r[3]} for r in rows]


def update_user(user_id: int, name: str, email: str, role: str, department: str = ""):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE users SET name=?, email=?, role=?, department=? WHERE id=?",
        (name.strip(), email.strip().lower(), role, department.strip(), user_id),
    )
    conn.commit()
    conn.close()


def reset_password(user_id: int, new_password: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE users SET password_hash=? WHERE id=?", (_hash(new_password), user_id))
    conn.commit()
    conn.close()


# ── EMAIL SETTINGS ─────────────────────────────────────────────────────────────

def _ensure_email_settings(c):
    c.execute("""
        CREATE TABLE IF NOT EXISTS email_settings (
            id INTEGER PRIMARY KEY DEFAULT 1,
            outlook_email TEXT NOT NULL DEFAULT '',
            outlook_password TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL DEFAULT ''
        )
    """)
    c.execute("INSERT OR IGNORE INTO email_settings (id, outlook_email, outlook_password, updated_at) VALUES (1,'','','')")


def get_email_settings() -> dict:
    conn = get_conn()
    c = conn.cursor()
    _ensure_email_settings(c)
    conn.commit()
    c.execute("SELECT outlook_email, outlook_password, updated_at FROM email_settings WHERE id=1")
    row = c.fetchone()
    conn.close()
    if row:
        return {"outlook_email": row[0], "outlook_password": row[1], "updated_at": row[2]}
    return {"outlook_email": "", "outlook_password": "", "updated_at": ""}


def save_email_settings(outlook_email: str, outlook_password: str):
    conn = get_conn()
    c = conn.cursor()
    _ensure_email_settings(c)
    c.execute(
        "UPDATE email_settings SET outlook_email=?, outlook_password=?, updated_at=? WHERE id=1",
        (outlook_email.strip().lower(), outlook_password, _now()),
    )
    conn.commit()
    conn.close()


# ── PER-USER EMAIL SETTINGS ────────────────────────────────────────────────────

def _ensure_user_email_settings(c):
    c.execute("""
        CREATE TABLE IF NOT EXISTS user_email_settings (
            user_id INTEGER PRIMARY KEY,
            outlook_email TEXT NOT NULL DEFAULT '',
            outlook_password TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL DEFAULT ''
        )
    """)


def get_user_email_settings(user_id: int) -> dict:
    conn = get_conn()
    c = conn.cursor()
    _ensure_user_email_settings(c)
    conn.commit()
    c.execute("SELECT outlook_email, outlook_password, updated_at FROM user_email_settings WHERE user_id=?", (user_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return {"outlook_email": row[0], "outlook_password": row[1], "updated_at": row[2]}
    return {"outlook_email": "", "outlook_password": "", "updated_at": ""}


def save_user_email_settings(user_id: int, outlook_email: str, outlook_password: str):
    conn = get_conn()
    c = conn.cursor()
    _ensure_user_email_settings(c)
    c.execute(
        "INSERT OR REPLACE INTO user_email_settings (user_id, outlook_email, outlook_password, updated_at) VALUES (?,?,?,?)",
        (user_id, outlook_email.strip().lower(), outlook_password, _now()),
    )
    conn.commit()
    conn.close()


def set_active(user_id: int, active: bool):
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE users SET is_active=? WHERE id=?", (1 if active else 0, user_id))
    conn.commit()
    conn.close()


def delete_user(user_id: int):
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit()
    conn.close()


# ── TASK CRUD ──────────────────────────────────────────────────────────────────

_TASK_SQL = """
    SELECT t.id, t.title, t.description, t.status, t.progress,
           t.due_date, t.start_date, t.created_at, t.updated_at, t.comment,
           u1.name, u1.email, u2.name, u2.email, t.department
    FROM tasks t
    JOIN users u1 ON t.assigned_to_id = u1.id
    JOIN users u2 ON t.assigned_by_id = u2.id
"""


def _task(r) -> dict:
    return {
        "id": r[0], "title": r[1], "description": r[2],
        "status": r[3], "progress": r[4], "due_date": r[5],
        "start_date": r[6], "created_at": r[7], "updated_at": r[8],
        "comment": r[9], "assigned_to": r[10], "assigned_to_email": r[11],
        "assigned_by": r[12], "assigned_by_email": r[13], "department": r[14] or "",
    }


def create_task(title: str, description: str, assigned_to_id: int,
                assigned_by_id: int, due_date: str, start_date: str = "",
                department: str = "") -> int:
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO tasks (title, description, assigned_to_id, assigned_by_id, "
        "status, progress, start_date, due_date, created_at, updated_at, department) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
        (title.strip(), description.strip(), assigned_to_id, assigned_by_id,
         "Not Started", 0, start_date, due_date, now, now, department.strip()),
    )
    conn.commit()
    tid = c.lastrowid
    conn.close()
    return tid


def get_all_tasks() -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(_TASK_SQL + " ORDER BY t.id DESC")
    rows = c.fetchall()
    conn.close()
    return [_task(r) for r in rows]


def get_all_tasks_asc() -> list:
    return list(reversed(get_all_tasks()))



def sync_tasks_from_df(df):
    """Import tasks from a DataFrame into SQLite using batch upsert."""
    if df is None or (hasattr(df, 'empty') and df.empty):
        return
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, email FROM users")
    email_map = {row[1].strip().lower(): row[0] for row in c.fetchall()}
    now = _now()
    records = df.to_dict('records')
    parsed = []
    for row in records:
        try:
            task_id = int(float(str(row.get("id", 0) or 0)))
            title = str(row.get("title", "")).strip()
            if not task_id or not title:
                continue
            to_id = email_map.get(str(row.get("assigned_to_email", "")).strip().lower())
            by_id = email_map.get(str(row.get("assigned_by_email", "")).strip().lower())
            if not to_id or not by_id:
                continue
            parsed.append((
                task_id,
                title,
                str(row.get("description", "")).strip(),
                to_id, by_id,
                str(row.get("status", "Not Started")).strip(),
                int(float(str(row.get("progress", 0) or 0))),
                str(row.get("start_date", "")).strip(),
                str(row.get("due_date", "")).strip(),
                str(row.get("comment", "")).strip(),
                str(row.get("created_at", now)).strip() or now,
                str(row.get("updated_at", now)).strip() or now,
            ))
        except Exception:
            continue
    # Batch upsert: INSERT OR REPLACE preserves id, sets created_at from incoming data
    c.executemany(
        "INSERT OR REPLACE INTO tasks "
        "(id,title,description,assigned_to_id,assigned_by_id,status,progress,"
        "start_date,due_date,comment,created_at,updated_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
        parsed,
    )
    conn.commit()
    conn.close()


def sync_users_from_df(df):
    """Import users from a DataFrame into SQLite (batch, skips existing emails)."""
    if df is None or (hasattr(df, 'empty') and df.empty):
        return
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT email FROM users")
    existing_emails = {row[0].strip().lower() for row in c.fetchall()}
    now = _now()
    new_users = []
    for row in df.to_dict('records'):
        try:
            name = str(row.get("Name", row.get("name", ""))).strip()
            email = str(row.get("Email", row.get("email", ""))).strip().lower()
            role = str(row.get("Role", row.get("role", "employee"))).strip()
            password = str(row.get("Password", row.get("password", ""))).strip()
            active_str = str(row.get("Active", row.get("is_active", "Yes"))).strip().lower()
            is_active = 1 if active_str in ("yes", "1", "true") else 0
            if not name or not email or not password or email in existing_emails:
                continue
            if role not in ROLES:
                role = "employee"
            new_users.append((name, email, _hash(password), role, is_active, now))
            existing_emails.add(email)
        except Exception:
            continue
    if new_users:
        c.executemany(
            "INSERT INTO users (name, email, password_hash, role, is_active, created_at) VALUES (?,?,?,?,?,?)",
            new_users,
        )
    conn.commit()
    conn.close()


def sync_comments_from_df(df):
    """Bidirectional sync from a DataFrame: import missing, delete removed."""
    if df is None or (hasattr(df, 'empty') and df.empty):
        return
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, email FROM users")
    email_map = {row[1].strip().lower(): row[0] for row in c.fetchall()}
    c.execute("SELECT id FROM tasks")
    valid_task_ids = {row[0] for row in c.fetchall()}
    excel_keys = set()
    for _, row in df.iterrows():
        try:
            task_id = int(float(str(row.get("task_id", 0) or 0)))
            user_email = str(row.get("employee_email", "")).strip().lower()
            comment_text = str(row.get("comment", "")).strip()
            week_start = str(row.get("week_start", "")).strip()
            created_at = str(row.get("created_at", "")).strip() or _now()
            if not task_id or not user_email or not comment_text or not week_start:
                continue
            user_id = email_map.get(user_email)
            if not user_id or task_id not in valid_task_ids:
                continue
            excel_keys.add((task_id, user_id, week_start))
            c.execute(
                "INSERT OR IGNORE INTO task_comments (task_id, user_id, comment, week_start, created_at) VALUES (?,?,?,?,?)",
                (task_id, user_id, comment_text, week_start, created_at),
            )
        except Exception:
            continue
    c.execute("SELECT id, task_id, user_id, week_start FROM task_comments")
    to_delete = [row[0] for row in c.fetchall() if (row[1], row[2], row[3]) not in excel_keys]
    if to_delete:
        c.execute(f"DELETE FROM task_comments WHERE id IN ({','.join('?' * len(to_delete))})", to_delete)
    conn.commit()
    conn.close()


def get_user_tasks(user_id: int) -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(_TASK_SQL + " WHERE t.assigned_to_id=? ORDER BY t.id DESC", (user_id,))
    rows = c.fetchall()
    conn.close()
    return [_task(r) for r in rows]


def get_tasks_assigned_by(user_id: int) -> list:
    conn = get_conn()
    c = conn.cursor()
    c.execute(_TASK_SQL + " WHERE t.assigned_by_id=? ORDER BY t.id DESC", (user_id,))
    rows = c.fetchall()
    conn.close()
    return [_task(r) for r in rows]


def update_task_progress(task_id: int, progress: int, status: str, comment: str = ""):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE tasks SET progress=?, status=?, comment=?, updated_at=? WHERE id=?",
        (progress, status, comment, _now(), task_id),
    )
    conn.commit()
    conn.close()


def delete_task(task_id: int):
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM task_comments WHERE task_id=?", (task_id,))
    c.execute("DELETE FROM tasks WHERE id=?", (task_id,))
    conn.commit()
    conn.close()


def update_task_meta(task_id: int, title: str, description: str, start_date: str, due_date: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE tasks SET title=?, description=?, start_date=?, due_date=?, updated_at=? WHERE id=?",
        (title.strip(), description.strip(), start_date, due_date, _now(), task_id),
    )
    conn.commit()
    conn.close()


# ── LICENSE CRUD ───────────────────────────────────────────────────────────────

def _ensure_license_extras():
    global _license_extras_ready
    if _license_extras_ready:
        return
    conn = get_conn()
    c = conn.cursor()
    try:
        c.execute("ALTER TABLE licenses ADD COLUMN client_email TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE licenses ADD COLUMN license_plan TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("""
            CREATE TABLE IF NOT EXISTS license_notification_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                license_id INTEGER NOT NULL,
                license_type TEXT NOT NULL,
                threshold TEXT NOT NULL,
                sent_at TEXT NOT NULL,
                UNIQUE(license_id, license_type, threshold)
            )
        """)
        conn.commit()
    except Exception:
        pass
    conn.close()
    _license_extras_ready = True


def has_notification_been_sent(license_id: int, license_type: str, threshold: str) -> bool:
    _ensure_license_extras()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT 1 FROM license_notification_log WHERE license_id=? AND license_type=? AND threshold=?",
        (license_id, license_type, threshold)
    )
    found = c.fetchone() is not None
    conn.close()
    return found


def mark_notification_sent(license_id: int, license_type: str, threshold: str):
    _ensure_license_extras()
    conn = get_conn()
    c = conn.cursor()
    try:
        c.execute(
            "INSERT OR IGNORE INTO license_notification_log (license_id, license_type, threshold, sent_at) VALUES (?,?,?,?)",
            (license_id, license_type, threshold, _now())
        )
        conn.commit()
    except Exception:
        pass
    conn.close()


def create_license(tool_name: str, no_of_licenses: int, start_date: str, end_date: str,
                   client_email: str = "", license_plan: str = "") -> int:
    _ensure_license_extras()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO licenses (tool_name, no_of_licenses, start_date, end_date, client_email, license_plan, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?)",
        (tool_name.strip(), no_of_licenses, start_date.strip(), end_date.strip(),
         client_email.strip(), license_plan.strip(), now, now),
    )
    conn.commit()
    lid = c.lastrowid
    conn.close()
    return lid


def get_all_licenses() -> list:
    _ensure_license_extras()
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, tool_name, no_of_licenses, start_date, end_date, client_email, license_plan, created_at FROM licenses ORDER BY id DESC")
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "tool_name": r[1], "no_of_licenses": r[2],
         "start_date": r[3], "end_date": r[4],
         "client_email": r[5] or "", "license_plan": r[6] or "", "created_at": r[7]}
        for r in rows
    ]


def update_license(license_id: int, tool_name: str, no_of_licenses: int,
                   start_date: str, end_date: str, client_email: str = "",
                   license_plan: str = ""):
    _ensure_license_extras()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE licenses SET tool_name=?, no_of_licenses=?, start_date=?, end_date=?, client_email=?, license_plan=?, updated_at=? WHERE id=?",
        (tool_name.strip(), no_of_licenses, start_date.strip(), end_date.strip(),
         client_email.strip(), license_plan.strip(), _now(), license_id),
    )
    conn.commit()
    conn.close()


def delete_license(license_id: int):
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM licenses WHERE id=?", (license_id,))
    conn.commit()
    conn.close()


# ── SOLD LICENSE CRUD ──────────────────────────────────────────────────────────

def _ensure_sold_licenses_table():
    global _sold_licenses_ready
    if _sold_licenses_ready:
        return
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS sold_licenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tool_name TEXT NOT NULL,
            client TEXT NOT NULL,
            no_of_licenses INTEGER NOT NULL DEFAULT 1,
            start_date TEXT DEFAULT '',
            end_date TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            client_email TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)
    # Migration: add client_email if missing from older DB
    try:
        c.execute("ALTER TABLE sold_licenses ADD COLUMN client_email TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE sold_licenses ADD COLUMN license_plan TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    conn.commit()
    conn.close()
    _sold_licenses_ready = True


def create_sold_license(tool_name: str, client: str, no_of_licenses: int,
                        start_date: str, end_date: str, notes: str = "",
                        client_email: str = "", license_plan: str = "") -> int:
    _ensure_sold_licenses_table()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO sold_licenses (tool_name, client, no_of_licenses, start_date, end_date, notes, client_email, license_plan, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (tool_name.strip(), client.strip(), no_of_licenses,
         start_date.strip(), end_date.strip(), notes.strip(),
         client_email.strip(), license_plan.strip(), now, now),
    )
    conn.commit()
    lid = c.lastrowid
    conn.close()
    return lid


def get_all_sold_licenses() -> list:
    _ensure_sold_licenses_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, tool_name, client, no_of_licenses, start_date, end_date, notes, client_email, license_plan, created_at "
        "FROM sold_licenses ORDER BY id DESC"
    )
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "tool_name": r[1], "client": r[2], "no_of_licenses": r[3],
         "start_date": r[4], "end_date": r[5], "notes": r[6],
         "client_email": r[7] or "", "license_plan": r[8] or "", "created_at": r[9]}
        for r in rows
    ]


def update_sold_license(lid: int, tool_name: str, client: str, no_of_licenses: int,
                        start_date: str, end_date: str, notes: str = "",
                        client_email: str = "", license_plan: str = ""):
    _ensure_sold_licenses_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE sold_licenses SET tool_name=?, client=?, no_of_licenses=?, "
        "start_date=?, end_date=?, notes=?, client_email=?, license_plan=?, updated_at=? WHERE id=?",
        (tool_name.strip(), client.strip(), no_of_licenses,
         start_date.strip(), end_date.strip(), notes.strip(),
         client_email.strip(), license_plan.strip(), _now(), lid),
    )
    conn.commit()
    conn.close()


def delete_sold_license(lid: int):
    _ensure_sold_licenses_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM sold_licenses WHERE id=?", (lid,))
    conn.commit()
    conn.close()


def get_week_start(dt=None) -> str:
    """Returns the Monday of the week for dt (or today) as YYYY-MM-DD."""
    d = dt or date.today()
    return (d - timedelta(days=d.weekday())).isoformat()


# ── CRM TABLES ─────────────────────────────────────────────────────────────────

CRM_LEAD_STATUSES  = ["New", "Contacted", "Qualified", "Proposal", "Negotiation", "Won", "Lost"]
CRM_LEAD_SOURCES   = ["Direct", "Referral", "Website", "LinkedIn", "Email Campaign", "Cold Call", "Other"]
CRM_OPP_STAGES     = ["Prospecting", "Qualification", "Proposal", "Negotiation", "Closed Won", "Closed Lost"]
CRM_ACTIVITY_TYPES = ["Call", "Email", "Meeting", "Demo", "Follow-up", "Other"]


def _ensure_crm_tables():
    global _crm_tables_ready
    if _crm_tables_ready:
        return
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS crm_leads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT NOT NULL,
            contact_name TEXT DEFAULT '',
            email TEXT DEFAULT '',
            phone TEXT DEFAULT '',
            source TEXT DEFAULT '',
            status TEXT NOT NULL DEFAULT 'New',
            notes TEXT DEFAULT '',
            assigned_to_id INTEGER,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (assigned_to_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS crm_opportunities (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lead_id INTEGER,
            title TEXT NOT NULL,
            value REAL DEFAULT 0,
            stage TEXT NOT NULL DEFAULT 'Prospecting',
            probability INTEGER DEFAULT 0,
            expected_close TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            assigned_to_id INTEGER,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (lead_id) REFERENCES crm_leads(id),
            FOREIGN KEY (assigned_to_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS crm_activities (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lead_id INTEGER,
            opportunity_id INTEGER,
            type TEXT DEFAULT 'Call',
            subject TEXT NOT NULL,
            notes TEXT DEFAULT '',
            activity_date TEXT DEFAULT '',
            is_done INTEGER DEFAULT 0,
            created_by_id INTEGER,
            created_at TEXT NOT NULL,
            FOREIGN KEY (lead_id) REFERENCES crm_leads(id),
            FOREIGN KEY (opportunity_id) REFERENCES crm_opportunities(id),
            FOREIGN KEY (created_by_id) REFERENCES users(id)
        );
    """)
    conn.commit()
    conn.close()
    _crm_tables_ready = True


# ── CRM LEAD CRUD ──────────────────────────────────────────────────────────────

def create_lead(company_name: str, contact_name: str, email: str, phone: str,
                source: str, status: str, notes: str, assigned_to_id=None) -> int:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO crm_leads (company_name, contact_name, email, phone, source, status, notes, assigned_to_id, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (company_name.strip(), contact_name.strip(), email.strip().lower(), phone.strip(),
         source, status, notes.strip(), assigned_to_id, now, now),
    )
    conn.commit()
    lid = c.lastrowid
    conn.close()
    return lid


def get_all_leads() -> list:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        SELECT l.id, l.company_name, l.contact_name, l.email, l.phone,
               l.source, l.status, l.notes, l.assigned_to_id,
               u.name, l.created_at, l.updated_at
        FROM crm_leads l
        LEFT JOIN users u ON l.assigned_to_id = u.id
        ORDER BY l.id DESC
    """)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "company_name": r[1], "contact_name": r[2], "email": r[3],
         "phone": r[4], "source": r[5], "status": r[6], "notes": r[7],
         "assigned_to_id": r[8], "assigned_to": r[9] or "",
         "created_at": r[10], "updated_at": r[11]}
        for r in rows
    ]


def update_lead(lead_id: int, company_name: str, contact_name: str, email: str,
                phone: str, source: str, status: str, notes: str, assigned_to_id=None):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE crm_leads SET company_name=?, contact_name=?, email=?, phone=?, "
        "source=?, status=?, notes=?, assigned_to_id=?, updated_at=? WHERE id=?",
        (company_name.strip(), contact_name.strip(), email.strip().lower(), phone.strip(),
         source, status, notes.strip(), assigned_to_id, _now(), lead_id),
    )
    conn.commit()
    conn.close()


def delete_lead(lead_id: int):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM crm_activities WHERE lead_id=?", (lead_id,))
    c.execute("DELETE FROM crm_opportunities WHERE lead_id=?", (lead_id,))
    c.execute("DELETE FROM crm_leads WHERE id=?", (lead_id,))
    conn.commit()
    conn.close()


# ── CRM OPPORTUNITY CRUD ───────────────────────────────────────────────────────

def create_opportunity(lead_id, title: str, value: float, stage: str,
                       probability: int, expected_close: str, notes: str,
                       assigned_to_id=None) -> int:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO crm_opportunities (lead_id, title, value, stage, probability, expected_close, notes, assigned_to_id, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (lead_id, title.strip(), value, stage, probability, expected_close, notes.strip(), assigned_to_id, now, now),
    )
    conn.commit()
    oid = c.lastrowid
    conn.close()
    return oid


def get_all_opportunities() -> list:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        SELECT o.id, o.lead_id, l.company_name, o.title, o.value, o.stage,
               o.probability, o.expected_close, o.notes, o.assigned_to_id,
               u.name, o.created_at, o.updated_at
        FROM crm_opportunities o
        LEFT JOIN crm_leads l ON o.lead_id = l.id
        LEFT JOIN users u ON o.assigned_to_id = u.id
        ORDER BY o.id DESC
    """)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "lead_id": r[1], "company_name": r[2] or "", "title": r[3],
         "value": r[4] or 0, "stage": r[5], "probability": r[6] or 0,
         "expected_close": r[7] or "", "notes": r[8] or "",
         "assigned_to_id": r[9], "assigned_to": r[10] or "",
         "created_at": r[11], "updated_at": r[12]}
        for r in rows
    ]


def update_opportunity(opp_id: int, lead_id, title: str, value: float, stage: str,
                       probability: int, expected_close: str, notes: str, assigned_to_id=None):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE crm_opportunities SET lead_id=?, title=?, value=?, stage=?, probability=?, "
        "expected_close=?, notes=?, assigned_to_id=?, updated_at=? WHERE id=?",
        (lead_id, title.strip(), value, stage, probability, expected_close,
         notes.strip(), assigned_to_id, _now(), opp_id),
    )
    conn.commit()
    conn.close()


def delete_opportunity(opp_id: int):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM crm_activities WHERE opportunity_id=?", (opp_id,))
    c.execute("DELETE FROM crm_opportunities WHERE id=?", (opp_id,))
    conn.commit()
    conn.close()


# ── CRM ACTIVITY CRUD ──────────────────────────────────────────────────────────

def create_activity(lead_id, opportunity_id, act_type: str, subject: str,
                    notes: str, activity_date: str, created_by_id=None) -> int:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "INSERT INTO crm_activities (lead_id, opportunity_id, type, subject, notes, activity_date, is_done, created_by_id, created_at) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        (lead_id, opportunity_id, act_type, subject.strip(), notes.strip(),
         activity_date, 0, created_by_id, _now()),
    )
    conn.commit()
    aid = c.lastrowid
    conn.close()
    return aid


def get_all_activities() -> list:
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        SELECT a.id, a.lead_id, l.company_name, a.opportunity_id, o.title,
               a.type, a.subject, a.notes, a.activity_date, a.is_done,
               a.created_by_id, u.name, a.created_at
        FROM crm_activities a
        LEFT JOIN crm_leads l ON a.lead_id = l.id
        LEFT JOIN crm_opportunities o ON a.opportunity_id = o.id
        LEFT JOIN users u ON a.created_by_id = u.id
        ORDER BY a.activity_date DESC, a.id DESC
    """)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "lead_id": r[1], "company_name": r[2] or "", "opportunity_id": r[3],
         "opportunity_title": r[4] or "", "type": r[5], "subject": r[6],
         "notes": r[7] or "", "activity_date": r[8] or "", "is_done": bool(r[9]),
         "created_by_id": r[10], "created_by": r[11] or "", "created_at": r[12]}
        for r in rows
    ]


def update_activity_done(activity_id: int, is_done: bool):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE crm_activities SET is_done=? WHERE id=?", (1 if is_done else 0, activity_id))
    conn.commit()
    conn.close()


def delete_activity(activity_id: int):
    _ensure_crm_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM crm_activities WHERE id=?", (activity_id,))
    conn.commit()
    conn.close()


# ── TASK COMMENT CRUD ──────────────────────────────────────────────────────────

def add_task_comment(task_id: int, user_id: int, comment: str, week_start: str) -> bool:
    """Insert or update weekly comment (one per user per task per week). Returns False on error."""
    conn = get_conn()
    c = conn.cursor()
    try:
        c.execute(
            "INSERT OR REPLACE INTO task_comments (task_id, user_id, comment, week_start, created_at) "
            "VALUES (?,?,?,?,?)",
            (task_id, user_id, comment.strip(), week_start, _now()),
        )
        conn.commit()
        return True
    except Exception as exc:
        conn.rollback()
        raise RuntimeError(f"Failed to save task comment: {exc}") from exc
    finally:
        conn.close()


def get_user_week_comment(task_id: int, user_id: int, week_start: str):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, comment, created_at FROM task_comments WHERE task_id=? AND user_id=? AND week_start=?",
        (task_id, user_id, week_start),
    )
    row = c.fetchone()
    conn.close()
    return {"id": row[0], "comment": row[1], "created_at": row[2]} if row else None


def get_task_comments_with_users(task_id: int = None, from_date: str = None, to_date: str = None) -> list:
    conn = get_conn()
    c = conn.cursor()
    sql = """
        SELECT tc.id, tc.task_id, t.title, tc.week_start, tc.comment,
               tc.created_at, u.name, u.email
        FROM task_comments tc
        JOIN tasks t ON tc.task_id = t.id
        JOIN users u ON tc.user_id = u.id
        WHERE 1=1
    """
    params = []
    if task_id is not None:
        sql += " AND tc.task_id=?"
        params.append(task_id)
    if from_date:
        sql += " AND tc.week_start>=?"
        params.append(from_date)
    if to_date:
        sql += " AND tc.week_start<=?"
        params.append(to_date)
    sql += " ORDER BY tc.week_start DESC, tc.task_id, u.name"
    c.execute(sql, params)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "task_id": r[1], "task_title": r[2], "week_start": r[3],
         "comment": r[4], "created_at": r[5], "user_name": r[6], "user_email": r[7]}
        for r in rows
    ]


def get_all_comments_for_excel() -> list:
    """Return all task comments with task title, employee info, and timestamps for Excel export."""
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        SELECT tc.id, t.id, t.title, u.name, u.email,
               tc.comment, tc.week_start, tc.created_at
        FROM task_comments tc
        JOIN tasks t ON tc.task_id = t.id
        JOIN users u ON tc.user_id = u.id
        ORDER BY tc.created_at DESC
    """)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "task_id": r[1], "task_title": r[2],
         "employee_name": r[3], "employee_email": r[4],
         "comment": r[5], "week_start": r[6], "created_at": r[7]}
        for r in rows
    ]


# ── PROJECT TYPE CRUD ──────────────────────────────────────────────────────────

def get_project_types() -> list:
    """Return all project types ordered by name."""
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, name, color, created_at FROM project_types ORDER BY name")
    rows = c.fetchall()
    conn.close()
    return [{"id": r[0], "name": r[1], "color": r[2], "created_at": r[3]} for r in rows]


def add_project_type(name: str, color: str = "#94A3B8") -> bool:
    """Insert a new project type. Returns False if name already exists."""
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute(
            "INSERT INTO project_types (name, color, created_at) VALUES (?, ?, ?)",
            (name.strip(), color.strip(), _now()),
        )
        conn.commit()
        conn.close()
        return True
    except Exception:
        return False


def update_project_type(type_id: int, name: str, color: str) -> bool:
    """Rename or recolor an existing project type."""
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute(
            "UPDATE project_types SET name=?, color=? WHERE id=?",
            (name.strip(), color.strip(), type_id),
        )
        conn.commit()
        conn.close()
        return True
    except Exception:
        return False


def delete_project_type(type_id: int):
    """Delete a project type by id."""
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM project_types WHERE id=?", (type_id,))
    conn.commit()
    conn.close()


# ── AUDIT LOG & PROJECT COMMENTS ───────────────────────────────────────────────

def _ensure_audit_project_tables():
    global _audit_tables_ready
    if _audit_tables_ready:
        return
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            user_name TEXT NOT NULL DEFAULT '',
            action TEXT NOT NULL,
            table_name TEXT NOT NULL DEFAULT '',
            record_id TEXT DEFAULT '',
            description TEXT DEFAULT '',
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS project_comments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            project_name TEXT NOT NULL DEFAULT '',
            user_id INTEGER NOT NULL,
            user_name TEXT NOT NULL DEFAULT '',
            comment TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_audit_log_created   ON audit_log(created_at);
        CREATE INDEX IF NOT EXISTS idx_audit_log_table     ON audit_log(table_name);
        CREATE INDEX IF NOT EXISTS idx_proj_comments_pid   ON project_comments(project_id);
    """)
    conn.commit()
    conn.close()
    _audit_tables_ready = True


def log_audit(user_id: int, user_name: str, action: str, table_name: str = "",
              record_id: str = "", description: str = ""):
    """Insert a single audit log entry. Silently swallows errors to never break callers."""
    try:
        _ensure_audit_project_tables()
        conn = get_conn()
        c = conn.cursor()
        c.execute(
            "INSERT INTO audit_log (user_id, user_name, action, table_name, record_id, description, created_at) "
            "VALUES (?,?,?,?,?,?,?)",
            (user_id, user_name, action, table_name, str(record_id), description, _now()),
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def get_audit_logs(table_filter: str = "", action_filter: str = "",
                   user_filter: str = "") -> list:
    _ensure_audit_project_tables()
    conn = get_conn()
    c = conn.cursor()
    sql = "SELECT id, user_id, user_name, action, table_name, record_id, description, created_at FROM audit_log WHERE 1=1"
    params = []
    if table_filter:
        sql += " AND table_name=?"
        params.append(table_filter)
    if action_filter:
        sql += " AND action=?"
        params.append(action_filter)
    if user_filter:
        sql += " AND user_name LIKE ?"
        params.append(f"%{user_filter}%")
    sql += " ORDER BY id DESC"
    c.execute(sql, params)
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "user_id": r[1], "user_name": r[2], "action": r[3],
         "table_name": r[4], "record_id": r[5], "description": r[6], "created_at": r[7]}
        for r in rows
    ]


def add_project_comment(project_id: int, project_name: str, user_id: int,
                        user_name: str, comment: str) -> int:
    _ensure_audit_project_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "INSERT INTO project_comments (project_id, project_name, user_id, user_name, comment, created_at) "
        "VALUES (?,?,?,?,?,?)",
        (project_id, project_name.strip(), user_id, user_name.strip(), comment.strip(), _now()),
    )
    conn.commit()
    cid = c.lastrowid
    conn.close()
    return cid


def get_project_comments(project_id: int = None) -> list:
    _ensure_audit_project_tables()
    conn = get_conn()
    c = conn.cursor()
    if project_id is not None:
        c.execute(
            "SELECT id, project_id, project_name, user_id, user_name, comment, created_at "
            "FROM project_comments WHERE project_id=? ORDER BY id DESC",
            (project_id,),
        )
    else:
        c.execute(
            "SELECT id, project_id, project_name, user_id, user_name, comment, created_at "
            "FROM project_comments ORDER BY id DESC"
        )
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "project_id": r[1], "project_name": r[2], "user_id": r[3],
         "user_name": r[4], "comment": r[5], "created_at": r[6]}
        for r in rows
    ]


def delete_project_comment(comment_id: int):
    _ensure_audit_project_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM project_comments WHERE id=?", (comment_id,))
    conn.commit()
    conn.close()


# ── APP SETTINGS (Anthropic API key, etc.) ─────────────────────────────────────

def _ensure_app_settings(c):
    c.execute("""
        CREATE TABLE IF NOT EXISTS app_settings (
            id INTEGER PRIMARY KEY DEFAULT 1,
            anthropic_api_key TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL DEFAULT ''
        )
    """)
    c.execute(
        "INSERT OR IGNORE INTO app_settings (id, anthropic_api_key, updated_at) VALUES (1,'','')"
    )


def get_anthropic_api_key() -> str:
    try:
        conn = get_conn()
        c = conn.cursor()
        _ensure_app_settings(c)
        conn.commit()
        c.execute("SELECT anthropic_api_key FROM app_settings WHERE id=1")
        row = c.fetchone()
        conn.close()
        return row[0] if row else ""
    except Exception:
        return ""


def save_anthropic_api_key(api_key: str):
    conn = get_conn()
    c = conn.cursor()
    _ensure_app_settings(c)
    c.execute(
        "UPDATE app_settings SET anthropic_api_key=?, updated_at=? WHERE id=1",
        (api_key.strip(), _now()),
    )
    conn.commit()
    conn.close()


# ── TIMESHEETS ─────────────────────────────────────────────────────────────────

def _ensure_timesheets_table():
    global _timesheets_ready
    if _timesheets_ready:
        return
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS timesheets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            project_name TEXT NOT NULL DEFAULT '',
            employee_name TEXT NOT NULL DEFAULT '',
            work_date TEXT NOT NULL DEFAULT '',
            hours REAL NOT NULL DEFAULT 0,
            description TEXT DEFAULT '',
            created_by_id INTEGER,
            created_at TEXT NOT NULL
        )
    """)
    try:
        c.execute("CREATE INDEX IF NOT EXISTS idx_timesheets_pid ON timesheets(project_id)")
    except Exception:
        pass
    conn.commit()
    conn.close()
    _timesheets_ready = True


def create_timesheet_entry(project_id: int, project_name: str, employee_name: str,
                           work_date: str, hours: float, description: str = "",
                           created_by_id: int = None) -> int:
    _ensure_timesheets_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "INSERT INTO timesheets (project_id, project_name, employee_name, work_date, "
        "hours, description, created_by_id, created_at) VALUES (?,?,?,?,?,?,?,?)",
        (project_id, project_name.strip(), employee_name.strip(), work_date,
         hours, description.strip(), created_by_id, _now()),
    )
    conn.commit()
    eid = c.lastrowid
    conn.close()
    return eid


def get_project_timesheets(project_id: int = None) -> list:
    _ensure_timesheets_table()
    conn = get_conn()
    c = conn.cursor()
    if project_id is not None:
        c.execute(
            "SELECT id, project_id, project_name, employee_name, work_date, hours, description, created_at "
            "FROM timesheets WHERE project_id=? ORDER BY work_date DESC, id DESC",
            (project_id,),
        )
    else:
        c.execute(
            "SELECT id, project_id, project_name, employee_name, work_date, hours, description, created_at "
            "FROM timesheets ORDER BY work_date DESC, id DESC"
        )
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "project_id": r[1], "project_name": r[2], "employee_name": r[3],
         "work_date": r[4], "hours": r[5], "description": r[6] or "", "created_at": r[7]}
        for r in rows
    ]


def delete_timesheet_entry(entry_id: int):
    _ensure_timesheets_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM timesheets WHERE id=?", (entry_id,))
    conn.commit()
    conn.close()


# ── WORKSOFT PUNCH SYSTEM ──────────────────────────────────────────────────────

def _ensure_worksoft_tables():
    global _worksoft_tables_ready
    if _worksoft_tables_ready:
        return
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS worksoft_project_assignments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            user_name TEXT DEFAULT '',
            assigned_at TEXT NOT NULL,
            UNIQUE(project_id, user_id)
        );
        CREATE TABLE IF NOT EXISTS worksoft_punches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            project_name TEXT DEFAULT '',
            user_id INTEGER NOT NULL,
            user_name TEXT DEFAULT '',
            punch_in TEXT NOT NULL,
            punch_out TEXT DEFAULT '',
            hours_worked REAL DEFAULT 0,
            work_date TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS worksoft_hours_alert_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            threshold INTEGER NOT NULL,
            sent_at TEXT NOT NULL,
            UNIQUE(project_id, threshold)
        );
    """)
    try:
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_assign_pid ON worksoft_project_assignments(project_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_assign_uid ON worksoft_project_assignments(user_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_punch_pid  ON worksoft_punches(project_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_punch_uid  ON worksoft_punches(user_id)")
    except Exception:
        pass
    # Migrations
    try:
        c.execute("ALTER TABLE worksoft_project_assignments ADD COLUMN daily_hours REAL DEFAULT 0")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE worksoft_punches ADD COLUMN description TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass
    conn.commit()
    conn.close()
    _worksoft_tables_ready = True


def assign_worksoft_employees(project_id: int, user_ids: list, id_to_name: dict = None, daily_hours_map: dict = None):
    """Replace all employee assignments for a Worksoft project."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM worksoft_project_assignments WHERE project_id=?", (project_id,))
    now = _now()
    for uid in user_ids:
        uname = (id_to_name or {}).get(uid, "")
        dhrs = float((daily_hours_map or {}).get(uid, 0) or 0)
        try:
            c.execute(
                "INSERT INTO worksoft_project_assignments "
                "(project_id, user_id, user_name, assigned_at, daily_hours) VALUES (?,?,?,?,?)",
                (project_id, uid, uname, now, dhrs),
            )
        except Exception:
            pass
    conn.commit()
    conn.close()


def get_worksoft_project_assignments(project_id: int) -> list:
    """Get users assigned to a Worksoft project with their daily_hours allocation."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT a.user_id, u.name, u.email, COALESCE(a.daily_hours, 0) "
        "FROM worksoft_project_assignments a "
        "JOIN users u ON a.user_id = u.id "
        "WHERE a.project_id = ? ORDER BY u.name",
        (project_id,),
    )
    rows = c.fetchall()
    conn.close()
    return [{"user_id": r[0], "name": r[1], "email": r[2], "daily_hours": float(r[3] or 0)} for r in rows]


def get_user_worksoft_projects(user_id: int) -> list:
    """Get active Worksoft projects assigned to a user, including their daily_hours allocation."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT p.id, p.name, p.client, p.status, "
        "COALESCE(p.allocated_hours, 0), COALESCE(p.project_lead_email, ''), COALESCE(p.lead, ''), "
        "COALESCE(a.daily_hours, 0) "
        "FROM worksoft_project_assignments a "
        "JOIN projects p ON a.project_id = p.id "
        "WHERE a.user_id = ? AND p.is_active = 1 ORDER BY p.id",
        (user_id,),
    )
    rows = c.fetchall()
    conn.close()
    return [{"id": r[0], "name": r[1], "client": r[2], "status": r[3],
             "allocated_hours": float(r[4] or 0), "project_lead_email": r[5] or "",
             "lead": r[6] or "", "daily_hours": float(r[7] or 0)} for r in rows]


def worksoft_punch_in(project_id: int, project_name: str, user_id: int, user_name: str) -> int:
    """Record punch-in. Returns the new punch ID."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO worksoft_punches (project_id, project_name, user_id, user_name, "
        "punch_in, work_date, created_at) VALUES (?,?,?,?,?,?,?)",
        (project_id, project_name, user_id, user_name, now, now[:10], now),
    )
    punch_id = c.lastrowid
    conn.commit()
    conn.close()
    return punch_id


def worksoft_punch_out(project_id: int, user_id: int) -> float:
    """Close the active punch for this user/project. Returns hours worked."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, punch_in FROM worksoft_punches "
        "WHERE project_id=? AND user_id=? AND punch_out='' ORDER BY id DESC LIMIT 1",
        (project_id, user_id),
    )
    row = c.fetchone()
    if not row:
        conn.close()
        return 0.0
    punch_id, punch_in_str = row
    now = _now()
    try:
        delta = datetime.fromisoformat(now) - datetime.fromisoformat(punch_in_str)
        hours = delta.total_seconds() / 3600
    except Exception:
        hours = 0.0
    c.execute(
        "UPDATE worksoft_punches SET punch_out=?, hours_worked=? WHERE id=?",
        (now, round(hours, 4), punch_id),
    )
    conn.commit()
    conn.close()
    return hours


def get_project_total_hours(project_id: int) -> float:
    """Sum of all completed punch hours for a project across all employees."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT COALESCE(SUM(hours_worked), 0) FROM worksoft_punches "
        "WHERE project_id=? AND punch_out != ''",
        (project_id,),
    )
    row = c.fetchone()
    conn.close()
    return float(row[0] or 0) if row else 0.0


def check_and_log_hours_alert(project_id: int, total_hours: float, budget_hours: float) -> list:
    """Return list of newly-crossed thresholds (50 and/or 100) and mark them sent.
    Uses a unique log so each threshold fires only once per project."""
    if budget_hours <= 0:
        return []
    _ensure_worksoft_tables()
    pct = (total_hours / budget_hours) * 100
    newly_crossed = []
    conn = get_conn()
    c = conn.cursor()
    for threshold in (50, 100):
        if pct >= threshold:
            try:
                c.execute(
                    "INSERT INTO worksoft_hours_alert_log (project_id, threshold, sent_at) VALUES (?,?,?)",
                    (project_id, threshold, _now()),
                )
                conn.commit()
                newly_crossed.append(threshold)
            except Exception:
                pass  # UNIQUE constraint = already logged, skip
    conn.close()
    return newly_crossed


def get_worksoft_employee_hours(project_id: int) -> dict:
    """Return {user_id: total_completed_hours} for a single project."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT user_id, COALESCE(SUM(hours_worked), 0) "
        "FROM worksoft_punches WHERE project_id=? AND punch_out != '' GROUP BY user_id",
        (project_id,),
    )
    rows = c.fetchall()
    conn.close()
    return {int(r[0]): float(r[1] or 0) for r in rows}


def get_all_worksoft_total_hours() -> dict:
    """Return {project_id: total_completed_hours} for all projects in one query."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT project_id, COALESCE(SUM(hours_worked), 0) "
        "FROM worksoft_punches WHERE punch_out != '' GROUP BY project_id"
    )
    rows = c.fetchall()
    conn.close()
    return {int(r[0]): float(r[1] or 0) for r in rows}


def get_employee_punch_history(user_name: str) -> list:
    """Return daily hours per project for an employee across all projects."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT project_name, work_date, COALESCE(SUM(hours_worked), 0) "
        "FROM worksoft_punches "
        "WHERE user_name=? AND punch_out != '' "
        "GROUP BY project_name, work_date "
        "ORDER BY work_date ASC",
        (user_name,),
    )
    rows = c.fetchall()
    conn.close()
    return [{"project_name": r[0], "work_date": r[1], "hours_worked": float(r[2] or 0)} for r in rows]


def get_active_punch(project_id: int, user_id: int) -> dict:
    """Return the open punch record for this user/project, or None."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, punch_in FROM worksoft_punches "
        "WHERE project_id=? AND user_id=? AND punch_out='' ORDER BY id DESC LIMIT 1",
        (project_id, user_id),
    )
    row = c.fetchone()
    conn.close()
    return {"id": row[0], "punch_in": row[1]} if row else None


def get_project_punches(project_id: int) -> list:
    """All time entries for a project, newest first."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, user_id, user_name, punch_in, punch_out, hours_worked, work_date, "
        "COALESCE(description, '') "
        "FROM worksoft_punches WHERE project_id=? ORDER BY work_date DESC, id DESC",
        (project_id,),
    )
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "user_id": r[1], "user_name": r[2], "punch_in": r[3],
         "punch_out": r[4] or "", "hours_worked": r[5] or 0.0, "work_date": r[6],
         "description": r[7] or ""}
        for r in rows
    ]


def add_worksoft_manual_entry(project_id: int, project_name: str,
                              user_id: int, user_name: str,
                              work_date: str, hours: float,
                              description: str = "") -> int:
    """Log manual hours for a Worksoft project. Returns the new entry ID."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    now = _now()
    c.execute(
        "INSERT INTO worksoft_punches "
        "(project_id, project_name, user_id, user_name, punch_in, punch_out, "
        "hours_worked, work_date, description, created_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?)",
        (project_id, project_name, user_id, user_name,
         "manual", "manual", round(float(hours), 4),
         work_date, description.strip(), now),
    )
    entry_id = c.lastrowid
    conn.commit()
    conn.close()
    return entry_id


def update_worksoft_entry(entry_id: int, work_date: str, hours: float, description: str = ""):
    """Update date, hours and description of a manual time entry."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "UPDATE worksoft_punches SET work_date=?, hours_worked=?, description=? WHERE id=?",
        (work_date, round(float(hours), 4), description.strip(), entry_id),
    )
    conn.commit()
    conn.close()


def get_user_worksoft_entries(project_id: int, user_id: int) -> list:
    """Time entries logged by a specific user for a project, newest first."""
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, work_date, hours_worked, COALESCE(description, '') "
        "FROM worksoft_punches "
        "WHERE project_id=? AND user_id=? AND punch_out != '' "
        "ORDER BY work_date DESC, id DESC",
        (project_id, user_id),
    )
    rows = c.fetchall()
    conn.close()
    return [{"id": r[0], "work_date": r[1], "hours_worked": float(r[2] or 0), "description": r[3]}
            for r in rows]


def delete_worksoft_punch(punch_id: int):
    _ensure_worksoft_tables()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM worksoft_punches WHERE id=?", (punch_id,))
    conn.commit()
    conn.close()


# ── WORKSOFT PROJECT COMMENTS ──────────────────────────────────────────────────

def _ensure_worksoft_comments_table():
    global _worksoft_comments_ready
    if _worksoft_comments_ready:
        return
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS worksoft_comments (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id  INTEGER NOT NULL,
            project_name TEXT DEFAULT '',
            user_id     INTEGER NOT NULL,
            user_name   TEXT DEFAULT '',
            comment     TEXT NOT NULL,
            created_at  TEXT NOT NULL
        )
    """)
    try:
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_comments_pid ON worksoft_comments(project_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_comments_uid ON worksoft_comments(user_id)")
    except Exception:
        pass
    conn.commit()
    conn.close()
    _worksoft_comments_ready = True


def add_worksoft_comment(project_id: int, project_name: str,
                         user_id: int, user_name: str, comment: str) -> int:
    _ensure_worksoft_comments_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "INSERT INTO worksoft_comments (project_id, project_name, user_id, user_name, comment, created_at) "
        "VALUES (?,?,?,?,?,?)",
        (project_id, project_name.strip(), user_id, user_name.strip(),
         comment.strip(), _now()),
    )
    conn.commit()
    cid = c.lastrowid
    conn.close()
    return cid


def get_worksoft_project_comments(project_id: int, user_id: int = None) -> list:
    """Return comments for a project. If user_id given, returns only that user's comments."""
    _ensure_worksoft_comments_table()
    conn = get_conn()
    c = conn.cursor()
    if user_id is not None:
        c.execute(
            "SELECT id, user_id, user_name, comment, created_at "
            "FROM worksoft_comments WHERE project_id=? AND user_id=? ORDER BY created_at ASC",
            (project_id, user_id),
        )
    else:
        c.execute(
            "SELECT id, user_id, user_name, comment, created_at "
            "FROM worksoft_comments WHERE project_id=? ORDER BY created_at ASC",
            (project_id,),
        )
    rows = c.fetchall()
    conn.close()
    return [
        {"id": r[0], "user_id": r[1], "user_name": r[2],
         "comment": r[3], "created_at": r[4]}
        for r in rows
    ]


def delete_worksoft_comment(comment_id: int):
    _ensure_worksoft_comments_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM worksoft_comments WHERE id=?", (comment_id,))
    conn.commit()
    conn.close()


# ── IN-APP NOTIFICATIONS ───────────────────────────────────────────────────────

_notifs_ready = False


def _ensure_notifications_table():
    global _notifs_ready
    if _notifs_ready:
        return
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS notifications (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id    INTEGER NOT NULL,
            title      TEXT NOT NULL DEFAULT '',
            body       TEXT NOT NULL DEFAULT '',
            notif_type TEXT NOT NULL DEFAULT 'info',
            is_read    INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL
        )
    """)
    try:
        c.execute("CREATE INDEX IF NOT EXISTS idx_notifs_uid ON notifications(user_id)")
    except Exception:
        pass
    conn.commit()
    conn.close()
    _notifs_ready = True


def add_notification(user_id: int, title: str, body: str, notif_type: str = "info"):
    _ensure_notifications_table()
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute(
            "INSERT INTO notifications (user_id, title, body, notif_type, is_read, created_at) VALUES (?,?,?,?,0,?)",
            (user_id, title.strip(), body.strip(), notif_type, _now()),
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def get_notifications(user_id: int, unread_only: bool = False) -> list:
    _ensure_notifications_table()
    conn = get_conn()
    c = conn.cursor()
    sql = "SELECT id, title, body, notif_type, is_read, created_at FROM notifications WHERE user_id=?"
    if unread_only:
        sql += " AND is_read=0"
    sql += " ORDER BY id DESC LIMIT 50"
    c.execute(sql, (user_id,))
    rows = c.fetchall()
    conn.close()
    return [{"id": r[0], "title": r[1], "body": r[2], "type": r[3], "is_read": bool(r[4]), "created_at": r[5]} for r in rows]


def mark_notification_read(notif_id: int):
    _ensure_notifications_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE notifications SET is_read=1 WHERE id=?", (notif_id,))
    conn.commit()
    conn.close()


def mark_all_notifications_read(user_id: int):
    _ensure_notifications_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE notifications SET is_read=1 WHERE user_id=?", (user_id,))
    conn.commit()
    conn.close()


def count_unread_notifications(user_id: int) -> int:
    _ensure_notifications_table()
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM notifications WHERE user_id=? AND is_read=0", (user_id,))
    row = c.fetchone()
    conn.close()
    return int(row[0]) if row else 0


# ── SLACK / WEBHOOK SETTINGS ───────────────────────────────────────────────────

def get_slack_webhook_url() -> str:
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT value FROM app_settings_kv WHERE key='slack_webhook_url'")
        row = c.fetchone()
        conn.close()
        return row[0] if row else ""
    except Exception:
        return ""


def save_slack_webhook_url(url: str):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS app_settings_kv (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT ''
            )
        """)
        c.execute(
            "INSERT OR REPLACE INTO app_settings_kv (key, value, updated_at) VALUES ('slack_webhook_url',?,?)",
            (url.strip(), _now()),
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def send_slack_notification(webhook_url: str, text: str) -> bool:
    """Post a message to a Slack/Teams incoming webhook. Returns True on success."""
    if not webhook_url or not webhook_url.startswith("http"):
        return False
    try:
        import requests as _req
        resp = _req.post(webhook_url, json={"text": text}, timeout=5)
        return resp.status_code in (200, 204)
    except Exception:
        return False

