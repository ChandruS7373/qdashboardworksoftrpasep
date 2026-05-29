import smtplib
import random
import os
import queue
import re as _re
import threading
import time
from datetime import date, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

PORTAL_URL = "https://projecttracker.qualesce.com"

# Thread-safe queue for surfacing background email errors back to the UI
_error_queue: queue.Queue = queue.Queue()

# Secrets cached from the main Streamlit thread so background threads can read them
_secrets_cache: dict = {}


def prime_secrets_cache():
    """Call once from the main Streamlit thread to make secrets available to background threads."""
    global _secrets_cache
    try:
        import streamlit as st
        _secrets_cache = {
            "OUTLOOK_EMAIL":    st.secrets.get("OUTLOOK_EMAIL", ""),
            "OUTLOOK_PASSWORD": st.secrets.get("OUTLOOK_PASSWORD", ""),
        }
    except Exception:
        pass


def pop_email_errors() -> list[str]:
    """Drain and return any email errors queued by background threads."""
    errors = []
    while not _error_queue.empty():
        try:
            errors.append(_error_queue.get_nowait())
        except queue.Empty:
            break
    return errors


def _smtp_creds():
    # 1. Global DB settings (configured via admin Settings tab)
    try:
        import auth as _auth
        cfg = _auth.get_email_settings()
        if cfg["outlook_email"] and cfg["outlook_password"]:
            return cfg["outlook_email"], cfg["outlook_password"]
    except Exception:
        pass
    # 2. Cached secrets — safe to read from background threads
    e = _secrets_cache.get("OUTLOOK_EMAIL", "")
    p = _secrets_cache.get("OUTLOOK_PASSWORD", "")
    if e and p:
        return e, p
    # 3. Live secrets.toml (main thread only; silently skipped in background threads)
    try:
        import streamlit as st
        e = st.secrets.get("OUTLOOK_EMAIL", "")
        p = st.secrets.get("OUTLOOK_PASSWORD", "")
        if e and p:
            return e, p
    except Exception:
        pass
    # 4. Environment variables
    return os.environ.get("OUTLOOK_EMAIL", ""), os.environ.get("OUTLOOK_PASSWORD", "")


def _smtp_creds_for_role(role: str, user_id: int = 0) -> tuple[str, str]:
    """Return (email, password) for a specific role/user, with fallback chain:
    1. Per-user DB  →  2. Role secrets  →  3. Global DB  →  4. Global secrets"""
    # 1. Per-user DB settings
    try:
        import auth as _auth
        cfg = _auth.get_user_email_settings(user_id)
        if cfg["outlook_email"] and cfg["outlook_password"]:
            return cfg["outlook_email"], cfg["outlook_password"]
    except Exception:
        pass
    # 2. Role-specific secrets (LEAD_OUTLOOK_EMAIL / MANAGER_OUTLOOK_EMAIL)
    try:
        import streamlit as st
        key = role.upper()
        r_email = st.secrets.get(f"{key}_OUTLOOK_EMAIL", "")
        r_pwd   = st.secrets.get(f"{key}_OUTLOOK_PASSWORD", "")
        if r_email and r_pwd:
            return r_email, r_pwd
    except Exception:
        pass
    # 3. Role-specific env vars
    key = role.upper()
    r_email = os.environ.get(f"{key}_OUTLOOK_EMAIL", "")
    r_pwd   = os.environ.get(f"{key}_OUTLOOK_PASSWORD", "")
    if r_email and r_pwd:
        return r_email, r_pwd
    # 4. Fall back to global settings
    return _smtp_creds()


def _smtp_send(sender: str, password: str, to_email: str, subject: str, html_body: str,
               reply_to: str = "", display_name: str = "") -> tuple[bool, str]:
    msg = MIMEMultipart("alternative")
    msg["From"] = f"{display_name} <{sender}>" if display_name else sender
    msg["To"] = to_email
    msg["Subject"] = subject
    if reply_to:
        msg["Reply-To"] = reply_to
    msg.attach(MIMEText(html_body, "html"))
    try:
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.ehlo()
            server.starttls()
            server.login(sender, password)
            server.sendmail(sender, to_email, msg.as_string())
        return True, ""
    except Exception as e:
        return False, str(e)


def send_email(to_email: str, subject: str, html_body: str) -> tuple[bool, str]:
    sender, password = _smtp_creds()
    if not sender or not password:
        return False, "SMTP credentials not configured in secrets.toml"
    return _smtp_send(sender, password, to_email, subject, html_body)


def generate_otp() -> str:
    return str(random.randint(100000, 999999))


def send_otp_email(to_email: str, user_name: str, otp: str) -> tuple[bool, str]:
    subject = "Qualesce – Password Reset Code"
    body = f"""
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     max-width:480px;margin:0 auto;padding:32px;background:#F8FAFC;border-radius:12px">
  <div style="font-size:28px;font-weight:900;color:#3B82F6;letter-spacing:-1px;margin-bottom:4px">Q</div>
  <h2 style="color:#0F172A;margin:0 0 12px">Password Reset Request</h2>
  <p style="color:#475569;margin:0 0 20px">Hi <b>{user_name}</b>,<br>
  Use the code below to reset your Qualesce password. It expires in <b>10 minutes</b>.</p>
  <div style="font-size:40px;font-weight:900;letter-spacing:12px;color:#3B82F6;text-align:center;
       background:#EFF6FF;border:2px dashed #BFDBFE;border-radius:10px;padding:24px;margin:0 0 20px">
    {otp}
  </div>
  <p style="color:#94A3B8;font-size:12px;margin:0">
    If you did not request a password reset, you can safely ignore this email.
  </p>
  <p style="color:#94A3B8;font-size:12px;margin:8px 0 0">
    <a href="{PORTAL_URL}" style="color:#3B82F6">Back to Qualesce Portal</a>
  </p>
</div>
"""
    return send_email(to_email, subject, body)


def send_task_assigned_email(emp_email: str, emp_name: str, task_title: str,
                              assigned_by: str, due_date: str,
                              sender_email: str = "", sender_password: str = "") -> tuple[bool, str]:
    subject = f"Qualesce – New Task Assigned: {task_title}"
    due_line = f"Due: <b>{_fmt_email_date(due_date)}</b>" if due_date else "No due date set"
    body = f"""
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     max-width:480px;margin:0 auto;padding:32px;background:#F8FAFC;border-radius:12px">
  <div style="font-size:28px;font-weight:900;color:#3B82F6;letter-spacing:-1px;margin-bottom:4px">Q</div>
  <h2 style="color:#0F172A;margin:0 0 12px">New Task Assigned to You</h2>
  <p style="color:#475569;margin:0 0 16px">Hi <b>{emp_name}</b>, a new task has been assigned to you.</p>
  <div style="background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:18px 22px;margin:0 0 16px">
    <div style="font-size:17px;font-weight:700;color:#0F172A;margin-bottom:8px">{task_title}</div>
    <div style="color:#64748B;font-size:13px;margin-bottom:4px">Assigned by: <b>{assigned_by}</b></div>
    <div style="color:#64748B;font-size:13px">{due_line}</div>
  </div>
  <p style="color:#94A3B8;font-size:12px;margin:0">
    Log in to <a href="{PORTAL_URL}" style="color:#3B82F6;font-weight:600">Qualesce Portal</a> to view your task details and update your progress.
  </p>
</div>
"""
    if sender_email and sender_password:
        ok, err = _smtp_send(sender_email, sender_password, emp_email, subject, body)
        if ok:
            return ok, err
        # SMTP AUTH disabled for this account — fall back to global credentials
        # Reply-To is set to the intended sender so employee replies reach them directly
        global_sender, global_pwd = _smtp_creds()
        if global_sender and global_pwd and global_sender != sender_email:
            return _smtp_send(global_sender, global_pwd, emp_email, subject, body,
                              reply_to=sender_email,
                              display_name=f"{assigned_by}")
        return ok, err
    return send_email(emp_email, subject, body)


def send_license_expiry_email(client_email: str, client_name: str, tool_name: str,
                               end_date: str, days_left: int) -> tuple[bool, str]:
    """Send a license expiry notification email.
    days_left < 0 → expired; 0-30 → 30-day warning; 31-90 → 90-day warning; else → active notice.
    """
    end_date = _fmt_email_date(end_date)
    if days_left < 0:
        subject = f"Action Required – Your {tool_name} License Has Expired"
        badge_bg, badge_color, badge_text = "#FEF2F2", "#991B1B", "EXPIRED"
        heading = "Your License Has Expired"
        intro = (f"Hi <b>{client_name}</b>,<br><br>"
                 f"Your <b>{tool_name}</b> license expired on <b>{end_date}</b>. "
                 f"Please renew at your earliest convenience to avoid service disruption.")
        cta_label, cta_color = "Renew Now", "#DC2626"
    elif days_left <= 30:
        subject = f"Urgent – {tool_name} License Expiring in {days_left} Day{'s' if days_left != 1 else ''}"
        badge_bg, badge_color, badge_text = "#FEF2F2", "#DC2626", f"EXPIRING IN {days_left}d"
        heading = f"30-Day Alert: License Expires in {days_left} Day{'s' if days_left != 1 else ''}"
        intro = (f"Hi <b>{client_name}</b>,<br><br>"
                 f"<b>Urgent:</b> Your <b>{tool_name}</b> license is expiring on <b>{end_date}</b> — "
                 f"only <b>{days_left} day{'s' if days_left != 1 else ''}</b> remaining. "
                 f"Please renew immediately to ensure uninterrupted access.")
        cta_label, cta_color = "Renew Now", "#DC2626"
    elif days_left <= 60:
        subject = f"Reminder – {tool_name} License Expiring in {days_left} Days (60-Day Notice)"
        badge_bg, badge_color, badge_text = "#FFFBEB", "#B45309", f"EXPIRING IN {days_left}d"
        heading = f"60-Day Notice: License Expires in {days_left} Days"
        intro = (f"Hi <b>{client_name}</b>,<br><br>"
                 f"Your <b>{tool_name}</b> license will expire on <b>{end_date}</b> — "
                 f"<b>{days_left} days</b> from today. "
                 f"This is your 60-day advance notice. Please begin your renewal process.")
        cta_label, cta_color = "Start Renewal", "#D97706"
    else:
        subject = f"Reminder – {tool_name} License Expiring in {days_left} Days (90-Day Notice)"
        badge_bg, badge_color, badge_text = "#FEF3C7", "#92400E", f"EXPIRING IN {days_left}d"
        heading = f"90-Day Notice: License Expires in {days_left} Days"
        intro = (f"Hi <b>{client_name}</b>,<br><br>"
                 f"This is your 90-day advance notice that your <b>{tool_name}</b> license will expire on "
                 f"<b>{end_date}</b> (<b>{days_left} days</b> from today). "
                 f"Please plan your renewal well in advance.")
        cta_label, cta_color = "Plan Renewal", "#2563EB"

    body = f"""
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     max-width:520px;margin:0 auto;padding:32px;background:#F8FAFC;border-radius:12px">
  <div style="font-size:28px;font-weight:900;color:#3B82F6;letter-spacing:-1px;margin-bottom:4px">Q</div>
  <h2 style="color:#0F172A;margin:0 0 12px">{heading}</h2>
  <p style="color:#475569;margin:0 0 20px;line-height:1.6">{intro}</p>
  <div style="background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:18px 22px;margin:0 0 20px">
    <div style="font-size:16px;font-weight:700;color:#0F172A;margin-bottom:10px">{tool_name}</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;color:#64748B">
      <tr>
        <td style="padding:4px 0;width:40%">Expiry Date</td>
        <td style="padding:4px 0;font-weight:600;color:#0F172A">{end_date}</td>
      </tr>
      <tr>
        <td style="padding:4px 0">Status</td>
        <td style="padding:4px 0">
          <span style="background:{badge_bg};color:{badge_color};font-size:10px;font-weight:700;
                padding:2px 8px;border-radius:10px">{badge_text}</span>
        </td>
      </tr>
    </table>
  </div>
  <a href="{PORTAL_URL}" style="display:inline-block;background:{cta_color};color:#fff;
     text-decoration:none;font-weight:700;font-size:13px;padding:10px 24px;
     border-radius:6px;margin-bottom:24px">{cta_label}</a>
  <p style="color:#94A3B8;font-size:11px;margin:0;border-top:1px solid #E2E8F0;padding-top:16px">
    This is an automated notification from Qualesce AI Project Manager.
    Please contact your account manager if you need assistance.
  </p>
</div>
"""
    return send_email(client_email, subject, body)


def send_worksoft_hours_alert(to_email: str, lead_name: str, project_name: str,
                               total_hours: float, allocated_hours: float) -> tuple[bool, str]:
    """Alert the project lead that allocated hours have been reached on a Worksoft project."""
    subject = f"Qualesce – Hours Alert: {project_name}"
    greeting = f"Hi <b>{lead_name}</b>," if lead_name else "Hi,"
    body = f"""
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     max-width:480px;margin:0 auto;padding:32px;background:#F8FAFC;border-radius:12px">
  <div style="font-size:28px;font-weight:900;color:#3B82F6;letter-spacing:-1px;margin-bottom:4px">Q</div>
  <h2 style="color:#0F172A;margin:0 0 12px">Allocated Hours Reached</h2>
  <p style="color:#475569;margin:0 0 16px">{greeting}
  The team has reached the allocated hours for a Worksoft project you are leading.</p>
  <div style="background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:18px 22px;margin:0 0 16px">
    <div style="font-size:17px;font-weight:700;color:#0F172A;margin-bottom:10px">{project_name}</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;color:#64748B">
      <tr>
        <td style="padding:4px 0;width:50%">Allocated Hours</td>
        <td style="padding:4px 0;font-weight:600;color:#0F172A">{allocated_hours:.1f}h</td>
      </tr>
      <tr>
        <td style="padding:4px 0">Total Logged Hours</td>
        <td style="padding:4px 0;font-weight:700;color:#DC2626">{total_hours:.2f}h</td>
      </tr>
    </table>
  </div>
  <p style="color:#94A3B8;font-size:12px;margin:0">
    Log in to <a href="{PORTAL_URL}" style="color:#3B82F6;font-weight:600">Qualesce Portal</a>
    to review the project timeline and make adjustments as needed.
  </p>
</div>
"""
    return send_email(to_email, subject, body)


def send_worksoft_50pct_alert(to_email: str, lead_name: str, project_name: str,
                              total_hours: float, budget_hours: float) -> tuple:
    """Alert the project lead that 50% of the budget hours have been consumed."""
    subject = f"Qualesce – 50% Hours Alert: {project_name}"
    greeting = f"Hi <b>{lead_name}</b>," if lead_name else "Hi,"
    used_pct = round((total_hours / budget_hours) * 100, 1) if budget_hours > 0 else 0
    remaining = budget_hours - total_hours
    body = f"""
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     max-width:480px;margin:0 auto;padding:32px;background:#F8FAFC;border-radius:12px">
  <div style="font-size:28px;font-weight:900;color:#3B82F6;letter-spacing:-1px;margin-bottom:4px">Q</div>
  <h2 style="color:#0F172A;margin:0 0 12px">50% Hours Milestone Reached</h2>
  <p style="color:#475569;margin:0 0 16px">{greeting}
  Your Worksoft project has consumed <b>50%</b> of its allocated budget hours.</p>
  <div style="background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:18px 22px;margin:0 0 16px">
    <div style="font-size:17px;font-weight:700;color:#0F172A;margin-bottom:10px">{project_name}</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;color:#64748B">
      <tr>
        <td style="padding:4px 0;width:55%">Total Budget Hours</td>
        <td style="padding:4px 0;font-weight:600;color:#0F172A">{budget_hours:.1f}h</td>
      </tr>
      <tr>
        <td style="padding:4px 0">Hours Logged So Far</td>
        <td style="padding:4px 0;font-weight:700;color:#F59E0B">{total_hours:.2f}h ({used_pct}%)</td>
      </tr>
      <tr>
        <td style="padding:4px 0">Remaining Budget</td>
        <td style="padding:4px 0;font-weight:600;color:#16A34A">{remaining:.2f}h</td>
      </tr>
    </table>
    <div style="margin-top:12px;background:#FEF3C7;border-radius:6px;padding:8px 12px;
         font-size:12px;color:#92400E">
      ⚠️ Half the project budget has been used. Consider reviewing the timeline.
    </div>
  </div>
  <p style="color:#94A3B8;font-size:12px;margin:0">
    Log in to <a href="{PORTAL_URL}" style="color:#3B82F6;font-weight:600">Qualesce Portal</a>
    to review the project and adjust allocations as needed.
  </p>
</div>
"""
    return send_email(to_email, subject, body)


def send_worksoft_100pct_employee_alert(to_email: str, emp_name: str, project_name: str,
                                         total_hours: float, budget_hours: float) -> tuple:
    """Notify an assigned employee that 100% of project budget hours have been reached."""
    subject = f"Qualesce – Project Hours Limit Reached: {project_name}"
    greeting = f"Hi <b>{emp_name}</b>," if emp_name else "Hi,"
    body = f"""
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     max-width:480px;margin:0 auto;padding:32px;background:#F8FAFC;border-radius:12px">
  <div style="font-size:28px;font-weight:900;color:#3B82F6;letter-spacing:-1px;margin-bottom:4px">Q</div>
  <h2 style="color:#0F172A;margin:0 0 12px">Project Hours Budget Exhausted</h2>
  <p style="color:#475569;margin:0 0 16px">{greeting}
  The Worksoft project you are assigned to has reached <b>100%</b> of its allocated budget hours.</p>
  <div style="background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:18px 22px;margin:0 0 16px">
    <div style="font-size:17px;font-weight:700;color:#0F172A;margin-bottom:10px">{project_name}</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;color:#64748B">
      <tr>
        <td style="padding:4px 0;width:55%">Allocated Budget</td>
        <td style="padding:4px 0;font-weight:600;color:#0F172A">{budget_hours:.1f}h</td>
      </tr>
      <tr>
        <td style="padding:4px 0">Total Hours Logged</td>
        <td style="padding:4px 0;font-weight:700;color:#DC2626">{total_hours:.2f}h</td>
      </tr>
    </table>
    <div style="margin-top:12px;background:#FEE2E2;border-radius:6px;padding:8px 12px;
         font-size:12px;color:#991B1B">
      🔴 The project budget is fully consumed. Please contact your lead before logging further hours.
    </div>
  </div>
  <p style="color:#94A3B8;font-size:12px;margin:0">
    Log in to <a href="{PORTAL_URL}" style="color:#3B82F6;font-weight:600">Qualesce Portal</a>
    for details. Contact your project lead if you have questions.
  </p>
</div>
"""
    return send_email(to_email, subject, body)


def send_task_updated_email(assigner_email: str, assigner_name: str, emp_name: str,
                             task_title: str, new_status: str, progress: int,
                             comment: str) -> tuple[bool, str]:
    subject = f"Qualesce – Task Update: {task_title}"
    comment_block = (
        f'<div style="background:#F8FAFC;border-left:3px solid #CBD5E1;padding:10px 14px;'
        f'border-radius:0 6px 6px 0;color:#475569;font-size:13px;font-style:italic;margin-top:10px">'
        f'"{comment}"</div>'
    ) if comment.strip() else ""
    body = f"""
<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
     max-width:480px;margin:0 auto;padding:32px;background:#F8FAFC;border-radius:12px">
  <div style="font-size:28px;font-weight:900;color:#3B82F6;letter-spacing:-1px;margin-bottom:4px">Q</div>
  <h2 style="color:#0F172A;margin:0 0 12px">Task Updated</h2>
  <p style="color:#475569;margin:0 0 16px">Hi <b>{assigner_name}</b>,
  <b>{emp_name}</b> has updated a task you assigned.</p>
  <div style="background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:18px 22px;margin:0 0 16px">
    <div style="font-size:17px;font-weight:700;color:#0F172A;margin-bottom:8px">{task_title}</div>
    <div style="display:flex;gap:16px;flex-wrap:wrap">
      <span style="background:#DBEAFE;color:#1D4ED8;padding:3px 10px;border-radius:99px;
            font-size:12px;font-weight:600">{new_status}</span>
      <span style="background:#DCFCE7;color:#16A34A;padding:3px 10px;border-radius:99px;
            font-size:12px;font-weight:600">{progress}% complete</span>
    </div>
    {comment_block}
  </div>
  <p style="color:#94A3B8;font-size:12px;margin:0">
    Log in to <a href="{PORTAL_URL}" style="color:#3B82F6;font-weight:600">Qualesce Portal</a> to view the full update and task history.
  </p>
</div>
"""
    return send_email(assigner_email, subject, body)


# ── LICENSE EXPIRY AUTO-NOTIFIER ──────────────────────────────────────────────

_scheduler_started = False
_scheduler_lock = threading.Lock()


def _parse_date(date_str: str):
    if not date_str:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            pass
    return None


def _fmt_email_date(date_str: str) -> str:
    """Format a date string as DD-JAN-YYYY for use in emails."""
    d = _parse_date(str(date_str).strip())
    if d:
        return d.strftime("%d-%b-%Y").upper()
    return str(date_str).strip()


def run_license_notification_check():
    """Check all licenses and send expiry emails at 90-day, 30-day, and expired thresholds."""
    try:
        import auth as _auth
        today = date.today()

        sources = [
            (_auth.get_all_licenses(),      "purchased"),
            (_auth.get_all_sold_licenses(), "sold"),
        ]

        for licenses, lic_type in sources:
            for lic in licenses:
                client_email = lic.get("client_email", "").strip()
                if not client_email:
                    continue

                end_date = _parse_date(lic.get("end_date", ""))
                if not end_date:
                    continue

                days_left = (end_date - today).days

                if days_left <= 0:
                    threshold = "expired"
                elif days_left <= 30:
                    threshold = "30d"
                elif days_left <= 90:
                    threshold = "90d"
                else:
                    continue

                if _auth.has_notification_been_sent(lic["id"], lic_type, threshold):
                    continue

                addrs = [a.strip() for a in _re.split(r"[,\n;]+", client_email)
                         if a.strip() and "@" in a.strip()]
                if not addrs:
                    continue

                client_name = lic.get("client", lic.get("tool_name", "Client"))
                tool_name   = lic.get("tool_name", "License")
                end_date_str = lic.get("end_date", "")

                all_ok = True
                for addr in addrs:
                    ok, _ = send_license_expiry_email(addr, client_name, tool_name, end_date_str, days_left)
                    if not ok:
                        all_ok = False

                if all_ok:
                    _auth.mark_notification_sent(lic["id"], lic_type, threshold)

    except Exception:
        pass


def _scheduler_loop():
    while True:
        run_license_notification_check()
        time.sleep(24 * 60 * 60)


def start_license_notification_scheduler():
    """Start the background daily license-expiry notification thread (once per process)."""
    global _scheduler_started
    with _scheduler_lock:
        if _scheduler_started:
            return
        _scheduler_started = True
    t = threading.Thread(target=_scheduler_loop, daemon=True, name="license-notif-scheduler")
    t.start()
