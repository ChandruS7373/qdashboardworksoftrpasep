import os
import base64
import threading
import time
import requests

_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "qualesce.db")
_PUSH_INTERVAL = 30   # seconds between change-checks
_lock = threading.Lock()
_db_downloaded = False
_sync_started  = False


def _cfg():
    import streamlit as st
    token  = st.secrets.get("GITHUB_TOKEN", "")
    repo   = st.secrets.get("GITHUB_REPO",  "")   # "owner/repo-name"
    branch = st.secrets.get("GITHUB_BRANCH", "main")
    path   = st.secrets.get("GITHUB_DB_PATH", "qualesce.db")
    return token, repo, branch, path


def _headers(token):
    return {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}


def download_db():
    try:
        token, repo, branch, db_path = _cfg()
        if not token or not repo:
            return False
        url  = f"https://api.github.com/repos/{repo}/contents/{db_path}?ref={branch}"
        resp = requests.get(url, headers=_headers(token), timeout=30)
        if resp.status_code == 200:
            raw_url  = resp.json()["download_url"]
            raw_resp = requests.get(raw_url, headers=_headers(token), timeout=60)
            if raw_resp.status_code == 200:
                with _lock:
                    with open(_DB_PATH, "wb") as f:
                        f.write(raw_resp.content)
                print("[GH Sync] Database downloaded from GitHub.")
                return True
        elif resp.status_code == 404:
            print("[GH Sync] No db in repo yet — will create on first push.")
    except Exception as e:
        print(f"[GH Sync] Download error: {e}")
    return False


def push_db():
    try:
        token, repo, branch, db_path = _cfg()
        if not token or not repo:
            return False
        url = f"https://api.github.com/repos/{repo}/contents/{db_path}"
        with _lock:
            if not os.path.exists(_DB_PATH):
                return False
            with open(_DB_PATH, "rb") as f:
                content = base64.b64encode(f.read()).decode()
        # need current SHA to update existing file
        resp = requests.get(f"{url}?ref={branch}", headers=_headers(token), timeout=30)
        sha  = resp.json().get("sha", "") if resp.status_code == 200 else ""
        payload = {"message": "chore: sync db", "content": content, "branch": branch}
        if sha:
            payload["sha"] = sha
        put = requests.put(url, json=payload, headers=_headers(token), timeout=60)
        if put.status_code in (200, 201):
            print("[GH Sync] Database pushed to GitHub.")
            return True
        else:
            print(f"[GH Sync] Push failed: {put.status_code} {put.text[:200]}")
    except Exception as e:
        print(f"[GH Sync] Push error: {e}")
    return False


def _sync_loop():
    last_mtime = os.path.getmtime(_DB_PATH) if os.path.exists(_DB_PATH) else 0
    while True:
        time.sleep(_PUSH_INTERVAL)
        try:
            if os.path.exists(_DB_PATH):
                mtime = os.path.getmtime(_DB_PATH)
                if mtime != last_mtime:
                    push_db()
                    last_mtime = mtime
        except Exception as e:
            print(f"[GH Sync] Loop error: {e}")


def ensure_db_downloaded():
    global _db_downloaded
    if not _db_downloaded:
        _db_downloaded = True
        if not os.path.exists(_DB_PATH):
            download_db()


def start_sync_thread():
    global _sync_started
    if not _sync_started:
        _sync_started = True
        threading.Thread(target=_sync_loop, daemon=True).start()
        print("[GH Sync] Background sync thread started.")
