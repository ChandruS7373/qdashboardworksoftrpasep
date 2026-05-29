"""
Qualesce AI Agent — One-Click Launcher
Starts Streamlit + opens a live public URL via ngrok automatically.

First run: paste your free ngrok authtoken when prompted (takes 30 seconds).
Every run after: just double-click the .bat file — no prompts, fully automatic.
"""

import subprocess
import sys
import os
import time
import webbrowser
import threading
import json

# ── PATHS ─────────────────────────────────────────────────────────────────────
BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
APP_FILE  = os.path.join(BASE_DIR, "app.py")
CFG_FILE  = os.path.join(BASE_DIR, ".ngrok_token")   # stored locally, gitignored
PORT      = 8502

# ─────────────────────────────────────────────────────────────────────────────

def get_or_ask_token():
    """Return saved token or ask user to paste one, then save it."""
    if os.path.exists(CFG_FILE):
        with open(CFG_FILE) as f:
            token = f.read().strip()
        if token:
            return token

    print("\n" + "=" * 55)
    print("  FIRST-TIME SETUP: ngrok authtoken needed (free)")
    print("=" * 55)
    print("  1. Open  https://dashboard.ngrok.com/signup")
    print("     (free account, no credit card)")
    print("  2. After signup go to:")
    print("     https://dashboard.ngrok.com/get-started/your-authtoken")
    print("  3. Copy your authtoken and paste it below.")
    print("=" * 55)
    token = input("\n  Paste authtoken here: ").strip()

    if not token:
        print("[WARNING] No token entered — public URL will not be available.")
        return ""

    with open(CFG_FILE, "w") as f:
        f.write(token)
    print("  Token saved. You won't be asked again.\n")
    return token


def check_dependencies():
    missing = []
    for pkg in ["streamlit", "anthropic", "pandas", "plotly", "openpyxl", "pyngrok"]:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"  Installing: {missing} ...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install"] + missing,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )


def start_streamlit():
    cmd = [
        sys.executable, "-m", "streamlit", "run", APP_FILE,
        f"--server.port={PORT}",
        "--browser.gatherUsageStats=false",
        "--server.headless=true",
    ]
    return subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)


def wait_for_app(timeout=30):
    import urllib.request
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(f"http://localhost:{PORT}", timeout=1)
            return True
        except Exception:
            time.sleep(1)
    return False


def start_tunnel(token):
    from pyngrok import ngrok, conf
    conf.get_default().auth_token = token
    tunnel = ngrok.connect(PORT, "http")
    return tunnel.public_url


def open_browser(url):
    def _open():
        time.sleep(1)
        webbrowser.open(url)
    threading.Thread(target=_open, daemon=True).start()


def main():
    os.chdir(BASE_DIR)

    print("\n" + "=" * 55)
    print("   Qualesce AI Project Manager")
    print("=" * 55)

    print("\n[1/4] Checking dependencies...")
    check_dependencies()

    token = get_or_ask_token()

    print("[2/4] Starting Streamlit app...")
    proc = start_streamlit()

    print(f"[3/4] Waiting for app on port {PORT}...")
    if not wait_for_app():
        print("[ERROR] Streamlit failed to start. See errors above.")
        proc.terminate()
        sys.exit(1)

    public_url = None
    if token:
        print("[4/4] Opening public tunnel via ngrok...")
        try:
            public_url = start_tunnel(token)
        except Exception as e:
            print(f"[WARNING] Tunnel failed: {e}")

    local_url = f"http://localhost:{PORT}"

    print("\n" + "=" * 55)
    print("   APP IS LIVE!")
    print("=" * 55)
    print(f"   Local  : {local_url}")
    if public_url:
        print(f"   Public : {public_url}")
        print(f"\n   Share the PUBLIC link with anyone worldwide.")
    else:
        print("   (No public URL — add ngrok token for public access)")
    print("=" * 55)
    print("\n   Browser opening automatically...")
    print("   Press Ctrl+C here to stop the app.\n")

    open_browser(public_url or local_url)

    try:
        proc.wait()
    except KeyboardInterrupt:
        print("\n  Shutting down...")
        try:
            from pyngrok import ngrok
            ngrok.kill()
        except Exception:
            pass
        proc.terminate()


if __name__ == "__main__":
    main()
