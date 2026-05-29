"""
Run this once to push app.py to GitHub:
    python push_fix.py YOUR_GITHUB_TOKEN
"""
import sys, os, base64, requests

TOKEN  = sys.argv[1] if len(sys.argv) > 1 else ""
REPO   = "ChandruS7373/qdashboardworksoftrpasep"
BRANCH = "main"
FILE   = "app.py"
LOCAL  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")

if not TOKEN:
    print("Usage: python push_fix.py YOUR_GITHUB_TOKEN")
    sys.exit(1)

headers = {
    "Authorization": f"token {TOKEN}",
    "Accept": "application/vnd.github.v3+json",
}

with open(LOCAL, "rb") as f:
    content = base64.b64encode(f.read()).decode()

url  = f"https://api.github.com/repos/{REPO}/contents/{FILE}"
resp = requests.get(f"{url}?ref={BRANCH}", headers=headers, timeout=30)
sha  = resp.json().get("sha", "") if resp.status_code == 200 else ""

payload = {
    "message": "fix: cast bot metric cols to object dtype before .loc to fix Arrow TypeError",
    "content": content,
    "branch":  BRANCH,
    "sha":     sha,
}

put = requests.put(url, json=payload, headers=headers, timeout=60)
if put.status_code in (200, 201):
    commit_sha = put.json().get("commit", {}).get("sha", "")[:10]
    print(f"✅ Pushed successfully! Commit: {commit_sha}")
    print("Streamlit will redeploy in ~30 seconds.")
else:
    print(f"❌ Failed: {put.status_code} — {put.text[:300]}")
