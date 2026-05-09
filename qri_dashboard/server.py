"""
Quant RI Throughput Dashboard — Local Server
=============================================
Fetches SharePoint Excel/CSV files using your Office 365 login,
parses completed rows, and serves the dashboard on http://localhost:5050

Requirements:
    pip install flask msal openpyxl requests flask-cors

Run:
    python server.py
"""

import json
import os
import threading
import webbrowser
from io import BytesIO

import msal
import openpyxl
import requests
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS

# ── Config ─────────────────────────────────────────────────────────────────
# Paste your SharePoint file URLs here (the direct link to the .xlsx file).
# Right-click the file in SharePoint → "Copy link" → paste below.
# You can also add/remove sheets from the dashboard UI at runtime.
SHAREPOINT_FILES = [
    # {
    #     "name": "Jeronimo Martins",
    #     "url": "https://yourcompany.sharepoint.com/sites/SITE/Shared Documents/transition_jeronimo.xlsx",
    # },
    # {
    #     "name": "Nestle Group",
    #     "url": "https://yourcompany.sharepoint.com/sites/SITE/Shared Documents/transition_nestle.xlsx",
    # },
]

# ── Azure App registration ─────────────────────────────────────────────────
# You need a registered Azure AD app with:
#   - "Files.Read.All" or "Sites.Read.All" API permission (delegated)
#   - Redirect URI: http://localhost:5050  (type: Public client / native)
#
# Steps:
#   1. Go to https://portal.azure.com → Azure Active Directory → App registrations → New
#   2. Name it anything, select "Accounts in this org only"
#   3. Add redirect URI: http://localhost:5050 (Mobile/Desktop platform)
#   4. Under API permissions → Add → Microsoft Graph → Delegated → Files.Read.All → Grant admin consent
#   5. Copy the Application (client) ID and your Tenant ID below

CLIENT_ID = "YOUR_CLIENT_ID_HERE"   # App registration client ID
TENANT_ID = "YOUR_TENANT_ID_HERE"   # Your Azure AD tenant ID

SCOPES = ["https://graph.microsoft.com/Files.Read.All",
          "https://graph.microsoft.com/Sites.Read.All"]

TOKEN_CACHE_FILE = "token_cache.json"

# ── MSAL auth ──────────────────────────────────────────────────────────────
def get_msal_app():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        cache.deserialize(open(TOKEN_CACHE_FILE).read())

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )
    return app, cache


def get_token():
    app, cache = get_msal_app()
    accounts = app.get_accounts()
    result = None

    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        # Device code flow — user logs in once via browser
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception("Failed to create device flow: " + str(flow))
        print("\n" + "="*60)
        print("SIGN IN REQUIRED")
        print("="*60)
        print(flow["message"])
        print("="*60 + "\n")
        result = app.acquire_token_by_device_flow(flow)

    # Save cache
    if cache.has_state_changed:
        open(TOKEN_CACHE_FILE, "w").write(cache.serialize())

    if "access_token" not in result:
        raise Exception("Auth failed: " + str(result.get("error_description")))

    return result["access_token"]


# ── SharePoint fetch ───────────────────────────────────────────────────────
def sharepoint_url_to_graph(url: str) -> str:
    """
    Convert a SharePoint file URL to a Microsoft Graph download URL.
    Works for both /sites/ and personal OneDrive URLs.
    """
    # Graph API: GET /v1.0/shares/{encoded-url}/driveItem/content
    import base64
    encoded = base64.b64encode(url.encode()).decode().rstrip("=").replace("/", "_").replace("+", "-")
    return f"https://graph.microsoft.com/v1.0/shares/u!{encoded}/driveItem/content"


def fetch_file_bytes(url: str, token: str) -> bytes:
    graph_url = sharepoint_url_to_graph(url)
    resp = requests.get(graph_url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    if resp.status_code == 200:
        return resp.content
    # Fallback: try direct download with auth header
    resp2 = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp2.raise_for_status()
    return resp2.content


# ── Excel / CSV parse ──────────────────────────────────────────────────────
def parse_excel(file_bytes: bytes, sheet_name: str) -> list:
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active  # first sheet; change if needed

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [str(h).strip().lower().replace(" ", "_") if h else "" for h in rows[0]]

    def col(*keys):
        for k in keys:
            for i, h in enumerate(headers):
                if h == k or k in h:
                    return i
        return -1

    i_date   = col("date")
    i_ri     = col("quant_ri", "quant ri")
    i_status = col("status")
    i_comp   = col("company_id", "company")
    i_gran   = col("granularity")
    i_period = col("period")
    i_table  = col("table_id", "table id")

    results = []
    for row in rows[1:]:
        def cell(i):
            return str(row[i]).strip() if i >= 0 and i < len(row) and row[i] is not None else ""

        if cell(i_status).lower() != "completed":
            continue

        # Normalise date
        raw_date = cell(i_date)
        date_str = ""
        if raw_date:
            # openpyxl may return a datetime object
            if hasattr(row[i_date], "strftime"):
                date_str = row[i_date].strftime("%Y-%m-%d")
            else:
                # Try to parse common formats
                for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
                    try:
                        from datetime import datetime
                        date_str = datetime.strptime(raw_date, fmt).strftime("%Y-%m-%d")
                        break
                    except ValueError:
                        pass
                if not date_str:
                    date_str = raw_date

        results.append({
            "date":    date_str,
            "ri":      cell(i_ri).lower(),
            "company": cell(i_comp),
            "gran":    cell(i_gran),
            "period":  cell(i_period),
            "table":   cell(i_table),
            "sheet":   sheet_name,
        })

    return results


# ── Flask app ──────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder="static")
CORS(app)

# In-memory store
_sheets_config = list(SHAREPOINT_FILES)   # [{name, url}]
_all_rows      = []
_token_lock    = threading.Lock()
_cached_token  = {"value": None}


def get_cached_token():
    with _token_lock:
        if not _cached_token["value"]:
            _cached_token["value"] = get_token()
        return _cached_token["value"]


def refresh_token_if_needed():
    """Force re-auth on next call if token expired."""
    _cached_token["value"] = None


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/sheets", methods=["GET"])
def list_sheets():
    return jsonify([{"name": s["name"], "url": s["url"]} for s in _sheets_config])


@app.route("/api/sheets", methods=["POST"])
def add_sheet():
    data = request.json
    name = (data.get("name") or "").strip()
    url  = (data.get("url")  or "").strip()
    if not name or not url:
        return jsonify({"error": "name and url required"}), 400
    _sheets_config.append({"name": name, "url": url})
    save_config()
    return jsonify({"ok": True})


@app.route("/api/sheets/<path:name>", methods=["DELETE"])
def remove_sheet(name):
    global _all_rows
    before = len(_sheets_config)
    _sheets_config[:] = [s for s in _sheets_config if s["name"] != name]
    _all_rows = [r for r in _all_rows if r["sheet"] != name]
    save_config()
    return jsonify({"removed": before - len(_sheets_config)})


@app.route("/api/refresh", methods=["POST"])
def refresh_all():
    global _all_rows
    token = get_cached_token()
    errors = []
    new_rows = [r for r in _all_rows if not any(r["sheet"] == s["name"] for s in _sheets_config)]

    for sheet in _sheets_config:
        try:
            file_bytes = fetch_file_bytes(sheet["url"], token)
            rows = parse_excel(file_bytes, sheet["name"])
            new_rows.extend(rows)
        except Exception as e:
            errors.append({"sheet": sheet["name"], "error": str(e)})

    _all_rows = new_rows
    return jsonify({"rows": len(_all_rows), "errors": errors})


@app.route("/api/data", methods=["GET"])
def get_data():
    return jsonify(_all_rows)


def save_config():
    with open("sheets_config.json", "w") as f:
        json.dump(_sheets_config, f, indent=2)


def load_config():
    global _sheets_config
    if os.path.exists("sheets_config.json"):
        try:
            saved = json.load(open("sheets_config.json"))
            # Merge with hardcoded list
            existing_urls = {s["url"] for s in _sheets_config}
            for s in saved:
                if s["url"] not in existing_urls:
                    _sheets_config.append(s)
        except Exception:
            pass


# ── Main ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    load_config()
    print("\n" + "="*60)
    print("Quant RI Throughput Dashboard")
    print("="*60)

    if CLIENT_ID == "YOUR_CLIENT_ID_HERE":
        print("\n⚠️  ACTION REQUIRED: Open server.py and fill in:")
        print("   CLIENT_ID  — from your Azure App registration")
        print("   TENANT_ID  — your company's Azure AD tenant ID")
        print("\nSee README.md for step-by-step setup instructions.\n")
    else:
        print("\nAuthenticating with Microsoft 365...")
        try:
            get_cached_token()
            print("✓ Authenticated successfully\n")
        except Exception as e:
            print(f"✗ Auth error: {e}\n")

    print("Starting server at http://localhost:5050")
    print("Opening browser...\n")
    threading.Timer(1.5, lambda: webbrowser.open("http://localhost:5050")).start()
    app.run(port=5050, debug=False)
