# Quant RI Throughput Dashboard — Setup Guide

## What this does
- Reads your SharePoint `.xlsx` transition sheets using your Office 365 login
- Shows how many tables each Quant RI marked "Completed" per day
- Auto-refreshes every 5 minutes
- First login is a one-time browser sign-in — token is cached after that

---

## Step 1 — Install Python dependencies

```bash
pip install flask flask-cors msal openpyxl requests
```

---

## Step 2 — Register an Azure AD App (one-time, ~5 minutes)

You need this so the app can access SharePoint files on your behalf.

1. Go to: https://portal.azure.com
2. Search for **"App registrations"** → click **New registration**
3. Fill in:
   - Name: `QRI Dashboard` (anything)
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: Platform = **Public client/native** → URI = `http://localhost:5050`
4. Click **Register**
5. Copy the **Application (client) ID** — you'll need this
6. On the left sidebar → **Overview** → copy the **Directory (tenant) ID**

### Grant permissions:
7. Left sidebar → **API permissions** → **Add a permission**
8. Choose **Microsoft Graph** → **Delegated permissions**
9. Search for and add: `Files.Read.All`
10. Click **Grant admin consent** (or ask your IT admin to do this)

---

## Step 3 — Configure server.py

Open `server.py` and fill in:

```python
CLIENT_ID = "paste-your-client-id-here"
TENANT_ID = "paste-your-tenant-id-here"
```

Optionally pre-load your SharePoint files:

```python
SHAREPOINT_FILES = [
    {
        "name": "Jeronimo Martins",
        "url": "https://yourcompany.sharepoint.com/sites/SITE/Shared%20Documents/transition_jeronimo.xlsx",
    },
    {
        "name": "Nestle Group",
        "url": "https://yourcompany.sharepoint.com/sites/SITE/Shared%20Documents/transition_nestle.xlsx",
    },
]
```

**How to get the SharePoint URL:**
- Go to SharePoint → open the Excel file in the browser
- Copy the URL from the address bar
- It should look like: `https://yourcompany.sharepoint.com/:x:/r/sites/...`

---

## Step 4 — Run the server

```bash
python server.py
```

The first time it runs:
- A message will appear in the terminal with a **device code** and a URL
- Go to https://microsoft.com/devicelogin in your browser
- Enter the code and sign in with your Office 365 account
- Done — the token is cached, you won't need to do this again

The dashboard opens automatically at **http://localhost:5050**

---

## Step 5 — Add more sheets (from the UI)

You can also add sheets from the dashboard itself:
- Click **+ Add sheet**
- Paste the SharePoint Excel file URL
- Click **Add sheet**

---

## Column name mapping

The server auto-detects these columns in your Excel file (case-insensitive):

| Dashboard field | Looks for column named... |
|---|---|
| Date | `date` |
| Quant RI | `quant_ri` or `quant ri` |
| Status | `status` |
| Company | `company_id` or `company` |
| Granularity | `granularity` |
| Period | `period` |
| Table ID | `table_id` or `table id` |

Only rows where **Status = "Completed"** (case-insensitive) are shown.

---

## Troubleshooting

**"Could not reach server"** — Make sure `python server.py` is running in a terminal.

**"HTTP 401 / 403"** — Your token expired or permissions weren't granted. Delete `token_cache.json` and restart the server to re-authenticate.

**"HTTP 404"** — The SharePoint URL is wrong. Make sure it points directly to the `.xlsx` file.

**Columns not found** — Check that your Excel sheet has the expected column headers in the first row. The server prints detected headers to the terminal when fetching.

**Admin consent required** — Ask your IT/Azure admin to grant consent for the `Files.Read.All` permission on the app registration.
