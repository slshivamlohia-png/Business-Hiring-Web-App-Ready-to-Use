# 🚀 Hiring Dashboard

A full-featured, **AI-powered hiring management web app** built entirely on **Google Apps Script** — no servers, no hosting costs, no external databases. Just a Google Sheet, a Script, and your team.

![Platform](https://img.shields.io/badge/platform-Google%20Apps%20Script-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![AI](https://img.shields.io/badge/AI-Claude%20(Anthropic)-purple)

---

## ✨ What It Does

| Feature | Details |
|---|---|
| **CV Upload + AI Extraction** | Upload PDF/DOC/DOCX — Claude AI auto-extracts name, phone, email, skills, experience, company, education, and a 2-3 line summary |
| **Duplicate Detection** | Blocks duplicate candidates by phone or email before they're saved |
| **Candidate Pipeline** | Track every candidate through statuses: New → Shortlisted → Called → Interview Scheduled → Offered |
| **Call Detail Fields** | Configurable per-candidate fields (Yes/No, dropdowns, text) that auto-save |
| **Workforce Tracker** | Transfer hired candidates to Workforce with Training / Probation / Intern / Full-time stages |
| **Remarks System** | Timestamped, multi-remark per candidate with inline edit + delete |
| **Reports** | CVs added, status changes, cost-per-CV, spend-by-platform — all filterable by date range |
| **Expense Tracker** | Log hiring spend by platform (LinkedIn, Naukri, etc.) |
| **Activity Log** | Full audit trail with date + user + search filtering |
| **Role-based Access** | Password-protected team login; public candidate self-submit form |
| **Dark / Light Mode** | System-aware theme toggle |
| **Mobile Responsive** | Fully usable on phones |

---

## 🗂️ File Structure

```
neodrift-hiring-dashboard/
├── Code.gs       ← Google Apps Script backend (all server-side logic)
├── Index.html    ← Frontend UI (single-file, no build step)
├── README.md     ← This file
└── LICENSE       ← MIT License
```

---

## ⚙️ Setup Guide

### Prerequisites

- A Google account
- An [Anthropic API key](https://console.anthropic.com/) (for AI CV extraction)

---

### Step 1 — Create a Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) → **Blank spreadsheet**
2. Name it anything — e.g. `Neodrift Hiring Data`
3. Keep this tab open. You'll need the URL later.

---

### Step 2 — Open Apps Script

1. In your Google Sheet, click **Extensions → Apps Script**
2. This opens the Apps Script editor
3. Delete all existing code in the default `Code.gs` file

---

### Step 3 — Add the Code

1. Copy the entire contents of **`Code.gs`** from this repo
2. Paste it into the Apps Script editor (replacing everything)
3. Click **File → New → HTML file** → name it exactly **`Index`** (no `.html` extension needed — Apps Script adds it)
4. Copy the entire contents of **`Index.html`** from this repo
5. Paste it into the `Index.html` file in the editor
6. Click **💾 Save** (Ctrl+S / Cmd+S)

---

### Step 4 — Enable Drive API (Advanced Service)

The app uses Google Drive to store CVs. You need to enable it:

1. In the Apps Script editor, click **Services** (the `+` icon in the left sidebar)
2. Scroll to **Drive API** → click **Add**
3. Leave version as `v2` → click **Add**

---

### Step 5 — Set Your Anthropic API Key

1. In Apps Script editor, click **Project Settings** (⚙️ gear icon, left sidebar)
2. Scroll to **Script Properties** → click **Add script property**
3. Set:
   - **Property:** `ANTHROPIC_API_KEY`
   - **Value:** `sk-ant-...` (your actual key)
4. Click **Save script properties**

---

### Step 6 — Initialize the Sheet

1. In the Apps Script editor, select the function `initializeSheet` from the function dropdown (top bar)
2. Click **▶ Run**
3. Accept any permissions Google asks for (Drive access, Sheets access, external URLs)
4. You should see `Initialized!` in the execution log

> This creates all the required sheets: Candidates, Positions, Activity Log, Expenses, Settings, Call Data, Workforce.

---

### Step 7 — Deploy as Web App

1. Click **Deploy → New deployment**
2. Click the **⚙️ gear** next to "Select type" → choose **Web app**
3. Configure:
   - **Description:** `Neodrift Hiring Dashboard v1`
   - **Execute as:** `Me`
   - **Who has access:** `Anyone` (so candidates can self-submit, and your team can login)
4. Click **Deploy**
5. Accept permissions
6. Copy the **Web App URL** — this is your dashboard link! 🎉

> **To update after code changes:** Go to Deploy → Manage deployments → Edit → change version to "New version" → Deploy.

---

### Step 8 — Set Your Password

The default password is a placeholder. Before deploying, open `Code.gs` and replace `YOUR_PASSWORD_HERE` with your actual password:

```js
var DASHBOARD_PASSWORD='YOUR_PASSWORD_HERE';
```

> **Tip:** For better security, store the password in Script Properties (same way as the API key) and use `PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD')` instead.

---

## 🖥️ Dashboard Usage Guide

### Logging In

When you open the Web App URL, you'll see two options:

- **I'm a Candidate** → Public CV submission form (no password needed)
- **Team Login** → Password-protected dashboard for your hiring team

Enter your name + the team password to access the full dashboard.

---

### 📊 Dashboard (Home)

The home screen shows:
- **Live stats** — Total candidates, added today, added this week, count by status
- **Pipeline by position** — Cards for each job role with per-status breakdowns (click any card/number to jump to that filtered view)
- **Workforce summary** — Quick count of who's in Training, Probation, Hired, etc.
- **Expenses this month** — Total spend + breakdown by platform
- **Recent activity** — Last 8 actions across the team

---

### 👤 Candidates

The main candidate list with filters:

**Quick Filters (pill buttons):**
- Click any **position pill** or **status pill** at the top to instantly filter
- Combines with the search box

**Search:**
- Searches across name, phone, email, and skills simultaneously

**Table Actions (per row):**
| Button | Action |
|---|---|
| Click name | Opens detail side panel |
| Remarks cell | Click to open floating remarks popover |
| **Update** | Change status + add a remark |
| **+Note** | Add a remark without changing status |
| **→Transfer** | Move to Workforce (removes from candidates) |

**Bulk Update:**
1. Tick checkboxes on multiple candidates
2. Click **Bulk Update** → set status + optional remark → applies to all

**Add Manually:**
Click **+ Add** to add a candidate without uploading a CV.

---

### 🗂️ Candidate Detail Panel

Click any candidate name to open the side panel:

- **Contact buttons** — direct Call / WhatsApp links
- **Professional info** — experience, company, education, skills
- **AI Summary** — 2-3 line summary extracted from the CV
- **Re-extract AI** button — re-runs AI extraction on the stored CV (useful if the first extraction missed data)
- **Call Detail Fields** — configurable fields (From Delhi?, Salary, Notice Period, etc.) that auto-save as you type
- **Remarks** — full list with edit + delete per remark
- **Inline CV Preview** — the CV renders right inside the panel (click to expand fullscreen)
- **Keyboard shortcuts:** `←` `→` to navigate between candidates, `Esc` to close

---

### 📤 Upload CVs

1. Select the **Position** from the dropdown
2. Drag and drop files OR click to browse
3. Upload multiple PDFs/DOCs at once
4. Each file is processed by AI — extracted data appears in real time
5. Duplicates are automatically blocked with a clear error message

---

### 👥 Workforce

Tracks candidates who have been transferred post-selection:

**Stages:** Training → Probation → Intern (Active) → Intern (Completed) → Full-time Hired

- Filter by stage or position
- Edit salary, end date, stage, add remarks
- CV preview available

**How to transfer a candidate:**
1. On the Candidates page, click **→Transfer** on any candidate
2. Choose stage, salary, join date, optional note
3. Confirm — the candidate moves from Candidates to Workforce

---

### 📁 CV Files

Shows all CV folders organized by position in Google Drive:
- Click **Open Root Folder** to see everything in Drive
- Click any position folder to browse individual CVs

---

### 📈 Reports

Generate reports for any date range:

- **CVs Added** — total count + daily breakdown chart
- **Status Changes** — horizontal bar chart (how many candidates moved to each status)
- **CVs by Position** — which roles got the most applicants
- **Cost / CV** — total spend ÷ CVs added
- **Spend by Platform** — LinkedIn vs Naukri vs other platforms

Quick buttons: **Today / Week / Month**

---

### 💰 Expenses

Log and track hiring spend:

- Add entries with: date, platform, amount, description
- Filters by date range
- Summary stats: total + breakdown by platform
- Delete any entry

---

### ⚙️ Settings

Configure everything without touching code:

| Setting | What it does |
|---|---|
| **Candidate Statuses** | Add/remove status options (shown in Update dropdown) |
| **Expense Platforms** | Add/remove platform options (LinkedIn, Naukri, etc.) |
| **Positions** | Add/remove job positions |
| **Call Detail Fields** | Add custom per-candidate fields (text, yes/no, dropdown) |

---

### 📋 Activity Log

Full audit trail of every action:

- Filter by date range + free-text search
- Shows: user, action, candidate ID, details, timestamp
- Quick buttons: **Today / Week / All**

---

## 🔧 Customization

### Change Default Positions

Run `initializeSheet` again after modifying the positions array in `Code.gs`, or just add them via **Settings → Positions** in the UI.

### Change Default Call Fields

Go to **Settings → Call Detail Fields** and add fields as needed. Types:
- `text` — free text input
- `yesno` — Yes/No dropdown
- `dropdown:Option1,Option2,Option3` — custom dropdown

### Fix Duplicate IDs (if any)

If you ever see duplicate candidate IDs (shouldn't happen normally):
1. In Apps Script editor, run the function `fixDuplicateIds`

---

## 🛡️ Security Notes

- The team password is stored as a plain string in `Code.gs`. For higher security, move it to Script Properties (same way as the API key).
- All CVs are stored in a private Google Drive folder — only accessible to you and anyone you share the folder with.
- The web app runs as "Me" (your Google account) — external users can't access your Drive directly.

---

## 📌 Known Limitations

- **Google Apps Script execution limit:** 6 minutes per run. Batch uploading very large files may time out.
- **CV extraction accuracy:** Depends on CV quality and layout. Use "Re-extract AI" if a field shows N/A.
- **No real-time collaboration:** If two people update the same candidate simultaneously, the last write wins.

---

## 🤝 Contributing

This was built as an internal hiring tool, open-sourced for anyone to use. If you find bugs or have feature ideas, open an Issue or a PR. All contributions welcome.

---

## 📄 License

MIT — use freely, modify as needed. Attribution appreciated but not required.

---

*Built with ❤️ using Google Apps Script + Claude AI.*
