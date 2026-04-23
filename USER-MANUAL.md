# G-Stack Asset Command User Manual

Version: 1.0  
Module: `asset-command`  
Last updated: March 8, 2026

## 1. What G-Stack Asset Command Does

G-Stack Asset Command is a Google Sheets-based fleet and asset operations system for:

- Asset inventory and status tracking
- Driver credential/compliance tracking
- Maintenance scheduling and overdue monitoring
- Fuel and activity logging
- Dashboard reporting and alert emails

It runs as a container-bound Google Apps Script project inside a Google Sheet.

## 2. Prerequisites

Before setup, make sure you have:

- A Google account with access to Google Sheets and Apps Script
- Permission to send email from Apps Script (`MailApp`)
- If using CLI install:
- `Node.js` (recommended v18+)
- `@google/clasp` installed globally

Install clasp:

```powershell
npm install -g @google/clasp
clasp login
```

## 3. Install Option A - Using clasp (Recommended)

Use this when you want repeatable deployment and versioned local files.

### 3.1 Prepare a clean deployment folder

Important: avoid reusing another person's `.clasp.json`.

```powershell
cd "<REPO_ROOT>"

$deployDir = "C:\TNDS\deployments\asset-command"
if (Test-Path $deployDir) { Remove-Item $deployDir -Recurse -Force }
New-Item -ItemType Directory -Path $deployDir | Out-Null

Copy-Item .\*.gs,.\*.html,.\appsscript.json -Destination $deployDir
cd $deployDir
```

### 3.2 Create a new Sheet + Apps Script in your Google Drive

```powershell
clasp show-authorized-user
clasp create --title "G-Stack Asset Command TEST" --type sheets
clasp push --force
```

This creates:

- A new Google Sheet in the authenticated user's Drive
- A bound Apps Script project linked to that sheet

### 3.3 Open the project

With clasp v3+, use:

```powershell
clasp open-script
clasp open-container
```

## 4. Install Option B - Manual Copy/Paste in Apps Script

Use this when the user does not want CLI tooling.

### 4.1 Create container

1. Create a new Google Sheet.
2. Open `Extensions > Apps Script`.
3. Rename default project if desired.

### 4.2 Create script and HTML files

In Apps Script editor, create these files and paste contents from local module:

- `Code.gs`
- `FunctionRunner.gs`
- `Dashboard.html`
- `Sidebar.html`
- `Help.html`

### 4.3 Update manifest (`appsscript.json`)

1. In Apps Script editor, enable `Show "appsscript.json" manifest file` in project settings if hidden.
2. Replace manifest content with the local `appsscript.json` content.
3. Save all files.

Required scopes in manifest:

- `https://www.googleapis.com/auth/spreadsheets`
- `https://www.googleapis.com/auth/drive`
- `https://www.googleapis.com/auth/script.container.ui`

### 4.4 Authorize and initialize menu

1. In Apps Script editor, run function `onOpen` once.
2. Accept authorization prompts.
3. Return to the sheet and refresh.
4. Confirm the custom menu appears (menu label may be `🚗 test` or branded variant).

## 5. First-Time Setup Checklist (All Install Methods)

Run these in order from the custom menu:

1. `1. Build Complete Template`
2. `2. Setup Dashboard 2`
3. `3. Add Test Data` (optional but recommended for validation)

Then open `Config` sheet and set:

- `B2` Business Name
- `B3` Email Address (required for alerts)
- `B4` Alert Threshold (days) for maintenance reminders
- `B5` Fuel Anomaly Threshold (%)

If you added test data, validate dashboard behavior, then optionally run:

- `Clear Test Data`

## 6. What Gets Created

After `Build Complete Template`, these sheets are created:

- `Dashboard`
- `Dashboard 2` (after step 2)
- `Assets Master`
- `Drivers`
- `Activity Log`
- `Maintenance Tracker`
- `Cost Tracking`
- `Config`
- `Setup Instructions`

## 7. Daily User Workflow

### 7.1 Add data quickly (Sidebar)

Use menu `📝 Data Entry` to open sidebar.

Primary actions:

- Add asset
- Add driver
- Add maintenance record
- Add activity/fuel record
- CSV import (bulk records)

### 7.2 Review dashboards

- `🖥️ Open Dashboard` for interactive dashboard
- `Refresh Dashboards` after large imports/changes
- Review maintenance/compliance/fuel KPIs

### 7.3 Send alerts

- `Send Daily Digest`
- `Check Maintenance Due`
- Driver compliance submenu:
- `Check Driver Credentials`
- `Check Fuel Anomalies`

## 8. Function Runner (Optional Advanced Control)

Use this if you want checkbox-based execution of common functions.

Setup:

1. Menu `⚙️ Function Runner > Initialize Function Runner`
2. Menu `⚙️ Function Runner > Setup Trigger`

Use:

- Check a box in `Function Runner` sheet to execute that function.
- Trigger can be removed with `Remove Trigger`.

## 9. Configuration Reference

Config sheet values used by system:

- `Config!B2` business name (email subject and dashboard labels)
- `Config!B3` alert email recipient
- `Config!B4` maintenance threshold in days
- `Config!B5` fuel anomaly percentage threshold

If `B3` is blank, alert functions stop with a warning.

## 10. Initial Validation Test (Recommended)

Run this once after installation:

1. Build template and Dashboard 2.
2. Add test data.
3. Open dashboard and confirm KPIs/charts render.
4. Set `Config!B3` to your email.
5. Run `Send Daily Digest` and confirm email delivery.
6. Run `Check Maintenance Due` and confirm behavior.
7. Clear test data if needed.

## 11. Troubleshooting

### Menu does not appear

- Run `onOpen` manually once from Apps Script editor, then refresh sheet.
- Confirm script is bound to the same spreadsheet.

### "Please run Build Complete Template first"

- You are calling functions before template sheets exist.
- Run `1. Build Complete Template`.

### Alert email functions fail

- Set `Config!B3` to a valid email.
- Re-run authorization if `MailApp` permission was denied.

### `clasp open` says unknown command

- Use:
- `clasp open-script`
- `clasp open-container`

### Data Entry sidebar/manual dialog does not load

- Confirm `Sidebar.html` and `Help.html` exist in Apps Script project.
- Save project and refresh spreadsheet.

## 12. Security and Deployment Notes

- Never commit personal `.clasp.json` to shared repos.
- Keep `.clasp.json.template` for distribution.
- Each installer user must run `clasp login` with their own Google account.
- `clasp create --type sheets` creates assets in the currently authenticated user's Drive.

## 13. Admin Handoff Checklist

Before handing to end users:

1. Confirm install method (clasp or manual).
2. Confirm custom menu appears and all items run.
3. Confirm Config fields are documented for the user.
4. Confirm at least one alert email test succeeded.
5. Share this manual and in-sheet `Setup Instructions` tab.


