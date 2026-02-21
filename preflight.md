# Chandlr Pre-flight Checklist

Before recording the demo, work through this checklist to verify everything works.

---

## 1. Prerequisites

- [ ] Microsoft 365 subscription (Business Standard or higher)
- [ ] Excel desktop app (Windows or Mac) — not Excel Online
- [ ] Copilot license enabled on your account (check: open Excel → Copilot icon visible in the ribbon?)
- [ ] Test data ready:
  - A PO file (CSV or PDF) — use `test-data/tesco-po-2025-0247.csv` or a real one
  - ERP price data — paste into a sheet called "ERP" with columns: SKU, Price (at minimum)

---

## 2. Sideload the Task Pane Add-in

This loads the button-driven UI (Reconcile tab + task pane).

### Windows
1. Open Excel → **Insert** tab → **My Add-ins** → **Upload My Add-in**
2. Browse to `manifest.xml` in the project root
3. Click **Upload** — you should see a "Chandlr" tab appear in the ribbon

### Mac
1. Open Excel → **Insert** tab → **My Add-ins** → **Upload My Add-in**
2. Browse to `manifest.xml`
3. If that option doesn't appear: copy `manifest.xml` to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
4. Restart Excel — the Chandlr tab should appear

### Verify
- [ ] "Chandlr" tab visible in the Excel ribbon
- [ ] Click "Open Chandlr" → task pane opens on the right

---

## 3. Test the Task Pane Flow (Button Path)

Work through this before touching Copilot — it validates the core logic.

### Extract
- [ ] Paste or open PO data in a sheet
- [ ] In the task pane, click **Upload PO** → select your CSV/PDF
- [ ] A "PO Staging" sheet should appear with clean, standardized columns
- [ ] Check: column headers detected correctly? Line count matches?

### Reconcile
- [ ] Paste ERP prices in a separate sheet (columns: SKU, Price at minimum)
- [ ] Select the ERP data range (including headers)
- [ ] Switch to the Reconcile tab in the task pane
- [ ] Click **Reconcile**
- [ ] A "Reconciliation" sheet should appear with color-coded rows:
  - Green = price match
  - Yellow = within tolerance
  - Red = exception (price mismatch)
- [ ] Check: SKU matching worked? Price differences calculated correctly?

### Credit Note
- [ ] In the Actions tab, click **Generate Credit Note**
- [ ] A "Credit Note" sheet should appear
- [ ] Check: all lines listed with negative amounts at PO prices?

### Re-Invoice
- [ ] Click **Generate Re-Invoice**
- [ ] A "Re-Invoice" sheet should appear
- [ ] Check: all lines listed with ERP (correct) prices? Changed prices highlighted yellow?

### ERP Staging
- [ ] Click **Generate ERP Staging**
- [ ] A "ERP Staging" sheet should appear with color coding:
  - Green = Ready (enter into ERP)
  - Yellow = Review (price changed, confirm first)
  - Red = Hold (SKU not found in ERP)

### Price Intelligence
- [ ] Go to the Tools tab in the task pane
- [ ] Click **Generate Dashboard**
- [ ] First time: should show basic metrics from the one reconciliation you just ran
- [ ] For a better test: reconcile a second PO (change a few prices), then generate dashboard again — should show trends

---

## 4. Sideload the Copilot Agent

This is separate from the task pane add-in. The Copilot agent uses the unified manifest (`appPackage/manifest.json`).

### Option A: Teams Toolkit (recommended)
1. Install **Teams Toolkit** extension in VS Code
2. Open the `po-reconciler` project folder in VS Code
3. In the Teams Toolkit sidebar → **Provision** (or **Preview in Teams**)
4. This uploads the manifest and registers the agent
5. Open Excel → Copilot panel → you should see "Chandlr" as an available agent

### Option B: Manual upload via Teams Admin
1. Zip the `appPackage/` folder contents (manifest.json, declarativeAgent.json, Office-API-local-plugin.json, assets/)
2. Go to Teams Admin Center → **Manage apps** → **Upload custom app**
3. Upload the zip
4. Open Excel → Copilot → Chandlr should appear

### Option C: Developer Portal
1. Go to https://dev.teams.microsoft.com/apps
2. Click **Import app** → upload the zipped appPackage
3. Click **Publish** → **Publish to your org**
4. Wait for admin approval, then it appears in Copilot

### Verify
- [ ] Open Excel with data loaded
- [ ] Open Copilot side panel (ribbon → Copilot)
- [ ] Chandlr appears as an agent you can chat with
- [ ] Conversation starters visible: "Extract PO data", "Reconcile PO", etc.

---

## 5. Test the Copilot Agent Flow

These are the exact prompts from the demo script. Test each one.

- [ ] With PO data in the active sheet, type: **"Extract the order data from this PO"**
  - Expected: staging sheet created, Copilot summarizes what it found
- [ ] Select ERP data, type: **"Reconcile the PO against the ERP data I've selected"**
  - Expected: reconciliation sheet with color coding, summary of matches/exceptions
- [ ] Type: **"Generate a credit note"**
  - Expected: credit note sheet created
- [ ] Type: **"Create an ERP-ready staging sheet"**
  - Expected: ERP staging sheet with Ready/Review/Hold color coding
- [ ] Type: **"Show me price trends from my reconciliation history"**
  - Expected: dashboard sheet with metrics (may be limited with only one run)

---

## 6. Common Issues

| Problem | Fix |
|---------|-----|
| Task pane is blank | Check browser console (F12 in desktop Excel). Likely a CORS or loading issue. Make sure GitHub Pages deployment is current. |
| "Chandlr" tab doesn't appear | Re-upload manifest.xml. On Mac, restart Excel after copying to wef folder. |
| Copilot doesn't show the agent | The unified manifest may need Teams Toolkit provisioning. Manual upload via admin center is the fallback. |
| SKU matching misses items | Check SKU format — the reconciler handles prefix matching (PO "1234" matches ERP "1234V001") but not arbitrary differences. |
| PDF extraction looks wrong | Check the PDF isn't scanned/image-based. The parser needs selectable text. |
| "No reconciliation data" on Actions tab | Run a reconciliation first — actions unlock after reconcile completes. |

---

## 7. Ready to Record?

Once all checkboxes above are ticked:
1. Close unnecessary windows and notifications
2. Set Excel zoom to 120%
3. Open `demo-script.md` on a second screen
4. Record with OBS, Loom, or QuickTime (Mac)
5. Do a dry run without recording first — get the muscle memory for clicks and timing
