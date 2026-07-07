# S4U Booking Task Pane Add-in

Excel JavaScript task pane add-in that replaces the legacy Google Sheets +
Apps Script workflow. Ships as a **single combined manifest** (POD Booking)
that adds one ribbon tab with four buttons to POD workbooks.

- Version: `manifest.xml` v1.0.6.0
- Hosting: `https://akhilabilling-s4u.github.io/s4u_pages/`
- Scope: runs inside `POD A.xlsx`, `POD B.xlsx`, `POD RoW.xlsx` only

---

## Ribbon layout

Single tab **POD Booking** with three groups and four buttons:

| Group | Button | Task pane | Purpose |
|---|---|---|---|
| Booking | **Fill Row** | `form.html` | Append one validated booking row to the active POD sheet |
| Time | **In Time** | `timepicker.html` | Append `HH:MM` history to the active cell |
| Time | **CSS / VA DL** | `daytimepicker.html` | Write `Mon, 09:00 AM` style value to the active cell |
| Client DB | **New Client** | `newClient.html` | Queue a new client / user mapping. Includes an in-pane **Sync** button to trigger the reconcile flow on demand |

---

## Architecture

```
User workbook (POD A/B/RoW)                         Silent master
============================                        ============================
POD workbook (SharePoint)                           S4U CLIENT DB.xlsx (SharePoint)
  |                                                   |
  |-- Database (hidden PQ mirror) <------ read-only mirror -----| Database (source of truth)
  |    driven by Power Query                                    | plain sheet, no Table needed
  |    refresh on open                                          |
  |                                                             |
  |-- tbl_NewClientQueue (hidden Excel Table)                   |
  |    written by: common.js > saveNewClient()                  |
  |    columns: QueuedAt, ClientName, UserName, POD,            |
  |             Sensitivity, Reviewing, Status,                 |
  |             ProcessedAt, Error                              |
  |                                                             |
  |-- Ribbon: POD Booking                                       |-- Office Script: "Reconcile New Clients"
       Fill Row / In Time / CSS-VA DL / New Client              |     signature: main(workbook, payload)
                                                                |
                                                                v
                            Power Automate flow (every 15 min or on-demand)
                              For each POD:
                                List tbl_NewClientQueue rows where Status='Pending'
                                For each row:
                                  Run script (master) with payload
                                  Update queue row: Status/ProcessedAt/Error
```

Users never open the master. The master is edited only by the Office Script,
invoked by Power Automate. Data flows one way (POD queue -> master); the
POD's own Database view is a read-only Power Query mirror that refreshes
from the master.

---

## Folder layout

```
TaskPaneAddin/
  manifest.xml               single combined manifest (v1.0.6.0)
  ReconcileNewClients.ts     copy of the Office Script installed in master
  README.md                  this file
  SETUP-RUNBOOK.md           extended runbook (mirror of the sections below)
  DEPLOY.md                  legacy notes (kept for history)
  src/
    common.js                shared Office.js helpers + guards + queue writer
    form.html                Fill Row pane
    timepicker.html          In Time pane
    daytimepicker.html       CSS / VA DL pane
    newClient.html           New Client pane (queue write + Sync button)
  assets/
    icon-16.png  icon-32.png  icon-80.png
```

The Office Script that runs inside the master lives at
`Excel_Migration/OfficeScripts/ReconcileNewClients.ts` and is duplicated here
for convenience.

---

## Workbook-name guard

Every task pane calls `POD_BOOKING.isPodWorkbook()` from `common.js` on
startup and refuses to render its form when the workbook name does not
normalise to one of `poda`, `podb`, `podrow`. Recognised names:

- `POD A.xlsx`, `POD-A.xlsx`, `PodA.xlsx`, `POD_A.xlsx`
- `POD B.xlsx`, `PodB.xlsx`, ...
- `POD RoW.xlsx`, `POD-ROW.xlsx`, `PODRow.xlsx`, ...

Any other workbook (including `S4U CLIENT DB.xlsx`) shows a red banner
instead of the form. This is enforced at runtime; the manifest cannot
conditionally hide ribbon tabs.

---

## Prerequisites

1. **Master workbook** `S4U CLIENT DB.xlsx` on SharePoint. Must have a plain
   sheet named exactly `Database` with row 1 headers including at minimum
   `User` and `Client` (aliases supported: `Username`, `User Name`, `Client
   Name`, `Client Id`, `Client Code`, `POD`, `Sensitivity`, `Reviewing` /
   `INC` / `Flag`). No Excel Table conversion required.

2. **POD workbooks** `POD A.xlsx`, `POD B.xlsx`, `POD RoW.xlsx` on SharePoint.
   Each has a hidden `Database` sheet fed by Power Query from the master.

3. **Static hosting** for `src/` and `assets/`. Currently
   `https://akhilabilling-s4u.github.io/s4u_pages/`.

4. **M365 admin centre** access to upload the manifest to Integrated apps.

5. **Power Automate** access with the standard `Excel Online (Business)`
   connector. No Premium licence required.

---

## Deployment

### STEP 1 - Publish the static files

Upload the contents of `src/` and `assets/` to
`https://akhilabilling-s4u.github.io/s4u_pages/`. Verify each URL is
reachable:

- `https://akhilabilling-s4u.github.io/s4u_pages/common.js`
- `https://akhilabilling-s4u.github.io/s4u_pages/form.html`
- `https://akhilabilling-s4u.github.io/s4u_pages/timepicker.html`
- `https://akhilabilling-s4u.github.io/s4u_pages/daytimepicker.html`
- `https://akhilabilling-s4u.github.io/s4u_pages/newClient.html`
- `https://akhilabilling-s4u.github.io/s4u_pages/assets/icon-32.png`

### STEP 2 - Upload the manifest to the M365 admin centre

1. `https://admin.microsoft.com` > **Settings** > **Integrated apps**.
2. **Upload custom apps** > **Office Add-in**.
3. Upload `manifest.xml` (v1.0.6.0). If updating an existing entry, click the
   listed app and use **Update**.
4. Assign to the POD user group. Save.

### STEP 3 - Build the Power Automate reconcile flow

The flow drains each POD's `tbl_NewClientQueue` into the master `Database`
sheet on a schedule and on demand. This is the **bridge** between POD
workbooks and the master; without it, queued rows never reach the master.

#### 3.1 Save the Office Script in the master

1. Open `S4U CLIENT DB.xlsx` in Excel for the Web (admin, one-time).
2. **Automate** > **New Script**.
3. Paste the entire contents of `ReconcileNewClients.ts`.
4. Name the script exactly `Reconcile New Clients`.
5. **Save** (Ctrl+S). Do not click Run. Close the workbook.

#### 3.2 Create the scheduled flow

1. Open `https://make.powerautomate.com` and sign in.
2. **+ Create** > **Scheduled cloud flow**.
3. Fill in:

   | Field | Value |
   |---|---|
   | Flow name | `S4U Reconcile New Clients` |
   | Starting | today, any time |
   | Repeat every | `15` **Minute** (use `5` for tighter latency; stay above `1` to respect the free-tier daily cap) |

4. **Create**. The flow designer opens with a single **Recurrence** card.

#### 3.3 Add one drain block per POD

For each POD (`POD A`, `POD B`, `POD RoW`), add three actions in sequence.
Start with POD A; duplicate for the other two.

##### Action 1 - List rows present in a table

- Connector: **Excel Online (Business)** > **List rows present in a table**
- Location: **SharePoint Site**
- Document Library: (library containing the POD file)
- File: `POD A.xlsx`
- Table: `tbl_NewClientQueue`
- **Show advanced options** > **Filter Query**: `Status eq 'Pending'`

> If the Table dropdown is empty, no one has clicked **Client DB > New
> Client > Add** in that POD yet. Do a dummy submission first so the add-in
> auto-creates `tbl_NewClientQueue`, then refresh the dropdown.

##### Action 2 - Apply to each pending row

- **+ New step** > **Control** > **Apply to each**
- Input: dynamic content **value** from Action 1

Inside the loop, add:

##### Action 2a - Run script (master)

- Connector: **Excel Online (Business)** > **Run script**
- Location: SharePoint Site
- Document Library: (library containing `S4U CLIENT DB.xlsx`)
- File: `S4U CLIENT DB.xlsx`
- Script: `Reconcile New Clients`
- **payload** (Object) - click **Add new item** to expand, then map:

  | Payload field | Dynamic content from Action 1 |
  |---|---|
  | clientName | `ClientName` |
  | userName | `UserName` |
  | pod | `POD` |
  | sensitivity | `Sensitivity` |
  | reviewing | `Reviewing` |

##### Action 2b - Update a row (this POD)

- Connector: **Excel Online (Business)** > **Update a row**
- Location: SharePoint Site
- File: `POD A.xlsx`
- Table: `tbl_NewClientQueue`
- Key Column: `QueuedAt`
- Key Value: dynamic content **QueuedAt** (from the Apply-to-each item)
- Row body:

  | Field | Value |
  |---|---|
  | Status | expression `outputs('Run_script')?['body/result/status']` |
  | ProcessedAt | expression `utcNow()` |
  | Error | expression `outputs('Run_script')?['body/result/message']` |

  Leave all other fields empty so they retain their existing values.

#### 3.4 Repeat for POD B and POD RoW

Right-click each of Actions 1 and 2 > **Copy to my clipboard**. Paste below
the previous POD's Apply-to-each block. Change the **File** in the pasted
"List rows" and "Update a row" actions to the next POD file. Do not change
the "Run script" action; all three PODs use the same script on the same
master file.

Final flow layout:

```
Recurrence (every 15 min)
+-- List rows (POD A)
+-- Apply to each (POD A)
|     +-- Run script (S4U CLIENT DB.xlsx / Reconcile New Clients)
|     +-- Update a row (POD A / tbl_NewClientQueue)
+-- List rows (POD B)
+-- Apply to each (POD B)
|     +-- Run script (S4U CLIENT DB.xlsx / Reconcile New Clients)
|     +-- Update a row (POD B / tbl_NewClientQueue)
+-- List rows (POD RoW)
+-- Apply to each (POD RoW)
      +-- Run script (S4U CLIENT DB.xlsx / Reconcile New Clients)
      +-- Update a row (POD RoW / tbl_NewClientQueue)
```

#### 3.5 Save and test

1. Top right > **Save**.
2. Top right > **Test** > **Manually** > **Test** > **Run flow**.
3. Open the **28-day run history** tab. Every action in the latest run
   should be green (**Succeeded**). Empty POD queues loop zero times and
   still show Succeeded.

#### 3.6 Wire the Sync button in `newClient.html`

1. In the browser, copy the URL of the flow's detail page. It looks like:

   ```
   https://make.powerautomate.com/environments/Default-<env>/flows/<flowId>/details
   ```

2. Edit `src/newClient.html` line ~185:

   ```javascript
   // before
   const FLOW_URL = "PASTE_YOUR_FLOW_URL_HERE";

   // after
   const FLOW_URL = "https://make.powerautomate.com/environments/Default-.../flows/.../details";
   ```

3. Save and re-upload `newClient.html` to the GitHub Pages host.
4. Reopen the New Client task pane in a POD. The **Sync** button in the top
   sync bar is now enabled. Clicking it opens the flow's page in the default
   browser, where users can click **Run** to trigger reconcile on demand.

### STEP 4 - Power Query mirror in each POD (one-time)

Requires Excel Desktop (or iPad Excel) once per POD; after authoring, all
subsequent refreshes work in Excel for the Web.

Per POD workbook (`POD A.xlsx` first, then B, then RoW):

1. Open the POD in Excel Desktop from SharePoint.
2. **Data** > **Get Data** > **From File** > **From SharePoint Folder**.
3. Site URL: the root of the site (not the file URL).
4. In the Navigator, find `S4U CLIENT DB.xlsx` > **Transform Data**.
5. Drill into `Content` > find the row where `Name = Database`, `Kind = Sheet`
   > click the **Table** link in the **Data** column.
6. **Home** > **Use First Row as Headers** if headers show as `Column1...`.
7. Remove any extra columns (`Source.Name`, etc.). Set data types.
8. Right rail > rename the query to `qMasterClients`.
9. **Home** > **Close & Load To...** > **Table** > **New worksheet**.
10. Rename the new sheet exactly `Database`. Right-click > **Hide**.
11. **Data** > **Queries & Connections** > right-click `qMasterClients` >
    **Properties** > tick **Refresh data when opening the file** and
    **Enable background refresh**. OK.
12. Save.

### STEP 5 - End-to-end smoke test

1. Open `POD A.xlsx` in Excel for the Web.
2. **POD Booking** > **New Client**. Sync bar shows `0 pending . 0 done . 0 err`.
3. Submit a brand-new client. Toast: "Queued: client X added." Sync bar
   flips to `1 pending`.
4. Click **Sync**. Browser tab opens on flow page. Click **Run** >
   **Run flow**. Wait ~10 seconds.
5. Back in Excel, close and reopen the New Client pane. Sync bar shows
   `0 pending . 1 done . 0 err`.
6. **Data** > **Refresh All** in the POD (or close and reopen the workbook).
7. **POD Booking** > **Fill Row**. The new client appears in the dropdown.

Repeat the drop-and-refresh cycle from POD B and POD RoW to confirm full
cross-POD visibility.

---

## Local edit / republish loop

```
1. Edit src/*.html or src/common.js
2. Commit and push to the hosting repo (akhilabilling-s4u.github.io/s4u_pages)
3. Hard-refresh (Ctrl+F5) each URL in a browser to confirm the new content
4. In Excel: close and reopen the workbook. If Desktop, clear the Wef cache:
   Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
   Remove-Item "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef" -Recurse -Force -ErrorAction SilentlyContinue
```

Manifests rarely change; bump `<Version>` and re-upload only when the
ribbon, source URLs, or permissions change.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Ribbon tab missing | Manifest not deployed yet, or Desktop cache stale | Wait up to 24 h after first deploy on Desktop; clear Wef cache; verify manifest is Deployed in admin centre. |
| "POD Booking is available only inside POD workbooks" | Workbook name not in whitelist | Rename to `POD A / POD B / POD RoW` (`.xlsx`). |
| New Client "No clients found" | POD's local Database mirror missing | Complete STEP 4 for that POD. |
| Sync button disabled with "Flow URL not configured" | `FLOW_URL` still placeholder | STEP 3.6 - paste the flow URL into `newClient.html` and re-upload. |
| Flow run history: "Table not found" | `tbl_NewClientQueue` doesn't exist in a POD | Have someone submit a New Client in that POD once; the add-in auto-creates the table. |
| Flow run history: "Script returned Error - Master workbook is missing worksheet 'Database'" | Database sheet renamed in master | Rename back to exactly `Database`. |

See `SETUP-RUNBOOK.md` for the extended troubleshooting matrix.

---

## Retired pieces

| Item | Status |
|---|---|
| `manifest-pod.xml`, `manifest-masterdb.xml` | Superseded by the combined `manifest.xml`. Remove from admin centre if still installed. |
| `getWorkbookContext()` role detection | Kept for compatibility; not used by any pane. Replaced by `isPodWorkbook()`. |
| Two-tab layout (POD Booking + Client DB) | Consolidated into single POD Booking tab. |
| Direct POD-to-master writes | Replaced by queue + Power Automate reconcile. |
| Standalone Sync Now pane (`syncNow.html`) | Folded into `newClient.html` as an inline Sync button. |
