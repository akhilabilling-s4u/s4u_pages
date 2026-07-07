# S4U Booking — Deployment Runbook

End-to-end setup for the **POD Booking + Client DB** Excel task-pane add-in and its scheduled reconcile pipeline.

> Total wall-clock to a green production state: **~90 minutes**, of which only ~15 minutes per POD requires Excel Desktop (Power Query authoring is the only desktop-bound step).

---

# Table of Contents

1. [Architecture](#architecture)
2. [Prerequisites](#prerequisites)
3. [Step 1 — Verify the master Database sheet](#step-1--verify-the-master-database-sheet)
4. [Step 2 — Create the Power Query mirror in each POD](#step-2--create-the-power-query-mirror-in-each-pod)
5. [Step 3 — Build the Power Automate reconcile flow](#step-3--build-the-power-automate-reconcile-flow)
6. [Step 4 — End-to-end smoke test](#step-4--end-to-end-smoke-test)
7. [Step 5 — Hand off to users](#step-5--hand-off-to-users)
8. [Troubleshooting](#troubleshooting)
9. [Operational rhythm](#operational-rhythm)
10. [Rollback](#rollback)

---

# Architecture

```text
                          ┌──────────────────────────────────────────────┐
                          │  S4U CLIENT DB.xlsx (silent, no user opens) │
                          │──────────────────────────────────────────────│
                          │ Database sheet (source of truth)            │
                          │ Office Script: Reconcile New Clients        │
                          └──────────────┬───────────────────────────────┘
                                         │
                          Power Query refresh on open / Refresh All
                                         ▲
                                         │

┌──────────────────────────┐             │            ┌──────────────────────────┐
│ POD A.xlsx               │─────────────┼────────────│ POD B.xlsx               │
│ - Database (hidden PQ)   │             │            │ - Database (hidden PQ)   │
│ - tbl_NewClientQueue     │             │            │ - tbl_NewClientQueue     │
│ - POD Booking ribbon     │             │            │ - POD Booking ribbon     │
│ - Client DB ribbon       │             │            │ - Client DB ribbon       │
└────────────┬─────────────┘             │            └────────────┬─────────────┘
             │                           │                         │
             └───────────────────────────┼─────────────────────────┘
                                         │
                        Power Automate (every 15 minutes)
                           "S4U Reconcile New Clients"
                                         │
                          Run script → Master Database
                          Update queue → Status = Processed
```

## Key Invariants

| What | Where | Why |
|------|--------|-----|
| `Database` sheet | Master workbook | Source of truth |
| `Database` mirror | POD workbooks | Dropdown source |
| `tbl_NewClientQueue` | POD workbooks | Required by Power Automate |
| Add-in HTML/JS | GitHub Pages | Static hosting |
| Manifest | M365 Admin Center | Ribbon deployment |

---

# Prerequisites

| Item | Required | Notes |
|------|----------|-------|
| M365 admin role | Yes | Office Apps Admin or higher |
| Power Automate access | Yes | Standard license only |
| Excel Desktop | Once | Required for Power Query authoring |
| SharePoint site | Yes | Hosts all workbooks |
| S4U CLIENT DB.xlsx | Yes | Master workbook |
| POD workbooks | Yes | User-facing files |
| GitHub Pages hosting | Yes | Add-in source |

## Files Already Deployed

- [x] HTML/JS published to GitHub Pages
- [x] `manifest.xml` uploaded to M365 Admin Center
- [x] `ReconcileNewClients.ts` saved in `S4U CLIENT DB.xlsx`

---

# Step 1 — Verify the Master Database Sheet

> Estimated time: **5 minutes**

1. Open `S4U CLIENT DB.xlsx` in Excel for the Web.
2. Verify the worksheet name is exactly:
   - `Database`
3. Confirm Row 1 contains at least:
   - `User`
   - `Client`

## Accepted Header Aliases

| Logical Field | Accepted Headers |
|--------------|-----------------|
| User | `User`, `Username`, `User Name` |
| Client | `Client`, `Client Name` |
| Client Code | `Client Id`, `Client Code` |
| POD | `POD` |
| Sensitivity | `Sensitivity` |
| Reviewing | `Reviewing`, `INC`, `Flag` |

> The reconciliation script validates these headers automatically.

---

# Step 2 — Create the Power Query Mirror

> Estimated time: **15 minutes per POD**

Repeat for:

- POD A.xlsx
- POD B.xlsx
- POD RoW.xlsx

## 2.1 Open Workbook

1. Open the POD workbook in SharePoint.
2. Select **Open in Desktop App**.

## 2.2 Create Power Query

1. Data → Get Data → From File → From SharePoint Folder.
2. Enter:

```text
https://<tenant>.sharepoint.com/sites/<sitename>
```

3. Sign in.
4. Select `S4U CLIENT DB.xlsx`.
5. Click **Transform Data**.

Continue until:

- Name = `Database`
- Kind = `Sheet`

If necessary:

- Home → Use First Row as Headers.
- Remove extra columns.
- Detect Data Types.

Rename the query:

```text
qMasterClients
```

## 2.3 Load Data

- Close & Load To...
- Table
- New Worksheet

## 2.4 Rename Worksheet

Rename to:

```text
Database
```

Hide the worksheet afterward.

## 2.5 Enable Refresh

Enable:

- Refresh data when opening the file
- Enable background refresh

## 2.6 Verify

- Save.
- Reopen workbook.
- Test:
  - POD Booking → Fill Row
  - Client DB → New Client

---

# Step 3 — Build the Power Automate Flow

> Estimated time: **20 minutes**

## Create Flow

1. Open:

```text
https://make.powerautomate.com
```

2. Create → Scheduled Cloud Flow.
3. Name:

```text
S4U Reconcile New Clients
```

4. Run every:

```text
15 minutes
```

---

## List Rows

Configure:

- Location: SharePoint
- File: POD workbook
- Table: `tbl_NewClientQueue`

Filter:

```text
Status eq 'Pending'
```

---

## Run Script Payload

```json
{
  "clientName": "@{items('Apply_to_each')?['ClientName']}",
  "userName": "@{items('Apply_to_each')?['UserName']}",
  "pod": "@{items('Apply_to_each')?['POD']}",
  "sensitivity": "@{items('Apply_to_each')?['Sensitivity']}",
  "reviewing": "@{items('Apply_to_each')?['Reviewing']}"
}
```

---

## Update Row

| Field | Value |
|------|--------|
| Key Column | QueuedAt |
| Status | Run script status |
| ProcessedAt | utcNow() |
| Error | Run script message |

Repeat for:

- POD A
- POD B
- POD RoW

---

# Step 4 — End-to-End Smoke Test

1. Open POD workbook.
2. Verify Fill Row dropdowns.
3. Add a new client.
4. Confirm queue row exists.
5. Run Power Automate.
6. Verify:
   - Status = Processed
   - ProcessedAt populated
7. Confirm row exists in master Database.
8. Refresh POD.
9. Verify dropdowns contain new values.

---

# Step 5 — Hand Off to Users

> The POD workbooks now contain two new ribbon tabs:
>
> - **POD Booking**
>   - Fill Row
>   - In Time
>   - CSS / VA DL
>
> - **Client DB**
>   - New Client submission
>
> New entries become available after reconciliation and a Data Refresh.

---

# Troubleshooting

## No Clients Found

Missing local Database mirror.

Complete Step 2.

---

## Empty Dropdowns

Refresh:

```text
Data → Refresh All
```

---

## Empty Table Dropdown

`tbl_NewClientQueue` has not yet been created.

Submit one test record.

---

## Missing Database Worksheet

Error:

```text
Master workbook is missing worksheet 'Database'
```

Rename the worksheet exactly:

```text
Database
```

---

## Missing Columns

Error:

```text
Database sheet must have at least 'User' and 'Client' columns
```

Add or rename headers.

---

## Status = Skipped

Duplicate User + Client already exists.

No action required.

---

## New Client Not Visible

Run:

```text
Data → Refresh All
```

---

## Empty New Client Form

Clear Office cache:

```powershell
Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue

Remove-Item `
"$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef" `
-Recurse -Force `
-ErrorAction SilentlyContinue
```

---

## Missing Ribbon

Propagation may take up to 24 hours.

Check:

```text
Insert → Get Add-ins → Admin Managed
```

---

# Operational Rhythm

| Event | Owner | Action |
|------|--------|--------|
| Workbook opens | User | Refreshes Database mirror |
| Booking added | User | Updates POD workbook |
| New client request | User | Creates queue record |
| Every 15 minutes | Power Automate | Processes queue |
| Refresh All | User | Updates dropdown data |
| Schema changes | Admin | Add columns safely |

---

# Rollback

## Remove Add-in

```text
Admin Center → Settings → Integrated Apps
```

Remove:

```text
POD Booking
```

## Disable Flow

Turn off or delete:

```text
S4U Reconcile New Clients
```

## Delete Script

```text
Excel → Automate → All Scripts
```

Delete:

```text
Reconcile New Clients
```

## Remove Power Query

```text
Data → Queries & Connections
```

Delete:

```text
qMasterClients
```

---

# Repository Layout

```text
Excel_Migration/
├── TaskPaneAddin/
│   ├── manifest.xml
│   ├── ReconcileNewClients.ts
│   ├── README.md
│   ├── SETUP-RUNBOOK.md
│   ├── DEPLOY.md
│   ├── assets/
│   │   ├── icon-16.png
│   │   ├── icon-32.png
│   │   └── icon-80.png
│   └── src/
│       ├── common.js
│       ├── form.html
│       ├── timepicker.html
│       ├── daytimepicker.html
│       └── newClient.html
└── OfficeScripts/
    └── ReconcileNewClients.ts
```

---

# Optional Enhancements

- Email notifications for Error/Skipped rows.
- Automatic Power Query refresh after reconciliation.
- Audit logging in the master workbook.

---

*Maintained by the S4U Migration Team.*

*Last updated for code base version `manifest.xml v1.0.2.0`.*
