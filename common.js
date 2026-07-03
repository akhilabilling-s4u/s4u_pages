/* common.js — shared Office.js helpers used by all task-pane HTML pages.
 *
 * Replaces the Apps Script `google.script.run.<fn>()` bridge with direct
 * Excel reads/writes via the Office.js / Excel JavaScript API.
 *
 * The original implementation referenced an external Google Sheet by ID
 * for master data. In Excel we expect a sheet named `Database` inside the
 * same workbook (see Excel_Migration/README.md for the consolidation step).
 */

const POD_BOOKING = (function () {
  const DB_SHEET_NAME = "Database";
  const DB_TABLE_NAME = "tbl_Database"; // exists only in the master S4U CLIENT DB workbook

  const DEFAULT_COLS = {
    user: 0,
    client: 1,
    clientId: 2,
    pod: 3,
    sensitivity: 4,
  };

  function normalizeText(value) {
    return String(value ?? "").trim().toLowerCase();
  }

  function normalizePod(value) {
    return normalizeText(value)
      .replace(/^pod[\s\-_]*/i, "")
      .replace(/[\s\-_]/g, "");
  }

  function detectColumns(rows) {
    if (!rows || !rows.length) {
      return { cols: DEFAULT_COLS, startRow: 0 };
    }

    const header = rows[0].map((v) => normalizeText(v));

    const hasHeader = header.some((h) =>
      ["user", "username", "user name", "client", "client name", "pod", "sensitivity"].includes(h)
    );

    const find = (candidates, fallback) => {
      const idx = header.findIndex((h) => candidates.includes(h));
      return idx >= 0 ? idx : fallback;
    };

    const cols = hasHeader
      ? {
          user: find(["user", "username", "user name"], DEFAULT_COLS.user),
          client: find(["client", "clientname", "client name"], DEFAULT_COLS.client),
          clientId: find(
            ["client id", "clientid", "client_id", "client code", "clientcode"],
            DEFAULT_COLS.clientId
          ),
          pod: find(["pod"], DEFAULT_COLS.pod),
          sensitivity: find(["sensitivity"], DEFAULT_COLS.sensitivity),
        }
      : DEFAULT_COLS;

    return { cols, startRow: hasHeader ? 1 : 0 };
  }

  /**
   * Detect workbook role:
   *  - master: contains tbl_Database
   *  - pod: contains Database sheet only
   *  - unknown: otherwise
   */
  async function getWorkbookContext() {
    return Excel.run(async (ctx) => {
      const wb = ctx.workbook;

      const table = wb.tables.getItemOrNullObject(DB_TABLE_NAME);
      const sheet = wb.worksheets.getItemOrNullObject(DB_SHEET_NAME);

      table.load("name");
      sheet.load("name");

      await ctx.sync();

      const hasTable = !table.isNullObject;
      const hasSheet = !sheet.isNullObject;

      let role = "unknown";
      if (hasTable) role = "master";
      else if (hasSheet) role = "pod";

      return { role, hasTable, hasSheet };
    });
  }

  async function loadDatabase(context) {
    const wb = context.workbook;
    const sheet = wb.worksheets.getItemOrNullObject(DB_SHEET_NAME);

    sheet.load("name");
    await context.sync();

    if (sheet.isNullObject) return [];

    const used = sheet.getUsedRange();
    used.load("values");
    await context.sync();

    return used.values || [];
  }

  async function getDropdownData(podFilter) {
    return Excel.run(async (ctx) => {
      const rows = await loadDatabase(ctx);
      const { cols, startRow } = detectColumns(rows);

      const users = new Set();
      const wantedPod = normalizePod(podFilter);

      for (let i = startRow; i < rows.length; i++) {
        const u = String(rows[i][cols.user] || "");
        const p = String(rows[i][cols.pod] || "");

        const podMatch = !wantedPod || normalizePod(p) === wantedPod;
        if (u && podMatch) users.add(u);
      }

      if (!users.size && wantedPod) {
        for (let i = startRow; i < rows.length; i++) {
          const u = String(rows[i][cols.user] || "");
          if (u) users.add(u);
        }
      }

      return { users: Array.from(users).sort() };
    });
  }

  async function getClientsBasedOnUser(user) {
    return Excel.run(async (ctx) => {
      const rows = await loadDatabase(ctx);
      const { cols, startRow } = detectColumns(rows);

      const target = normalizeText(user);
      const clients = [];

      for (let i = startRow; i < rows.length; i++) {
        if (normalizeText(rows[i][cols.user]) === target) {
          const c = String(rows[i][cols.client] ?? "").trim();
          if (c) clients.push(c);
        }
      }

      return Array.from(new Set(clients)).sort();
    });
  }

  async function getClientDetails(client) {
    return Excel.run(async (ctx) => {
      const rows = await loadDatabase(ctx);
      const { cols, startRow } = detectColumns(rows);

      const target = normalizeText(client);

      for (let i = startRow; i < rows.length; i++) {
        if (normalizeText(rows[i][cols.client]) === target) {
          return {
            clientId: String(rows[i][cols.clientId] ?? "").trim(),
            pod: String(rows[i][cols.pod] ?? "").trim(),
          };
        }
      }

      return { clientId: "", pod: "" };
    });
  }

  async function getSensitivity(user) {
    return Excel.run(async (ctx) => {
      const rows = await loadDatabase(ctx);
      const { cols, startRow } = detectColumns(rows);

      const target = normalizeText(user);

      for (let i = startRow; i < rows.length; i++) {
        if (normalizeText(rows[i][cols.user]) === target) {
          return String(rows[i][cols.sensitivity] ?? "").trim();
        }
      }

      return "";
    });
  }

  async function diagnose() {
    return Excel.run(async (ctx) => {
      const wb = ctx.workbook;

      const sheet = wb.worksheets.getItemOrNullObject(DB_SHEET_NAME);
      const active = wb.worksheets.getActiveWorksheet();

      sheet.load("name");
      active.load("name");

      await ctx.sync();

      const out = {
        databaseSheetFound: !sheet.isNullObject,
        activeSheet: active.name,
        rowCount: 0,
        detectedColumns: null,
        startRow: 0,
        header: [],
        sample: [],
      };

      if (sheet.isNullObject) return out;

      const used = sheet.getUsedRange();
      used.load("values");

      await ctx.sync();

      const rows = used.values || [];
      const det = detectColumns(rows);

      out.rowCount = rows.length;
      out.detectedColumns = det.cols;
      out.startRow = det.startRow;
      out.header = rows[0] || [];
      out.sample = rows.slice(det.startRow, det.startRow + 3);

      return out;
    });
  }

  async function getClientListWithPOD() {
    return Excel.run(async (ctx) => {
      const rows = await loadDatabase(ctx);
      const { cols, startRow } = detectColumns(rows);

      const seen = new Set();
      const out = [];

      for (let i = startRow; i < rows.length; i++) {
        const c = String(rows[i][cols.client] || "");
        const p = String(rows[i][cols.pod] || "POD X");

        if (c && !seen.has(c)) {
          seen.add(c);
          out.push({ client: c, pod: p });
        }
      }

      return out;
    });
  }

  async function submitForm(form) {
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRangeOrNullObject();

      used.load("rowCount");
      await ctx.sync();

      const nextRowIdx = used.isNullObject ? 0 : used.rowCount;

      const range = sheet.getRangeByIndexes(nextRowIdx, 0, 1, 17);

      range.values = [[
        false,
        form.date,
        form.softwareId,
        form.inTime,
        form.user,
        form.client,
        form.clientCode,
        form.pod,
        form.sensitivity,
        form.projectCode,
        form.jobDescription,
        form.newValue,
        form.edits,
        form.rework,
        "",
        "TBC",
        "TBC",
      ]];

      if (form.sensitivity === "High") {
        const userCell = sheet.getRangeByIndexes(nextRowIdx, 4, 1, 1);
        userCell.format.font.color = "#FF0000";
      }

      await ctx.sync();

      return { ok: true, row: nextRowIdx + 1 };
    });
  }

  const QUEUE_SHEET_NAME = "_NewClientQueue";
  const QUEUE_TABLE_NAME = "tbl_NewClientQueue";

  const QUEUE_HEADERS = [
    "QueuedAt",
    "ClientName",
    "UserName",
    "POD",
    "Sensitivity",
    "Reviewing",
    "Status",
    "ProcessedAt",
    "Error",
  ];

  async function ensureQueueTable(ctx) {
    const wb = ctx.workbook;

    let table = wb.tables.getItemOrNullObject(QUEUE_TABLE_NAME);
    table.load("name");
    await ctx.sync();

    if (!table.isNullObject) return table;

    let sheet = wb.worksheets.getItemOrNullObject(QUEUE_SHEET_NAME);
    sheet.load("name");
    await ctx.sync();

    if (sheet.isNullObject) {
      sheet = wb.worksheets.add(QUEUE_SHEET_NAME);
    }

    sheet.visibility = Excel.SheetVisibility.hidden;

    const headerRange = sheet.getRangeByIndexes(0, 0, 1, QUEUE_HEADERS.length);
    headerRange.values = [QUEUE_HEADERS];

    const tableRange = sheet.getRangeByIndexes(0, 0, 1, QUEUE_HEADERS.length);

    table = wb.tables.add(tableRange, true);
    table.name = QUEUE_TABLE_NAME;

    await ctx.sync();

    return table;
  }

  async function saveNewClient(payload) {
    return Excel.run(async (ctx) => {
      const table = await ensureQueueTable(ctx);

      table.rows.add(null, [[
        new Date().toISOString(),
        payload.clientName || "",
        payload.userName || "",
        payload.pod || "",
        payload.sensitivity || "High",
        payload.reviewing || "INC",
        "Pending",
        "",
        "",
      ]]);

      await ctx.sync();
      return { ok: true };
    });
  }

  async function appendTime(timeValue) {
    return Excel.run(async (ctx) => {
      const cell = ctx.workbook.getActiveCell();
      cell.load("values");

      await ctx.sync();

      const old = String(cell.values?.[0]?.[0] || "");
      const matches = old.match(/(\d{1,2}:\d{2})/g) || [];

      matches.push(timeValue);

      cell.values = [[matches.join(" / ")]];
      await ctx.sync();

      return matches.join(" / ");
    });
  }

  async function saveDateTime(day, hour, minute, ampm) {
    return Excel.run(async (ctx) => {
      const cell = ctx.workbook.getActiveCell();

      const formatted =
        day.substring(0, 3) +
        ", " +
        hour +
        ":" +
        String(minute).padStart(2, "0") +
        " " +
        ampm;

      cell.values = [[formatted]];
      await ctx.sync();

      return formatted;
    });
  }

  return {
    getWorkbookContext,
    getDropdownData,
    getClientsBasedOnUser,
    getClientDetails,
    getSensitivity,
    getClientListWithPOD,
    submitForm,
    saveNewClient,
    appendTime,
    saveDateTime,
    diagnose,
  };
})();