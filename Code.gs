/**
 * GOOGLE DRIVE OWNERSHIP AUDITOR
 * Version 3.1
 *
 * Recursive, resumable, parallel audit of who owns every item
 * under a Google Drive folder.
 *
 * Source Drive: READ-ONLY
 * Report: Google Sheets
 *
 * Setup, usage and troubleshooting: see README.md.
 */

const CONFIG = {
  // --- Required: replace these placeholders ---
  ROOT_FOLDER_ID: "YOUR_ROOT_FOLDER_ID",  // Folder to audit.
  MY_EMAIL: "you@example.com",            // Your account.
  INTERNAL_DOMAINS: ["example.com"],      // Your organisation's domains.

  PAGE_SIZE: 1000,
  CONCURRENCY: 25,          // Parallel Drive API requests per wave.
  MAX_RUNTIME_MS: 280000,   // Stop starting new waves after this.
  FLUSH_EVERY_MS: 60000,    // Save progress to the sheet this often.
  MAX_RETRIES: 5,           // Per folder, per run, for 429/5xx.

  CHECK_SHORTCUT_TARGETS: true,
  AUTO_RESUME: true,
  RESUME_EVERY_MINUTES: 1,

  REPORT_NAME: "Google Drive - External Ownership Audit"
};

const MIME_FOLDER = "application/vnd.google-apps.folder";
const MIME_SHORTCUT = "application/vnd.google-apps.shortcut";
const STATE_KEY = "OWNERSHIP_AUDIT_REPORT_ID";
const DRIVE_API = "https://www.googleapis.com/drive/v3/files";

const LIST_FIELDS =
  "nextPageToken,incompleteSearch,files(" +
  "id,name,mimeType,owners(emailAddress)," +
  "driveId,webViewLink," +
  "shortcutDetails(targetId,targetMimeType))";

const AUDIT_HEADERS = [
  "Path", "Name", "Type", "Ownership Status",
  "Owner Email", "Owner Domain", "File ID", "File URL",
  "Shortcut Target ID", "Shortcut Target Owner",
  "Shortcut Target Status", "Remarks"
];

const QUEUE_HEADERS = [
  "Folder ID", "Folder Path", "Page Token", "Status"
];

/**
 * START A NEW AUDIT.
 * Run this function only once per audit.
 */
function startOwnershipAudit() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error("Another audit execution is running.");
  }

  try {
    assertConfigured();

    const props = PropertiesService.getScriptProperties();

    if (props.getProperty(STATE_KEY)) {
      throw new Error(
        "An audit already exists. Run resumeOwnershipAudit(). " +
        "To start a new report, run resetOwnershipAudit() first."
      );
    }

    const root = Drive.Files.get(CONFIG.ROOT_FOLDER_ID, {
      fields: "id,name,mimeType,owners(emailAddress),driveId,webViewLink",
      supportsAllDrives: true
    });

    if (root.mimeType !== MIME_FOLDER) {
      throw new Error("ROOT_FOLDER_ID must be a folder.");
    }

    const ss = SpreadsheetApp.create(
      CONFIG.REPORT_NAME + " - " +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm"
      )
    );

    const audit = ss.getActiveSheet();
    audit.setName("Ownership Audit");
    audit.getRange(1, 1, 1, AUDIT_HEADERS.length)
      .setValues([AUDIT_HEADERS]);
    styleHeader(audit, AUDIT_HEADERS.length);

    const queue = ss.insertSheet("Scan Queue");
    queue.getRange(1, 1, 1, QUEUE_HEADERS.length)
      .setValues([QUEUE_HEADERS]);
    styleHeader(queue, QUEUE_HEADERS.length);

    appendRows(queue, [[root.id, root.name, "", "PENDING"]]);

    // Include the root folder itself.
    appendRows(audit, [makeAuditRow(root, root.name, null)]);

    const summary = ss.insertSheet("Summary");
    summary.getRange("A1:B1")
      .setValues([["Ownership Audit", ""]]);
    summary.getRange("A1:B1")
      .setBackground("#17365D")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");

    summary.getRange("A2:B5").setValues([
      ["Root Folder", root.name],
      ["Started At", new Date()],
      ["Status", "RUNNING"],
      ["Report URL", ss.getUrl()]
    ]);

    summary.getRange("A7:B7")
      .setValues([["Ownership Status", "Count"]]);

    const statuses = [
      "OWNED BY ME",
      "OTHER ACCOUNT",
      "EXTERNAL DOMAIN",
      "SHARED DRIVE",
      "UNKNOWN OWNER",
      "SCAN INCOMPLETE"
    ];

    statuses.forEach((status, i) => {
      const row = i + 8;
      summary.getRange(row, 1).setValue(status);
      summary.getRange(row, 2).setFormula(
        '=COUNTIF(\'Ownership Audit\'!D:D,A' + row + ')'
      );
    });

    summary.getRange("A16:B19").setValues([
      ["Scanned Records", ""],
      ["Completed Folders", ""],
      ["Pending Folders", ""],
      ["Error Folders", ""]
    ]);

    summary.getRange("B16").setFormula(
      "=COUNTA('Ownership Audit'!G:G)-1"
    );
    summary.getRange("B17").setFormula(
      '=COUNTIF(\'Scan Queue\'!D:D,"DONE")'
    );
    summary.getRange("B18").setFormula(
      '=COUNTIF(\'Scan Queue\'!D:D,"PENDING")'
    );
    summary.getRange("B19").setFormula(
      '=COUNTIF(\'Scan Queue\'!D:D,"ERROR")'
    );

    summary.autoResizeColumns(1, 2);

    props.setProperty(STATE_KEY, ss.getId());

    if (CONFIG.AUTO_RESUME) installResumeTrigger();

    console.log("Report: " + ss.getUrl());

  } finally {
    lock.releaseLock();
  }

  // Start first batch immediately.
  resumeOwnershipAudit();
}

/**
 * PROCESS THE NEXT BATCH.
 * Safe to run manually or from a time trigger.
 */
function resumeOwnershipAudit() {
  const lock = LockService.getScriptLock();

  // Avoid two simultaneous executions.
  if (!lock.tryLock(1000)) {
    console.log("Audit already running. Skipping.");
    return;
  }

  try {
    const startTime = Date.now();
    const ss = openReport();
    const audit = ss.getSheetByName("Ownership Audit");
    const queue = ss.getSheetByName("Scan Queue");
    const summary = ss.getSheetByName("Summary");
    const token = ScriptApp.getOAuthToken();

    // Load the whole queue ONCE. qData[i] is sheet row i + 2.
    const lastRow = queue.getLastRow();
    const qData = lastRow >= 2
      ? queue.getRange(2, 1, lastRow - 1, 4).getValues()
        .map(r => r.map(v => String(v)))
      : [];
    let rowsInSheet = qData.length;

    const pending = [];
    qData.forEach((r, i) => {
      if (r[3] === "PENDING") pending.push(i);
    });

    let auditBuf = [];
    let lastFlush = Date.now();
    let processed = 0;
    let waves = 0;
    let throttleStreak = 0;
    let throttleCount = 0;
    let shortcutLookups = 0;
    const timing = {
      load: Date.now() - startTime,
      list: 0, shortcuts: 0, sleep: 0, flush: 0
    };
    const retries = {};
    const targetCache = new Map();

    const flush = () => {
      const t0 = Date.now();
      // Audit rows first: if the run dies mid-flush we may get
      // duplicate rows (dedupe by File ID), never missing rows.
      appendRows(audit, auditBuf);
      auditBuf = [];

      if (rowsInSheet > 0) {
        queue.getRange(2, 3, rowsInSheet, 2).setValues(
          qData.slice(0, rowsInSheet)
            .map(r => [safeCell(r[2]), safeCell(r[3])])
        );
      }
      if (qData.length > rowsInSheet) {
        appendRows(queue, qData.slice(rowsInSheet));
        rowsInSheet = qData.length;
      }

      SpreadsheetApp.flush();
      lastFlush = Date.now();
      timing.flush += lastFlush - t0;
    };

    const markError = (i, message) => {
      const [folderId, folderPath] = qData[i];
      qData[i][3] = "ERROR";
      console.error("Folder error: " + folderPath + " - " + message);
      auditBuf.push([
        folderPath,
        folderPath.split("/").pop(),
        "SCAN ERROR",
        "SCAN INCOMPLETE",
        "", "",
        folderId,
        "https://drive.google.com/drive/folders/" + folderId,
        "", "", "",
        message
      ]);
    };

    while (
      pending.length &&
      Date.now() - startTime < CONFIG.MAX_RUNTIME_MS
    ) {
      const batch = pending.splice(0, CONFIG.CONCURRENCY);
      let t0 = Date.now();
      const responses = UrlFetchApp.fetchAll(
        batch.map(i => buildListRequest(qData[i][0], qData[i][2], token))
      );
      waves++;
      timing.list += Date.now() - t0;

      const ok = [];
      const shortcutIds = new Set();
      let throttled = false;

      responses.forEach((res, k) => {
        const i = batch[k];
        const code = res.getResponseCode();
        const text = res.getContentText();

        if (code === 200) {
          const body = JSON.parse(text);
          if (body.incompleteSearch) {
            markError(i, "Drive API returned an incomplete search.");
            return;
          }
          const files = body.files || [];
          files.forEach(f => {
            if (
              f.mimeType === MIME_SHORTCUT &&
              f.shortcutDetails && f.shortcutDetails.targetId
            ) {
              shortcutIds.add(f.shortcutDetails.targetId);
            }
          });
          ok.push({ i, files, next: body.nextPageToken || "" });

        } else if (isRetryable(code, text)) {
          throttled = true;
          throttleCount++;
          retries[i] = (retries[i] || 0) + 1;
          if (retries[i] > CONFIG.MAX_RETRIES) {
            markError(i, "HTTP " + code + ": " + errorMessage(text));
          } else {
            pending.push(i); // Try again later in this run.
          }

        } else {
          markError(i, "HTTP " + code + ": " + errorMessage(text));
        }
      });

      if (CONFIG.CHECK_SHORTCUT_TARGETS && shortcutIds.size) {
        t0 = Date.now();
        const before = targetCache.size;
        fetchShortcutTargets([...shortcutIds], token, targetCache);
        shortcutLookups += targetCache.size - before;
        timing.shortcuts += Date.now() - t0;
      }

      ok.forEach(({ i, files, next }) => {
        const folderPath = qData[i][1];

        files.forEach(file => {
          const path = folderPath + "/" + file.name;
          auditBuf.push(makeAuditRow(file, path, targetCache));

          if (file.mimeType === MIME_FOLDER) {
            qData.push([file.id, path, "", "PENDING"]);
            pending.push(qData.length - 1);
          }
        });

        if (next) {
          qData[i][2] = next;       // Checkpoint: next page.
          pending.push(i);
        } else {
          qData[i][2] = "";
          qData[i][3] = "DONE";
        }

        processed += files.length;
      });

      if (throttled) {
        throttleStreak = Math.min(throttleStreak + 1, 5);
        Utilities.sleep(1000 * throttleStreak);
        timing.sleep += 1000 * throttleStreak;
      } else {
        throttleStreak = 0;
      }

      if (Date.now() - lastFlush > CONFIG.FLUSH_EVERY_MS) flush();
    }

    flush();

    const count = status =>
      qData.filter(r => r[3] === status).length;
    const pendingCount = count("PENDING");
    const errors = count("ERROR");

    if (pendingCount === 0) {
      const finalStatus = errors > 0
        ? "COMPLETED WITH ERRORS"
        : "COMPLETED";

      summary.getRange("B4").setValue(finalStatus);
      summary.getRange("A21:B21")
        .setValues([["Completed At", new Date()]]);

      removeResumeTrigger();

      // Format once at the end.
      finishReport(audit);

      console.log("Audit finished: " + finalStatus);

    } else {
      summary.getRange("B4").setValue("RUNNING");
    }

    console.log(
      "Run: " + Math.round((Date.now() - startTime) / 1000) + "s" +
      " | Waves: " + waves +
      " | Items: " + processed +
      " | Queue: " + qData.length +
      " | Pending folders: " + pendingCount +
      " | Errors: " + errors
    );
    const sec = ms => Math.round(ms / 1000) + "s";
    console.log(
      "Timing: load " + sec(timing.load) +
      " | list " + sec(timing.list) +
      " | shortcuts " + sec(timing.shortcuts) +
      " (" + shortcutLookups + " lookups)" +
      " | throttle sleep " + sec(timing.sleep) +
      " (" + throttleCount + " throttled)" +
      " | flush " + sec(timing.flush)
    );
    console.log("Report: " + ss.getUrl());

  } finally {
    lock.releaseLock();
  }
}

/**
 * VIEW AUDIT PROGRESS.
 */
function checkOwnershipAuditStatus() {
  const reportId = PropertiesService
    .getScriptProperties()
    .getProperty(STATE_KEY);

  if (!reportId) {
    console.log("No active audit.");
    return;
  }

  const ss = SpreadsheetApp.openById(reportId);
  const queue = ss.getSheetByName("Scan Queue");
  const audit = ss.getSheetByName("Ownership Audit");

  const states = queue.getRange(
    2, 4, Math.max(queue.getLastRow() - 1, 1), 1
  ).getValues().flat();

  const count = status =>
    states.filter(s => s === status).length;

  console.log(JSON.stringify({
    scannedRecords: audit.getLastRow() - 1,
    completedFolders: count("DONE"),
    pendingFolders: count("PENDING"),
    errorFolders: count("ERROR"),
    reportUrl: ss.getUrl()
  }, null, 2));
}

/**
 * RE-QUEUE FOLDERS THAT FAILED.
 * Their saved page token is kept, so scanning resumes at the
 * failed page. The old "SCAN ERROR" rows stay in the report.
 */
function retryErrorFolders() {
  const ss = openReport();
  const queue = ss.getSheetByName("Scan Queue");
  const n = queue.getLastRow() - 1;
  if (n < 1) return;

  const range = queue.getRange(2, 4, n, 1);
  const values = range.getValues();
  let requeued = 0;

  values.forEach(r => {
    if (r[0] === "ERROR") {
      r[0] = "PENDING";
      requeued++;
    }
  });

  range.setValues(values);
  ss.getSheetByName("Summary").getRange("B4").setValue("RUNNING");

  if (CONFIG.AUTO_RESUME) installResumeTrigger();

  console.log("Re-queued " + requeued + " folders.");
}

/**
 * RESET CONTROL STATE.
 *
 * Does NOT delete the existing spreadsheet.
 * Required before starting a completely new audit.
 */
function resetOwnershipAudit() {
  removeResumeTrigger();

  PropertiesService.getScriptProperties()
    .deleteProperty(STATE_KEY);

  console.log(
    "Audit state reset. Previous report remains in Drive."
  );
}

/**
 * SET UP AUTOMATIC RESUMPTION.
 */
function installResumeTrigger() {
  removeResumeTrigger();

  ScriptApp.newTrigger("resumeOwnershipAudit")
    .timeBased()
    .everyMinutes(CONFIG.RESUME_EVERY_MINUTES)
    .create();
}

function removeResumeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(
      t => t.getHandlerFunction() === "resumeOwnershipAudit"
    )
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function assertConfigured() {
  const placeholders = [
    CONFIG.ROOT_FOLDER_ID === "YOUR_ROOT_FOLDER_ID" && "ROOT_FOLDER_ID",
    CONFIG.MY_EMAIL === "you@example.com" && "MY_EMAIL",
    CONFIG.INTERNAL_DOMAINS.includes("example.com") && "INTERNAL_DOMAINS"
  ].filter(Boolean);

  if (placeholders.length) {
    throw new Error(
      "Set these in CONFIG before starting: " + placeholders.join(", ")
    );
  }
}

function openReport() {
  const reportId = PropertiesService
    .getScriptProperties()
    .getProperty(STATE_KEY);

  if (!reportId) {
    throw new Error(
      "No active audit. Run startOwnershipAudit()."
    );
  }

  return SpreadsheetApp.openById(reportId);
}

/**
 * DRIVE REST HELPERS (parallel via UrlFetchApp.fetchAll).
 */
function buildListRequest(folderId, pageToken, token) {
  const params = {
    q: "'" + folderId + "' in parents and trashed = false",
    fields: LIST_FIELDS,
    pageSize: CONFIG.PAGE_SIZE,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  };

  if (pageToken) params.pageToken = pageToken;

  return {
    url: DRIVE_API + "?" + toQueryString(params),
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  };
}

/**
 * Looks up shortcut targets in parallel. Results are cached in
 * `cache` (targetId -> {ok, owner} or {ok: false, error}).
 */
function fetchShortcutTargets(ids, token, cache) {
  let todo = ids.filter(id => !cache.has(id));
  const query = toQueryString({
    fields: "id,owners(emailAddress),driveId",
    supportsAllDrives: true
  });

  for (let attempt = 0; todo.length && attempt < 3; attempt++) {
    const retry = [];

    for (let s = 0; s < todo.length; s += CONFIG.CONCURRENCY) {
      const chunk = todo.slice(s, s + CONFIG.CONCURRENCY);
      const responses = UrlFetchApp.fetchAll(chunk.map(id => ({
        url: DRIVE_API + "/" + encodeURIComponent(id) + "?" + query,
        headers: { Authorization: "Bearer " + token },
        muteHttpExceptions: true
      })));

      responses.forEach((res, k) => {
        const id = chunk[k];
        const code = res.getResponseCode();
        const text = res.getContentText();

        if (code === 200) {
          cache.set(id, {
            ok: true,
            owner: classifyOwner(JSON.parse(text))
          });
        } else if (isRetryable(code, text) && attempt < 2) {
          retry.push(id);
        } else {
          cache.set(id, {
            ok: false,
            error: "HTTP " + code + ": " + errorMessage(text)
          });
        }
      });
    }

    todo = retry;
    if (todo.length) Utilities.sleep(2000 * (attempt + 1));
  }
}

function isRetryable(code, text) {
  return code === 429 || code >= 500 ||
    (code === 403 &&
      /rateLimitExceeded|userRateLimitExceeded|backendError/
        .test(text));
}

function errorMessage(text) {
  try {
    return JSON.parse(text).error.message;
  } catch (e) {
    return String(text).slice(0, 200);
  }
}

function toQueryString(params) {
  return Object.keys(params)
    .map(k =>
      encodeURIComponent(k) + "=" + encodeURIComponent(params[k])
    )
    .join("&");
}

/**
 * OWNERSHIP CLASSIFICATION.
 */
function classifyOwner(file) {
  if (file.driveId) {
    return {
      status: "SHARED DRIVE",
      email: "",
      domain: ""
    };
  }

  const owners = file.owners || [];

  if (!owners.length || !owners[0].emailAddress) {
    return {
      status: "UNKNOWN OWNER",
      email: "",
      domain: ""
    };
  }

  const email = owners[0].emailAddress
    .toLowerCase().trim();

  const domain = email.split("@")[1] || "";

  let status = "EXTERNAL DOMAIN";

  if (email === CONFIG.MY_EMAIL.toLowerCase()) {
    status = "OWNED BY ME";

  } else if (
    CONFIG.INTERNAL_DOMAINS
      .map(d => d.toLowerCase())
      .includes(domain)
  ) {
    status = "OTHER ACCOUNT";
  }

  return { status, email, domain };
}

/**
 * CONSTRUCT OWNERSHIP AUDIT ROW.
 * Shortcut targets come from `targetCache`, filled beforehand by
 * fetchShortcutTargets (no API calls in here).
 */
function makeAuditRow(file, path, targetCache) {
  const owner = classifyOwner(file);

  let type = "FILE";

  if (file.mimeType === MIME_FOLDER) {
    type = "FOLDER";
  } else if (file.mimeType === MIME_SHORTCUT) {
    type = "SHORTCUT";
  }

  let targetId = "";
  let targetOwner = "";
  let targetStatus = "";
  let remarks = "";

  if (type === "SHORTCUT") {
    targetId = (file.shortcutDetails || {}).targetId || "";

    if (CONFIG.CHECK_SHORTCUT_TARGETS && targetId) {
      const target = targetCache && targetCache.get(targetId);

      if (target && target.ok) {
        targetOwner = target.owner.email;
        targetStatus = target.owner.status;

        if (target.owner.status === "EXTERNAL DOMAIN") {
          remarks = "Externally owned shortcut target";
        }
      } else {
        targetStatus = "TARGET UNVERIFIED";
        remarks = target ? target.error : "Target not looked up";
      }
    }
  }

  const url = file.webViewLink ||
    (type === "FOLDER"
      ? "https://drive.google.com/drive/folders/" + file.id
      : "https://drive.google.com/open?id=" + file.id);

  return [
    path, file.name, type,
    owner.status, owner.email, owner.domain,
    file.id, url,
    targetId, targetOwner, targetStatus, remarks
  ];
}

/**
 * APPEND MULTIPLE ROWS IN ONE WRITE.
 */
function appendRows(sheet, rows) {
  if (!rows.length) return;

  const safeRows = rows.map(row => row.map(safeCell));
  const firstRow = sheet.getLastRow() + 1;
  const needed = firstRow + safeRows.length - 1;
  const maxRows = sheet.getMaxRows();

  if (needed > maxRows) {
    sheet.insertRowsAfter(maxRows, needed - maxRows);
  }

  sheet.getRange(
    firstRow,
    1,
    safeRows.length,
    safeRows[0].length
  ).setValues(safeRows);
}

// Prevent spreadsheet formula injection.
function safeCell(value) {
  const s = String(value == null ? "" : value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/**
 * REPORT FORMATTING.
 */
function styleHeader(sheet, columns) {
  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, columns)
    .setBackground("#17365D")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");
}

function finishReport(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, lastRow, AUDIT_HEADERS.length)
      .createFilter();
  }

  const statusRange = sheet.getRange(2, 4, lastRow - 1, 1);

  const colors = {
    "EXTERNAL DOMAIN": "#F4CCCC",
    "OTHER ACCOUNT": "#FFF2CC",
    "UNKNOWN OWNER": "#FCE5CD",
    "SCAN INCOMPLETE": "#F4CCCC"
  };

  const rules = Object.entries(colors).map(
    ([status, color]) =>
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status)
        .setBackground(color)
        .setRanges([statusRange])
        .build()
  );

  sheet.setConditionalFormatRules(rules);

  sheet.setColumnWidth(1, 400);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 250);
  sheet.setColumnWidth(8, 300);
}
