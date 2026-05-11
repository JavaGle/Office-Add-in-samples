/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, localStorage, fetch */

// ─── Constants ──────────────────────────────────────────────────────────────

const STORAGE_KEY_API_URL   = "azuresql_api_url";
const STORAGE_KEY_API_KEY   = "azuresql_api_key";
const STORAGE_KEY_BATCH_SIZE = "azuresql_batch_size";
const STORAGE_KEY_SYNC_MODE = "azuresql_sync_mode";
const DEFAULT_BATCH_SIZE    = 200;

// ─── Office Initialization ──────────────────────────────────────────────────

Office.onReady(() => {
  loadSavedConfig();

  document.getElementById("saveConfig").onclick    = saveConfig;
  document.getElementById("refreshSheets").onclick = refreshSheets;
  document.getElementById("syncButton").onclick    = startSync;
  document.getElementById("clearLog").onclick      = clearLog;
});

// ─── Configuration helpers ──────────────────────────────────────────────────

function loadSavedConfig() {
  document.getElementById("apiUrl").value    = localStorage.getItem(STORAGE_KEY_API_URL)    || "";
  document.getElementById("apiKey").value    = localStorage.getItem(STORAGE_KEY_API_KEY)    || "";
  document.getElementById("batchSize").value = localStorage.getItem(STORAGE_KEY_BATCH_SIZE) || DEFAULT_BATCH_SIZE;
  document.getElementById("syncMode").value  = localStorage.getItem(STORAGE_KEY_SYNC_MODE)  || "upsert";
}

function saveConfig() {
  const apiUrl    = document.getElementById("apiUrl").value.trim();
  const apiKey    = document.getElementById("apiKey").value.trim();
  const batchSize = parseInt(document.getElementById("batchSize").value, 10) || DEFAULT_BATCH_SIZE;
  const syncMode  = document.getElementById("syncMode").value;

  localStorage.setItem(STORAGE_KEY_API_URL,    apiUrl);
  localStorage.setItem(STORAGE_KEY_API_KEY,    apiKey);
  localStorage.setItem(STORAGE_KEY_BATCH_SIZE, batchSize);
  localStorage.setItem(STORAGE_KEY_SYNC_MODE,  syncMode);

  showStatus("Configuration saved.", "info");
  addLog("Configuration saved.");
}

// ─── Sheet List ─────────────────────────────────────────────────────────────

async function refreshSheets() {
  const sheetListEl = document.getElementById("sheetList");
  sheetListEl.innerHTML = "<p class='hint'>Loading…</p>";

  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name,items/visibility");
      await context.sync();

      // Count rows per sheet for the meta label (capped read to avoid perf hit)
      sheetListEl.innerHTML = "";
      for (const sheet of sheets.items) {
        if (sheet.visibility !== Excel.SheetVisibility.visible) continue;

        const item = document.createElement("div");
        item.className = "sheet-item";

        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.id   = `sheet_cb_${sheet.name}`;
        cb.value = sheet.name;
        cb.checked = true;

        const lbl = document.createElement("label");
        lbl.htmlFor = cb.id;
        lbl.textContent = sheet.name;

        const meta = document.createElement("span");
        meta.className = "sheet-meta";
        meta.textContent = "loading rows…";

        item.appendChild(cb);
        item.appendChild(lbl);
        item.appendChild(meta);
        sheetListEl.appendChild(item);

        // Async row count without blocking main render
        loadSheetMeta(sheet.name, meta);
      }

      if (sheetListEl.children.length === 0) {
        sheetListEl.innerHTML = "<p class='hint'>No visible worksheets found.</p>";
      }
    });
  } catch (err) {
    sheetListEl.innerHTML = "<p class='hint'>Error loading sheets.</p>";
    showStatus(`Error: ${err.message}`, "error");
  }
}

async function loadSheetMeta(sheetName, metaEl) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const usedRange = sheet.getUsedRangeOrNullObject(true);
      usedRange.load("rowCount,columnCount");
      await context.sync();

      if (usedRange.isNullObject) {
        metaEl.textContent = "(empty)";
      } else {
        // Subtract 1 for header row
        const dataRows = Math.max(0, usedRange.rowCount - 1);
        metaEl.textContent = `${dataRows} row(s), ${usedRange.columnCount} col(s)`;
      }
    });
  } catch {
    metaEl.textContent = "";
  }
}

// ─── Sync orchestration ─────────────────────────────────────────────────────

async function startSync() {
  clearStatus();
  clearLog();

  const apiUrl    = localStorage.getItem(STORAGE_KEY_API_URL)    || "";
  const apiKey    = localStorage.getItem(STORAGE_KEY_API_KEY)    || "";
  const batchSize = parseInt(localStorage.getItem(STORAGE_KEY_BATCH_SIZE) || DEFAULT_BATCH_SIZE, 10);
  const syncMode  = localStorage.getItem(STORAGE_KEY_SYNC_MODE)  || "upsert";

  if (!apiUrl) {
    showStatus("Please enter and save the API Endpoint URL before syncing.", "error");
    return;
  }

  // Collect selected sheets
  const selectedSheets = getSelectedSheets();
  if (selectedSheets.length === 0) {
    showStatus("Please select at least one sheet to sync.", "error");
    return;
  }

  setButtonsDisabled(true);
  showProgress(true);
  addLog(`Starting sync → ${apiUrl}`);
  addLog(`Mode: ${syncMode} | Batch size: ${batchSize} rows`);

  let totalBatches  = 0;
  let doneBatches   = 0;
  let hasError      = false;
  const sheetSummary = [];

  try {
    // Phase 1: read all selected sheets from Excel
    const sheetData = await readSheets(selectedSheets);

    // Phase 2: count total batches across all sheets for progress calculation
    for (const { name, headers, rows } of sheetData) {
      const batches = Math.ceil(rows.length / batchSize);
      totalBatches += batches || 1; // at least 1 (even for empty → still send schema)
      addLog(`Sheet "${name}": ${rows.length} data row(s), ${headers.length} col(s)`);
    }

    setProgress(0, `0 / ${totalBatches} batches`);

    // Phase 3: send batches for each sheet
    for (const { name, headers, rows } of sheetData) {
      addLog(`\nSyncing sheet: "${name}"…`);
      const result = await syncSheet({ name, headers, rows, apiUrl, apiKey, batchSize, syncMode, onBatch: () => {
        doneBatches++;
        const pct = Math.round((doneBatches / totalBatches) * 100);
        setProgress(pct, `${doneBatches} / ${totalBatches} batches`);
      }});

      sheetSummary.push(`"${name}": ${result.rowsSynced} row(s) synced`);
      if (result.error) {
        addLog(`  ✗ Error: ${result.error}`, "err");
        hasError = true;
      } else {
        addLog(`  ✓ Done — ${result.rowsSynced} row(s) synced.`, "ok");
      }
    }

    setProgress(100, "Complete");

    const summaryMsg = sheetSummary.join("; ");
    if (hasError) {
      showStatus(`Sync completed with errors. ${summaryMsg}`, "error");
    } else {
      showStatus(`Sync successful! ${summaryMsg}`, "success");
    }

  } catch (err) {
    showStatus(`Sync failed: ${err.message}`, "error");
    addLog(`Fatal error: ${err.message}`, "err");
    setProgress(0, "Failed");
  } finally {
    setButtonsDisabled(false);
  }
}

// ─── Excel data reading ─────────────────────────────────────────────────────

/**
 * Read all selected sheets. Returns an array of { name, headers, rows } objects.
 * Reads each sheet's used range in one Excel.run call to minimise round-trips.
 * For sheets with many columns the used range is loaded as a flat 2D array
 * (Excel.Range.values) which is the most efficient approach.
 */
async function readSheets(sheetNames) {
  const result = [];

  for (const name of sheetNames) {
    addLog(`Reading sheet "${name}"…`);
    const data = await readSheet(name);
    result.push(data);
  }

  return result;
}

async function readSheet(sheetName) {
  return Excel.run(async (context) => {
    const sheet    = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRangeOrNullObject(true);

    // Load everything in one round-trip: values + dimensions.
    usedRange.load("values,rowCount,columnCount");
    await context.sync();

    if (usedRange.isNullObject || usedRange.rowCount < 1) {
      return { name: sheetName, headers: [], rows: [] };
    }

    const allValues = usedRange.values; // 2D array [row][col]
    const headers   = allValues[0].map((h) => (h == null || h === "" ? `Col${allValues[0].indexOf(h)}` : String(h)));

    // Build row objects. Empty trailing rows are skipped for cleanliness.
    const rows = [];
    for (let r = 1; r < allValues.length; r++) {
      const rowArr = allValues[r];
      // Skip completely blank rows
      if (rowArr.every((cell) => cell === null || cell === "")) continue;
      const rowObj = {};
      for (let c = 0; c < headers.length; c++) {
        rowObj[headers[c]] = rowArr[c] !== undefined ? rowArr[c] : null;
      }
      rows.push(rowObj);
    }

    return { name: sheetName, headers, rows };
  });
}

// ─── HTTP sync ──────────────────────────────────────────────────────────────

/**
 * Send rows for one sheet to the backend API in batches.
 * Returns { rowsSynced, error }.
 */
async function syncSheet({ name, headers, rows, apiUrl, apiKey, batchSize, syncMode, onBatch }) {
  let rowsSynced = 0;
  const tableName = sanitizeTableName(name);

  // If no rows, still send a schema-only request so the backend can ensure the table exists.
  const batches = rows.length === 0 ? [[]] : chunkArray(rows, batchSize);

  for (let i = 0; i < batches.length; i++) {
    const batch = batches[i];
    const isFirst = i === 0;

    const payload = {
      tableName,
      headers,
      rows: batch,
      syncMode: isFirst ? syncMode : "insert", // only first batch carries mode (replace only on first)
      isFirstBatch: isFirst,
      isLastBatch:  i === batches.length - 1,
    };

    try {
      const response = await fetchWithRetry(apiUrl, {
        method:  "POST",
        headers: buildRequestHeaders(apiKey),
        body:    JSON.stringify(payload),
      });

      if (!response.ok) {
        const errText = await response.text();
        return { rowsSynced, error: `HTTP ${response.status}: ${errText}` };
      }

      const json = await response.json();
      rowsSynced += json.rowsAffected || batch.length;

    } catch (fetchErr) {
      return { rowsSynced, error: fetchErr.message };
    }

    onBatch();
    addLog(`  Batch ${i + 1}/${batches.length} sent (${batch.length} rows)`);
  }

  return { rowsSynced, error: null };
}

/**
 * Fetch with one automatic retry on network failure.
 */
async function fetchWithRetry(url, options, retries = 1) {
  try {
    return await fetch(url, options);
  } catch (err) {
    if (retries > 0) {
      addLog(`  Network error, retrying… (${err.message})`, "warn");
      await sleep(1000);
      return fetchWithRetry(url, options, retries - 1);
    }
    throw err;
  }
}

function buildRequestHeaders(apiKey) {
  const headers = { "Content-Type": "application/json" };
  if (apiKey) {
    headers["x-api-key"] = apiKey;
  }
  return headers;
}

// ─── Utilities ──────────────────────────────────────────────────────────────

function getSelectedSheets() {
  const checkboxes = document.querySelectorAll("#sheetList input[type=checkbox]:checked");
  return Array.from(checkboxes).map((cb) => cb.value);
}

/** Split an array into chunks of the given size. */
function chunkArray(arr, size) {
  const chunks = [];
  for (let i = 0; i < arr.length; i += size) {
    chunks.push(arr.slice(i, i + size));
  }
  return chunks;
}

/** Sanitize a sheet name so it is safe as a SQL table name. */
function sanitizeTableName(name) {
  return name.replace(/[^a-zA-Z0-9_]/g, "_").replace(/^([0-9])/, "_$1");
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ─── UI helpers ─────────────────────────────────────────────────────────────

function showStatus(message, type) {
  const area = document.getElementById("statusArea");
  area.innerHTML = "";
  const card = document.createElement("div");
  card.className = `status-card ${type}`;
  card.textContent = message;
  area.appendChild(card);
}

function clearStatus() {
  document.getElementById("statusArea").innerHTML = "";
}

function addLog(text, level) {
  const logArea = document.getElementById("logArea");
  const line = document.createElement("p");
  line.className = `log-line${level ? " log-" + level : ""}`;
  const ts = new Date().toLocaleTimeString();
  line.textContent = `[${ts}] ${text}`;
  logArea.appendChild(line);
  logArea.scrollTop = logArea.scrollHeight;
}

function clearLog() {
  document.getElementById("logArea").innerHTML = "";
}

function setButtonsDisabled(disabled) {
  document.getElementById("syncButton").disabled   = disabled;
  document.getElementById("refreshSheets").disabled = disabled;
  document.getElementById("saveConfig").disabled   = disabled;
}

function showProgress(visible) {
  const area = document.getElementById("progressArea");
  if (visible) {
    area.classList.remove("hidden");
  } else {
    area.classList.add("hidden");
  }
}

function setProgress(pct, label) {
  document.getElementById("progressBar").style.width = `${pct}%`;
  document.getElementById("progressText").textContent = label;
  document.getElementById("progressPct").textContent  = `${pct}%`;
}
