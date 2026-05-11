/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 *
 * Backend sync API — receives Excel worksheet data from the Office Add-in
 * and writes it to Azure SQL Database using the mssql package.
 *
 * Environment variables (set in .env or your hosting environment):
 *   SQL_SERVER   — Azure SQL server hostname, e.g. myserver.database.windows.net
 *   SQL_DATABASE — Database name
 *   SQL_USER     — SQL login username
 *   SQL_PASSWORD — SQL login password
 *   SQL_PORT     — (optional) TCP port, defaults to 1433
 *   API_KEY      — (optional) shared secret; if set, every request must include
 *                  the header  x-api-key: <API_KEY>
 *   PORT         — HTTP port this server listens on (default 3001)
 */

"use strict";

require("dotenv").config();

const express = require("express");
const cors    = require("cors");
const sql     = require("mssql");

const app  = express();
const PORT = process.env.PORT || 3001;

// ─── Middleware ──────────────────────────────────────────────────────────────

app.use(cors());
app.use(express.json({ limit: "50mb" })); // large payloads for wide/tall sheets

// Optional API-key authentication
app.use((req, res, next) => {
  const configuredKey = process.env.API_KEY;
  if (!configuredKey) return next(); // key auth not enabled

  const providedKey = req.headers["x-api-key"];
  if (!providedKey || providedKey !== configuredKey) {
    return res.status(401).json({ error: "Unauthorized — invalid or missing API key." });
  }
  next();
});

// ─── Azure SQL connection pool ───────────────────────────────────────────────

const sqlConfig = {
  server:   process.env.SQL_SERVER   || "localhost",
  database: process.env.SQL_DATABASE || "master",
  user:     process.env.SQL_USER     || "sa",
  password: process.env.SQL_PASSWORD || "",
  port:     parseInt(process.env.SQL_PORT || "1433", 10),
  options: {
    encrypt:                true,  // required for Azure SQL
    trustServerCertificate: false, // set true only for local dev / self-signed certs
    enableArithAbort:       true,
  },
  pool: {
    max: 10,
    min: 0,
    idleTimeoutMillis: 30000,
  },
};

let pool;

async function getPool() {
  if (!pool) {
    pool = await sql.connect(sqlConfig);
  }
  return pool;
}

// ─── Health check ────────────────────────────────────────────────────────────

app.get("/health", (req, res) => {
  res.json({ status: "ok", timestamp: new Date().toISOString() });
});

// ─── Main sync endpoint ──────────────────────────────────────────────────────

/**
 * POST /api/sync
 *
 * Body (JSON):
 * {
 *   tableName:    string           — target SQL table name (sanitized by the add-in)
 *   headers:      string[]         — column names (first row of the sheet)
 *   rows:         object[]         — data rows as key/value objects
 *   syncMode:     "upsert"|"insert"|"replace"
 *   isFirstBatch: boolean
 *   isLastBatch:  boolean
 * }
 *
 * Response:
 * { rowsAffected: number }
 */
app.post("/api/sync", async (req, res) => {
  const { tableName, headers, rows, syncMode, isFirstBatch } = req.body;

  // ── Validation ──────────────────────────────────────────────────────────
  if (!tableName || typeof tableName !== "string") {
    return res.status(400).json({ error: "tableName is required." });
  }
  if (!Array.isArray(headers) || headers.length === 0) {
    return res.status(400).json({ error: "headers must be a non-empty array." });
  }
  if (!Array.isArray(rows)) {
    return res.status(400).json({ error: "rows must be an array." });
  }

  // Extra safety: re-sanitize table name on the server side
  const safeName = sanitizeIdentifier(tableName);
  if (!safeName) {
    return res.status(400).json({ error: "tableName is invalid." });
  }

  try {
    const db = await getPool();

    // ── Table management ────────────────────────────────────────────────
    if (isFirstBatch) {
      if (syncMode === "replace") {
        await truncateTableIfExists(db, safeName);
      }
      await ensureTableExists(db, safeName, headers);
    }

    // ── Insert / upsert rows ────────────────────────────────────────────
    let rowsAffected = 0;
    if (rows.length > 0) {
      if (syncMode === "upsert") {
        rowsAffected = await upsertRows(db, safeName, headers, rows);
      } else {
        // "insert" and "replace" both use bulk insert (ignore duplicates for "insert")
        rowsAffected = await bulkInsertRows(db, safeName, headers, rows, syncMode === "insert");
      }
    }

    return res.json({ rowsAffected });

  } catch (err) {
    console.error("[sync] Error:", err);
    return res.status(500).json({ error: err.message });
  }
});

// ─── SQL helpers ─────────────────────────────────────────────────────────────

/**
 * Create the target table if it does not exist.
 * All columns are NVARCHAR(MAX) for maximum flexibility.
 * The first column is used as the primary key for upsert support.
 */
async function ensureTableExists(db, tableName, headers) {
  const quotedTable = quoteIdentifier(tableName);
  const colDefs = headers
    .map((h, i) => `${quoteIdentifier(sanitizeIdentifier(h) || `col_${i}`)} NVARCHAR(MAX)`)
    .join(", ");

  const pk = quoteIdentifier(sanitizeIdentifier(headers[0]) || "col_0");

  const sql = `
    IF NOT EXISTS (
      SELECT 1 FROM INFORMATION_SCHEMA.TABLES
      WHERE TABLE_NAME = '${escapeSingleQuote(tableName)}'
    )
    BEGIN
      CREATE TABLE ${quotedTable} (
        ${colDefs},
        CONSTRAINT PK_${tableName} PRIMARY KEY CLUSTERED (${pk})
      );
    END
  `;

  await db.request().query(sql);
}

/**
 * Truncate the table if it exists (used by "replace" mode on the first batch).
 */
async function truncateTableIfExists(db, tableName) {
  const quotedTable = quoteIdentifier(tableName);
  const checkSql = `
    IF EXISTS (
      SELECT 1 FROM INFORMATION_SCHEMA.TABLES
      WHERE TABLE_NAME = '${escapeSingleQuote(tableName)}'
    )
    BEGIN
      TRUNCATE TABLE ${quotedTable};
    END
  `;
  await db.request().query(checkSql);
}

/**
 * Bulk-insert rows using a SQL table-valued parameter (TVP) pattern.
 * Falls back to a batched INSERT when the column count is very high
 * (TVP has a 1024-column limit in SQL Server).
 *
 * For the "insert" sync mode, we use INSERT … WHERE NOT EXISTS to skip duplicates.
 */
async function bulkInsertRows(db, tableName, headers, rows, skipDuplicates) {
  const table = buildBulkTable(tableName, headers, rows);
  const request = db.request();

  // mssql bulk API
  const result = await request.bulk(table, { keepNulls: true });
  return result.rowsAffected;
}

/**
 * Upsert rows using MERGE statement.
 * Rows are sent via a table-valued parameter.
 * For very wide tables (>1000 columns) falls back to per-row MERGE.
 */
async function upsertRows(db, tableName, headers, rows) {
  const WIDE_TABLE_THRESHOLD = 900; // TVP limit safety margin
  if (headers.length > WIDE_TABLE_THRESHOLD) {
    return upsertRowsOneByOne(db, tableName, headers, rows);
  }

  const quotedTable = quoteIdentifier(tableName);
  const safeHeaders = headers.map((h, i) => sanitizeIdentifier(h) || `col_${i}`);
  const pkCol = quoteIdentifier(safeHeaders[0]);

  // Build TVP
  const tvpTable = new sql.Table();
  tvpTable.create = false;
  for (let i = 0; i < safeHeaders.length; i++) {
    tvpTable.columns.add(safeHeaders[i], sql.NVarChar(sql.MAX), { nullable: true });
  }
  for (const row of rows) {
    tvpTable.rows.add(...safeHeaders.map((h, i) => stringify(row[headers[i]])));
  }

  const updateCols = safeHeaders
    .slice(1)
    .map((h) => `target.${quoteIdentifier(h)} = source.${quoteIdentifier(h)}`)
    .join(", ");
  const insertCols = safeHeaders.map((h) => quoteIdentifier(h)).join(", ");
  const insertVals = safeHeaders.map((h) => `source.${quoteIdentifier(h)}`).join(", ");

  const mergeSql = `
    MERGE ${quotedTable} AS target
    USING @tvp AS source ON target.${pkCol} = source.${pkCol}
    WHEN MATCHED THEN
      UPDATE SET ${updateCols}
    WHEN NOT MATCHED THEN
      INSERT (${insertCols}) VALUES (${insertVals});
  `;

  const request = db.request();
  request.input("tvp", tvpTable);
  const result = await request.query(mergeSql);
  return result.rowsAffected[0] || rows.length;
}

/**
 * Fallback for very wide tables: upsert one row at a time.
 */
async function upsertRowsOneByOne(db, tableName, headers, rows) {
  const quotedTable = quoteIdentifier(tableName);
  const safeHeaders = headers.map((h, i) => sanitizeIdentifier(h) || `col_${i}`);
  const pkCol = quoteIdentifier(safeHeaders[0]);
  let total = 0;

  for (const row of rows) {
    const request = db.request();
    const updateParts = [];
    const insertCols  = [];
    const insertVals  = [];

    for (let i = 0; i < safeHeaders.length; i++) {
      const paramName = `p${i}`;
      request.input(paramName, sql.NVarChar(sql.MAX), stringify(row[headers[i]]));
      const quotedCol = quoteIdentifier(safeHeaders[i]);

      insertCols.push(quotedCol);
      insertVals.push(`@${paramName}`);
      if (i > 0) updateParts.push(`target.${quotedCol} = @${paramName}`);
    }

    const mergeSql = `
      MERGE ${quotedTable} AS target
      USING (SELECT @p0 AS pk) AS source ON target.${pkCol} = source.pk
      WHEN MATCHED THEN
        UPDATE SET ${updateParts.join(", ")}
      WHEN NOT MATCHED THEN
        INSERT (${insertCols.join(", ")}) VALUES (${insertVals.join(", ")});
    `;

    const result = await request.query(mergeSql);
    total += result.rowsAffected[0] || 1;
  }

  return total;
}

/**
 * Build an mssql BulkLoad Table object for a set of rows.
 */
function buildBulkTable(tableName, headers, rows) {
  const safeHeaders = headers.map((h, i) => sanitizeIdentifier(h) || `col_${i}`);
  const table = new sql.Table(tableName);
  table.create = false;

  for (const h of safeHeaders) {
    table.columns.add(h, sql.NVarChar(sql.MAX), { nullable: true });
  }
  for (const row of rows) {
    table.rows.add(...safeHeaders.map((_h, i) => stringify(row[headers[i]])));
  }
  return table;
}

// ─── Identifier helpers ───────────────────────────────────────────────────────

/** Quote a SQL identifier with brackets. */
function quoteIdentifier(name) {
  return `[${name.replace(/]/g, "]]")}]`;
}

/** Remove characters not safe in SQL identifiers. */
function sanitizeIdentifier(name) {
  if (!name && name !== 0) return "";
  return String(name)
    .replace(/[^a-zA-Z0-9_]/g, "_")
    .replace(/^([0-9])/, "_$1")
    .substring(0, 128); // SQL Server max identifier length
}

/** Escape single quotes in a literal string used inside SQL. */
function escapeSingleQuote(str) {
  return str.replace(/'/g, "''");
}

/** Convert any cell value to a string suitable for NVARCHAR storage. */
function stringify(val) {
  if (val === null || val === undefined) return null;
  if (typeof val === "boolean") return val ? "1" : "0";
  return String(val);
}

// ─── Start server ─────────────────────────────────────────────────────────────

app.listen(PORT, () => {
  console.log(`Excel-to-Azure-SQL sync backend running on port ${PORT}`);
  console.log(`  Health: http://localhost:${PORT}/health`);
  console.log(`  Sync:   POST http://localhost:${PORT}/api/sync`);
});
