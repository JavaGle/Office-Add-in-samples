# Excel to Azure SQL Sync — Office Add-in Sample

This sample demonstrates an Excel Web Add-in that syncs one or more Excel worksheet's data to **Azure SQL Database** in real time. It supports:

- **Multiple sheets** — choose which worksheets to sync via a checklist UI.
- **Large column sets** — all data is read in a single `Excel.run` call using the range `values` 2D array; HTTP requests are chunked into configurable row batches to avoid payload size and timeout issues.
- **Three sync modes** — *Upsert*, *Insert only*, and *Replace* (truncate + reload).
- **Progress tracking** — a progress bar updates after each batch.
- **Retry on transient network error** — one automatic retry per batch.

---

## Architecture

```
┌──────────────────────────────────────────────────┐
│  Excel Desktop / Web                             │
│                                                  │
│  ┌─────────────────────────────────────────────┐ │
│  │  Task Pane (Add-in)                         │ │
│  │  • Reads sheet data via Office JS API       │ │
│  │  • Batches rows (default: 200 / request)    │ │
│  │  • POST /api/sync  →  Backend               │ │
│  └─────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────┘
            │  HTTPS (JSON)
            ▼
┌───────────────────────────────────┐
│  Backend — Node.js / Express      │
│  backend/server.js                │
│  • Validates payload              │
│  • Creates table if not exists    │
│  • Bulk-insert / MERGE (upsert)   │
│  • Handles wide tables (>900 col) │
└───────────────────────────────────┘
            │  TCP / TDS
            ▼
┌───────────────────────────┐
│  Azure SQL Database       │
└───────────────────────────┘
```

---

## Quick start

### 1 — Deploy the backend

```bash
cd backend
cp .env.example .env
# Edit .env with your Azure SQL credentials
npm install
npm start
```

The server starts on `http://localhost:3001`. Deploy this to Azure App Service, Azure Container Apps, or any host reachable from your Excel Add-in.

### 2 — Install the add-in

```bash
# From the root of this sample
npm install
npm run dev-server       # starts webpack-dev-server on https://localhost:3000
```

Then sideload `manifest.xml` in Excel:

- **Excel on the web**: go to **Insert → Add-ins → Upload My Add-in** and select `manifest.xml`.
- **Excel desktop (Windows)**: follow the [sideloading guide](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in).

### 3 — Use the add-in

1. Open Excel with your data. Each sheet you want to sync should have **column headers in row 1**.
2. Open the task pane: **Home → Azure SQL Sync → Sync to Azure SQL**.
3. **Step 1**: Enter the backend API URL (e.g. `http://localhost:3001/api/sync`) and optionally an API key.  Click **Save Configuration**.
4. **Step 2**: Click **Refresh Sheet List** and tick the sheets you want to sync.
5. **Step 3**: Click **Sync to Azure SQL**.

---

## Performance considerations

| Scenario | Recommendation |
|---|---|
| Many columns (> 200) | The add-in reads the entire used range in **one** `Excel.run` call — only the HTTP transfer size matters. Reduce batch size if requests time out. |
| Large number of rows | Increase the **Row Batch Size** (up to 2000) to reduce the number of HTTP round-trips. |
| Very wide tables (> 900 columns) | The backend automatically falls back from a TVP MERGE to a per-row MERGE to stay within SQL Server's column limits. |
| Slow network | Reduce batch size so individual requests complete faster. The add-in retries once on transient network failures. |

---

## Security notes

- The backend accepts a shared-secret API key via the `x-api-key` header. Set `API_KEY` in the backend `.env` file and the matching key in the add-in task pane.
- All column and table names are sanitized on **both** the client and the server before being used in SQL statements. Identifiers are bracket-quoted to prevent SQL injection.
- All cell values are stored as `NVARCHAR(MAX)` — no dynamic SQL is constructed from cell values.
- For production, deploy the backend behind HTTPS and restrict CORS origins.

---

## Project structure

```
excel-add-in-azure-sql-sync/
├── manifest.xml              ← Office Add-in manifest
├── package.json              ← Add-in build dependencies
├── webpack.config.js
├── babel.config.json
├── assets/                   ← Add-in icons
├── src/
│   └── taskpane/
│       ├── taskpane.html     ← Task pane UI
│       ├── taskpane.css
│       └── taskpane.js       ← Office JS logic + fetch calls
└── backend/
    ├── package.json          ← Backend dependencies (express, mssql)
    ├── server.js             ← Express + Azure SQL sync API
    └── .env.example          ← Environment variable template
```

---

## Related resources

- [Office Add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/)
- [Excel JavaScript API reference](https://learn.microsoft.com/javascript/api/excel)
- [mssql npm package](https://www.npmjs.com/package/mssql)
- [Azure SQL Database quickstart](https://learn.microsoft.com/azure/azure-sql/database/single-database-create-quickstart)
