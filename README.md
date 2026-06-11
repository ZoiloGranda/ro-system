# ro-system

Node.js workspace for cleaning the `ZELLE` Excel sheet, generating grouped `CUENTA` totals, and now serving a JavaScript dashboard prototype backed by mock API data.

## Requirements

- Node.js 18 or newer
- npm

## Setup

1. Open a terminal in the project root.
2. Install dependencies:

```bash
npm install
```

## Prototype Dashboard

`npm start` now runs a Node.js backend that serves a browser-based JavaScript dashboard from `public/` and exposes mock JSON endpoints for the demo UI.

Default URL:

```text
http://127.0.0.1:3000
```

If port `3000` is already in use:

```powershell
$env:PORT=3001
npm start
```

Prototype features:

- Overview dashboard with KPI cards and charts
- Transaction explorer with search and filters
- New-entry form aligned to the `ZELLE` sheet fields
- In-memory mock submission flow that refreshes reports instantly

Prototype API routes:

- `GET /api/transactions`
- `GET /api/summary/cuentas`
- `GET /api/reports/overview`
- `GET /api/filters/metadata`
- `POST /api/transactions`

The prototype uses seeded mock data shaped from the current spreadsheet outputs. It is intended for presentation use, not production persistence.

## Spreadsheet Utility Workflow

The default workflow uses the files in the `files/` directory:

- `files/MAYO.xlsx`: source Excel file
- `files/output.csv`: cleaned CSV exported from the Excel sheet
- `files/cuenta-totals.csv`: grouped totals by `CUENTA`
- `files/cuenta-totals-chart.html`: generated doughnut chart

Run the commands in this order.

### 1. Convert Excel to a cleaned CSV

```bash
npm run convert
```

Default behavior:

- Source file: `files/MAYO.xlsx`
- Sheet name: `ZELLE`
- Output file: `files/output.csv`

Custom arguments:

```bash
node convertToCsv.js <source-file> <sheet-name> <output-file>
```

Example:

```bash
node convertToCsv.js files/MAYO.xlsx ZELLE files/output.csv
```

### 2. Summarize totals by `CUENTA`

```bash
npm run summarize
```

Default behavior:

- Source file: `files/output.csv`
- Output file: `files/cuenta-totals.csv`

Custom arguments:

```bash
node summarizeCuentaTotals.js <source-file> <output-file>
```

Example:

```bash
node summarizeCuentaTotals.js files/output.csv files/cuenta-totals.csv
```

### 3. Generate the chart HTML

```bash
npm run chart
```

Default behavior:

- Source file: `files/cuenta-totals.csv`
- Output file: `files/cuenta-totals-chart.html`

Custom arguments:

```bash
node generateCuentasTotales.js <source-file> <output-file>
```

Example:

```bash
node generateCuentasTotales.js files/cuenta-totals.csv files/cuenta-totals-chart.html
```

### 4. Serve the dashboard locally

```bash
npm start
```

This starts the Node.js prototype server at:

```text
http://127.0.0.1:3000
```

You can override the port with the `PORT` environment variable.

PowerShell example:

```powershell
$env:PORT=4000
npm start
```

## Quick Start

If your source file is already at `files/MAYO.xlsx` and the sheet is named `ZELLE`, run:

```bash
npm install
npm run convert
npm run summarize
npm run chart
npm start
```

Then open `http://127.0.0.1:3000` in your browser.

## Notes

- The conversion step looks for a header row that matches the expected sheet structure.
- The summary step groups rows by normalized `CUENTA` values, so name casing and extra spaces do not create separate groups.
- The chart uses the `TOTAL GENERAL` column from the summary output.
- The dashboard prototype is served from `public/` and uses mock API responses generated in `server.js` and `mockData.js`.