# ro-system

Small Node.js utility project for cleaning an Excel sheet, generating grouped `CUENTA` totals, building an HTML chart, and serving that chart locally.

## Requirements

- Node.js 18 or newer
- npm

## Setup

1. Open a terminal in the project root.
2. Install dependencies:

```bash
npm install
```

## Project Workflow

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

### 4. Serve the generated chart locally

```bash
npm start
```

This starts a small HTTP server at:

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