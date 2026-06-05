const fs = require('fs');
const XLSX = require('xlsx');

const DEFAULT_SOURCE_FILE = 'files/output.csv';
const DEFAULT_OUTPUT_FILE = 'files/cuenta-totals.csv';

const [, , sourceArg, outputArg] = process.argv;
const SOURCE_FILE = sourceArg || DEFAULT_SOURCE_FILE;
const OUTPUT_FILE = outputArg || DEFAULT_OUTPUT_FILE;

const normalizeHeader = (value) => String(value ?? '').replace(/\s+/g, ' ').trim().toUpperCase();
const cleanText = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();
const normalizeCuenta = (value) => cleanText(value).toUpperCase();

const parseAmount = (value) => {
	const text = cleanText(value);

	if (!text) {
		return null;
	}

	const normalized = text
		.replace(/\((.*)\)/, '-$1')
		.replace(/[$,]/g, '')
		.replace(/\s+/g, '');

	const amount = Number(normalized);
	return Number.isFinite(amount) ? amount : null;
};

const formatAmount = (value) => {
	if (!value) {
		return '';
	}

	return new Intl.NumberFormat('en-US', {
		minimumFractionDigits: 2,
		maximumFractionDigits: 2,
	}).format(value);
};

const workbook = XLSX.readFile(SOURCE_FILE, { raw: false });
const [firstSheetName] = workbook.SheetNames;
const worksheet = workbook.Sheets[firstSheetName];

if (!worksheet) {
	throw new Error(`Could not read a worksheet from ${SOURCE_FILE}`);
}

const rows = XLSX.utils.sheet_to_json(worksheet, {
	header: 1,
	raw: false,
	blankrows: false,
});

if (rows.length === 0) {
	throw new Error(`No rows found in ${SOURCE_FILE}`);
}

const headers = rows[0].map(cleanText);
const normalizedHeaders = headers.map(normalizeHeader);
const cuentaIndex = normalizedHeaders.indexOf('CUENTA');

if (cuentaIndex === -1) {
	throw new Error(`Column "CUENTA" was not found in ${SOURCE_FILE}`);
}

const textColumns = new Set(['FECHA', 'CLIENTE', 'ENVIA', 'CUENTA', 'CLIENTE2']);
const amountColumns = normalizedHeaders.reduce((indexes, header, index) => {
	if (!textColumns.has(header)) {
		indexes.push(index);
	}

	return indexes;
}, []);

const totalsByCuenta = new Map();

for (const row of rows.slice(1)) {
	const cuenta = normalizeCuenta(row[cuentaIndex]);

	if (!cuenta) {
		continue;
	}

	if (!totalsByCuenta.has(cuenta)) {
		totalsByCuenta.set(cuenta, {
			cuenta,
			totals: amountColumns.reduce((accumulator, index) => {
				accumulator[index] = 0;
				return accumulator;
			}, {}),
		});
	}

	const entry = totalsByCuenta.get(cuenta);

	for (const index of amountColumns) {
		const amount = parseAmount(row[index]);
		if (amount !== null) {
			entry.totals[index] += amount;
		}
	}
}

const summaryRows = [[
	'CUENTA',
	...amountColumns.map((index) => headers[index] || `COLUMN_${index + 1}`),
	'TOTAL GENERAL',
]];

const sortedEntries = [...totalsByCuenta.values()].sort((left, right) => left.cuenta.localeCompare(right.cuenta));

for (const { cuenta, totals } of sortedEntries) {
	const rowTotals = amountColumns.map((index) => totals[index]);
	const grandTotal = rowTotals.reduce((sum, value) => sum + value, 0);

	summaryRows.push([
		cuenta,
		...rowTotals.map(formatAmount),
		formatAmount(grandTotal),
	]);
}

const summaryWorksheet = XLSX.utils.aoa_to_sheet(summaryRows);
const csv = XLSX.utils.sheet_to_csv(summaryWorksheet, { blankrows: false });

fs.writeFileSync(OUTPUT_FILE, csv, 'utf8');

console.log(`Cuenta totals written to ${OUTPUT_FILE}`);
console.log(`Source: ${SOURCE_FILE}`);
console.log(`Grouped accounts: ${totalsByCuenta.size}`);