const fs = require('fs');
const XLSX = require('xlsx');

const DEFAULT_SOURCE_FILE = 'files/MAYO.xlsx';
const DEFAULT_SHEET_NAME = 'ZELLE';
const DEFAULT_OUTPUT_FILE = 'files/output.csv';

const [, , sourceArg, sheetArg, outputArg] = process.argv;
const SOURCE_FILE = sourceArg || DEFAULT_SOURCE_FILE;
const SHEET_NAME = sheetArg || DEFAULT_SHEET_NAME;
const OUTPUT_FILE = outputArg || DEFAULT_OUTPUT_FILE;

const EXPECTED_HEADER = [
	'FECHA',
	'CLIENTE',
	'ENVIA',
	'CUENTA',
	'CLIENTE2',
	'ZELLE',
	'EFECTIVO',
	'BOLIVARES COMPRAS',
	'OTROS METODOS',
	'COMISIONES POR TERCEROS',
	'INGRESOS POR COMISIONES EN BOLIVARES',
	'ENTREGA PENDIENTE',
	'BOLIVARES VENTA',
];

const workbook = XLSX.readFile(SOURCE_FILE);
const worksheet = workbook.Sheets[SHEET_NAME];

if (!worksheet) {
	throw new Error(`Sheet "${SHEET_NAME}" was not found in ${SOURCE_FILE}`);
}

const rows = XLSX.utils.sheet_to_json(worksheet, {
	header: 1,
	raw: false,
	blankrows: false,
});

const normalizeCell = (value) => String(value ?? '').replace(/\s+/g, ' ').trim().toUpperCase();

const headerRowIndex = rows.findIndex((row) => {
	const normalizedRow = row.map(normalizeCell);
	return EXPECTED_HEADER.every((header, index) => normalizedRow[index] === header);
});

if (headerRowIndex === -1) {
	throw new Error(`Could not find the expected header row in sheet "${SHEET_NAME}"`);
}

const dataRows = rows.slice(headerRowIndex);

const logicalWidth = dataRows.reduce((maxWidth, row, index) => {
	if (index >= 50) {
		return maxWidth;
	}

	const lastNonEmptyIndex = row.reduce((lastIndex, cell, cellIndex) => {
		return cell !== null && cell !== undefined && String(cell).trim() !== ''
			? cellIndex
			: lastIndex;
	}, -1);

	const nonEmptyCount = row.reduce((count, cell) => {
		return cell !== null && cell !== undefined && String(cell).trim() !== ''
			? count + 1
			: count;
	}, 0);

	if (nonEmptyCount >= 4 && lastNonEmptyIndex + 1 > maxWidth) {
		return lastNonEmptyIndex + 1;
	}

	return maxWidth;
}, 0);

const cleanedRows = dataRows
	.map((row) => {
		const normalizedRow = logicalWidth > 0 ? row.slice(0, logicalWidth) : [...row];

		while (normalizedRow.length > 0) {
			const lastCell = normalizedRow[normalizedRow.length - 1];
			if (lastCell !== null && lastCell !== undefined && String(lastCell).trim() !== '') {
				break;
			}
			normalizedRow.pop();
		}

		return normalizedRow;
	})
	.filter((row) => row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== ''));

const cleanedWorksheet = XLSX.utils.aoa_to_sheet(cleanedRows);
const csv = XLSX.utils.sheet_to_csv(cleanedWorksheet, { blankrows: false });

fs.writeFileSync(OUTPUT_FILE, csv, 'utf8');

console.log(`Clean CSV written to ${OUTPUT_FILE}`);
console.log(`Source: ${SOURCE_FILE}`);
console.log(`Sheet: ${SHEET_NAME}`);
