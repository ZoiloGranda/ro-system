const fs = require('fs');
const XLSX = require('xlsx');

const DEFAULT_SOURCE_FILE = 'files/cuenta-totals.csv';
const DEFAULT_OUTPUT_FILE = 'files/cuenta-totals-chart.html';

const [, , sourceArg, outputArg] = process.argv;
const SOURCE_FILE = sourceArg || DEFAULT_SOURCE_FILE;
const OUTPUT_FILE = outputArg || DEFAULT_OUTPUT_FILE;

const cleanText = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();
const normalizeHeader = (value) => cleanText(value).toUpperCase();

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

const createChartColors = (count) => Array.from({ length: count }, (_, index) => {
	const hue = Math.round((index * 360) / Math.max(count, 1));
	return `hsl(${hue} 68% 55%)`;
});

const buildChartHtml = (labels, values) => {
	const colors = createChartColors(labels.length);
	const chartData = JSON.stringify({ labels, values, colors });

	return `<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<title>Cuenta Totals Doughnut Chart</title>
	<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
	<style>
		:root {
			color-scheme: light;
			font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
			background: linear-gradient(160deg, #f7f2e8 0%, #fffdf8 45%, #eef4f8 100%);
			color: #1f2933;
		}

		* {
			box-sizing: border-box;
		}

		body {
			margin: 0;
			min-height: 100vh;
			display: grid;
			place-items: center;
			padding: 24px;
		}

		.card {
			width: min(1100px, 100%);
			background: rgba(255, 255, 255, 0.92);
			border: 1px solid rgba(31, 41, 51, 0.08);
			border-radius: 24px;
			box-shadow: 0 24px 64px rgba(31, 41, 51, 0.12);
			padding: 28px;
		}

		h1 {
			margin: 0 0 8px;
			font-size: clamp(1.75rem, 2vw + 1rem, 2.5rem);
			letter-spacing: -0.03em;
		}

		p {
			margin: 0 0 24px;
			color: #52606d;
		}

		.chart-wrap {
			position: relative;
			min-height: 420px;
		}

		.footer {
			margin-top: 20px;
			font-size: 0.95rem;
			color: #7b8794;
		}
	</style>
</head>
<body>
	<main class="card">
		<h1>Cuenta Totals</h1>
		<p>Doughnut chart generated from the summarized TOTAL GENERAL values.</p>
		<div class="chart-wrap">
			<canvas id="cuentaChart"></canvas>
		</div>
		<div class="footer">Open this file in a browser after running the chart generator.</div>
	</main>

	<script>
		const data = ${chartData};
		const currencyFormatter = new Intl.NumberFormat('en-US', {
			minimumFractionDigits: 2,
			maximumFractionDigits: 2,
		});

		new Chart(document.getElementById('cuentaChart'), {
			type: 'doughnut',
			data: {
				labels: data.labels,
				datasets: [{
					data: data.values,
					backgroundColor: data.colors,
					borderColor: '#ffffff',
					borderWidth: 2,
					hoverOffset: 12,
				}],
			},
			options: {
				responsive: true,
				maintainAspectRatio: false,
				plugins: {
					legend: {
						position: 'right',
						labels: {
							boxWidth: 14,
							boxHeight: 14,
							usePointStyle: true,
							pointStyle: 'circle',
						},
					},
					tooltip: {
						callbacks: {
							label(context) {
								const label = context.label || '';
								const value = Number(context.parsed) || 0;
								return label + ': ' + currencyFormatter.format(value);
							},
						},
					},
				},
				cutout: '58%',
			},
		});
	</script>
</body>
</html>`;
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
const totalGeneralIndex = normalizedHeaders.indexOf('TOTAL GENERAL');

if (cuentaIndex === -1 || totalGeneralIndex === -1) {
	throw new Error(`Columns "CUENTA" and "TOTAL GENERAL" are required in ${SOURCE_FILE}`);
}

const chartLabels = [];
const chartValues = [];

for (const row of rows.slice(1)) {
	const cuenta = cleanText(row[cuentaIndex]);
	const totalGeneral = parseAmount(row[totalGeneralIndex]);

	if (!cuenta || totalGeneral === null || totalGeneral === 0) {
		continue;
	}

	chartLabels.push(cuenta);
	chartValues.push(totalGeneral);
}

const chartHtml = buildChartHtml(chartLabels, chartValues);

fs.writeFileSync(OUTPUT_FILE, chartHtml, 'utf8');

console.log(`Chart written to ${OUTPUT_FILE}`);
console.log(`Source: ${SOURCE_FILE}`);
console.log(`Chart slices: ${chartLabels.length}`);