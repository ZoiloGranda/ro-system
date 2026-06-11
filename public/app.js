const currencyFormatter = new Intl.NumberFormat('es-VE', {
	minimumFractionDigits: 2,
	maximumFractionDigits: 2,
});

const state = {
	transactions: [],
	pendingOutgoing: [],
	metadata: { cuentas: [], paymentMethods: [] },
	overview: null,
	summary: [],
	charts: {},
};

const elements = {
	kpiGrid: document.getElementById('kpiGrid'),
	transactionsTableBody: document.getElementById('transactionsTableBody'),
	pendingOutgoingTableBody: document.getElementById('pendingOutgoingTableBody'),
	accountFilter: document.getElementById('accountFilter'),
	methodFilter: document.getElementById('methodFilter'),
	searchInput: document.getElementById('searchInput'),
	entryForm: document.getElementById('entryForm'),
	formStatus: document.getElementById('formStatus'),
	leaderboard: document.getElementById('leaderboard'),
	navButtons: [...document.querySelectorAll('[data-nav-target]')],
	views: [...document.querySelectorAll('[data-view]')],
};

const kpiConfig = [
	{ key: 'transactionCount', label: 'Transacciones', type: 'count' },
	{ key: 'activeCuentas', label: 'Cuentas activas', type: 'count' },
	{ key: 'totalVolume', label: 'Actividad total', type: 'currency' },
	{ key: 'pendingAmount', label: 'Entrega pendiente', type: 'currency' },
	{ key: 'commissionIncome', label: 'Ingresos por comision', type: 'currency' },
	{ key: 'bolivaresCompras', label: 'Bolivares compras', type: 'currency' },
];

const chartColors = ['#0f766e', '#fb7185', '#38bdf8', '#f59e0b', '#6366f1', '#16a34a'];

const fetchJson = async (url, options) => {
	const response = await fetch(url, options);
	const payload = await response.json().catch(() => ({}));

	if (!response.ok) {
		throw new Error(payload.error || payload.message || 'La solicitud fallo');
	}

	return payload.data;
};

const formatValue = (value, type) => {
	if (type === 'currency') {
		return currencyFormatter.format(Number(value || 0));
	}

	return String(value || 0);
};

const renderKpis = () => {
	if (!state.overview) {
		return;
	}

	elements.kpiGrid.innerHTML = kpiConfig.map(({ key, label, type }) => `
		<article class="kpi-card">
			<p>${label}</p>
			<strong>${formatValue(state.overview.kpis[key], type)}</strong>
		</article>
	`).join('');
};

const renderFilters = () => {
	elements.accountFilter.innerHTML = ['<option value="">Todas las cuentas</option>']
		.concat(state.metadata.cuentas.map((cuenta) => `<option value="${cuenta}">${cuenta}</option>`))
		.join('');

	elements.methodFilter.innerHTML = ['<option value="">Todos los metodos</option>']
		.concat(state.metadata.paymentMethods.map((method) => `<option value="${method}">${method}</option>`))
		.join('');
};

const filteredTransactions = () => {
	const query = elements.searchInput.value.trim().toLowerCase();
	const cuenta = elements.accountFilter.value;
	const method = elements.methodFilter.value;

	return state.transactions.filter((transaction) => {
		const text = [transaction.cliente, transaction.envia, transaction.cuenta].join(' ').toLowerCase();
		const matchesQuery = !query || text.includes(query);
		const matchesCuenta = !cuenta || transaction.cuenta.toUpperCase() === cuenta;
		const matchesMethod = !method || transaction.paymentMethod === method;
		return matchesQuery && matchesCuenta && matchesMethod;
	});
};

const renderTransactions = () => {
	const rows = filteredTransactions();

	elements.transactionsTableBody.innerHTML = rows.map((transaction) => `
		<tr>
			<td>${transaction.fecha}</td>
			<td>${transaction.cliente || '-'}</td>
			<td>${transaction.envia || '-'}</td>
			<td>${transaction.cuenta || '-'}</td>
			<td>${transaction.paymentMethod}</td>
			<td>${currencyFormatter.format(transaction.zelle)}</td>
			<td>${currencyFormatter.format(transaction.efectivo)}</td>
			<td>${currencyFormatter.format(transaction.entregaPendiente)}</td>
			<td>${currencyFormatter.format(transaction.totalGeneral)}</td>
		</tr>
	`).join('');
};

const renderPendingOutgoing = () => {
	if (!state.pendingOutgoing.length) {
		elements.pendingOutgoingTableBody.innerHTML = `
			<tr>
				<td colspan="7" class="empty-cell">No se detectaron salientes pendientes en la hoja X PAGAR.</td>
			</tr>
		`;
		return;
	}

	elements.pendingOutgoingTableBody.innerHTML = state.pendingOutgoing.map((transaction) => `
		<tr>
			<td>${transaction.reference}</td>
			<td>${transaction.sourceAccount}</td>
			<td>${transaction.beneficiary}</td>
			<td>${currencyFormatter.format(transaction.amount)}</td>
			<td>${transaction.reason}</td>
			<td>${transaction.sheet}</td>
			<td><span class="status-pill">${transaction.status}</span></td>
		</tr>
	`).join('');
};

const destroyChart = (name) => {
	if (state.charts[name]) {
		state.charts[name].destroy();
	}
};

const buildBarChart = (canvasId, labels, values, color) => {
	return new Chart(document.getElementById(canvasId), {
		type: 'bar',
		data: {
			labels,
			datasets: [{
				data: values,
				backgroundColor: color,
				borderRadius: 12,
			}]
		},
		options: {
			plugins: { legend: { display: false } },
			maintainAspectRatio: false,
			scales: { y: { beginAtZero: true } },
		}
	});
};

const renderCharts = () => {
	if (!state.overview) {
		return;
	}

	destroyChart('topAccounts');
	destroyChart('paymentMix');
	destroyChart('dailyTrend');

	state.charts.topAccounts = buildBarChart(
		'topAccountsChart',
		state.overview.topCuentas.map((item) => item.cuenta),
		state.overview.topCuentas.map((item) => item.totalGeneral),
		chartColors[0]
	);

	state.charts.paymentMix = new Chart(document.getElementById('paymentMixChart'), {
		type: 'doughnut',
		data: {
			labels: state.overview.paymentMix.map((item) => item.label),
			datasets: [{
				data: state.overview.paymentMix.map((item) => item.value),
				backgroundColor: chartColors,
				hoverOffset: 12,
			}],
		},
		options: {
			maintainAspectRatio: false,
		}
	});

	state.charts.dailyTrend = new Chart(document.getElementById('dailyTrendChart'), {
		type: 'line',
		data: {
			labels: state.overview.dailyTrend.map((item) => item.day),
			datasets: [{
				label: 'Actividad diaria',
				data: state.overview.dailyTrend.map((item) => item.total),
				borderColor: '#111827',
				backgroundColor: 'rgba(17, 24, 39, 0.08)',
				fill: true,
				tension: 0.35,
			}],
		},
		options: {
			plugins: { legend: { display: false } },
			maintainAspectRatio: false,
			scales: { y: { beginAtZero: true } },
		}
	});
};

const renderLeaderboard = () => {
	elements.leaderboard.innerHTML = state.summary.slice(0, 10).map((entry, index) => `
		<article class="leaderboard-row">
			<div>
				<span class="leaderboard-rank">${index + 1}</span>
				<strong>${entry.cuenta}</strong>
				<p>${entry.transactions} transacciones</p>
			</div>
			<strong>${currencyFormatter.format(entry.totalGeneral)}</strong>
		</article>
	`).join('');
};

const renderAll = () => {
	renderKpis();
	renderFilters();
	renderTransactions();
	renderPendingOutgoing();
	renderCharts();
	renderLeaderboard();
};

const loadAllData = async () => {
	const [transactions, pendingOutgoing, metadata, overview, summary] = await Promise.all([
		fetchJson('/api/transactions'),
		fetchJson('/api/transactions/pending-outgoing'),
		fetchJson('/api/filters/metadata'),
		fetchJson('/api/reports/overview'),
		fetchJson('/api/summary/cuentas'),
	]);

	state.transactions = transactions;
	state.pendingOutgoing = pendingOutgoing;
	state.metadata = metadata;
	state.overview = overview;
	state.summary = summary;
	renderAll();
};

const switchView = (viewName) => {
	elements.views.forEach((view) => {
		view.classList.toggle('is-visible', view.dataset.view === viewName);
	});

	document.querySelectorAll('.nav-link').forEach((button) => {
		button.classList.toggle('is-active', button.dataset.navTarget === viewName);
	});
};

elements.navButtons.forEach((button) => {
	button.addEventListener('click', () => switchView(button.dataset.navTarget));
});

[elements.searchInput, elements.accountFilter, elements.methodFilter].forEach((element) => {
	element.addEventListener('input', renderTransactions);
	element.addEventListener('change', renderTransactions);
});

elements.entryForm.addEventListener('submit', async (event) => {
	event.preventDefault();
	elements.formStatus.textContent = 'Guardando registro de demo...';

	const payload = Object.fromEntries(new FormData(elements.entryForm).entries());

	try {
		await fetchJson('/api/transactions', {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify(payload),
		});
		elements.entryForm.reset();
		elements.formStatus.textContent = 'Registro agregado. Se actualizaron el resumen, la tabla y el ranking.';
		await loadAllData();
		switchView('overview');
	} catch (error) {
		elements.formStatus.textContent = error.message;
	}
});

loadAllData().catch((error) => {
	elements.formStatus.textContent = error.message;
});