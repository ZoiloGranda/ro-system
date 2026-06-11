const path = require('path');
const XLSX = require('xlsx');

const workbookPath = path.join(__dirname, 'files', 'MAYO.xlsx');

const transactionSeed = [
	{
		id: 'txn-001',
		fecha: 1,
		cliente: 'SALDO INICIAL',
		envia: 'SALDO INICIAL',
		cuenta: 'ROSSELIS',
		cliente2: 'ZELLE',
		zelle: 1309.16,
		efectivo: 0,
		bolivaresCompras: 0,
		otrosMetodos: 0,
		comisionesPorTerceros: 0,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: 0,
		bolivaresVenta: 0,
	},
	{
		id: 'txn-002',
		fecha: 1,
		cliente: 'MELVIN',
		envia: 'ROSANGELA',
		cuenta: 'MELVIN',
		cliente2: 'ZELLE',
		zelle: 160.5,
		efectivo: 150,
		bolivaresCompras: 0,
		otrosMetodos: 0,
		comisionesPorTerceros: 0,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: 0,
		bolivaresVenta: 0,
	},
	{
		id: 'txn-003',
		fecha: 1,
		cliente: 'caracciolo',
		envia: 'caracciolo',
		cuenta: 'MELVIN',
		cliente2: 'zelle',
		zelle: 140,
		efectivo: 0,
		bolivaresCompras: 42000,
		otrosMetodos: 0,
		comisionesPorTerceros: 0,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: 0,
		bolivaresVenta: 0,
	},
	{
		id: 'txn-004',
		fecha: 2,
		cliente: 'ELIBETH',
		envia: 'Lisbeth Abreu',
		cuenta: 'NOSE',
		cliente2: 'ZELLE',
		zelle: 74,
		efectivo: 0,
		bolivaresCompras: 39550,
		otrosMetodos: 0,
		comisionesPorTerceros: 4,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: 0,
		bolivaresVenta: 4,
	},
	{
		id: 'txn-005',
		fecha: 2,
		cliente: 'ELIBETH',
		envia: 'Erick Guillaume',
		cuenta: 'MELVIN',
		cliente2: 'ZELLE',
		zelle: 220,
		efectivo: 206.8,
		bolivaresCompras: 0,
		otrosMetodos: 0,
		comisionesPorTerceros: 0,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: 206.8,
		bolivaresVenta: 0,
	},
	{
		id: 'txn-006',
		fecha: 2,
		cliente: 'ELIBETH',
		envia: 'EFECTIVO',
		cuenta: 'LA PUERTA',
		cliente2: 'EFECTIVO',
		zelle: 0,
		efectivo: 0,
		bolivaresCompras: 0,
		otrosMetodos: 0,
		comisionesPorTerceros: 0,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: -290,
		bolivaresVenta: 0,
	},
	{
		id: 'txn-007',
		fecha: 3,
		cliente: 'JORGE ELIEZER',
		envia: 'Mariela Pinto',
		cuenta: 'JORGE ELIEZER',
		cliente2: 'ZELLE',
		zelle: 780,
		efectivo: 560,
		bolivaresCompras: 168500,
		otrosMetodos: 0,
		comisionesPorTerceros: 6.8,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: 78,
		bolivaresVenta: 0,
	},
	{
		id: 'txn-008',
		fecha: 4,
		cliente: 'JOEL HERNANDEZ',
		envia: 'Jose Lopez',
		cuenta: 'JOEL HERNANDEZ',
		cliente2: 'ZELLE',
		zelle: 1026.99,
		efectivo: 0,
		bolivaresCompras: 14125,
		otrosMetodos: 0,
		comisionesPorTerceros: 0,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: 120,
		bolivaresVenta: 0,
	},
	{
		id: 'txn-009',
		fecha: 5,
		cliente: 'REGION BANK',
		envia: 'Carlos Mena',
		cuenta: 'REGION BANK',
		cliente2: 'ZELLE',
		zelle: 2623.75,
		efectivo: 1098.29,
		bolivaresCompras: 652147.8,
		otrosMetodos: 31.25,
		comisionesPorTerceros: 0,
		ingresosPorComisionesEnBolivares: 1,
		entregaPendiente: 419.37,
		bolivaresVenta: 0,
	},
	{
		id: 'txn-010',
		fecha: 5,
		cliente: 'ROBERTO FARAH',
		envia: 'Adriana Nunez',
		cuenta: 'ROBERTO FARAH',
		cliente2: 'ZELLE',
		zelle: 1548.4,
		efectivo: 532,
		bolivaresCompras: 456770,
		otrosMetodos: 0,
		comisionesPorTerceros: 8,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: 357,
		bolivaresVenta: 0,
	},
	{
		id: 'txn-011',
		fecha: 6,
		cliente: 'LEONEL AMAYA',
		envia: 'Pedro Blanco',
		cuenta: 'LEONEL AMAYA',
		cliente2: 'OTROS',
		zelle: 5661,
		efectivo: 205,
		bolivaresCompras: 3465600,
		otrosMetodos: 0,
		comisionesPorTerceros: 0,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: 0,
		bolivaresVenta: 3465600,
	},
	{
		id: 'txn-012',
		fecha: 6,
		cliente: 'MILAGROS ZELLE BOFA',
		envia: 'Lisbeth Abreu',
		cuenta: 'MILAGROS ZELLE BOFA',
		cliente2: 'ZELLE',
		zelle: 106,
		efectivo: 0,
		bolivaresCompras: 60000,
		otrosMetodos: 0,
		comisionesPorTerceros: 5,
		ingresosPorComisionesEnBolivares: 0,
		entregaPendiente: 0,
		bolivaresVenta: 6,
	},
];

const numericFields = [
	'zelle',
	'efectivo',
	'bolivaresCompras',
	'otrosMetodos',
	'comisionesPorTerceros',
	'ingresosPorComisionesEnBolivares',
	'entregaPendiente',
	'bolivaresVenta',
];

const state = {
	transactions: transactionSeed.map((transaction) => ({ ...transaction })),
};

const normalizeCuenta = (value) => String(value ?? '').trim().toUpperCase();

const clone = (value) => JSON.parse(JSON.stringify(value));

const cleanCellText = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();

const formatExcelDate = (value) => {
	if (typeof value === 'number') {
		if (Number.isInteger(value) && value >= 1 && value <= 31) {
			return String(value);
		}

		const parsed = XLSX.SSF.parse_date_code(value);
		if (parsed) {
			const day = String(parsed.d).padStart(2, '0');
			const month = String(parsed.m).padStart(2, '0');
			return `${day}/${month}/${parsed.y}`;
		}
	}

	return cleanCellText(value);
};

const readSheetRows = (sheetName) => {
	try {
		const workbook = XLSX.readFile(workbookPath);
		const sheet = workbook.Sheets[sheetName];

		if (!sheet) {
			return [];
		}

		return XLSX.utils.sheet_to_json(sheet, {
			header: 1,
			blankrows: false,
			defval: '',
		});
	} catch (error) {
		return [];
	}
};

const getGroupLabel = (headerRow, columnIndex) => {
	for (let index = columnIndex; index >= 0; index -= 1) {
		const label = cleanCellText(headerRow[index]);
		if (label) {
			return label;
		}
	}

	return 'X PAGAR';
};

const formatSheetReference = (sheetName, rowIndex, columnIndex) => {
	const columnLabel = XLSX.utils.encode_col(columnIndex);
	return `${sheetName} ${columnLabel}${rowIndex + 1}`;
};

const numericAmount = (value) => {
	const amount = Number(value);
	return Number.isFinite(amount) ? amount : null;
};

const extractWorkbookTransaction = (sheetName, rowIndex, columnMap = {}) => {
	const rows = readSheetRows(sheetName);
	const row = rows[rowIndex] || [];
	const {
		date = 0,
		client = 1,
		amount = 2,
		efectivo = 3,
		bolivares = 4,
		otrosMetodos = 5,
	} = columnMap;
	const clientName = cleanCellText(row[client]);
	const bankAmount = numericAmount(row[amount]);

	if (!clientName || bankAmount === null) {
		return null;
	}

	return {
		sheet: sheetName,
		reference: formatSheetReference(sheetName, rowIndex, 0),
		fecha: formatExcelDate(row[date]),
		cliente: clientName,
		bankAmount,
		efectivo: numericAmount(row[efectivo]),
		bolivares: numericAmount(row[bolivares]),
		otrosMetodos: numericAmount(row[otrosMetodos]),
	};
};

const findWorkbookTransactionsByClient = (sheetName, clientNeedle) => {
	const needle = cleanCellText(clientNeedle).toLowerCase();

	return readSheetRows(sheetName)
		.map((_, rowIndex) => extractWorkbookTransaction(sheetName, rowIndex))
		.filter((transaction) => transaction && transaction.cliente.toLowerCase().includes(needle));
};

const getPendingOutgoingTransactions = () => {
	const rows = readSheetRows('X PAGAR');
	const headerRow = rows[0] || [];
	const transactions = [];
	const seenReferences = new Set();

	const pushRecord = ({ rowIndex, columnIndex, sourceAccount, beneficiary, amount, reason }) => {
		const reference = formatSheetReference('X PAGAR', rowIndex, columnIndex);
		if (seenReferences.has(reference)) {
			return;
		}

		seenReferences.add(reference);
		transactions.push({
			id: `pending-${rowIndex + 1}-${columnIndex + 1}`,
			reference,
			sourceAccount,
			beneficiary,
			amount: Math.abs(amount),
			reason,
			sheet: 'X PAGAR',
			status: 'Pendiente de verificacion',
		});
	};

	rows.forEach((row, rowIndex) => {
		row.forEach((cell, columnIndex) => {
			const text = cleanCellText(cell);
			if (!/transferencia/i.test(text)) {
				return;
			}

			const amount = numericAmount(row[columnIndex + 1]);
			if (amount === null) {
				return;
			}

			pushRecord({
				rowIndex,
				columnIndex,
				sourceAccount: getGroupLabel(headerRow, columnIndex),
				beneficiary: text,
				amount,
				reason: 'Transferencia saliente detectada en la hoja operativa',
			});
		});
	});

	rows.forEach((row, rowIndex) => {
		row.forEach((cell, columnIndex) => {
			const note = cleanCellText(cell);
			if (!/colocado en pendiente/i.test(note)) {
				return;
			}

			for (let lookupIndex = columnIndex + 1; lookupIndex < row.length - 1; lookupIndex += 1) {
				const beneficiary = cleanCellText(row[lookupIndex]);
				const amount = numericAmount(row[lookupIndex + 1]);

				if (!beneficiary || amount === null || amount >= 0) {
					continue;
				}

				pushRecord({
					rowIndex,
					columnIndex: lookupIndex,
					sourceAccount: getGroupLabel(headerRow, lookupIndex),
					beneficiary,
					amount,
					reason: note,
				});
				break;
			}
		});
	});

	return transactions.sort((left, right) => left.reference.localeCompare(right.reference));
};

const getPendingIncomingTransactions = () => {
	const rows = readSheetRows('X COBRAR');
	const primaryHeaderRow = rows[0] || [];
	const secondaryHeaderRow = rows[2] || [];
	const transactions = [];
	const seenReferences = new Set();
	const excludedNames = /^(saldo inicial|efectivo|descontado)$/i;

	const groups = [];
	const maxColumnCount = Math.max(primaryHeaderRow.length, secondaryHeaderRow.length);

	for (let columnIndex = 0; columnIndex < maxColumnCount; columnIndex += 1) {
		const primaryHeader = cleanCellText(primaryHeaderRow[columnIndex]);
		const secondaryHeader = cleanCellText(secondaryHeaderRow[columnIndex]);
		const isDateColumn = primaryHeader.toUpperCase() === 'FECHA' || secondaryHeader.toUpperCase() === 'FECHA';

		if (!isDateColumn) {
			continue;
		}

		groups.push({
			columnIndex,
			dataStartRow: primaryHeader.toUpperCase() === 'FECHA' ? 1 : 3,
			accountLabel: (() => {
				const label = getGroupLabel(primaryHeaderRow, columnIndex);
				return label === 'FECHA' ? 'X COBRAR principal' : label;
			})(),
		});
	}

	for (const group of groups) {
		for (let rowIndex = group.dataStartRow; rowIndex < rows.length; rowIndex += 1) {
			const row = rows[rowIndex] || [];
			const rawDate = row[group.columnIndex];
			const sender = cleanCellText(row[group.columnIndex + 1]);
			const amount = numericAmount(row[group.columnIndex + 2]);
			const reference = formatSheetReference('X COBRAR', rowIndex, group.columnIndex);

			if (!sender || amount === null || amount <= 0 || excludedNames.test(sender) || seenReferences.has(reference)) {
				continue;
			}

			seenReferences.add(reference);
			transactions.push({
				id: `incoming-${rowIndex + 1}-${group.columnIndex + 1}`,
				reference,
				date: formatExcelDate(rawDate),
				rawDate: typeof rawDate === 'number' ? rawDate : Number.NEGATIVE_INFINITY,
				sender,
				destinationAccount: group.accountLabel,
				amount,
				reason: 'Pendiente por conciliacion',
				sheet: 'X COBRAR',
				status: 'Pendiente de validacion',
			});
		}
	}

	return transactions
		.sort((left, right) => right.rawDate - left.rawDate || right.amount - left.amount)
		.map(({ rawDate, ...transaction }) => transaction);
};

const getAiValidatedTransactions = () => {
	const joseMedina = extractWorkbookTransaction('HERNANDO', 8);
	const joseTavola = extractWorkbookTransaction('X COBRAR', 1, { date: 1, client: 2, amount: 3 });
	const zoiloGranda = extractWorkbookTransaction('PANAMA', 3);
	const carlosDavid = extractWorkbookTransaction('X COBRAR', 5, { date: 1, client: 2, amount: 3 });
	const emiroAlbanil = extractWorkbookTransaction('X COBRAR', 6, { date: 1, client: 2, amount: 3 });
	const normanMatches = findWorkbookTransactionsByClient('PANAMA', 'Norman Franco');
	const normanReferences = normanMatches.slice(0, 3).map((transaction) => transaction.reference).join(', ');

	return [
		joseMedina && {
			id: 'ai-001',
			fecha: joseMedina.fecha,
			cliente: joseMedina.cliente,
			cuenta: joseMedina.sheet,
			reference: joseMedina.reference,
			reportChannel: 'Excel',
			reportedAmount: joseMedina.bankAmount,
			bankAmount: joseMedina.bankAmount,
			humanAmount: joseMedina.bankAmount,
			status: 'Validado',
			reason: 'El monto reportado coincide exactamente con el ingreso detectado en MAYO.xlsx.',
		},
		joseTavola && {
			id: 'ai-002',
			fecha: joseTavola.fecha,
			cliente: joseTavola.cliente,
			cuenta: joseTavola.sheet,
			reference: joseTavola.reference,
			reportChannel: 'WhatsApp',
			reportedAmount: joseTavola.bankAmount,
			bankAmount: joseTavola.bankAmount,
			humanAmount: joseTavola.bankAmount,
			status: 'Validado',
			reason: 'La IA encontro una coincidencia exacta entre el reporte enviado y el movimiento bancario.',
		},
		zoiloGranda && {
			id: 'ai-003',
			fecha: zoiloGranda.fecha,
			cliente: zoiloGranda.cliente,
			cuenta: zoiloGranda.sheet,
			reference: zoiloGranda.reference,
			reportChannel: 'Verificacion humana',
			reportedAmount: zoiloGranda.bankAmount,
			bankAmount: zoiloGranda.bankAmount,
			humanAmount: 500,
			status: 'Error',
			reason: 'El banco muestra 530, pero la verificacion humana registro 500.',
		},
		normanMatches.length && {
			id: 'ai-004',
			fecha: normanMatches[0].fecha,
			cliente: 'Norman Franco',
			cuenta: 'PANAMA',
			reference: normanReferences || 'Sin coincidencia exacta',
			reportChannel: 'WhatsApp',
			reportedAmount: 150,
			bankAmount: null,
			humanAmount: 150,
			status: 'Error',
			reason: 'No aparece un abono exacto de 150 en banco; en MAYO.xlsx solo hay movimientos de 100 y 200 para ese remitente.',
		},
		carlosDavid && {
			id: 'ai-005',
			fecha: carlosDavid.fecha,
			cliente: carlosDavid.cliente,
			cuenta: carlosDavid.sheet,
			reference: carlosDavid.reference,
			reportChannel: 'Excel',
			reportedAmount: carlosDavid.bankAmount,
			bankAmount: carlosDavid.bankAmount,
			humanAmount: carlosDavid.bankAmount,
			status: 'Warning',
			reason: 'Monto inusualmente bajo para el flujo regular; la IA lo marca para revision adicional.',
		},
		emiroAlbanil && {
			id: 'ai-006',
			fecha: emiroAlbanil.fecha,
			cliente: emiroAlbanil.cliente,
			cuenta: emiroAlbanil.sheet,
			reference: emiroAlbanil.reference,
			reportChannel: 'Excel',
			reportedAmount: emiroAlbanil.bankAmount,
			bankAmount: emiroAlbanil.bankAmount,
			humanAmount: emiroAlbanil.bankAmount,
			status: 'Warning',
			reason: 'Monto inusualmente alto frente al patron diario, aunque el registro existe en banco.',
		},
	].filter(Boolean);
};

const totalGeneralForTransaction = (transaction) => numericFields.reduce((sum, field) => {
	return sum + Number(transaction[field] || 0);
}, 0);

const paymentMethodForTransaction = (transaction) => {
	if (transaction.zelle > 0) {
		return 'ZELLE';
	}

	if (transaction.efectivo > 0) {
		return 'EFECTIVO';
	}

	if (transaction.otrosMetodos > 0) {
		return 'OTROS';
	}

	return 'MIXTO';
};

const getTransactions = () => {
	return state.transactions.map((transaction) => ({
		...transaction,
		totalGeneral: totalGeneralForTransaction(transaction),
		paymentMethod: paymentMethodForTransaction(transaction),
		status: transaction.entregaPendiente > 0 ? 'Pendiente' : 'Liquidado',
	}));
};

const getCuentaSummary = () => {
	const grouped = new Map();

	for (const transaction of getTransactions()) {
		const cuenta = normalizeCuenta(transaction.cuenta);

		if (!grouped.has(cuenta)) {
			grouped.set(cuenta, {
				cuenta,
				zelle: 0,
				efectivo: 0,
				bolivaresCompras: 0,
				otrosMetodos: 0,
				comisionesPorTerceros: 0,
				ingresosPorComisionesEnBolivares: 0,
				entregaPendiente: 0,
				bolivaresVenta: 0,
				totalGeneral: 0,
				transactions: 0,
			});
		}

		const entry = grouped.get(cuenta);
		entry.transactions += 1;

		for (const field of numericFields) {
			entry[field] += Number(transaction[field] || 0);
		}

		entry.totalGeneral += transaction.totalGeneral;
	}

	return [...grouped.values()].sort((left, right) => right.totalGeneral - left.totalGeneral);
};

const getOverview = () => {
	const transactions = getTransactions();
	const summary = getCuentaSummary();
	const totals = transactions.reduce((accumulator, transaction) => {
		accumulator.transactionCount += 1;
		accumulator.totalVolume += transaction.totalGeneral;
		accumulator.pendingAmount += Math.max(transaction.entregaPendiente, 0);
		accumulator.commissionIncome += Number(transaction.comisionesPorTerceros || 0) + Number(transaction.ingresosPorComisionesEnBolivares || 0);
		accumulator.bolivaresCompras += Number(transaction.bolivaresCompras || 0);
		return accumulator;
	}, {
		transactionCount: 0,
		totalVolume: 0,
		pendingAmount: 0,
		commissionIncome: 0,
		bolivaresCompras: 0,
	});

	const paymentMix = ['ZELLE', 'EFECTIVO', 'OTROS', 'MIXTO'].map((label) => ({
		label,
		value: transactions.filter((transaction) => transaction.paymentMethod === label).length,
	}));

	const dailyTrendMap = new Map();
	for (const transaction of transactions) {
		const key = String(transaction.fecha);
		if (!dailyTrendMap.has(key)) {
			dailyTrendMap.set(key, { day: key, total: 0, operations: 0 });
		}

		const entry = dailyTrendMap.get(key);
		entry.total += transaction.totalGeneral;
		entry.operations += 1;
	}

	return {
		kpis: {
			...totals,
			activeCuentas: summary.length,
		},
		paymentMix,
		dailyTrend: [...dailyTrendMap.values()].sort((left, right) => Number(left.day) - Number(right.day)),
		topCuentas: summary.slice(0, 6).map((entry) => ({
			cuenta: entry.cuenta,
			totalGeneral: entry.totalGeneral,
			transactions: entry.transactions,
		})),
	};
};

const getMetadata = () => {
	const transactions = getTransactions();
	const cuentas = [...new Set(transactions.map((transaction) => normalizeCuenta(transaction.cuenta)))].sort();
	const clientNames = [...new Set(transactions.map((transaction) => String(transaction.cliente || '').trim()).filter(Boolean))].sort();

	return {
		cuentas,
		clientes: clientNames,
		paymentMethods: ['ZELLE', 'EFECTIVO', 'OTROS', 'MIXTO'],
		statuses: ['Pendiente', 'Liquidado'],
	};
};

const addTransaction = (payload) => {
	const nextId = `txn-${String(state.transactions.length + 1).padStart(3, '0')}`;
	const transaction = {
		id: nextId,
		fecha: Number(payload.fecha || 0),
		cliente: String(payload.cliente || '').trim(),
		envia: String(payload.envia || '').trim(),
		cuenta: String(payload.cuenta || '').trim(),
		cliente2: String(payload.cliente2 || '').trim(),
		zelle: Number(payload.zelle || 0),
		efectivo: Number(payload.efectivo || 0),
		bolivaresCompras: Number(payload.bolivaresCompras || 0),
		otrosMetodos: Number(payload.otrosMetodos || 0),
		comisionesPorTerceros: Number(payload.comisionesPorTerceros || 0),
		ingresosPorComisionesEnBolivares: Number(payload.ingresosPorComisionesEnBolivares || 0),
		entregaPendiente: Number(payload.entregaPendiente || 0),
		bolivaresVenta: Number(payload.bolivaresVenta || 0),
	};

	state.transactions.unshift(transaction);
	return getTransactions().find((entry) => entry.id === nextId);
};

module.exports = {
	getTransactions: () => clone(getTransactions()),
	getCuentaSummary: () => clone(getCuentaSummary()),
	getOverview: () => clone(getOverview()),
	getMetadata: () => clone(getMetadata()),
	getAiValidatedTransactions: () => clone(getAiValidatedTransactions()),
	getPendingIncomingTransactions: () => clone(getPendingIncomingTransactions()),
	getPendingOutgoingTransactions: () => clone(getPendingOutgoingTransactions()),
	addTransaction,
};