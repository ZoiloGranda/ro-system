const http = require('http');
const fs = require('fs');
const path = require('path');

const {
	getTransactions,
	getCuentaSummary,
	getOverview,
	getMetadata,
	addTransaction,
} = require('./mockData');

const host = '127.0.0.1';
const port = process.env.PORT || 3000;
const publicDir = path.join(__dirname, 'public');

const contentTypes = {
	'.css': 'text/css; charset=utf-8',
	'.html': 'text/html; charset=utf-8',
	'.js': 'application/javascript; charset=utf-8',
	'.json': 'application/json; charset=utf-8',
	'.svg': 'image/svg+xml',
};

const sendJson = (response, statusCode, payload) => {
	response.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' });
	response.end(JSON.stringify(payload));
};

const sendText = (response, statusCode, message) => {
	response.writeHead(statusCode, { 'Content-Type': 'text/plain; charset=utf-8' });
	response.end(message);
};

const readBody = (request) => new Promise((resolve, reject) => {
	let body = '';

	request.on('data', (chunk) => {
		body += chunk;
		if (body.length > 1024 * 1024) {
			reject(new Error('Request body too large'));
			request.destroy();
		}
	});

	request.on('end', () => resolve(body));
	request.on('error', reject);
});

const serveStaticFile = (requestPath, response) => {
	const normalizedPath = requestPath === '/' ? '/index.html' : requestPath;
	const filePath = path.normalize(path.join(publicDir, normalizedPath));

	if (!filePath.startsWith(publicDir)) {
		sendText(response, 403, 'Forbidden');
		return;
	}

	fs.readFile(filePath, (error, content) => {
		if (error) {
			sendText(response, error.code === 'ENOENT' ? 404 : 500, error.code === 'ENOENT' ? 'Not found' : 'Unable to load file');
			return;
		}

		const extension = path.extname(filePath).toLowerCase();
		response.writeHead(200, {
			'Content-Type': contentTypes[extension] || 'application/octet-stream',
		});
		response.end(content);
	});
};

const server = http.createServer((request, response) => {
	const url = new URL(request.url, `http://${request.headers.host || `${host}:${port}`}`);

	if (request.method === 'GET' && url.pathname === '/api/transactions') {
		sendJson(response, 200, { data: getTransactions() });
		return;
	}

	if (request.method === 'GET' && url.pathname === '/api/summary/cuentas') {
		sendJson(response, 200, { data: getCuentaSummary() });
		return;
	}

	if (request.method === 'GET' && url.pathname === '/api/reports/overview') {
		sendJson(response, 200, { data: getOverview() });
		return;
	}

	if (request.method === 'GET' && url.pathname === '/api/filters/metadata') {
		sendJson(response, 200, { data: getMetadata() });
		return;
	}

	if (request.method === 'POST' && url.pathname === '/api/transactions') {
		readBody(request)
			.then((body) => {
				const payload = body ? JSON.parse(body) : {};
				if (!payload.cuenta || !payload.cliente2 || !payload.fecha) {
					sendJson(response, 400, {
						error: 'fecha, cuenta, and cliente2 are required for the prototype entry flow.',
					});
					return;
				}

				const transaction = addTransaction(payload);
				sendJson(response, 201, {
					data: transaction,
					overview: getOverview(),
					summary: getCuentaSummary(),
				});
			})
			.catch((error) => {
				sendJson(response, 400, { error: error.message || 'Invalid request body' });
			});
		return;
	}

	if (['GET', 'HEAD'].includes(request.method)) {
		serveStaticFile(url.pathname, response);
		return;
	}

	sendText(response, 404, 'Not found');
});

server.listen(port, host, () => {
	console.log(`Serving prototype dashboard at http://${host}:${port}`);
});