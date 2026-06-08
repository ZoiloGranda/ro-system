const http = require('http');
const fs = require('fs');
const path = require('path');

const host = '127.0.0.1';
const port = process.env.PORT || 3000;
const chartPath = path.join(__dirname, 'files', 'cuenta-totals-chart.html');

const server = http.createServer((request, response) => {
	if (!['GET', 'HEAD'].includes(request.method) || request.url !== '/') {
		response.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
		response.end('Not found');
		return;
	}

	fs.readFile(chartPath, 'utf8', (error, html) => {
		if (error) {
			response.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
			response.end('Unable to load chart file.');
			return;
		}

		response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
		response.end(request.method === 'HEAD' ? undefined : html);
	});
});

server.listen(port, host, () => {
	console.log(`Serving cuenta-totals-chart.html at http://${host}:${port}`);
});