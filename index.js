/*
 * Test server for experimenting and observing Microsoft Office URL link handling
 *
 * This server allows for experimenting with server-side redirects and inspect the 
 * requests made by the client, Microsoft Office or otherwise.
 *
 * Note that this logic can be used to perform reflection and cross-site request 
 * forgery attacks, so NEVER DEPLOY THIS SERVER OUTSIDE OF LOCALHOST!
 */

const http = require('http');
const fs = require('fs');
const path = require('path');

const hostname = '127.0.0.1';
const port = 3000;

// regex to identify a user agent as a Microsoft Office product
const isOfficeUA = new RegExp("(Microsoft Office)|(ms-office)");

// page to present as a help page
const helpContentFileName = path.join(__dirname, 'index.html');

// we're handling all requests in one huge code block
// in order to manage a variety of combinations of 
// requests for redirect, etc.
const server = http.createServer((req, res) => {

	const url = new URL(req.url, `http://${req.headers.host}`);
	const path = url.pathname;
	const params = url.searchParams;
	const headers = req.headers;
	const ua = headers['user-agent'];

	// assume 200 response
	var status = 200;
	var content = `Default response for ${path}`;
	var contentType = 'text/plain';
	var debug = true;
	var redirectUrl = '';

	// return the help page only if the root doc with no parameters is requested
	if (path === '/' && params.entries().next().done) {
		
		res.setHeader('Content-Type', 'text/html');
		res.statusCode = 200;

		fs.readFile(helpContentFileName, function (err,data) {
			if (err) {
				res.end(`<h1>Error</h1><p>Error loading help page from ${helpContentFileName}</p>`);
				return;
			}
			res.end(data);
		});
		
		// bypass any other processing
		return;
	}

	// ignore favicon and return Not Found
	if (path == "/favicon.ico") {
		debug = false;
		status = 404;
	}

	if (debug) {
		console.log(req.method + "\t" + path + "\t" + url.search + "\t" + ua);
	}

	// redirect if told so by the client...
	if (params.has('redirect')) {
		redirectUrl = params.get('redirect');
		status = 302;	// moved temporarily
	}

	if (params.has('contentType')) {
		contentType = params.get('contentType');
	}

	if (params.has('content')) {
		content = params.get('content');
	}

	if (params.has('status')) {
		status = params.get('status');
	}

	if (params.has('bypassOffice')) {
		if (isOfficeUA.test(ua)) {
			
			contentType = 'text/plain';
			status = 200;
			content = 'Go away, Office. You\'re drunk!';
			redirectUrl = '';
		}
		else {
			content = 'Office bypass requested but no Office detected';
		}
	}
	
	res.setHeader('Content-Type', contentType);
	res.statusCode = status;

	debugOut = " => "+status;
	if (status >= 300 && status <= 399 && redirectUrl.length > 0) {
		res.setHeader('Location', redirectUrl);
		debugOut += `\tLocation: ${redirectUrl}`;
	}

	debugOut += `\t${content}`;

	if (debug) {
		console.log(debugOut);
	}

	res.end(content);
});

server.listen(port, hostname, () => {
	console.log(`Server running at http://${hostname}:${port}/`);
});

