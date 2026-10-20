#!/usr/bin/env node
/**
 * Static file server plus a small Microsoft Graph mail API.
 * Mail is sent with the OAuth2 client-credentials flow (application permissions), so the client
 * secret lives only on this server and is never exposed to the browser.
 *
 *   node server.js [port]
 */
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = __dirname;

function loadEnvFile(file) {
  let text;
  try {
    text = fs.readFileSync(file, 'utf8');
  } catch {
    return;
  }
  for (const line of text.split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (!match) {
      continue;
    }
    let value = match[2];
    if (/^(["']).*\1$/.test(value)) {
      value = value.slice(1, -1);
    }
    if (!(match[1] in process.env)) {
      process.env[match[1]] = value;
    }
  }
}

loadEnvFile(path.join(ROOT, '.env'));

const PORT = Number(process.argv[2] || process.env.PORT || 3000);
const HOST = process.env.HOST || '127.0.0.1';
const GRAPH_URL = 'https://graph.microsoft.com/v1.0';
const MAX_BODY_BYTES = 2 * 1024 * 1024;
const SENDER_PATTERN = /^[^\s@/?#%;,]+@[^\s@/?#%;,]+$|^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const cfg = {
  tenantId: (process.env.TENANT_ID || '').trim(),
  clientId: (process.env.CLIENT_ID || '').trim(),
  clientSecret: (process.env.CLIENT_SECRET || '').trim(),
  defaultSender: (process.env.DEFAULT_SENDER || '').trim(),
  allowedSenders: (process.env.ALLOWED_SENDERS || '')
    .split(/[;,]/)
    .map((address) => address.trim().toLowerCase())
    .filter(Boolean)
};

const isConfigured = () => Boolean(cfg.tenantId && cfg.clientId && cfg.clientSecret);

// Only these paths are served, so .env and server.js can never be downloaded.
const PUBLIC_DIRS = new Set(['css', 'js', 'samples']);

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.map': 'application/json; charset=utf-8'
};

class ApiError extends Error {
  constructor(status, code, message, requestId) {
    super(message);
    this.status = status;
    this.code = code;
    this.requestId = requestId;
  }
}

function sendJson(res, status, payload) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(payload));
}

let tokenCache = { value: null, expiresAt: 0 };

async function getAppToken() {
  if (tokenCache.value && Date.now() < tokenCache.expiresAt - 60000) {
    return tokenCache.value;
  }
  const response = await fetch(
    'https://login.microsoftonline.com/' + encodeURIComponent(cfg.tenantId) + '/oauth2/v2.0/token',
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: cfg.clientId,
        client_secret: cfg.clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      })
    }
  );
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new ApiError(
      502,
      payload.error || 'TokenRequestFailed',
      payload.error_description || 'Could not acquire an application token.'
    );
  }
  tokenCache = { value: payload.access_token, expiresAt: Date.now() + Number(payload.expires_in || 3600) * 1000 };
  return tokenCache.value;
}

async function graphRequest(method, graphPath, body) {
  const token = await getAppToken();
  const response = await fetch(GRAPH_URL + graphPath, {
    method,
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body)
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => null);
    throw new ApiError(
      response.status,
      (payload && payload.error && payload.error.code) || 'GraphError',
      (payload && payload.error && payload.error.message) || response.statusText,
      response.headers.get('request-id') || undefined
    );
  }
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

const allowedHosts = new Set(['localhost:' + PORT, '127.0.0.1:' + PORT]);

// Blocks DNS-rebinding and cross-site calls, since this API can send mail as any allowed mailbox.
function assertLocalRequest(req) {
  const origin = req.headers.origin;
  if (!allowedHosts.has(req.headers.host) || (origin && !allowedHosts.has(origin.replace(/^https?:\/\//, '')))) {
    throw new ApiError(403, 'Forbidden', 'Requests are only accepted from this application.');
  }
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size <= MAX_BODY_BYTES) {
        chunks.push(chunk);
      }
    });
    req.on('end', () => {
      if (size > MAX_BODY_BYTES) {
        reject(new ApiError(413, 'PayloadTooLarge', 'Request body is too large.'));
        return;
      }
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString('utf8')));
      } catch {
        reject(new ApiError(400, 'InvalidJson', 'Request body is not valid JSON.'));
      }
    });
    req.on('error', reject);
  });
}

function resolveSender(from) {
  const sender = String(from || cfg.defaultSender || '').trim();
  if (!SENDER_PATTERN.test(sender)) {
    throw new ApiError(400, 'InvalidSender', 'A valid sender mailbox (email address or user object ID) is required.');
  }
  if (cfg.allowedSenders.length && !cfg.allowedSenders.includes(sender.toLowerCase())) {
    throw new ApiError(403, 'SenderNotAllowed', sender + ' is not listed in ALLOWED_SENDERS.');
  }
  return sender;
}

function toRecipients(list) {
  if (!Array.isArray(list)) {
    return [];
  }
  return list
    .map((entry) => String((entry && entry.emailAddress && entry.emailAddress.address) || '').trim())
    .filter(Boolean)
    .map((address) => ({ emailAddress: { address } }));
}

// Rebuilds the message from known fields so callers cannot pass arbitrary Graph properties.
function sanitizeMessage(message) {
  if (!message || typeof message !== 'object') {
    throw new ApiError(400, 'InvalidMessage', 'A message object is required.');
  }
  const to = toRecipients(message.toRecipients);
  const cc = toRecipients(message.ccRecipients);
  const html = message.body && message.body.content;
  if (!to.length) {
    throw new ApiError(400, 'InvalidMessage', 'At least one recipient is required.');
  }
  if (typeof message.subject !== 'string' || !message.subject.trim()) {
    throw new ApiError(400, 'InvalidMessage', 'A subject is required.');
  }
  if (typeof html !== 'string' || !html) {
    throw new ApiError(400, 'InvalidMessage', 'A message body is required.');
  }
  const clean = {
    subject: message.subject.trim(),
    body: { contentType: 'HTML', content: html },
    toRecipients: to
  };
  if (cc.length) {
    clean.ccRecipients = cc;
  }
  return clean;
}

async function handleApi(req, res, url) {
  assertLocalRequest(req);

  if (req.method === 'GET' && url.pathname === '/api/status') {
    sendJson(res, 200, {
      configured: isConfigured(),
      tenantId: cfg.tenantId,
      clientId: cfg.clientId,
      defaultSender: cfg.defaultSender,
      allowedSenders: cfg.allowedSenders
    });
    return;
  }

  if (req.method === 'POST' && (url.pathname === '/api/send' || url.pathname === '/api/draft')) {
    if (!(req.headers['content-type'] || '').startsWith('application/json')) {
      throw new ApiError(415, 'UnsupportedMediaType', 'Content-Type must be application/json.');
    }
    if (!isConfigured()) {
      throw new ApiError(503, 'NotConfigured', 'Set TENANT_ID, CLIENT_ID and CLIENT_SECRET in .env and restart the server.');
    }
    const body = await readJson(req);
    const sender = resolveSender(body.from);
    const message = sanitizeMessage(body.message);
    const mailbox = '/users/' + encodeURIComponent(sender);

    if (url.pathname === '/api/send') {
      await graphRequest('POST', mailbox + '/sendMail', {
        message,
        saveToSentItems: body.saveToSentItems !== false
      });
      sendJson(res, 200, { ok: true, from: sender });
    } else {
      const draft = await graphRequest('POST', mailbox + '/messages', message);
      sendJson(res, 200, { ok: true, from: sender, id: draft.id });
    }
    return;
  }

  throw new ApiError(404, 'NotFound', 'Unknown API route.');
}

function serveStatic(res, pathname) {
  let relative;
  try {
    relative = decodeURIComponent(pathname);
  } catch {
    res.writeHead(400).end('Bad request');
    return;
  }
  relative = relative === '/' ? 'index.html' : relative.replace(/^\/+/, '');
  const filePath = path.resolve(ROOT, relative);
  const [topLevel] = path.relative(ROOT, filePath).split(path.sep);

  if (!filePath.startsWith(ROOT + path.sep) || (relative !== 'index.html' && !PUBLIC_DIRS.has(topLevel))) {
    res.writeHead(403).end('Forbidden');
    return;
  }

  fs.readFile(filePath, (error, data) => {
    if (error) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }).end('Not found');
      return;
    }
    res.writeHead(200, {
      'Content-Type': MIME[path.extname(filePath).toLowerCase()] || 'application/octet-stream',
      'Cache-Control': 'no-cache'
    });
    res.end(data);
  });
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, 'http://localhost');
    if (url.pathname.startsWith('/api/')) {
      await handleApi(req, res, url);
    } else {
      serveStatic(res, url.pathname);
    }
  } catch (error) {
    if (error instanceof ApiError) {
      sendJson(res, error.status, { error: { code: error.code, message: error.message }, requestId: error.requestId });
    } else {
      console.error(error);
      sendJson(res, 500, { error: { code: 'ServerError', message: 'Internal server error.' } });
    }
  }
});

server.listen(PORT, HOST, () => {
  console.log('Actionable Message Sender running at http://localhost:' + PORT);
  console.log(
    isConfigured()
      ? 'App-only (client credentials) auth is configured for client ' + cfg.clientId + '.'
      : 'Missing TENANT_ID, CLIENT_ID or CLIENT_SECRET - copy .env.example to .env and fill them in.'
  );
});
