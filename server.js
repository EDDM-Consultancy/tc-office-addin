// server.js — hosts the add-in task pane over HTTPS and proxies Teamcenter.
//
//   /                 -> src/ (task pane), /assets, /manifest.xml
//   /config.json      -> "client" section of config.json
//   /vendor/jszip.min.js
//   /tc/*             -> <tcWebUrl>/<tcContext>/*   (SOA JsonRestServices, session cookie)
//   /fms/*            -> <fscUrl>/fms/*             (FMS file download/upload)
//
// Usage:  node server.js [path/to/config.json]
// Env overrides: PORT, TC_WEB_URL, TC_CONTEXT, FSC_URL, CERT_FILE, KEY_FILE, HTTP=1 (plain http, tests only)

const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');
const https = require('https');
const express = require('express');
const { createProxyMiddleware } = require('http-proxy-middleware');

function loadConfig(file) {
  const cfg = JSON.parse(fs.readFileSync(file, 'utf8'));
  const s = cfg.server;
  s.port = Number(process.env.PORT || s.port || 3000);
  s.tcWebUrl = process.env.TC_WEB_URL || s.tcWebUrl;
  s.tcContext = (process.env.TC_CONTEXT || s.tcContext || 'tc').replace(/^\/|\/$/g, '');
  s.fscUrl = process.env.FSC_URL || s.fscUrl;
  s.certFile = process.env.CERT_FILE || s.certFile;
  s.keyFile = process.env.KEY_FILE || s.keyFile;
  return cfg;
}

function createApp(cfg) {
  const s = cfg.server;
  const app = express();
  app.disable('x-powered-by');

  // --- Teamcenter web tier proxy (keeps browser same-origin, rewrites cookie path)
  app.use(createProxyMiddleware({
    target: s.tcWebUrl,
    changeOrigin: true,
    pathFilter: '/tc',
    pathRewrite: (p) => p.replace(/^\/tc/, `/${s.tcContext}`),
    cookiePathRewrite: { [`/${s.tcContext}`]: '/tc', '*': '/tc' },
    cookieDomainRewrite: '',
    proxyTimeout: 300000,
  }));

  // --- FSC proxy (file transfer)
  app.use(createProxyMiddleware({
    target: s.fscUrl,
    changeOrigin: true,
    pathFilter: '/fms',
    proxyTimeout: 600000,
  }));

  app.get('/config.json', (_req, res) => res.json(cfg.client));
  app.get('/vendor/jszip.min.js', (_req, res) => res.sendFile(require.resolve('jszip/dist/jszip.min.js')));
  app.get('/manifest.xml', (_req, res) => res.sendFile(path.join(__dirname, 'manifest.xml')));
  app.use('/assets', express.static(path.join(__dirname, 'assets')));
  app.use('/', express.static(path.join(__dirname, 'src'), { index: 'taskpane.html' }));
  return app;
}

function tlsOptions(s) {
  const dev = path.join(os.homedir(), '.office-addin-dev-certs');
  const cert = s.certFile || path.join(dev, 'localhost.crt');
  const key = s.keyFile || path.join(dev, 'localhost.key');
  if (!fs.existsSync(cert) || !fs.existsSync(key)) {
    throw new Error(`TLS certificate not found (${cert}). Run "npx office-addin-dev-certs install" or set certFile/keyFile in config.json.`);
  }
  return { cert: fs.readFileSync(cert), key: fs.readFileSync(key) };
}

if (require.main === module) {
  const cfg = loadConfig(process.argv[2] || path.join(__dirname, 'config.json'));
  const app = createApp(cfg);
  const s = cfg.server;
  const server = process.env.HTTP === '1' ? http.createServer(app) : https.createServer(tlsOptions(s), app);
  server.listen(s.port, () => {
    console.log(`Teamcenter Office add-in on ${process.env.HTTP === '1' ? 'http' : 'https'}://localhost:${s.port}`);
    console.log(`  /tc  -> ${s.tcWebUrl}/${s.tcContext}`);
    console.log(`  /fms -> ${s.fscUrl}/fms`);
  });
}

module.exports = { createApp, loadConfig };
