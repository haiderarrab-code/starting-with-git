'use strict';

const { app, BrowserWindow, Menu, shell } = require('electron');
const path = require('path');
const http = require('http');
const fs = require('fs');

// Increase JS heap limit for large Excel imports
app.commandLine.appendSwitch('js-flags', '--max-old-space-size=4096');
// GPU acceleration causes black-screen/crash on some Windows setups; the UI
// is a simple table view so software rendering is plenty fast.
app.disableHardwareAcceleration();

let server = null;
let serverPort = 0;
let win = null;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'text/javascript; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.wasm': 'application/wasm',
  '.svg':  'image/svg+xml',
  '.png':  'image/png',
  '.woff': 'font/woff',
  '.woff2':'font/woff2',
};

// Internal static server using Node's built-in http (NO dependencies) so Web
// Workers + wasm load over http like a normal site. Avoids any missing-module
// black screen when packaged.
function startServer() {
  return new Promise(resolve => {
    server = http.createServer((req, res) => {
      try {
        let urlPath = decodeURIComponent((req.url || '/').split('?')[0]);
        if (urlPath === '/' || urlPath === '') urlPath = '/index.html';
        const safePath = path.normalize(urlPath).replace(/^(\.\.[/\\])+/, '');
        const filePath = path.join(__dirname, safePath);
        if (!filePath.startsWith(__dirname)) { res.writeHead(403); res.end('forbidden'); return; }
        fs.readFile(filePath, (err, data) => {
          if (err) { res.writeHead(404); res.end('not found'); return; }
          const ext = path.extname(filePath).toLowerCase();
          res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
          res.end(data);
        });
      } catch (e) {
        res.writeHead(500); res.end('error');
      }
    });
    server.listen(0, '127.0.0.1', () => {
      serverPort = server.address().port;
      resolve(serverPort);
    });
  });
}

async function createWindow() {
  const port = await startServer();

  win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 900,
    minHeight: 600,
    backgroundColor: '#0f172a',
    title: 'نظام البحث الموحّد في قواعد البيانات',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  // Hide the default menu bar (cleaner desktop-app feel)
  Menu.setApplicationMenu(null);

  win.loadURL(`http://127.0.0.1:${port}/index.html`);

  // If renderer crashes, reload instead of showing blank screen
  win.webContents.on('render-process-gone', (event, details) => {
    console.error('Renderer crashed:', details.reason);
    setTimeout(() => win.loadURL(`http://127.0.0.1:${port}/index.html`), 1000);
  });

  win.webContents.on('unresponsive', () => {
    // Give it 10 seconds to recover before reloading
    setTimeout(() => {
      if (win && !win.isDestroyed()) win.webContents.reload();
    }, 10000);
  });

  // Open external links in the system browser, not inside the app
  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (server) server.close();
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
