'use strict';

const { app, BrowserWindow, Menu, shell, protocol } = require('electron');
const path = require('path');
const fs = require('fs');

// Increase JS heap limit for large Excel imports
app.commandLine.appendSwitch('js-flags', '--max-old-space-size=4096');
// GPU acceleration causes black-screen/crash on some Windows setups; the UI
// is a simple table view so software rendering is plenty fast.
app.disableHardwareAcceleration();

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

const APP_ORIGIN = 'app://local';

// A custom scheme gives a STABLE origin (app://local) with no port. This is
// essential: IndexedDB is keyed by origin, so a fixed origin lets imported data
// persist across restarts. (A random http port would change the origin each
// launch and lose all stored data.) It's also registered as standard+secure so
// Web Workers and IndexedDB work in a secure context.
protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true } },
]);

function registerProtocol() {
  protocol.handle('app', async (request) => {
    try {
      const url = new URL(request.url);
      let p = decodeURIComponent(url.pathname);
      if (!p || p === '/') p = '/index.html';
      const safe = path.normalize(p).replace(/^(\.\.[/\\])+/, '');
      const filePath = path.join(__dirname, safe);
      if (!filePath.startsWith(__dirname)) return new Response('forbidden', { status: 403 });
      const data = await fs.promises.readFile(filePath);
      const ext = path.extname(filePath).toLowerCase();
      return new Response(data, { headers: { 'content-type': MIME[ext] || 'application/octet-stream' } });
    } catch (e) {
      return new Response('not found', { status: 404 });
    }
  });
}

function createWindow() {
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

  Menu.setApplicationMenu(null);
  win.loadURL(`${APP_ORIGIN}/index.html`);

  // If renderer crashes, reload instead of showing a blank screen
  win.webContents.on('render-process-gone', (event, details) => {
    console.error('Renderer crashed:', details.reason);
    setTimeout(() => win.loadURL(`${APP_ORIGIN}/index.html`), 1000);
  });

  win.webContents.on('unresponsive', () => {
    setTimeout(() => { if (win && !win.isDestroyed()) win.webContents.reload(); }, 10000);
  });

  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
}

app.whenReady().then(() => {
  registerProtocol();
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
