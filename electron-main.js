'use strict';

const { app, BrowserWindow, Menu, shell } = require('electron');
const express = require('express');
const path = require('path');
const http = require('http');

// Increase JS heap limit for large Excel imports
app.commandLine.appendSwitch('js-flags', '--max-old-space-size=4096');

let server = null;
let serverPort = 0;
let win = null;

// Start an internal static server so Web Workers + wasm load like a normal site
function startServer() {
  return new Promise(resolve => {
    const expApp = express();
    expApp.use(express.static(__dirname));
    server = http.createServer(expApp);
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
