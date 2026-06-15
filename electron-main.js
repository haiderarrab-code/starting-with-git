'use strict';

const { app, BrowserWindow, Menu, shell } = require('electron');
const express = require('express');
const path = require('path');
const http = require('http');

let server = null;
let serverPort = 0;

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

  const win = new BrowserWindow({
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
