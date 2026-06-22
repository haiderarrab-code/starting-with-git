/* parse-worker.js — disposable Excel/CSV parser.
   This worker ONLY parses one buffer at a time and returns full rows.
   It is intentionally separate from data-worker.js (the storage worker)
   so that if a single corrupt / huge / password-protected file makes
   XLSX.read hang, the renderer can terminate THIS worker and recreate it
   without losing any data already stored in data-worker.
*/
'use strict';

importScripts('vendor/xlsx.full.min.js');

const ROW_LIMIT = 120000;

self.postMessage({ type: 'ready' });

self.onmessage = function (e) {
  const { id, buffer, name } = e.data;
  try {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true, WTF: false });
    const tables = wb.SheetNames.map(sheetName => {
      try {
        const json = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
        if (!json.length) return null;
        const rows    = json.slice(0, ROW_LIMIT);
        const columns = Object.keys(json[0]);
        return { name: sheetName, columns, rows };
      } catch { return null; }
    }).filter(Boolean);
    self.postMessage({ id, name, tables, error: null });
  } catch (err) {
    self.postMessage({ id, name, tables: null, error: err.message });
  }
};
