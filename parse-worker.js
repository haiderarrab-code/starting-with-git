/* parse-worker.js — disposable Excel/CSV parser.
   Parses one buffer and posts the parsed tables DIRECTLY to data-worker via a
   MessageChannel port (rows never pass through the renderer heap).
   `maxRows` limits how many rows are read per sheet (partial import = fast).
*/
'use strict';

importScripts('vendor/xlsx.full.min.js');

self.postMessage({ type: 'ready' });

self.onmessage = function (e) {
  const { id, buffer, name, port, maxRows } = e.data;
  try {
    const opts = { type: 'array', cellDates: true, WTF: false };
    // sheetRows limits parsing to the first N data rows (+ header) → faster, less memory.
    if (maxRows && maxRows > 0) opts.sheetRows = maxRows + 1;

    const wb = XLSX.read(buffer, opts);
    const tables = wb.SheetNames.map(sheetName => {
      try {
        const json = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
        if (!json.length) return null;
        const rows = (maxRows && maxRows > 0) ? json.slice(0, maxRows) : json;
        const columns = Object.keys(json[0]);
        return { name: sheetName, columns, rows };
      } catch { return null; }
    }).filter(Boolean);

    port.postMessage({ tables, error: null });
    port.close();
    self.postMessage({ id, error: null });
  } catch (err) {
    try { port.postMessage({ tables: null, error: err.message }); port.close(); } catch (_) {}
    self.postMessage({ id, error: err.message });
  }
};
