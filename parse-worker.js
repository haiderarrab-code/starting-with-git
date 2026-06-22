/* parse-worker.js — disposable Excel/CSV parser.
   Parses one buffer at a time and posts results DIRECTLY to data-worker
   via a MessageChannel port (rows never pass through the renderer heap).
*/
'use strict';

importScripts('vendor/xlsx.full.min.js');

const ROW_LIMIT = 120000;

self.postMessage({ type: 'ready' });

self.onmessage = function (e) {
  const { id, buffer, name, port } = e.data;
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

    // Send parsed tables DIRECTLY to data-worker via the port (zero renderer involvement).
    port.postMessage({ tables, error: null });
    port.close();
    self.postMessage({ id, error: null });           // tell renderer parse is done
  } catch (err) {
    port.postMessage({ tables: null, error: err.message });
    port.close();
    self.postMessage({ id, error: err.message });
  }
};
