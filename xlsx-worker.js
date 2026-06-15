/* xlsx-worker.js — parses Excel/CSV off the main thread */
const ROW_LIMIT = 10000; // max rows per sheet transferred to main thread

try {
  importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
  self.postMessage({ ready: true });
} catch (e) {
  self.postMessage({ ready: false, error: e.message });
  self.close();
}

self.onmessage = function (e) {
  const { id, buffer, name } = e.data;
  try {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, WTF: false });
    const tables = workbook.SheetNames
      .map(sheetName => {
        try {
          const ws    = workbook.Sheets[sheetName];
          const json  = XLSX.utils.sheet_to_json(ws, { defval: '' });
          if (!json.length) return null;
          const columns = Object.keys(json[0]);
          const total   = json.length;
          const rows    = json.slice(0, ROW_LIMIT);
          return { name: sheetName, columns, rows, total };
        } catch { return null; }
      })
      .filter(Boolean);
    self.postMessage({ id, name, tables, error: null });
  } catch (err) {
    self.postMessage({ id, name, tables: null, error: err.message });
  }
};
