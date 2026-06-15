/* xlsx-worker.js — parses Excel/CSV off the main thread */
try {
  importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
  self.postMessage({ ready: true }); // signal to main thread that XLSX loaded
} catch (e) {
  // CDN failed — tell main thread to use fallback
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
          const ws   = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
          if (!json.length) return null;
          return { name: sheetName, columns: Object.keys(json[0]), rows: json };
        } catch { return null; }
      })
      .filter(Boolean);
    self.postMessage({ id, name, tables, error: null });
  } catch (err) {
    self.postMessage({ id, name, tables: null, error: err.message });
  }
};
