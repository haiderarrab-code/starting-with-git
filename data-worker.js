/* data-worker.js
   Holds ALL row data and search indexes.
   The renderer process never stores actual rows — only metadata.
   This keeps the renderer heap small and prevents OOM crashes.
*/
'use strict';

importScripts('vendor/xlsx.full.min.js');

// sourceId -> [{name, columns, rows, _index}]
const store = new Map();
const ROW_LIMIT = 120_000;

self.postMessage({ type: 'ready' });

self.onmessage = function(e) {
  const msg = e.data;
  switch (msg.op) {
    case 'parse':       return handleParse(msg);
    case 'parse_word':  return handleParseWord(msg);
    case 'rows':        return handleGetRows(msg);
    case 'search':      return handleSearch(msg);
    case 'remove':  store.delete(msg.sourceId); self.postMessage({ op: 'removed', id: msg.id }); break;
    case 'clear':   store.clear();               self.postMessage({ op: 'cleared',  id: msg.id }); break;
  }
};

// ── Parse an Excel/CSV buffer and store rows in worker memory ──
function handleParse({ id, sourceId, buffer, name }) {
  try {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true, WTF: false });
    const tables = wb.SheetNames.map(sheetName => {
      try {
        const json = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
        if (!json.length) return null;
        const rows    = json.slice(0, ROW_LIMIT);
        const columns = Object.keys(json[0]);
        const _index  = rows.map(r => Object.values(r).join('\x00').toLowerCase());
        return { name: sheetName, columns, rows, _index };
      } catch { return null; }
    }).filter(Boolean);

    store.set(sourceId, tables);

    // Only send metadata back — NO rows in this message
    self.postMessage({
      op: 'parsed', id, sourceId, name, error: null,
      tables: tables.map(t => ({ name: t.name, columns: t.columns, totalRows: t.rows.length })),
    });
  } catch (err) {
    self.postMessage({ op: 'parsed', id, sourceId, name, error: err.message, tables: null });
  }
}

// ── Accept pre-parsed Word tables encoded as JSON buffer ──
function handleParseWord({ id, sourceId, buffer, name }) {
  try {
    const text = new TextDecoder().decode(buffer);
    const { _wordTables } = JSON.parse(text);
    const tables = _wordTables.map(t => {
      const _index = t.rows.map(r => Object.values(r).join('\x00').toLowerCase());
      return { name: t.name, columns: t.columns, rows: t.rows, _index };
    });
    store.set(sourceId, tables);
    self.postMessage({
      op: 'parsed', id, sourceId, name, error: null,
      tables: tables.map(t => ({ name: t.name, columns: t.columns, totalRows: t.rows.length })),
    });
  } catch(err) {
    self.postMessage({ op: 'parsed', id, sourceId, name, error: err.message, tables: null });
  }
}

// ── Return a page of rows for virtual scroll / tab display ──
function handleGetRows({ id, sourceId, tableName, offset, limit }) {
  const tables = store.get(sourceId);
  const table  = tables && tables.find(t => t.name === tableName);
  const rows   = table ? table.rows.slice(offset, offset + limit) : [];
  self.postMessage({ op: 'rows', id, sourceId, tableName, offset, rows });
}

// ── Full-text search across requested sources ──
function handleSearch({ id, query, filterIds }) {
  const q = query.toLowerCase();
  const results = [];

  for (const [sourceId, tables] of store) {
    if (filterIds && filterIds.length && !filterIds.includes(sourceId)) continue;
    for (const t of tables) {
      const rows = [];
      for (let i = 0; i < t.rows.length; i++) {
        if (t._index[i].includes(q)) rows.push(t.rows[i]);
      }
      if (rows.length) results.push({ sourceId, tableName: t.name, columns: t.columns, rows });
    }
  }

  self.postMessage({ op: 'search_results', id, results });
}
