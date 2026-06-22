/* data-worker.js — persistent storage layer (IndexedDB).
   Rows live on disk, NOT in RAM, so file size / total volume is limited by
   disk, not memory. Data also survives app restarts.

   Schema (DB "uds_db"):
     store "rows": keyPath "pk"
        pk = `${sourceId}|${tableIndex(2)}|${rowIndex(10)}`  (lexicographic = row order)
        record = { pk, sourceId, t: tableIndex, s: searchText, r: rowObject }
     store "meta": keyPath "sourceId"
        record = { sourceId, name, type, color, tables:[{name,columns,totalRows}] }
*/
'use strict';

importScripts('vendor/xlsx.full.min.js');

const DB_NAME = 'uds_db';
const DB_VER  = 1;
const CHUNK   = 4000;     // rows written per IndexedDB transaction
let   _db     = null;
const metaCache = new Map(); // sourceId -> meta (kept in RAM, tiny)

// ── open / init DB ────────────────────────────────────────────
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VER);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains('rows')) db.createObjectStore('rows', { keyPath: 'pk' });
      if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta', { keyPath: 'sourceId' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error);
  });
}

let _dbReady = false;
const _msgQueue = [];

async function init() {
  _db = await openDB();
  const metas = await idbGetAll('meta');
  for (const m of metas) metaCache.set(m.sourceId, m);
  _dbReady = true;
  self.postMessage({ type: 'ready', sources: metas });
  for (const m of _msgQueue.splice(0)) dispatch(m);   // drain queued work
}

// ── tiny IDB helpers ──────────────────────────────────────────
function idbGetAll(store, range, count) {
  return new Promise((resolve, reject) => {
    const tx = _db.transaction(store, 'readonly');
    const req = tx.objectStore(store).getAll(range, count);
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error);
  });
}

function pad(n, w) { return String(n).padStart(w, '0'); }
function rangeFor(sourceId, ti, fromIdx) {
  const base = `${sourceId}|${pad(ti, 2)}|`;
  const lo = base + pad(fromIdx || 0, 10);
  const hi = base + '9999999999';
  return IDBKeyRange.bound(lo, hi);
}
function sourceRange(sourceId) {
  return IDBKeyRange.bound(`${sourceId}|`, `${sourceId}|￿`);
}

self.onmessage = function (e) {
  if (!_dbReady) { _msgQueue.push(e.data); return; }
  dispatch(e.data);
};

function dispatch(msg) {
  switch (msg.op) {
    case 'listen_port': return handleListenPort(msg);
    case 'store':       return handleStore(msg);
    case 'parse_word':  return handleParseWord(msg);
    case 'rows':        return handleGetRows(msg);
    case 'search':      return handleSearch(msg);
    case 'remove':      return handleRemove(msg);
    case 'clear':       return handleClear(msg);
  }
};

// ── Receive parsed tables from parse-worker over a MessageChannel port,
//    then write them to IndexedDB. Rows never touch the renderer. ──
function handleListenPort({ id, sourceId, name, type, color, port }) {
  port.onmessage = e => {
    const { tables, error } = e.data;
    if (error || !tables || !tables.length) {
      self.postMessage({ op: 'stored', id, sourceId, name, error: error || 'empty', tables: null });
      return;
    }
    handleStore({ id, sourceId, name, type, color, tables });
  };
}

// ── Write pre-parsed tables to IndexedDB in chunks (never blocks long) ──
async function handleStore({ id, sourceId, name, type, color, tables }) {
  try {
    const meta = { sourceId, name, type: type || 'excel', color, tables: [] };
    const totalRows = (tables || []).reduce((s, t) => s + t.rows.length, 0);
    let written = 0;

    for (let ti = 0; ti < tables.length; ti++) {
      const t = tables[ti];
      meta.tables.push({ name: t.name, columns: t.columns, totalRows: t.rows.length });

      for (let off = 0; off < t.rows.length; off += CHUNK) {
        const slice = t.rows.slice(off, off + CHUNK);
        await new Promise((resolve, reject) => {
          const tx = _db.transaction('rows', 'readwrite');
          const os = tx.objectStore('rows');
          for (let i = 0; i < slice.length; i++) {
            const row = slice[i];
            const idx = off + i;
            os.put({
              pk: `${sourceId}|${pad(ti, 2)}|${pad(idx, 10)}`,
              sourceId, t: ti,
              s: Object.values(row).join('\x00').toLowerCase(),
              r: row,
            });
          }
          tx.oncomplete = resolve;
          tx.onerror    = () => reject(tx.error);
        });
        written += slice.length;
        self.postMessage({ op: 'store_progress', id, sourceId, name, written, total: totalRows });
      }
    }

    // Persist metadata
    await new Promise((resolve, reject) => {
      const tx = _db.transaction('meta', 'readwrite');
      tx.objectStore('meta').put(meta);
      tx.oncomplete = resolve;
      tx.onerror    = () => reject(tx.error);
    });
    metaCache.set(sourceId, meta);

    self.postMessage({
      op: 'stored', id, sourceId, name, error: null,
      tables: meta.tables.map(t => ({ name: t.name, columns: t.columns, totalRows: t.totalRows })),
    });
  } catch (err) {
    self.postMessage({ op: 'stored', id, sourceId, name, error: err.message, tables: null });
  }
}

// ── Word (.docx) tables arrive pre-parsed as JSON buffer ──
async function handleParseWord({ id, sourceId, buffer, name, type, color }) {
  try {
    const { _wordTables } = JSON.parse(new TextDecoder().decode(buffer));
    return handleStore({ id, sourceId, name, type, color, tables: _wordTables });
  } catch (err) {
    self.postMessage({ op: 'stored', id, sourceId, name, error: err.message, tables: null });
  }
}

// ── Return a page of rows for virtual scroll ──
async function handleGetRows({ id, sourceId, tableName, offset, limit }) {
  try {
    const meta = metaCache.get(sourceId);
    const ti = meta ? meta.tables.findIndex(t => t.name === tableName) : -1;
    if (ti < 0) { self.postMessage({ op: 'rows', id, sourceId, tableName, offset, rows: [] }); return; }
    const recs = await idbGetAll('rows', rangeFor(sourceId, ti, offset), limit);
    self.postMessage({ op: 'rows', id, sourceId, tableName, offset, rows: recs.map(x => x.r) });
  } catch (err) {
    self.postMessage({ op: 'rows', id, sourceId, tableName, offset, rows: [] });
  }
}

// ── Full-text search via cursor (streams from disk, never loads everything) ──
function handleSearch({ id, query, filterIds }) {
  const q = query.toLowerCase();
  const MAX_RESULTS = 5000;
  const groups = new Map(); // key sourceId|ti -> { sourceId, tableName, columns, rows }
  let total = 0;

  const sources = [...metaCache.keys()].filter(sid =>
    !filterIds || !filterIds.length || filterIds.includes(sid));

  let si = 0;
  const nextSource = () => {
    if (si >= sources.length || total >= MAX_RESULTS) {
      self.postMessage({ op: 'search_results', id, results: [...groups.values()], capped: total >= MAX_RESULTS });
      return;
    }
    const sid = sources[si++];
    const meta = metaCache.get(sid);
    const tx = _db.transaction('rows', 'readonly');
    const cursorReq = tx.objectStore('rows').openCursor(sourceRange(sid));
    cursorReq.onsuccess = ev => {
      const cur = ev.target.result;
      if (!cur || total >= MAX_RESULTS) { nextSource(); return; }
      const rec = cur.value;
      if (rec.s.includes(q)) {
        const key = sid + '|' + rec.t;
        let g = groups.get(key);
        if (!g) {
          const tm = meta.tables[rec.t];
          g = { sourceId: sid, tableName: tm.name, columns: tm.columns, rows: [] };
          groups.set(key, g);
        }
        g.rows.push(rec.r);
        total++;
      }
      cur.continue();
    };
    cursorReq.onerror = () => nextSource();
  };
  nextSource();
}

// ── Delete a source ──
async function handleRemove({ id, sourceId }) {
  try {
    await new Promise((resolve, reject) => {
      const tx = _db.transaction('rows', 'readwrite');
      tx.objectStore('rows').delete(sourceRange(sourceId));
      tx.oncomplete = resolve; tx.onerror = () => reject(tx.error);
    });
    await new Promise((resolve, reject) => {
      const tx = _db.transaction('meta', 'readwrite');
      tx.objectStore('meta').delete(sourceId);
      tx.oncomplete = resolve; tx.onerror = () => reject(tx.error);
    });
    metaCache.delete(sourceId);
    self.postMessage({ op: 'removed', id, sourceId });
  } catch (err) {
    self.postMessage({ op: 'removed', id, sourceId, error: err.message });
  }
}

// ── Clear everything ──
async function handleClear({ id }) {
  try {
    await new Promise((resolve, reject) => {
      const tx = _db.transaction(['rows', 'meta'], 'readwrite');
      tx.objectStore('rows').clear();
      tx.objectStore('meta').clear();
      tx.oncomplete = resolve; tx.onerror = () => reject(tx.error);
    });
    metaCache.clear();
    self.postMessage({ op: 'cleared', id });
  } catch (err) {
    self.postMessage({ op: 'cleared', id, error: err.message });
  }
}

init();
