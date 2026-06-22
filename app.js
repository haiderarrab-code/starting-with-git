/* =========================================================
   Unified Database Search System — app.js
   Architecture: data-worker.js holds ALL row data.
   Renderer only holds metadata (name, columns, totalRows).
   ========================================================= */
'use strict';

// ── Data Worker (holds all rows + search index) ──────────────
let _dataWorker    = null;
let _workerReady   = false;
let _workerMsgId   = 0;
const _pending     = new Map(); // id -> { resolve, reject }

function initDataWorker() {
  _dataWorker = new Worker('data-worker.js');
  _dataWorker.onmessage = e => {
    const msg = e.data;
    if (msg.type === 'ready') { _workerReady = true; return; }
    if (msg.op === 'ready')   { _workerReady = true; return; }
    const p = _pending.get(msg.id);
    if (p) { _pending.delete(msg.id); p.resolve(msg); }
  };
  _dataWorker.onerror = err => {
    console.error('data-worker error:', err);
    for (const [id, p] of _pending) { _pending.delete(id); p.reject(err); }
  };
}

function workerSend(payload, transfer) {
  return new Promise((resolve, reject) => {
    const id = ++_workerMsgId;
    _pending.set(id, { resolve, reject });
    _dataWorker.postMessage({ ...payload, id }, transfer || []);
  });
}

// ── Disposable parse worker (recoverable if a file hangs XLSX.read) ──
let _parseWorker     = null;
let _parseMsgId      = 0;
const _parsePending  = new Map();   // id -> { resolve, reject }
const PARSE_TIMEOUT  = 30000;       // ms before we give up on a single file
let _currentParseAbort = null;      // abort fn for the in-flight parse

function initParseWorker() {
  _parseWorker = new Worker('parse-worker.js');
  _parseWorker.onmessage = e => {
    const msg = e.data;
    if (msg.type === 'ready') return;
    const p = _parsePending.get(msg.id);
    if (p) { _parsePending.delete(msg.id); p.resolve(msg); }
  };
  _parseWorker.onerror = () => {
    for (const [id, p] of _parsePending) { _parsePending.delete(id); p.reject(new Error('parse-worker error')); }
  };
}

// Kill the parse worker (aborts whatever it's stuck on) and start a fresh one.
// Data already stored in data-worker is NOT affected.
function recreateParseWorker() {
  try { if (_parseWorker) _parseWorker.terminate(); } catch (_) {}
  for (const [id, p] of _parsePending) { _parsePending.delete(id); p.reject(new Error('aborted')); }
  initParseWorker();
}

function parseWorkerSend(buffer, name, transfer) {
  return new Promise((resolve, reject) => {
    const id = ++_parseMsgId;
    _parsePending.set(id, { resolve, reject });
    _parseWorker.postMessage({ id, buffer, name }, transfer || []);
  });
}

// Parse one buffer with a hard timeout. On timeout / abort the parse worker
// is destroyed and recreated so the next file gets a clean worker.
function parseWithTimeout(buffer, name) {
  return new Promise((resolve, reject) => {
    let done = false;
    const finish = (fn, arg) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      _currentParseAbort = null;
      fn(arg);
    };
    const timer = setTimeout(() => {
      recreateParseWorker();
      finish(reject, new Error('TIMEOUT'));
    }, PARSE_TIMEOUT);

    // Allow skip/cancel buttons to abort the current parse immediately.
    _currentParseAbort = () => {
      recreateParseWorker();
      finish(reject, new Error('ABORTED'));
    };

    parseWorkerSend(buffer, name, [buffer])
      .then(msg => {
        if (msg.error) finish(reject, new Error(msg.error));
        else           finish(resolve, msg.tables || []);
      })
      .catch(err => finish(reject, err));
  });
}

// Parse a file: rows flow parse-worker → data-worker directly via MessageChannel.
// The renderer never holds or transfers any row data.
async function parseFile(file) {
  const isTxt = /\.txt$/i.test(file.name);
  let buffer;
  if (isTxt) {
    const text = await new Promise((res, rej) => {
      const r = new FileReader();
      r.onload  = e => res(e.target.result);
      r.onerror = rej;
      r.readAsText(file, 'windows-1256');
    });
    buffer = new TextEncoder().encode(text).buffer;
  } else {
    buffer = await file.arrayBuffer();
  }

  const sourceId = uid();

  // Create a direct pipe between the two workers.
  const { port1, port2 } = new MessageChannel();

  // Tell data-worker to listen on port1 and store whatever arrives.
  const storeId = ++_workerMsgId;
  const storePromise = new Promise((resolve, reject) => {
    _pending.set(storeId, { resolve, reject });
  });
  _dataWorker.postMessage({ op: 'listen_port', id: storeId, sourceId, name: file.name, port: port1 }, [port1]);

  // Tell parse-worker to parse the buffer and send results to data-worker via port2.
  // Kill it (and reject storePromise) if it doesn't finish within PARSE_TIMEOUT.
  return new Promise((resolve, reject) => {
    let done = false;
    const finish = (fn, arg) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      _currentParseAbort = null;
      fn(arg);
    };

    const timer = setTimeout(() => {
      recreateParseWorker();
      _pending.delete(storeId);
      finish(reject, new Error('TIMEOUT'));
    }, PARSE_TIMEOUT);

    _currentParseAbort = () => {
      recreateParseWorker();
      _pending.delete(storeId);
      finish(reject, new Error('ABORTED'));
    };

    // When data-worker finishes indexing, we're done.
    storePromise.then(msg => finish(resolve, msg)).catch(err => finish(reject, err));

    // Kick parse-worker (buffer + port2 transferred zero-copy).
    const parseId = ++_parseMsgId;
    _parsePending.set(parseId, {
      resolve: msg => { if (msg.error) { _pending.delete(storeId); finish(reject, new Error(msg.error)); } },
      reject:  err => { _pending.delete(storeId); finish(reject, err); },
    });
    _parseWorker.postMessage({ id: parseId, buffer, name: file.name, port: port2 }, [buffer, port2]);
  });
}

// ── State (metadata only — no rows) ──────────────────────────
const state = {
  sources:      [],  // { id, name, type, color, totalRows, tables:[{name,columns,totalRows}] }
  activeTab:    'all',
  searchQuery:  '',
  activeFilters: new Set(),
};

// ── Colors ───────────────────────────────────────────────────
const TYPE_COLOR = { excel: '#16a34a' };
const TYPE_LABEL = { excel: 'Excel'   };

// ── Unique ID ─────────────────────────────────────────────────
let _uid = Date.now();
function uid() { return 'src_' + (++_uid); }

// ═══════════════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  initDataWorker();
  initParseWorker();
  initTheme();
  loadFromStorage();
  renderAll();

  let _searchTimer = null;
  const si = document.getElementById('searchInput');
  si.addEventListener('input', () => {
    const val = si.value.trim();
    document.getElementById('searchClear').style.display = val ? '' : 'none';
    clearTimeout(_searchTimer);
    _searchTimer = setTimeout(() => {
      state.searchQuery = val;
      renderTabContent();
    }, 250);
  });

  document.getElementById('folderFileInput').addEventListener('change', e => handleFolderImport(e));
  document.getElementById('excelFileInput').addEventListener('change',  e => handleExcelFiles(e));
  document.getElementById('exportBtn').addEventListener('click', exportUnified);
  document.getElementById('clearAllBtn').addEventListener('click', clearAll);
  document.getElementById('themeToggle').addEventListener('click', toggleTheme);
});

// ═══════════════════════════════════════════════════════════
//  THEME
// ═══════════════════════════════════════════════════════════
function initTheme() {
  applyTheme(localStorage.getItem('uds_theme') || 'dark');
}
function toggleTheme() {
  applyTheme(document.body.classList.contains('light') ? 'dark' : 'light');
}
function applyTheme(t) {
  document.body.classList.toggle('light', t === 'light');
  document.getElementById('themeToggle').querySelector('.theme-icon').textContent = t === 'light' ? '🌙' : '☀️';
  localStorage.setItem('uds_theme', t);
}

// ═══════════════════════════════════════════════════════════
//  FILE HANDLERS
// ═══════════════════════════════════════════════════════════
async function handleExcelFiles(e) {
  const files = Array.from(e.target.files);
  if (!files.length) return;
  e.target.value = '';

  if (files.length === 1) {
    if (/\.docx$/i.test(files[0].name)) {
      await parseWordFile(files[0]);
      renderAll();
      scheduleSave();
      return;
    }
    await importSingleFile(files[0]);
    return;
  }
  await importFileList(files);
}

async function handleFolderImport(e) {
  const all = Array.from(e.target.files);
  e.target.value = '';
  const files = all.filter(f => /\.(xlsx|xls|csv|txt|docx)$/i.test(f.name));
  if (!files.length) {
    notify('لا توجد ملفات مدعومة في المجلد', 'warning', 4000);
    return;
  }
  await importFileList(files);
}

// ── Single file ───────────────────────────────────────────────
async function importSingleFile(file) {
  _importCancelled = false;
  _importSkip      = false;
  showProgress(`جاري قراءة "${file.name}"...`, 30);
  try {
    const msg = await parseFile(file);
    if (_importCancelled || _importSkip) { hideProgress(); return; }
    if (msg.error) {
      hideProgress();
      notify(`خطأ في "${file.name}": ${msg.error}`, 'error');
    } else {
      addSource({ id: msg.sourceId, name: msg.name, type: 'excel',
        tables: msg.tables, totalRows: msg.tables.reduce((s, t) => s + t.totalRows, 0) });
      renderAll();
      scheduleSave();
      showProgress('تم التحميل ✓', 100);
      await sleep(500);
      hideProgress();
    }
  } catch (err) {
    hideProgress();
    notify(`خطأ في "${file.name}": ${err.message}`, 'error');
  }
}

// ── Multiple files (sequential, with cancel/skip) ─────────────
async function importFileList(files) {
  let done = 0, skipped = 0, errors = 0;
  const failedFiles = [];
  const total = files.length;
  _importCancelled = false;
  _importSkip      = false;
  showProgress(`جاري استيراد ${total} ملف...`, 0);

  const MAX_ROWS = 500_000;

  for (const file of files) {
    if (_importCancelled) break;

    const currentTotal = state.sources.reduce((s, src) => s + src.totalRows, 0);
    if (currentTotal >= MAX_ROWS) {
      notify(`وصلت للحد الأقصى (${MAX_ROWS.toLocaleString('ar')} سجل).`, 'warning', 6000);
      break;
    }

    _importSkip = false;
    showProgress(`(${done + 1}/${total}) ${file.name}`, (done / total) * 100);
    await sleep(10); // minimal yield for skip/cancel buttons
    if (_importSkip) { skipped++; done++; continue; }
    if (_importCancelled) break;

    try {
      if (/\.docx$/i.test(file.name)) {
        await parseWordFile(file);
      } else {
        const msg = await parseFile(file);
        if (_importSkip) { skipped++; done++; continue; }
        if (msg.error) {
          errors++;
          failedFiles.push(file.name);
        } else if (msg.tables && msg.tables.length) {
          addSource({ id: msg.sourceId, name: msg.name, type: 'excel',
            tables: msg.tables, totalRows: msg.tables.reduce((s, t) => s + t.totalRows, 0) });
          renderStatsBar();   // live counter feedback (cheap text update)
          renderBadges();
        }
      }
    } catch (err) {
      if (_importSkip || err.message === 'ABORTED') {
        skipped++;
      } else {
        errors++;
        failedFiles.push(file.name + (err.message === 'TIMEOUT' ? ' (تجاوز الوقت المسموح)' : ''));
      }
    }
    done++;
  }

  renderAll();
  scheduleSave();
  if (!_importCancelled) hideProgress();

  const ok = done - errors - skipped;
  const parts = [`تم استيراد ${ok} ملف بنجاح`];
  if (skipped)   parts.push(`${skipped} مُتخطى`);
  if (errors)    parts.push(`${errors} فشل`);
  notify(parts.join(' — ') + ' ✓', errors ? 'warning' : 'success', 6000);
  if (failedFiles.length) {
    setTimeout(() => notify(`ملفات فاشلة: ${failedFiles.slice(0,5).join(', ')}${failedFiles.length > 5 ? '...' : ''}`, 'error', 8000), 800);
  }
}

// ── Word (.docx) ──────────────────────────────────────────────
async function parseWordFile(file) {
  if (typeof JSZip === 'undefined') { notify('JSZip غير متاحة', 'error'); return; }
  if (!_importCancelled) showProgress(`جاري قراءة "${file.name}"...`, 20);
  try {
    const buf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);
    const xmlFile = zip.file('word/document.xml');
    if (!xmlFile) throw new Error('ملف Word غير صالح');

    const xmlText = await xmlFile.async('string');
    const doc = new DOMParser().parseFromString(xmlText, 'application/xml');
    const NS  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const tbls = doc.getElementsByTagNameNS(NS, 'tbl');
    if (!tbls.length) { notify(`${file.name}: لا توجد جداول.`, 'warning'); return; }

    const getCellText = cell => Array.from(cell.getElementsByTagNameNS(NS, 't')).map(t => t.textContent).join('').trim();

    const tables = [];
    for (let ti = 0; ti < tbls.length; ti++) {
      const rows = tbls[ti].getElementsByTagNameNS(NS, 'tr');
      if (rows.length < 2) continue;
      const columns  = Array.from(rows[0].getElementsByTagNameNS(NS, 'tc')).map((c, i) => getCellText(c) || `عمود ${i + 1}`);
      const dataRows = [];
      for (let ri = 1; ri < rows.length; ri++) {
        const cells = Array.from(rows[ri].getElementsByTagNameNS(NS, 'tc'));
        const obj   = {};
        columns.forEach((col, ci) => { obj[col] = cells[ci] ? getCellText(cells[ci]) : ''; });
        dataRows.push(obj);
      }
      if (dataRows.length) tables.push({ name: `جدول ${ti + 1}`, columns, rows: dataRows });
    }

    if (!tables.length) { notify(`${file.name}: الجداول فارغة.`, 'warning'); return; }

    // Send Word tables to data-worker too (as pre-parsed data)
    const sourceId = uid();
    // Word files are small enough to parse inline; store via a special message
    // We'll encode them as a fake "already-parsed" source for consistency
    const tableMeta = tables.map(t => ({ name: t.name, columns: t.columns, totalRows: t.rows.length }));

    // Store inline in data-worker by sending buffer trick: send JSON-encoded rows
    // as a UTF-8 buffer that data-worker reads back
    const encoded = new TextEncoder().encode(JSON.stringify({ _wordTables: tables }));
    const msg = await workerSend({ op: 'parse_word', sourceId, buffer: encoded.buffer, name: file.name }, [encoded.buffer]);

    addSource({ id: sourceId, name: file.name, type: 'excel', tables: tableMeta,
      totalRows: tableMeta.reduce((s, t) => s + t.totalRows, 0) });
  } catch (err) {
    notify(`خطأ في "${file.name}": ${err.message}`, 'error');
  }
}

// ═══════════════════════════════════════════════════════════
//  SOURCE MANAGEMENT  (metadata only — no rows in renderer)
// ═══════════════════════════════════════════════════════════
function addSource(src) {
  const existing = state.sources.findIndex(s => s.name === src.name);
  if (existing !== -1) {
    workerSend({ op: 'remove', sourceId: state.sources[existing].id });
    state.sources.splice(existing, 1);
  }
  state.sources.push({ color: TYPE_COLOR[src.type], ...src });
  state.activeTab = 'all';
}

function deleteSource(id) {
  workerSend({ op: 'remove', sourceId: id });
  state.sources = state.sources.filter(s => s.id !== id);
  state.activeFilters.delete(id);
  if (state.activeTab === id) state.activeTab = 'all';
  scheduleSave();
  renderAll();
}

function clearAll() {
  if (!state.sources.length) return;
  if (!confirm('هل تريد مسح جميع البيانات المستوردة؟')) return;
  workerSend({ op: 'clear' });
  state.sources = [];
  state.activeTab = 'all';
  state.activeFilters.clear();
  state.searchQuery = '';
  document.getElementById('searchInput').value = '';
  document.getElementById('searchClear').style.display = 'none';
  saveToStorage();
  renderAll();
  notify('تم مسح جميع البيانات', 'info');
}

// ═══════════════════════════════════════════════════════════
//  STORAGE  (metadata + first 200 rows per table for preview)
// ═══════════════════════════════════════════════════════════
let _saveTimer = null;
function scheduleSave() {
  clearTimeout(_saveTimer);
  _saveTimer = setTimeout(saveToStorage, 2000);
}

function saveToStorage() {
  try {
    const slim = state.sources.map(s => ({
      id: s.id, name: s.name, type: s.type, color: s.color, totalRows: s.totalRows,
      tables: s.tables.map(t => ({ name: t.name, columns: t.columns, totalRows: t.totalRows })),
    }));
    localStorage.setItem('uds_sources_meta', JSON.stringify(slim));
  } catch(e) {
    try { localStorage.removeItem('uds_sources_meta'); } catch(_) {}
  }
}

function loadFromStorage() {
  try {
    const raw = localStorage.getItem('uds_sources_meta');
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      state.sources = parsed.map(s => ({ color: TYPE_COLOR[s.type] || '#16a34a', ...s }));
      // Note: rows are gone after restart — data-worker is empty.
      // Sources show as "cached" with 0 rows available until reimported.
      // We mark them so user knows.
      for (const s of state.sources) s._cached = true;
    }
  } catch(e) {
    try { localStorage.removeItem('uds_sources_meta'); } catch(_) {}
  }
}

// ═══════════════════════════════════════════════════════════
//  RENDER ALL
// ═══════════════════════════════════════════════════════════
function renderAll() {
  renderChips();
  renderBadges();
  renderStatsBar();
  renderSearchSection();
  renderTabs();
  renderTabContent();
}

// ── Source chips ─────────────────────────────────────────────
function renderChips() {
  const container = document.getElementById('excelChips');
  const toggle    = document.getElementById('chipsToggle');
  const srcs = state.sources.filter(s => s.type === 'excel');

  container.innerHTML = srcs.map(s => `
    <div class="source-chip${s._cached ? ' chip-cached' : ''}" onclick="switchTab('${s.id}')">
      <span class="chip-dot" style="background:${s.color}"></span>
      <span class="chip-name" title="${esc(s.name)}">${esc(s.name)}</span>
      <span class="chip-count" style="background:${s.color}">${s.totalRows.toLocaleString('ar')}</span>
      <button class="chip-delete" title="حذف" onclick="event.stopPropagation();deleteSource('${s.id}')">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" width="12" height="12">
          <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
        </svg>
      </button>
    </div>
  `).join('');

  if (srcs.length) {
    toggle.classList.add('visible');
    document.getElementById('chipsToggleLabel').textContent = `${srcs.length} ملف — عرض / إخفاء`;
  } else {
    toggle.classList.remove('visible');
    container.classList.remove('open');
    toggle.classList.remove('open');
  }
}

function toggleChips() {
  const container = document.getElementById('excelChips');
  const toggle    = document.getElementById('chipsToggle');
  const isOpen = container.classList.contains('open');
  container.classList.toggle('open', !isOpen);
  toggle.classList.toggle('open', !isOpen);
}

// ── Badges ───────────────────────────────────────────────────
function renderBadges() {
  const badge = document.getElementById('excelBadge');
  const srcs  = state.sources.filter(s => s.type === 'excel');
  if (srcs.length) {
    badge.textContent = srcs.reduce((s, x) => s + x.totalRows, 0).toLocaleString('ar');
    badge.style.display = '';
  } else {
    badge.style.display = 'none';
  }
}

// ── Stats bar ─────────────────────────────────────────────────
function renderStatsBar() {
  const bar = document.getElementById('statsBar');
  const exportBtn = document.getElementById('exportBtn');
  if (!state.sources.length) { bar.style.display = 'none'; exportBtn.style.display = 'none'; return; }
  bar.style.display = '';
  exportBtn.style.display = '';
  document.getElementById('totalSources').textContent = state.sources.length;
  document.getElementById('totalTables').textContent  = state.sources.reduce((s, src) => s + src.tables.length, 0);
  document.getElementById('totalRecords').textContent = state.sources.reduce((s, src) => s + src.totalRows, 0).toLocaleString('ar');
}

// ── Search section ────────────────────────────────────────────
function renderSearchSection() {
  const sec = document.getElementById('searchSection');
  if (!state.sources.length) { sec.style.display = 'none'; return; }
  sec.style.display = '';
  renderFilterPills();
}

function renderFilterPills() {
  const container = document.getElementById('searchFilters');
  const toggle    = document.getElementById('filtersToggle');
  const pills = [
    { id: 'all_pill', label: 'الكل', color: '#6366f1' },
    ...state.sources.map(s => ({ id: s.id, label: s.name, color: s.color })),
  ];
  container.innerHTML = pills.map(p => {
    const isAll    = p.id === 'all_pill';
    const isActive = isAll ? state.activeFilters.size === 0 : state.activeFilters.has(p.id);
    return `<button class="filter-pill ${isActive ? 'active' : ''}"
      style="${isActive ? `background:${p.color};border-color:${p.color}` : `border-color:${p.color};color:${p.color}`}"
      onclick="toggleFilter('${p.id}')">
      <span class="pill-dot" style="background:${isActive ? 'rgba(255,255,255,0.7)' : p.color}"></span>
      ${esc(p.label)}
    </button>`;
  }).join('');
  if (state.sources.length) {
    toggle.style.display = '';
    document.getElementById('filtersToggleLabel').textContent = `فلترة المصادر (${state.sources.length})`;
  } else { toggle.style.display = 'none'; }
}

function toggleFilters() {
  const c = document.getElementById('searchFilters');
  const t = document.getElementById('filtersToggle');
  const open = c.classList.contains('open');
  c.classList.toggle('open', !open);
  t.classList.toggle('open', !open);
}

function toggleFilter(id) {
  if (id === 'all_pill') state.activeFilters.clear();
  else {
    if (state.activeFilters.has(id)) state.activeFilters.delete(id);
    else state.activeFilters.add(id);
  }
  renderFilterPills();
  renderTabContent();
}

// ── Tabs ──────────────────────────────────────────────────────
function renderTabs() {
  const tabsBar   = document.getElementById('tabsBar');
  const dataArea  = document.getElementById('dataArea');
  const emptyState = document.getElementById('emptyState');
  if (!state.sources.length) {
    emptyState.style.display = '';
    dataArea.style.display   = 'none';
    tabsBar.innerHTML = '';
    return;
  }
  emptyState.style.display = 'none';
  dataArea.style.display   = '';
  const tabs = [
    { id: 'all', label: 'جميع البيانات' },
    ...state.sources.map(s => ({ id: s.id, label: s.name, color: s.color, type: s.type })),
  ];
  tabsBar.innerHTML = tabs.map(t => `
    <button class="tab-btn ${state.activeTab === t.id ? 'active' : ''}" onclick="switchTab('${t.id}')">
      ${t.color ? `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${t.color};margin-left:6px;vertical-align:middle"></span>` : ''}
      ${esc(t.label)}
    </button>`).join('');
}

function switchTab(id) {
  state.activeTab = id;
  renderTabs();
  renderTabContent();
}

// ═══════════════════════════════════════════════════════════
//  TAB CONTENT  — async virtual scroll
//  Rows are fetched from data-worker on demand.
// ═══════════════════════════════════════════════════════════
const VT_ROW_H   = 34;
const VT_VIEW_H  = 440;
const VT_BUFFER  = 5;
const VT_PAGE    = 300;   // rows fetched per request
let   _vtId      = 0;
// vtId -> { sourceId, tableName, columns, totalRows, query, cache: Map<pageKey,rows[]>, loading:Set }
const _vtState   = new Map();

function renderTabContent() {
  const container = document.getElementById('tabContent');
  _vtState.clear();
  if (!state.sources.length) { container.innerHTML = ''; return; }

  const q = state.searchQuery.toLowerCase();
  const activeSrcs = state.activeFilters.size > 0
    ? state.sources.filter(s => state.activeFilters.has(s.id))
    : state.sources;

  if (q) {
    // Search: async — send to worker
    renderSearchResults(activeSrcs, q, container);
    return;
  }

  if (state.activeTab === 'all') {
    let html = '';
    for (const src of activeSrcs) {
      for (const t of src.tables) html += tableShellHTML(src, t, t.totalRows, '');
    }
    container.innerHTML = html || noResultsHTML('لا توجد مصادر');
  } else {
    const src = state.sources.find(s => s.id === state.activeTab);
    container.innerHTML = src ? src.tables.map(t => tableShellHTML(src, t, t.totalRows, '')).join('') : '';
  }
  mountVirtualTables();
}

// ── Async search via data-worker ──────────────────────────────
function renderSearchResults(activeSrcs, q, container) {
  container.innerHTML = `<div class="loading"><div class="spinner"></div><span>جاري البحث...</span></div>`;
  const filterIds = state.activeFilters.size ? [...state.activeFilters] : [];
  workerSend({ op: 'search', query: q, filterIds }).then(msg => {
    if (state.searchQuery.toLowerCase() !== q) return; // stale
    const { results } = msg;
    if (!results || !results.length) {
      container.innerHTML = noResultsHTML(`لا توجد نتائج لـ "${esc(q)}"`);
      return;
    }
    const totalMatches = results.reduce((s, g) => s + g.rows.length, 0);
    let html = `<div class="results-header"><div class="results-count">
      نتائج البحث: <strong>${totalMatches.toLocaleString('ar')}</strong> سجل
      في <strong>${results.length}</strong> جدول لـ "<strong>${esc(q)}</strong>"
    </div></div>`;

    for (const g of results) {
      const src = state.sources.find(s => s.id === g.sourceId);
      if (!src) continue;
      const tableMeta = { name: g.tableName, columns: g.columns, totalRows: g.rows.length };
      const vtId = 'vt_' + (++_vtId);
      _vtState.set(vtId, {
        sourceId: g.sourceId, tableName: g.tableName, columns: g.columns,
        totalRows: g.rows.length, query: q,
        cache: new Map([[0, g.rows]]), // search results arrive all at once
        loading: new Set(),
      });
      html += tableShellHTMLById(vtId, src, tableMeta, g.rows.length, q);
    }
    container.innerHTML = html;
    mountVirtualTables();
  }).catch(() => {
    container.innerHTML = noResultsHTML('خطأ في البحث');
  });
}

// ── Build table shell HTML ────────────────────────────────────
function tableShellHTML(src, tableMeta, totalRows, query) {
  if (!tableMeta.columns || !tableMeta.columns.length) return '';
  const vtId = 'vt_' + (++_vtId);
  _vtState.set(vtId, {
    sourceId: src.id, tableName: tableMeta.name, columns: tableMeta.columns,
    totalRows, query, cache: new Map(), loading: new Set(),
  });
  return tableShellHTMLById(vtId, src, tableMeta, totalRows, query);
}

function tableShellHTMLById(vtId, src, tableMeta, totalRows, query) {
  const gridCols = tableMeta.columns.map(() => 'minmax(140px, 1fr)').join(' ');
  const headCells = tableMeta.columns.map(c => `<div class="vt-cell vt-th">${esc(String(c))}</div>`).join('');
  const scrollH  = Math.min(totalRows * VT_ROW_H, VT_VIEW_H);
  return `<div class="table-wrapper">
    <div class="table-header"><div class="table-title">
      <span class="table-name">${esc(tableMeta.name)}</span>
      <span class="table-count" style="background:${src.color}">${totalRows.toLocaleString('ar')}</span>
      <span class="result-source-badge badge-${src.type}">${TYPE_LABEL[src.type]}</span>
      <span style="font-size:0.78rem;color:var(--color-text-dim)">${esc(src.name)}</span>
    </div></div>
    <div class="vt-outer">
      <div class="vt-track" style="--vt-cols:${gridCols}">
        <div class="vt-head">${headCells}</div>
        <div class="vt-scroll" id="${vtId}" style="height:${scrollH}px">
          <div class="vt-canvas" style="height:${totalRows * VT_ROW_H}px">
            <div class="vt-win"></div>
          </div>
        </div>
      </div>
    </div>
  </div>`;
}

// ── Virtual table mount — requests rows from data-worker ──────
function mountVirtualTables() {
  for (const [vtId, st] of _vtState) {
    const scrollEl = document.getElementById(vtId);
    if (!scrollEl) continue;
    const winEl = scrollEl.querySelector('.vt-win');

    const render = () => {
      const scrollTop = scrollEl.scrollTop;
      const startRow  = Math.max(0, Math.floor(scrollTop / VT_ROW_H) - VT_BUFFER);
      const endRow    = Math.min(st.totalRows, startRow + Math.ceil(VT_VIEW_H / VT_ROW_H) + VT_BUFFER * 2);

      // Which page covers startRow?
      const pageKey = Math.floor(startRow / VT_PAGE) * VT_PAGE;
      const rows    = st.cache.get(pageKey);

      if (!rows) {
        // Fetch this page from worker
        if (!st.loading.has(pageKey)) {
          st.loading.add(pageKey);
          workerSend({ op: 'rows', sourceId: st.sourceId, tableName: st.tableName,
            offset: pageKey, limit: VT_PAGE * 2 })
            .then(msg => {
              st.loading.delete(pageKey);
              st.cache.set(pageKey, msg.rows || []);
              render(); // re-render with fresh data
            }).catch(() => { st.loading.delete(pageKey); });
        }
        // Show placeholders while loading
        let ph = '';
        for (let i = startRow; i < endRow; i++) {
          ph += `<div class="vt-row">${st.columns.map(() => '<div class="vt-cell" style="opacity:0.3">...</div>').join('')}</div>`;
        }
        winEl.style.transform = `translateY(${startRow * VT_ROW_H}px)`;
        winEl.innerHTML = ph;
        return;
      }

      // Rows are cached — render them
      let html = '';
      for (let i = startRow; i < endRow; i++) {
        const localIdx = i - pageKey;
        const row = rows[localIdx];
        if (!row) continue;
        let cells = '';
        for (const col of st.columns) {
          const val = row[col] != null ? String(row[col]) : '';
          cells += `<div class="vt-cell" title="${esc(val)}">${st.query ? highlightText(val, st.query) : esc(val)}</div>`;
        }
        html += `<div class="vt-row">${cells}</div>`;
      }
      winEl.style.transform = `translateY(${startRow * VT_ROW_H}px)`;
      winEl.innerHTML = html;
    };

    let _raf = null;
    scrollEl.addEventListener('scroll', () => {
      if (_raf) return;
      _raf = requestAnimationFrame(() => { _raf = null; render(); });
    });
    render(); // initial render — triggers first page fetch
  }
}

// ═══════════════════════════════════════════════════════════
//  EXPORT
// ═══════════════════════════════════════════════════════════
async function exportUnified() {
  if (!state.sources.length) return;
  const btn = document.getElementById('exportBtn');
  btn.disabled = true;
  const totalTables = state.sources.reduce((s, src) => s + src.tables.length, 0);
  const resetBtn = () => {
    btn.disabled = false;
    btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> تصدير موحّد`;
  };
  try {
    const wb = XLSX.utils.book_new();
    const usedNames = new Set();
    let done = 0;

    for (const src of state.sources) {
      for (const table of src.tables) {
        done++;
        btn.textContent = `جاري التصدير... ${done}/${totalTables}`;
        await sleep(0);

        // Fetch ALL rows for this table from worker
        const msg = await workerSend({ op: 'rows', sourceId: src.id, tableName: table.name,
          offset: 0, limit: 150_000 });

        let sheetName = `${src.name.replace(/\.[^.]+$/, '')}_${table.name}`.slice(0, 31);
        let base = sheetName, n = 2;
        while (usedNames.has(sheetName)) sheetName = `${base.slice(0, 28)}_${n++}`;
        usedNames.add(sheetName);

        const ws = XLSX.utils.json_to_sheet(msg.rows || []);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      }
    }
    btn.textContent = 'جاري حفظ الملف...';
    await sleep(0);
    XLSX.writeFile(wb, `بيانات_موحدة_${new Date().toISOString().slice(0,10)}.xlsx`);
    notify('تم تصدير الملف بنجاح ✓', 'success');
  } catch (err) {
    notify(`خطأ في التصدير: ${err.message}`, 'error');
  } finally { resetBtn(); }
}

// ═══════════════════════════════════════════════════════════
//  IMPORT PROGRESS + CANCEL/SKIP
// ═══════════════════════════════════════════════════════════
let _importCancelled = false;
let _importSkip      = false;

function cancelImport() {
  _importCancelled = true;
  if (_currentParseAbort) _currentParseAbort();   // abort the file being parsed right now
  hideProgress();
  notify('تم إلغاء الاستيراد', 'info', 3000);
}
function skipImport() {
  _importSkip = true;
  if (_currentParseAbort) _currentParseAbort();    // abort the stuck file immediately
}

function showProgress(title, pct) {
  if (_importCancelled) return;
  document.getElementById('progressTitle').textContent = title;
  document.getElementById('progressText').textContent  = Math.round(pct) + '%';
  document.getElementById('progressFill').style.width  = pct + '%';
  document.getElementById('importProgress').classList.add('visible');
}
function hideProgress() {
  _importCancelled = false;
  _importSkip      = false;
  document.getElementById('importProgress').classList.remove('visible');
  document.getElementById('progressFill').style.width = '0%';
}

// ═══════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════
function highlightText(text, query) {
  if (!query || !text) return esc(text);
  return esc(text).replace(new RegExp(escapeRegex(query), 'gi'), m => `<mark>${m}</mark>`);
}
function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function esc(str) {
  if (str == null) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function noResultsHTML(title, desc = '') {
  return `<div class="no-results">
    <svg viewBox="0 0 48 48" fill="none" width="64" height="64">
      <circle cx="24" cy="24" r="22" stroke="currentColor" stroke-width="1.5" stroke-dasharray="5 3"/>
      <path d="M16 24h16M24 16v16" stroke="currentColor" stroke-width="1.5" opacity="0.5" transform="rotate(45 24 24)"/>
    </svg>
    <h3>${title}</h3>${desc ? `<p>${desc}</p>` : ''}
  </div>`;
}
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ── Notifications ─────────────────────────────────────────────
let _notifTimer = null;
function notify(msg, type = 'info', duration = 3500) {
  const el = document.getElementById('notification');
  el.className = `notification ${type}`;
  el.textContent = msg;
  el.style.display = '';
  clearTimeout(_notifTimer);
  _notifTimer = setTimeout(() => { el.style.display = 'none'; }, duration);
}

function clearSearch() {
  state.searchQuery = '';
  document.getElementById('searchInput').value = '';
  document.getElementById('searchClear').style.display = 'none';
  renderTabContent();
}

// ── Expose globals ────────────────────────────────────────────
window.clearSearch    = clearSearch;
window.switchTab      = switchTab;
window.deleteSource   = deleteSource;
window.toggleFilter   = toggleFilter;
window.toggleChips    = toggleChips;
window.toggleFilters  = toggleFilters;
window.cancelImport   = cancelImport;
window.skipImport     = skipImport;
