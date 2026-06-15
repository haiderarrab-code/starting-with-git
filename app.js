/* =========================================================
   Unified Database Search System — app.js
   ========================================================= */

'use strict';

// ── Web Worker for Excel parsing (keeps UI responsive) ──────
let _xlsxWorker     = null;
let _workerReady    = false;   // true once Worker confirms XLSX is loaded
const _workerCallbacks = new Map();
let   _workerMsgId  = 0;
const WORKER_TIMEOUT_MS = 120_000; // 2 min max per file

function getWorker() {
  if (_xlsxWorker && _workerReady) return _xlsxWorker;
  if (_xlsxWorker) return null; // still initialising — use fallback
  try {
    _xlsxWorker = new Worker('xlsx-worker.js');
    _xlsxWorker.onmessage = e => {
      if (e.data.ready) { _workerReady = true; return; }
      const cb = _workerCallbacks.get(e.data.id);
      if (cb) { _workerCallbacks.delete(e.data.id); cb(e.data); }
    };
    _xlsxWorker.onerror = err => {
      console.warn('Worker failed, using main thread:', err.message);
      // Reject all pending callbacks so they fall through to the fallback
      for (const [id, cb] of _workerCallbacks) {
        _workerCallbacks.delete(id);
        cb({ id, name: '', tables: null, error: '__worker_failed__' });
      }
      _xlsxWorker.terminate();
      _xlsxWorker = null;
      _workerReady = false;
    };
  } catch { _xlsxWorker = null; }
  return null; // not ready yet on first call
}

// Initialise worker early so it's warm by the time the user picks a file
getWorker();

function parseExcelOnMainThread(buffer, name) {
  try {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true, WTF: false });
    const tables = wb.SheetNames.map(s => {
      try {
        const json = XLSX.utils.sheet_to_json(wb.Sheets[s], { defval: '' });
        if (!json.length) return null;
        return { name: s, columns: Object.keys(json[0]), rows: json };
      } catch { return null; }
    }).filter(Boolean);
    return { name, tables, error: null };
  } catch (err) {
    return { name, tables: null, error: err.message };
  }
}

function parseExcelViaWorker(file, type) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = ev => {
      const buffer = ev.target.result;
      const worker = getWorker();

      if (!worker) {
        // Worker not ready — run on main thread (may freeze briefly for large files)
        const result = parseExcelOnMainThread(buffer, file.name);
        resolve({ ...result, type });
        return;
      }

      const id = ++_workerMsgId;
      let settled = false;

      // Timeout guard — if Worker goes silent, fall back
      const timer = setTimeout(() => {
        if (settled) return;
        settled = true;
        _workerCallbacks.delete(id);
        console.warn(`Worker timed out for ${file.name}, falling back`);
        const result = parseExcelOnMainThread(buffer.slice ? buffer.slice(0) : buffer, file.name);
        resolve({ ...result, type });
      }, WORKER_TIMEOUT_MS);

      _workerCallbacks.set(id, data => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        if (data.error === '__worker_failed__') {
          // Worker crashed mid-flight — re-parse on main thread
          const result = parseExcelOnMainThread(buffer, file.name);
          resolve({ ...result, type });
        } else {
          resolve({ ...data, type });
        }
      });

      worker.postMessage({ id, buffer, name: file.name }, [buffer]);
    };
    reader.onerror = () => resolve({ id: 0, name: file.name, tables: null, error: 'فشل قراءة الملف', type });
    reader.readAsArrayBuffer(file);
  });
}

// ── State ──────────────────────────────────────────────────
const state = {
  sources: [],        // { id, name, type:'access'|'sqlite'|'excel', tables:[{name,columns,rows}], color, totalRows }
  activeTab: 'all',   // tab id or 'all' or 'search'
  searchQuery: '',
  activeFilters: new Set(), // source ids to filter by (empty = all)
  sqlReady: false,
};

let sqlJs = null;     // sql.js SQL namespace

// ── Colors ─────────────────────────────────────────────────
const TYPE_COLOR = {
  access: '#7c3aed',
  sqlite: '#0891b2',
  excel:  '#16a34a',
};
const TYPE_LABEL = {
  access: 'Access',
  sqlite: 'SQLite',
  excel:  'Excel',
};

// ── Unique ID ───────────────────────────────────────────────
let _uid = Date.now();
function uid() { return 'src_' + (++_uid); }

// ═══════════════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initSqlJs();
  loadFromStorage();
  renderAll();

  // Search — debounced 250ms so it doesn't fire on every keystroke
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

  // File inputs
  document.getElementById('sqliteFileInput').addEventListener('change', e => handleSqliteFiles(e));
  document.getElementById('folderFileInput').addEventListener('change', e => handleFolderImport(e));
  document.getElementById('excelFileInput').addEventListener('change', e => handleExcelFiles(e));

  // Export unified Excel
  document.getElementById('exportBtn').addEventListener('click', exportUnified);

  // Clear all
  document.getElementById('clearAllBtn').addEventListener('click', clearAll);

  // Theme
  document.getElementById('themeToggle').addEventListener('click', toggleTheme);
});

// ── sql.js init (local wasm — offline) ───────────────────────
function initSqlJs() {
  const loader = window.initSqlJs || window.SQL;
  if (!loader) { state.sqlReady = false; return; }

  const cfg = { locateFile: f => `vendor/${f}` };
  try {
    loader(cfg).then(SQL => {
      sqlJs = SQL;
      state.sqlReady = true;
    }).catch(() => { state.sqlReady = false; });
  } catch(e) { state.sqlReady = false; }
}

// ═══════════════════════════════════════════════════════════
//  THEME
// ═══════════════════════════════════════════════════════════
function initTheme() {
  const saved = localStorage.getItem('uds_theme') || 'dark';
  applyTheme(saved);
}
function toggleTheme() {
  const isLight = document.body.classList.contains('light');
  applyTheme(isLight ? 'dark' : 'light');
}
function applyTheme(t) {
  document.body.classList.toggle('light', t === 'light');
  const btn = document.getElementById('themeToggle');
  btn.querySelector('.theme-icon').textContent = t === 'light' ? '🌙' : '☀️';
  localStorage.setItem('uds_theme', t);
}

// ═══════════════════════════════════════════════════════════
//  SQLITE FILE HANDLER
// ═══════════════════════════════════════════════════════════
async function handleSqliteFiles(e) {
  const files = Array.from(e.target.files);
  if (!files.length) return;
  e.target.value = '';

  if (!state.sqlReady || !sqlJs) {
    notify('sql.js لم يُحمَّل بعد. تأكد من الاتصال بالإنترنت وأعد المحاولة.', 'error');
    return;
  }

  for (const file of files) {
    showLoading();
    try {
      const buf = await file.arrayBuffer();
      const arr = new Uint8Array(buf);
      const db = new sqlJs.Database(arr);

      const tablesRes = db.exec("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'");
      if (!tablesRes.length || !tablesRes[0].values.length) {
        notify(`${file.name}: لا توجد جداول في قاعدة البيانات.`, 'warning');
        hideLoading();
        continue;
      }

      const tableNames = tablesRes[0].values.map(r => r[0]);
      const tables = [];

      for (const tname of tableNames) {
        try {
          const res = db.exec(`SELECT * FROM "${tname}" LIMIT 5000`);
          if (res.length) {
            tables.push({
              name: tname,
              columns: res[0].columns,
              rows: res[0].values.map(r => Object.fromEntries(res[0].columns.map((c,i) => [c, r[i]]))),
            });
          } else {
            tables.push({ name: tname, columns: [], rows: [] });
          }
        } catch(te) {
          tables.push({ name: tname, columns: [], rows: [] });
        }
      }
      db.close();

      addSource({
        name: file.name,
        type: 'sqlite',
        tables,
        totalRows: tables.reduce((s, t) => s + t.rows.length, 0),
      });
    } catch(err) {
      notify(`خطأ في قراءة ${file.name}: ${err.message}`, 'error');
    }
    hideLoading();
  }
}

// ═══════════════════════════════════════════════════════════
//  EXCEL / CSV FILE HANDLER
// ═══════════════════════════════════════════════════════════
function handleExcelFiles(e) {
  const files = Array.from(e.target.files);
  if (!files.length) return;
  e.target.value = '';
  files.forEach(file => parseExcelFile(file, 'excel'));
}

// ═══════════════════════════════════════════════════════════
//  FOLDER IMPORT  (parallel via Worker — UI stays responsive)
// ═══════════════════════════════════════════════════════════
async function handleFolderImport(e) {
  const all = Array.from(e.target.files);
  e.target.value = '';

  const excelFiles = all.filter(f => /\.(xlsx|xls|csv)$/i.test(f.name));
  if (!excelFiles.length) {
    notify('لا توجد ملفات Excel في المجلد المختار (.xlsx / .xls / .csv)', 'warning', 5000);
    return;
  }

  notify(`جاري استيراد ${excelFiles.length} ملف...`, 'info', 120000);
  let done = 0, errors = 0;

  // Fire all parses in parallel — Worker handles them off the main thread
  const promises = excelFiles.map(file =>
    parseExcelViaWorker(file, 'excel').then(result => {
      done++;
      if (result.error) {
        errors++;
        notify(`خطأ في "${result.name}": ${result.error}`, 'error', 4000);
      } else if (result.tables && result.tables.length) {
        addSource({
          name: result.name, type: 'excel',
          tables: result.tables,
          totalRows: result.tables.reduce((s, t) => s + t.rows.length, 0),
        });
      }
      notify(`جاري الاستيراد... ${done} / ${excelFiles.length}`, 'info', 120000);
    })
  );

  await Promise.all(promises);
  const ok = done - errors;
  notify(`تم استيراد ${ok} ملف بنجاح${errors ? ` (${errors} بها أخطاء)` : ''} ✓`, 'success', 5000);
}

// ── Single Excel file — via Worker ─────────────────────────
async function parseExcelFile(file, type) {
  showLoading();
  const result = await parseExcelViaWorker(file, type);
  if (result.error) {
    notify(`خطأ في قراءة "${file.name}": ${result.error}`, 'error');
  } else if (!result.tables || !result.tables.length) {
    notify(`${file.name}: لا توجد بيانات في الملف.`, 'warning');
  } else {
    addSource({
      name: result.name, type,
      tables: result.tables,
      totalRows: result.tables.reduce((s, t) => s + t.rows.length, 0),
    });
  }
  hideLoading();
}

// ═══════════════════════════════════════════════════════════
//  SOURCE MANAGEMENT
// ═══════════════════════════════════════════════════════════
function buildSearchIndex(tables) {
  for (const t of tables) {
    t._index = t.rows.map(r => Object.values(r).join('\x00').toLowerCase());
  }
}

// Builds index in 500-row chunks using setTimeout so the UI stays responsive
function buildSearchIndexAsync(tables, onDone) {
  const work = tables.map(t => ({ t, i: 0 }));
  function step() {
    const CHUNK = 500;
    let busy = false;
    for (const item of work) {
      if (item.i >= item.t.rows.length) continue;
      busy = true;
      if (!item.t._index) item.t._index = [];
      const end = Math.min(item.i + CHUNK, item.t.rows.length);
      for (let r = item.i; r < end; r++) {
        item.t._index[r] = Object.values(item.t.rows[r]).join('\x00').toLowerCase();
      }
      item.i = end;
      break; // one table chunk per tick
    }
    if (busy) setTimeout(step, 0);
    else onDone && onDone();
  }
  setTimeout(step, 0);
}

// Debounced save — avoids re-serializing everything on every imported file
let _saveTimer = null;
function scheduleSave() {
  clearTimeout(_saveTimer);
  _saveTimer = setTimeout(saveToStorage, 1500);
}

function addSource(src) {
  const existing = state.sources.findIndex(s => s.name === src.name);
  if (existing !== -1) state.sources.splice(existing, 1);

  const id = uid();
  state.sources.push({ id, color: TYPE_COLOR[src.type], ...src });
  state.activeTab = 'all';
  renderAll();
  notify(`تم استيراد "${src.name}" بنجاح (${src.totalRows.toLocaleString('ar')} سجل)`, 'success');

  // Defer heavy work so the UI renders first
  scheduleSave();
  setTimeout(() => buildSearchIndexAsync(src.tables), 50);
}

function deleteSource(id) {
  state.sources = state.sources.filter(s => s.id !== id);
  state.activeFilters.delete(id);
  if (state.activeTab === id) state.activeTab = 'all';
  scheduleSave();
  renderAll();
}

function clearAll() {
  if (!state.sources.length) return;
  if (!confirm('هل تريد مسح جميع البيانات المستوردة؟')) return;
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
//  LOCAL STORAGE
// ═══════════════════════════════════════════════════════════
function saveToStorage() {
  try {
    // Limit storage size: store up to 2000 rows per table
    const slim = state.sources.map(s => ({
      ...s,
      tables: s.tables.map(t => ({
        ...t,
        rows: t.rows.slice(0, 2000),
      })),
    }));
    localStorage.setItem('uds_sources', JSON.stringify(slim));
  } catch(e) {
    // Storage quota exceeded — skip silently
  }
}

function loadFromStorage() {
  try {
    const raw = localStorage.getItem('uds_sources');
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      state.sources = parsed.map(s => {
        buildSearchIndex(s.tables || []);
        return { color: TYPE_COLOR[s.type] || '#64748b', ...s };
      });
    }
  } catch(e) {}
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

// ── Source chips under buttons ──────────────────────────────
function renderChips() {
  ['sqlite','excel'].forEach(type => {
    const container = document.getElementById(type + 'Chips');
    const srcs = state.sources.filter(s => s.type === type);
    container.innerHTML = srcs.map(s => `
      <div class="source-chip" onclick="switchTab('${s.id}')">
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
  });
}

// ── Record badges on import buttons ────────────────────────
function renderBadges() {
  ['sqlite','excel'].forEach(type => {
    const badge = document.getElementById(type + 'Badge');
    const srcs = state.sources.filter(s => s.type === type);
    if (srcs.length) {
      const total = srcs.reduce((s, x) => s + x.totalRows, 0);
      badge.textContent = total.toLocaleString('ar');
      badge.style.display = '';
    } else {
      badge.style.display = 'none';
    }
  });
}

// ── Stats bar ───────────────────────────────────────────────
function renderStatsBar() {
  const bar = document.getElementById('statsBar');
  const exportBtn = document.getElementById('exportBtn');
  if (!state.sources.length) {
    bar.style.display = 'none';
    exportBtn.style.display = 'none';
    return;
  }
  bar.style.display = '';
  exportBtn.style.display = '';
  const totalTables = state.sources.reduce((s, src) => s + src.tables.length, 0);
  const totalRecords = state.sources.reduce((s, src) => s + src.totalRows, 0);
  document.getElementById('totalSources').textContent = state.sources.length;
  document.getElementById('totalTables').textContent = totalTables;
  document.getElementById('totalRecords').textContent = totalRecords.toLocaleString('ar');
}

// ── Search section ──────────────────────────────────────────
function renderSearchSection() {
  const sec = document.getElementById('searchSection');
  if (!state.sources.length) { sec.style.display = 'none'; return; }
  sec.style.display = '';
  renderFilterPills();
}

function renderFilterPills() {
  const container = document.getElementById('searchFilters');
  const pills = [
    { id: 'all_pill', label: 'الكل', color: '#6366f1' },
    ...state.sources.map(s => ({ id: s.id, label: s.name, color: s.color })),
  ];

  container.innerHTML = pills.map(p => {
    const isAll = p.id === 'all_pill';
    const isActive = isAll ? state.activeFilters.size === 0 : state.activeFilters.has(p.id);
    return `
      <button class="filter-pill ${isActive ? 'active' : ''}"
        style="${isActive ? `background:${p.color};border-color:${p.color}` : `border-color:${p.color};color:${p.color}`}"
        onclick="toggleFilter('${p.id}')">
        <span class="pill-dot" style="background:${isActive ? 'rgba(255,255,255,0.7)' : p.color}"></span>
        ${esc(p.label)}
      </button>
    `;
  }).join('');
}

function toggleFilter(id) {
  if (id === 'all_pill') {
    state.activeFilters.clear();
  } else {
    if (state.activeFilters.has(id)) state.activeFilters.delete(id);
    else state.activeFilters.add(id);
  }
  renderFilterPills();
  renderTabContent();
}

// ── Tabs ────────────────────────────────────────────────────
function renderTabs() {
  const tabsBar = document.getElementById('tabsBar');
  const dataArea = document.getElementById('dataArea');
  const emptyState = document.getElementById('emptyState');

  if (!state.sources.length) {
    emptyState.style.display = '';
    dataArea.style.display = 'none';
    tabsBar.innerHTML = '';
    return;
  }
  emptyState.style.display = 'none';
  dataArea.style.display = '';

  const tabs = [
    { id: 'all', label: 'جميع البيانات' },
    ...state.sources.map(s => ({ id: s.id, label: s.name, color: s.color, type: s.type })),
  ];

  tabsBar.innerHTML = tabs.map(t => `
    <button class="tab-btn ${state.activeTab === t.id ? 'active' : ''}"
      onclick="switchTab('${t.id}')">
      ${t.color ? `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${t.color};margin-left:6px;vertical-align:middle"></span>` : ''}
      ${esc(t.label)}
    </button>
  `).join('');
}

function switchTab(id) {
  state.activeTab = id;
  renderTabs();
  renderTabContent();
}

// ═══════════════════════════════════════════════════════════
//  TAB CONTENT  (virtual-scroll based — fast for huge tables)
// ═══════════════════════════════════════════════════════════
const VT_ROW_H = 34;   // px per row
const VT_VIEW_H = 440; // visible viewport height
const VT_BUFFER = 6;   // extra rows above/below window
let   _vtId = 0;
let   _vtRegistry = {}; // id -> { columns, rows, query }

function renderTabContent() {
  const container = document.getElementById('tabContent');
  // Detach old scroll listeners
  _vtRegistry = {};
  if (!state.sources.length) { container.innerHTML = ''; return; }

  const q = state.searchQuery.toLowerCase();

  const activeSrcs = state.activeFilters.size > 0
    ? state.sources.filter(s => state.activeFilters.has(s.id))
    : state.sources;

  let html;
  if (q) {
    html = buildSearchHTML(activeSrcs, q);
  } else if (state.activeTab === 'all') {
    html = activeSrcs.length
      ? activeSrcs.map(src => src.tables.map(t => tableShellHTML(t, src, t.rows, '')).join('')).join('')
      : noResultsHTML('لا توجد مصادر محددة');
  } else {
    const src = state.sources.find(s => s.id === state.activeTab);
    html = src ? src.tables.map(t => tableShellHTML(t, src, t.rows, '')).join('') : '';
  }

  container.innerHTML = html;
  mountVirtualTables();
}

// ── Build the static shell for a table; rows filled by virtualizer ──
function tableShellHTML(table, src, rows, query) {
  if (!table.columns.length) {
    return `<div class="table-wrapper">
      <div class="table-header"><div class="table-title">
        <span class="table-name">${esc(table.name)}</span>
        <span class="table-count" style="background:${src.color}">0</span>
        <span class="result-source-badge badge-${src.type}">${TYPE_LABEL[src.type]}</span>
      </div></div>
      <div class="no-results"><p>الجدول فارغ</p></div>
    </div>`;
  }

  const id = 'vt_' + (++_vtId);
  _vtRegistry[id] = { columns: table.columns, rows, query };

  const gridCols = table.columns.map(() => 'minmax(140px, 1fr)').join(' ');
  const headCells = table.columns.map(c => `<div class="vt-cell vt-th">${esc(String(c))}</div>`).join('');
  const count = rows.length;
  const scrollH = Math.min(count * VT_ROW_H, VT_VIEW_H);

  return `
    <div class="table-wrapper">
      <div class="table-header"><div class="table-title">
        <span class="table-name">${esc(table.name)}</span>
        <span class="table-count" style="background:${src.color}">${count.toLocaleString('ar')}</span>
        <span class="result-source-badge badge-${src.type}">${TYPE_LABEL[src.type]}</span>
        <span style="font-size:0.78rem;color:var(--color-text-dim)">${esc(src.name)}</span>
      </div></div>
      <div class="vt-outer">
        <div class="vt-track" style="--vt-cols:${gridCols}">
          <div class="vt-head">${headCells}</div>
          <div class="vt-scroll" id="${id}" style="height:${scrollH}px">
            <div class="vt-canvas" style="height:${count * VT_ROW_H}px">
              <div class="vt-win"></div>
            </div>
          </div>
        </div>
      </div>
    </div>`;
}

// ── Attach virtual scrollers to all mounted tables ──────────
function mountVirtualTables() {
  for (const id of Object.keys(_vtRegistry)) {
    const scrollEl = document.getElementById(id);
    if (!scrollEl) continue;
    const winEl = scrollEl.querySelector('.vt-win');
    const data  = _vtRegistry[id];

    const render = () => {
      const scrollTop = scrollEl.scrollTop;
      const total = data.rows.length;
      let start = Math.floor(scrollTop / VT_ROW_H) - VT_BUFFER;
      if (start < 0) start = 0;
      const visible = Math.ceil(VT_VIEW_H / VT_ROW_H) + VT_BUFFER * 2;
      let end = start + visible;
      if (end > total) end = total;

      let html = '';
      for (let i = start; i < end; i++) {
        const row = data.rows[i];
        let cells = '';
        for (const col of data.columns) {
          const val = row[col] != null ? String(row[col]) : '';
          cells += `<div class="vt-cell" title="${esc(val)}">${data.query ? highlightText(val, data.query) : esc(val)}</div>`;
        }
        html += `<div class="vt-row">${cells}</div>`;
      }
      winEl.style.transform = `translateY(${start * VT_ROW_H}px)`;
      winEl.innerHTML = html;
    };

    scrollEl.addEventListener('scroll', () => {
      // rAF-throttle so scrolling stays smooth
      if (scrollEl._raf) return;
      scrollEl._raf = requestAnimationFrame(() => { scrollEl._raf = null; render(); });
    });
    render();
  }
}

// ── Search: filter rows then reuse the virtual table shell ──
function buildSearchHTML(srcs, q) {
  let totalMatches = 0;
  const groups = [];

  for (const src of srcs) {
    for (const table of src.tables) {
      const idx = table._index;
      const matched = [];
      for (let i = 0; i < table.rows.length; i++) {
        if (rowMatches(table.rows[i], q, idx ? idx[i] : undefined)) matched.push(table.rows[i]);
      }
      if (matched.length) {
        totalMatches += matched.length;
        groups.push({ src, table, rows: matched });
      }
    }
  }

  if (!totalMatches) {
    return noResultsHTML(`لا توجد نتائج لـ "${esc(q)}"`, `جرّب كلمات مختلفة أو اختر مصادر أخرى`);
  }

  const header = `
    <div class="results-header">
      <div class="results-count">
        نتائج البحث: <strong>${totalMatches.toLocaleString('ar')}</strong> سجل
        في <strong>${groups.length}</strong> جدول
        لـ "<strong>${esc(q)}</strong>"
      </div>
    </div>`;

  return header + groups.map(g => tableShellHTML(g.table, g.src, g.rows, q)).join('');
}

// ═══════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════
function rowMatches(row, q, indexStr) {
  // Use pre-built index string when available (much faster for large tables)
  if (indexStr !== undefined) return indexStr.includes(q);
  return Object.values(row).some(v => v != null && String(v).toLowerCase().includes(q));
}

function highlightText(text, query) {
  if (!query || !text) return esc(text);
  const escaped = esc(text);
  const re = new RegExp(escapeRegex(query), 'gi');
  return escaped.replace(re, m => `<mark>${m}</mark>`);
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function esc(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function noResultsHTML(title, desc = '') {
  return `
    <div class="no-results">
      <svg viewBox="0 0 48 48" fill="none" width="64" height="64">
        <circle cx="24" cy="24" r="22" stroke="currentColor" stroke-width="1.5" stroke-dasharray="5 3"/>
        <path d="M16 24h16M24 16v16" stroke="currentColor" stroke-width="1.5" opacity="0.5" transform="rotate(45 24 24)"/>
      </svg>
      <h3>${title}</h3>
      ${desc ? `<p>${desc}</p>` : ''}
    </div>`;
}

// ── Loading indicator ───────────────────────────────────────
let _loadingCount = 0;
function showLoading() {
  _loadingCount++;
  let el = document.getElementById('globalLoading');
  if (!el) {
    el = document.createElement('div');
    el.id = 'globalLoading';
    el.className = 'loading';
    el.innerHTML = '<div class="spinner"></div><span>جاري المعالجة...</span>';
    document.getElementById('tabContent').prepend(el);
  }
}
function hideLoading() {
  _loadingCount = Math.max(0, _loadingCount - 1);
  if (!_loadingCount) {
    const el = document.getElementById('globalLoading');
    if (el) el.remove();
  }
}

// ── Notifications ───────────────────────────────────────────
let _notifTimer = null;
function notify(msg, type = 'info', duration = 3500) {
  const el = document.getElementById('notification');
  el.className = `notification ${type}`;
  el.textContent = msg;
  el.style.display = '';
  clearTimeout(_notifTimer);
  _notifTimer = setTimeout(() => { el.style.display = 'none'; }, duration);
}

// ── clearSearch ─────────────────────────────────────────────
function clearSearch() {
  state.searchQuery = '';
  document.getElementById('searchInput').value = '';
  document.getElementById('searchClear').style.display = 'none';
  renderTabContent();
}

// ── Export unified Excel ────────────────────────────────────
function exportUnified() {
  if (!state.sources.length) return;
  if (typeof XLSX === 'undefined') { notify('SheetJS غير متاح', 'error'); return; }

  const btn = document.getElementById('exportBtn');
  btn.disabled = true;
  btn.textContent = 'جاري التصدير...';

  try {
    const wb = XLSX.utils.book_new();
    const usedNames = new Set();

    for (const src of state.sources) {
      for (const table of src.tables) {
        // Build a unique sheet name ≤31 chars (Excel limit)
        let sheetName = `${src.name.replace(/\.[^.]+$/, '')}_${table.name}`.slice(0, 31);
        // Deduplicate
        let base = sheetName, n = 2;
        while (usedNames.has(sheetName)) sheetName = `${base.slice(0, 28)}_${n++}`;
        usedNames.add(sheetName);

        const ws = XLSX.utils.json_to_sheet(table.rows);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      }
    }

    const date = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `بيانات_موحدة_${date}.xlsx`);
    notify('تم تصدير الملف بنجاح ✓', 'success');
  } catch (err) {
    notify(`خطأ في التصدير: ${err.message}`, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> تصدير موحّد`;
  }
}

// ── Expose globals needed by inline onclick ─────────────────
window.clearSearch = clearSearch;
window.switchTab = switchTab;
window.deleteSource = deleteSource;
window.toggleFilter = toggleFilter;
