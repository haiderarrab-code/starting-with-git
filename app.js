/* =========================================================
   Unified Database Search System — app.js
   ========================================================= */

'use strict';

// ── State ──────────────────────────────────────────────────
const state = {
  sources: [],        // { id, name, type:'access'|'sqlite'|'excel', tables:[{name,columns,rows}], color, totalRows }
  activeTab: 'all',   // tab id or 'all' or 'search'
  searchQuery: '',
  activeFilters: new Set(), // source ids to filter by (empty = all)
  sqlReady: false,
};

let sqlJs = null;     // sql.js SQL namespace
const BACKEND = 'http://localhost:3000';
let backendAvailable = false;

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
  checkBackend();
  loadFromStorage();
  renderAll();

  // Search — debounced 350ms so it doesn't fire on every keystroke
  let _searchTimer = null;
  const si = document.getElementById('searchInput');
  si.addEventListener('input', () => {
    const val = si.value.trim();
    document.getElementById('searchClear').style.display = val ? '' : 'none';
    clearTimeout(_searchTimer);
    _searchTimer = setTimeout(() => {
      state.searchQuery = val;
      renderTabContent();
    }, 350);
  });

  // File inputs
  document.getElementById('accessFileInput').addEventListener('change', e => handleAccessFiles(e));
  document.getElementById('sqliteFileInput').addEventListener('change', e => handleSqliteFiles(e));
  document.getElementById('excelFileInput').addEventListener('change', e => handleExcelFiles(e));

  // Access button: open file picker directly (backend handles .mdb/.accdb)
  const accessBtn = document.getElementById('accessBtn');
  accessBtn.onclick = null; // remove inline
  accessBtn.addEventListener('click', () => {
    document.getElementById('accessFileInput').click();
  });

  // Clear all
  document.getElementById('clearAllBtn').addEventListener('click', clearAll);

  // Theme
  document.getElementById('themeToggle').addEventListener('click', toggleTheme);
});

// ── Backend health check ─────────────────────────────────────
async function checkBackend() {
  try {
    const r = await fetch(`${BACKEND}/api/health`, { signal: AbortSignal.timeout(2000) });
    backendAvailable = r.ok;
  } catch {
    backendAvailable = false;
  }
  updateAccessBtnTitle();
}

function updateAccessBtnTitle() {
  const btn = document.getElementById('accessBtn');
  btn.title = backendAvailable
    ? 'استيراد ملف Access (.mdb / .accdb)'
    : 'الخادم غير متاح — شغّل: npm start';
}

// ── sql.js init ─────────────────────────────────────────────
function initSqlJs() {
  // sql.js CDN exposes window.initSqlJs
  const loader = window.initSqlJs || window.SQL;
  if (!loader) { state.sqlReady = false; return; }

  const cfg = { locateFile: f => `https://cdn.jsdelivr.net/npm/sql.js@1.10.2/dist/${f}` };
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
//  ACCESS FILE HANDLER
// ═══════════════════════════════════════════════════════════
async function handleAccessFiles(e) {
  const files = Array.from(e.target.files);
  if (!files.length) return;
  e.target.value = '';

  const nativeFiles = files.filter(f => /\.(mdb|accdb)$/i.test(f.name));
  const csvFiles    = files.filter(f => /\.csv$/i.test(f.name));

  // CSV exports from Access → parsed directly in the browser
  csvFiles.forEach(file => parseExcelFile(file, 'access'));

  // Native Access files → upload to Node.js backend
  if (!nativeFiles.length) return;

  // Re-check backend availability before attempting upload
  await checkBackend();

  if (!backendAvailable) {
    document.getElementById('accessModal').style.display = 'flex';
    return;
  }

  for (const file of nativeFiles) {
    showLoading();
    try {
      const formData = new FormData();
      formData.append('file', file);

      notify(`جاري رفع ومعالجة "${file.name}"... قد يستغرق بعض الوقت للملفات الكبيرة`, 'info', 30000);
      const res  = await fetch(`${BACKEND}/api/access`, { method: 'POST', body: formData });
      const text = await res.text();

      let json;
      try {
        json = JSON.parse(text);
      } catch {
        // Server returned non-JSON (e.g. HTML error page from a proxy)
        notify(
          `استجابة غير متوقعة من الخادم أثناء معالجة "${file.name}". تأكد أن npm start يعمل على المنفذ 3000.`,
          'error', 7000
        );
        hideLoading();
        continue;
      }

      if (!res.ok) {
        notify(`خطأ في "${file.name}": ${json.error || res.statusText}`, 'error', 6000);
        hideLoading();
        continue;
      }

      addSource({ name: json.name, type: 'access', tables: json.tables, totalRows: json.totalRows });
    } catch (err) {
      // Network-level failure (server not running, CORS, etc.)
      backendAvailable = false;
      updateAccessBtnTitle();
      notify(
        `تعذّر الاتصال بالخادم أثناء معالجة "${file.name}". شغّل: npm start`,
        'error', 7000
      );
    }
    hideLoading();
  }
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

function parseExcelFile(file, type) {
  if (typeof XLSX === 'undefined') {
    notify('SheetJS لم يُحمَّل. تأكد من الاتصال بالإنترنت.', 'error');
    return;
  }

  showLoading();
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const data = ev.target.result;
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });

      const tables = workbook.SheetNames.map(sheetName => {
        const ws = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
        if (!json.length) return { name: sheetName, columns: [], rows: [] };
        const columns = Object.keys(json[0]);
        return { name: sheetName, columns, rows: json };
      }).filter(t => t.columns.length > 0);

      if (!tables.length) {
        notify(`${file.name}: لا توجد بيانات في الملف.`, 'warning');
        hideLoading();
        return;
      }

      addSource({
        name: file.name,
        type,
        tables,
        totalRows: tables.reduce((s, t) => s + t.rows.length, 0),
      });
    } catch(err) {
      notify(`خطأ في قراءة ${file.name}: ${err.message}`, 'error');
    }
    hideLoading();
  };
  reader.onerror = () => { notify(`فشل قراءة الملف ${file.name}`, 'error'); hideLoading(); };
  reader.readAsArrayBuffer(file);
}

// ═══════════════════════════════════════════════════════════
//  SOURCE MANAGEMENT
// ═══════════════════════════════════════════════════════════
function buildSearchIndex(tables) {
  // Pre-build one lowercase string per row for fast includes() search
  for (const t of tables) {
    t._index = t.rows.map(r => Object.values(r).join('\x00').toLowerCase());
  }
}

function addSource(src) {
  // Avoid duplicate file names
  const existing = state.sources.findIndex(s => s.name === src.name);
  if (existing !== -1) {
    state.sources.splice(existing, 1);
  }

  buildSearchIndex(src.tables);

  const id = uid();
  state.sources.push({ id, color: TYPE_COLOR[src.type], ...src });
  state.activeTab = 'all';
  saveToStorage();
  renderAll();
  notify(`تم استيراد "${src.name}" بنجاح (${src.totalRows.toLocaleString('ar')} سجل)`, 'success');
}

function deleteSource(id) {
  state.sources = state.sources.filter(s => s.id !== id);
  state.activeFilters.delete(id);
  if (state.activeTab === id) state.activeTab = 'all';
  saveToStorage();
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
  ['access','sqlite','excel'].forEach(type => {
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
  ['access','sqlite','excel'].forEach(type => {
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
  if (!state.sources.length) { bar.style.display = 'none'; return; }
  bar.style.display = '';
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
//  TAB CONTENT
// ═══════════════════════════════════════════════════════════
function renderTabContent() {
  const container = document.getElementById('tabContent');
  if (!state.sources.length) { container.innerHTML = ''; return; }

  const q = state.searchQuery.toLowerCase();

  // Active sources based on filters
  const activeSrcs = state.activeFilters.size > 0
    ? state.sources.filter(s => state.activeFilters.has(s.id))
    : state.sources;

  if (q) {
    container.innerHTML = renderSearchResults(activeSrcs, q);
    return;
  }

  if (state.activeTab === 'all') {
    container.innerHTML = renderAllTables(activeSrcs);
  } else {
    const src = state.sources.find(s => s.id === state.activeTab);
    if (src) {
      container.innerHTML = renderSourceTables(src);
    }
  }
}

// ── Render all tables (merged view) ────────────────────────
function renderAllTables(srcs) {
  if (!srcs.length) return noResultsHTML('لا توجد مصادر محددة');
  return srcs.map(src => renderSourceTables(src)).join('');
}

// ── Render tables for one source ───────────────────────────
function renderSourceTables(src) {
  return src.tables.map(t => renderTable(t, src)).join('');
}

// ── Render a single table ───────────────────────────────────
function renderTable(table, src, highlightQuery = '') {
  if (!table.columns.length) {
    return `<div class="table-wrapper" style="margin-bottom:16px">
      <div class="table-header">
        <div class="table-title">
          <span class="table-name">${esc(table.name)}</span>
          <span class="table-count" style="background:${src.color}">0</span>
          <span class="result-source-badge badge-${src.type}">${TYPE_LABEL[src.type]}</span>
        </div>
      </div>
      <div class="no-results"><p>الجدول فارغ</p></div>
    </div>`;
  }

  // In search mode rows are already filtered upstream; just cap display
  const rows = highlightQuery ? table.rows : table.rows;
  const limit = highlightQuery ? MAX_RESULTS_PER_TABLE : 500;
  const displayRows = rows.slice(0, limit);

  const headerCells = table.columns.map(c => `<th>${esc(String(c))}</th>`).join('');
  const bodyRows = displayRows.map(row => {
    const cells = table.columns.map(col => {
      const val = row[col] != null ? String(row[col]) : '';
      return `<td title="${esc(val)}">${highlightQuery ? highlightText(val, highlightQuery) : esc(val)}</td>`;
    }).join('');
    return `<tr>${cells}</tr>`;
  }).join('');

  const moreRows = rows.length > 500 ? `<div style="padding:10px 16px;font-size:0.82rem;color:var(--color-text-dim);border-top:1px solid var(--color-border)">يُعرض أول 500 من ${rows.length.toLocaleString('ar')} سجل</div>` : '';

  return `
    <div class="table-wrapper" style="margin-bottom:16px">
      <div class="table-header">
        <div class="table-title">
          <span class="table-name">${esc(table.name)}</span>
          <span class="table-count" style="background:${src.color}">${rows.length.toLocaleString('ar')}</span>
          <span class="result-source-badge badge-${src.type}">${TYPE_LABEL[src.type]}</span>
          <span style="font-size:0.78rem;color:var(--color-text-dim)">${esc(src.name)}</span>
        </div>
      </div>
      <div class="table-scroll">
        <table>
          <thead><tr>${headerCells}</tr></thead>
          <tbody>${bodyRows}</tbody>
        </table>
      </div>
      ${moreRows}
    </div>`;
}

// ── Search results ──────────────────────────────────────────
const MAX_RESULTS_PER_TABLE = 100;

function renderSearchResults(srcs, q) {
  let totalMatches = 0;
  const groups = [];

  for (const src of srcs) {
    for (const table of src.tables) {
      const idx = table._index;
      const matched = [];
      for (let i = 0; i < table.rows.length; i++) {
        if (rowMatches(table.rows[i], q, idx ? idx[i] : undefined)) {
          matched.push(table.rows[i]);
        }
      }
      if (matched.length) {
        totalMatches += matched.length;
        groups.push({ src, table: { ...table, rows: matched } });
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
        ${totalMatches > MAX_RESULTS_PER_TABLE * groups.length ? `<span style="color:var(--color-text-dim);font-size:0.8rem">(يُعرض أول ${MAX_RESULTS_PER_TABLE} لكل جدول)</span>` : ''}
      </div>
    </div>`;

  const tables = groups.map(g => renderTable(g.table, g.src, q)).join('');
  return header + tables;
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

// ── Modal ───────────────────────────────────────────────────
function closeAccessModal() {
  document.getElementById('accessModal').style.display = 'none';
}

// Close modal on overlay click
document.addEventListener('click', e => {
  if (e.target.id === 'accessModal') closeAccessModal();
});

// ── clearSearch ─────────────────────────────────────────────
function clearSearch() {
  state.searchQuery = '';
  document.getElementById('searchInput').value = '';
  document.getElementById('searchClear').style.display = 'none';
  renderTabContent();
}

// ── Expose globals needed by inline onclick ─────────────────
window.closeAccessModal = closeAccessModal;
window.clearSearch = clearSearch;
window.switchTab = switchTab;
window.deleteSource = deleteSource;
window.toggleFilter = toggleFilter;
