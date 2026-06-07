'use strict';

const express = require('express');
const multer  = require('multer');
const path    = require('path');
const MDBReader = require('mdb-reader').default;

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 200 * 1024 * 1024 } });

// Allow cross-origin requests from file:// or other origins during development
app.use((_req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  next();
});

// ── POST /api/access ────────────────────────────────────────
// Accepts a single .mdb or .accdb file, returns JSON with all tables.
app.post('/api/access', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'لم يتم إرفاق ملف' });
  }

  const ext = path.extname(req.file.originalname).toLowerCase();
  if (!['.mdb', '.accdb'].includes(ext)) {
    return res.status(400).json({ error: 'نوع الملف غير مدعوم. يُقبل .mdb و .accdb فقط.' });
  }

  let reader;
  try {
    reader = new MDBReader(req.file.buffer);
  } catch (err) {
    return res.status(422).json({
      error: `تعذّر فتح الملف — قد يكون تالفاً أو بصيغة غير مدعومة: ${err.message}`,
    });
  }

  let tableNames;
  try {
    tableNames = reader.getTableNames({ includeSystemTables: false });
  } catch (err) {
    return res.status(422).json({ error: `تعذّر قراءة أسماء الجداول: ${err.message}` });
  }

  if (!tableNames.length) {
    return res.status(422).json({ error: 'لا توجد جداول في قاعدة البيانات.' });
  }

  const tables = [];
  for (const name of tableNames) {
    try {
      const table   = reader.getTable(name);
      const columns = table.getColumnNames();
      // getData() returns an array of plain objects keyed by column name
      const rawRows = table.getData({ rowLimit: 5000 });

      const rows = rawRows.map(r => {
        const obj = {};
        for (const col of columns) {
          const val = r[col];
          if (val == null)              obj[col] = '';
          else if (val instanceof Date) obj[col] = val.toISOString();
          else if (Buffer.isBuffer(val)) obj[col] = `[Binary ${val.length}B]`;
          else                           obj[col] = String(val);
        }
        return obj;
      });

      if (columns.length) tables.push({ name, columns, rows });
    } catch (err) {
      // Skip tables that cannot be read but continue with the rest
      tables.push({ name, columns: [], rows: [], error: err.message });
    }
  }

  const readable = tables.filter(t => t.columns.length > 0);
  if (!readable.length) {
    return res.status(422).json({ error: 'لم يتم العثور على أي جدول قابل للقراءة في الملف.' });
  }

  res.json({
    name:      req.file.originalname,
    tables:    readable,
    totalRows: readable.reduce((s, t) => s + t.rows.length, 0),
  });
});

// ── Health check ────────────────────────────────────────────
app.get('/api/health', (_req, res) => res.json({ ok: true }));

// Serve the static front-end — registered AFTER API routes so Express
// never intercepts API requests with the static file handler.
app.use(express.static(path.join(__dirname)));

// Global error handler — always returns JSON so the browser never sees HTML
app.use((err, _req, res, _next) => {
  console.error(err);
  res.status(err.status || 500).json({ error: err.message || 'خطأ داخلي في الخادم' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✓ Unified DB Search server running → http://localhost:${PORT}`);
});
