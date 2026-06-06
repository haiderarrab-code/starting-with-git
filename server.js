'use strict';

const express = require('express');
const multer  = require('multer');
const path    = require('path');
const MDBReader = require('mdb-reader');

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
      const rawRows = table.getData();

      // Normalise each row from array → object, cap at 5000 rows
      const rows = rawRows.slice(0, 5000).map(r =>
        Object.fromEntries(columns.map((c, i) => [c, r[i] != null ? r[i] : '']))
      );

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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✓ Unified DB Search server running → http://localhost:${PORT}`);
});
