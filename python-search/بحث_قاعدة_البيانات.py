# -*- coding: utf-8 -*-
"""
نظام البحث الموحّد في قواعد البيانات — قسم المعلومات المركزية
==============================================================

يفتح Excel / CSV / SQLite / Access، يبني فهرساً مقلوباً على القرص
(SQLite FTS5 trigram) مرة واحدة، فيصبح البحث الجزئي في كل الأعمدة
بأقل من ثانية حتى على 10 ملايين صف وأكثر، بذاكرة منخفضة.

المتطلبات: Python 3.11+ ، pandas ، openpyxl   (pyodbc اختيارياً لملفات Access)
"""
import csv
import hashlib
import json
import os
import queue
import sqlite3
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from urllib.request import pathname2url

import pandas as pd


def ro_connect(path):
    """اتصال SQLite للقراءة فقط يعمل حتى مع مسارات تحوي % ? # (ترميز URI صحيح)."""
    uri = "file:" + pathname2url(os.path.abspath(path)) + "?mode=ro"
    return sqlite3.connect(uri, uri=True)

APP_TITLE = "نظام البحث الموحّد في قواعد البيانات"
CREDIT = "اعداد قسم المعلومات المركزي"
WATERMARK = "قسم المعلومات المركزية"

PAGE = 1000                 # عدد الصفوف المعروضة في الجدول
EXPORT_CAP = 1_000_000      # أقصى عدد صفوف يُصدَّر دفعة واحدة
COUNT_CAP = 100_000         # سقف العدّ (يُعرض «أكثر من» بعده لإبقاء الواجهة سريعة)
MIN_FAST_QUERY = 3          # أقل طول للبحث المفهرس التلقائي (trigram)
ALL_KEY = "🔍 كل المصادر (بحث شامل)"
SOURCE_COL = "📄 المصدر"

EXCEL_EXTS = (".xlsx", ".xls", ".xlsm")
CSV_EXTS = (".csv",)
SQLITE_EXTS = (".db", ".sqlite", ".sqlite3")
ACCESS_EXTS = (".mdb", ".accdb")
FOLDER_EXTS = EXCEL_EXTS + CSV_EXTS          # ما يُستورد من مجلد كامل

SEP = "\x01"                # فاصل داخلي بين قيم الأعمدة
IDX_SUFFIX = ".udsidx"      # امتداد ملف الفهرس
INDEX_VERSION = "2"         # عند تغيير بنية الفهرس يرتفع الرقم فيُعاد البناء
CONFIG_NAME = "uds_config.json"
DATA_DIRS = ["قواعد الاكسل", "قواعد البيانات", "البيانات", "data"]


# ════════════════════════════════════════════════════════════════
#  مسارات ومساعدات عامة
# ════════════════════════════════════════════════════════════════
def base_dir():
    """مجلد البرنامج (بجانب السكربت أو ملف exe)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def appdata_dir():
    """مجلد بيانات المستخدم — يُستخدم للفهرس عندما يكون مجلد المصدر للقراءة فقط."""
    if os.name == "nt":
        root = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    else:
        root = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
    d = os.path.join(root, "UnifiedSearch")
    os.makedirs(d, exist_ok=True)
    return d


def _dir_writable(d):
    try:
        probe = os.path.join(d, ".uds_write_test")
        with open(probe, "w") as f:
            f.write("x")
        os.remove(probe)
        return True
    except Exception:
        return False


def index_path_for(source_path):
    """موقع ملف الفهرس: بجانب المصدر إن أمكن، وإلا في مجلد بيانات المستخدم."""
    if os.path.isdir(source_path):
        cand_dir, cand_name = source_path, "_uds_index" + IDX_SUFFIX
    else:
        cand_dir = os.path.dirname(source_path) or "."
        cand_name = os.path.basename(source_path) + IDX_SUFFIX
    if _dir_writable(cand_dir):
        return os.path.join(cand_dir, cand_name)
    h = hashlib.sha1(os.path.abspath(source_path).encode("utf-8")).hexdigest()[:16]
    return os.path.join(appdata_dir(), f"idx_{h}{IDX_SUFFIX}")


def list_data_files(folder):
    """كل ملفات Excel/CSV في المجلد وفروعه، مرتبة."""
    files = []
    for root, _dirs, names in os.walk(folder):
        for fn in names:
            if fn.lower().endswith(FOLDER_EXTS) and not fn.startswith("~$"):
                files.append(os.path.join(root, fn))
    files.sort()
    return files


def fingerprint(source_path):
    """بصمة للمصدر تتغيّر عند أي إضافة/تعديل/حذف — لتحديد صلاحية الفهرس."""
    items = [INDEX_VERSION]
    try:
        if os.path.isdir(source_path):
            for fp in list_data_files(source_path):
                st = os.stat(fp)
                items.append(f"{os.path.relpath(fp, source_path)}|{st.st_size}|{st.st_mtime_ns}")
        else:
            st = os.stat(source_path)
            items.append(f"{st.st_size}|{st.st_mtime_ns}")
    except OSError:
        items.append(str(time.time()))       # تعذّر القراءة → اعتبره متغيراً
    return hashlib.sha256("\n".join(items).encode("utf-8")).hexdigest()


def _cell(v):
    """تحويل قيمة خلية إلى نص آمن للفهرسة."""
    if v is None:
        return ""
    s = str(v)
    return s.replace(SEP, " ") if SEP in s else s


def _like_escape(q):
    """تهريب أحرف البدل في LIKE."""
    return q.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


# ════════════════════════════════════════════════════════════════
#  محرّك الفهرسة والبحث (SQLite + FTS5 trigram على القرص)
# ════════════════════════════════════════════════════════════════
class Engine:
    """فهرس مقلوب على القرص. القراءة آمنة من أي خيط (اتصال لكل استدعاء)."""

    def __init__(self, idx_path):
        self.idx_path = idx_path
        self.sources = {}    # source -> {"columns": [...], "nrows": int}

    def _ro(self):
        return ro_connect(self.idx_path)

    # ---- بيانات وصفية ----
    def load_meta(self):
        con = self._ro()
        try:
            self.sources = {
                source: {"columns": json.loads(cols), "nrows": nrows}
                for source, cols, nrows in con.execute(
                    "SELECT source, columns, nrows FROM meta ORDER BY source")
            }
        finally:
            con.close()
        return self.sources

    @staticmethod
    def stored_fingerprint(idx_path):
        try:
            con = ro_connect(idx_path)
            try:
                row = con.execute("SELECT v FROM kv WHERE k='fingerprint'").fetchone()
                return row[0] if row else None
            finally:
                con.close()
        except Exception:
            return None

    # ---- البناء ----
    @staticmethod
    def build(idx_path, sources_iter, fp, progress_cb=None, source_errors=None):
        """يبني الفهرس في ملف مؤقت ثم يستبدله ذرّياً.

        sources_iter يُنتج (source_name, columns, row_iterable).
        أسماء المصادر المكررة تُميَّز تلقائياً بلاحقة (2)، (3)…
        أي مصدر يفشل أثناء القراءة يُتخطّى (يُسجَّل في source_errors) دون
        إسقاط بقية الفهرس.
        """
        tmp = idx_path + ".tmp"
        for leftover in (tmp, tmp + "-wal", tmp + "-shm"):
            if os.path.exists(leftover):
                os.remove(leftover)
        con = sqlite3.connect(tmp)
        con.execute("PRAGMA journal_mode=WAL")
        con.execute("PRAGMA synchronous=OFF")
        con.execute("PRAGMA temp_store=MEMORY")
        con.execute("PRAGMA cache_size=-65536")          # 64MB
        con.execute("CREATE TABLE kv(k TEXT PRIMARY KEY, v TEXT)")
        con.execute("CREATE TABLE meta(source TEXT PRIMARY KEY, columns TEXT, nrows INT)")
        con.execute("CREATE TABLE docs(id INTEGER PRIMARY KEY, source TEXT, vals TEXT)")
        con.execute("CREATE VIRTUAL TABLE fts USING fts5(content, content='', tokenize='trigram')")
        rid = 0
        seen = {}
        try:
            for source, columns, rowiter in sources_iter:
                # تمييز الأسماء المكررة
                if source in seen:
                    seen[source] += 1
                    source = f"{source} ({seen[source]})"
                seen.setdefault(source, 1)

                n = 0
                dbatch, fbatch = [], []
                try:
                    for values in rowiter:
                        rid += 1
                        n += 1
                        s = SEP.join(_cell(v) for v in values)
                        dbatch.append((rid, source, s))
                        fbatch.append((rid, s))
                        if len(dbatch) >= 20000:
                            con.executemany("INSERT INTO docs VALUES(?,?,?)", dbatch)
                            con.executemany("INSERT INTO fts(rowid,content) VALUES(?,?)", fbatch)
                            dbatch.clear(); fbatch.clear()
                            if progress_cb:
                                progress_cb(source, rid)
                    if dbatch:
                        con.executemany("INSERT INTO docs VALUES(?,?,?)", dbatch)
                        con.executemany("INSERT INTO fts(rowid,content) VALUES(?,?)", fbatch)
                    con.execute("INSERT INTO meta VALUES(?,?,?)",
                                (source, json.dumps(list(columns), ensure_ascii=False), n))
                    con.commit()
                except Exception as e:
                    # مصدر تالف أثناء القراءة: تجاهل صفوفه غير المثبّتة وتابع
                    con.rollback()
                    if source_errors is not None:
                        source_errors.append(f"{source}: {e}")
                    continue
                if progress_cb:
                    progress_cb(source, rid)
            if progress_cb:
                progress_cb("ضغط الفهرس", rid)
            con.execute("INSERT INTO fts(fts) VALUES('optimize')")
            con.execute("INSERT INTO kv VALUES('fingerprint',?)", (fp,))
            con.execute("INSERT INTO kv VALUES('version',?)", (INDEX_VERSION,))
            con.commit()
            # تحويل الفهرس النهائي إلى وضع DELETE حتى لا تُنشأ ملفات -wal/-shm
            # جانبية عند فتحه للقراءة (أنظف، ويعمل حتى من مواقع للقراءة فقط)
            con.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            con.execute("PRAGMA journal_mode=DELETE")
            con.commit()
        finally:
            con.close()
        # استبدال ذرّي (مع إعادة محاولة على ويندوز إذا كان الملف القديم مفتوحاً)
        for attempt in range(10):
            try:
                os.replace(tmp, idx_path)
                break
            except PermissionError:
                if attempt == 9:
                    raise
                time.sleep(0.3)
        for leftover in (tmp + "-wal", tmp + "-shm", tmp + "-journal"):
            if os.path.exists(leftover):
                try:
                    os.remove(leftover)
                except OSError:
                    pass

    # ---- البحث ----
    def _where(self, name, q):
        """يبني (sql_from_where, params) حسب نوع الاستعلام."""
        if q and len(q) >= MIN_FAST_QUERY:
            match = '"' + q.replace('"', '""') + '"'
            base = "FROM fts JOIN docs d ON d.id = fts.rowid WHERE fts.content MATCH ?"
            if name == ALL_KEY:
                return base, [match]
            return base + " AND d.source = ?", [match, name]
        if q:   # استعلام قصير: مسح LIKE (يُستدعى يدوياً بـ Enter)
            like = "%" + _like_escape(q) + "%"
            base = "FROM docs d WHERE d.vals LIKE ? ESCAPE '\\'"
            if name == ALL_KEY:
                return base, [like]
            return base + " AND d.source = ?", [like, name]
        base = "FROM docs d"
        if name == ALL_KEY:
            return base, []
        return base + " WHERE d.source = ?", [name]

    def search(self, name, q, limit=PAGE):
        """يرجع (total, capped, rows). capped=True يعني العدد بلغ السقف
        (مصطلح شائع جداً) فيُعرض كـ «أكثر من …» بدل انتظار عدٍّ كامل بطيء."""
        con = self._ro()
        try:
            fw, params = self._where(name, q)
            capped = False
            if not q:   # العدد معروف من البيانات الوصفية — بلا مسح
                if name == ALL_KEY:
                    total = sum(s["nrows"] for s in self.sources.values())
                else:
                    total = self.sources.get(name, {}).get("nrows", 0)
            else:
                # عدّ محدود بسقف: يتوقّف مبكراً للمصطلحات الشائعة جداً
                total = con.execute(
                    f"SELECT count(*) FROM (SELECT 1 {fw} LIMIT ?)",
                    params + [COUNT_CAP + 1]).fetchone()[0]
                if total > COUNT_CAP:
                    total = COUNT_CAP
                    capped = True
            rows = con.execute(
                f"SELECT d.source, d.vals {fw} LIMIT ?", params + [limit]).fetchall()
            return total, capped, rows
        finally:
            con.close()

    def iter_all(self, name, q, cap=EXPORT_CAP):
        """يتدفّق بكل نتائج الاستعلام (للتصدير)."""
        con = self._ro()
        try:
            fw, params = self._where(name, q)
            cur = con.execute(f"SELECT d.source, d.vals {fw} LIMIT ?", params + [cap])
            yield from cur
        finally:
            con.close()


# ════════════════════════════════════════════════════════════════
#  قُرّاء المصادر — كلٌّ يُنتج (source, columns, row_iterable)
# ════════════════════════════════════════════════════════════════
def _detect_csv_encoding(path):
    """يكتشف الترميز عبر BOM أولاً ثم بالتجربة على عيّنة كاملة القراءة."""
    with open(path, "rb") as f:
        head = f.read(4)
    if head.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    if head.startswith(b"\xff\xfe\x00\x00"):
        return "utf-32-le"
    if head.startswith(b"\x00\x00\xfe\xff"):
        return "utf-32-be"
    if head.startswith(b"\xff\xfe"):
        return "utf-16-le"
    if head.startswith(b"\xfe\xff"):
        return "utf-16-be"
    for enc in ("utf-8", "cp1256", "latin-1"):
        try:
            with open(path, encoding=enc) as f:
                while f.read(1 << 20):     # اقرأ الملف كاملاً للتحقق
                    pass
            return enc
        except (UnicodeDecodeError, LookupError):
            continue
    return "cp1256"


def read_csv_source(path):
    base = os.path.splitext(os.path.basename(path))[0]
    enc = _detect_csv_encoding(path)
    # encoding_errors='replace' يمنع توقّف البناء إذا ظهر بايت تالف متأخر
    common = dict(dtype=str, encoding=enc, encoding_errors="replace",
                  keep_default_na=False, on_bad_lines="skip")
    header = pd.read_csv(path, nrows=0, **common)
    cols = [str(c) for c in header.columns]

    def rows():
        for chunk in pd.read_csv(path, chunksize=50_000, **common):
            yield from chunk.itertuples(index=False, name=None)
    yield base, cols, rows()


def read_excel_source(path):
    base = os.path.splitext(os.path.basename(path))[0]
    xls = pd.ExcelFile(path)
    multi = len(xls.sheet_names) > 1
    for sheet in xls.sheet_names:
        df = xls.parse(sheet, dtype=str)
        if df.empty:
            continue
        df = df.fillna("")
        name = f"{base} ▸ {sheet}" if multi else base
        cols = [str(c) for c in df.columns]
        yield name, cols, df.itertuples(index=False, name=None)


def read_folder_sources(folder, skipped):
    """كل ملفات Excel/CSV في المجلد؛ الملفات التالفة تُسجَّل وتُتخطّى."""
    for fp in list_data_files(folder):
        ext = os.path.splitext(fp)[1].lower()
        try:
            reader = read_csv_source if ext in CSV_EXTS else read_excel_source
            yield from reader(fp)
        except Exception as e:
            skipped.append(f"{os.path.basename(fp)}: {e}")


def read_sqlite_sources(con):
    names = [r[0] for r in con.execute(
        "SELECT name FROM sqlite_master WHERE type='table' "
        "AND name NOT LIKE 'sqlite\\_%' ESCAPE '\\' ORDER BY name")]
    for nm in names:
        q = '"' + nm.replace('"', '""') + '"'      # تهريب الاسم لمنع كسر الاستعلام
        cols = [r[1] for r in con.execute(f"PRAGMA table_info({q})")]
        yield nm, cols, con.execute(f"SELECT * FROM {q}")


def read_access_sources(con):
    cur = con.cursor()
    tables = sorted(r.table_name for r in cur.tables(tableType="TABLE")
                    if not r.table_name.startswith("MSys"))
    for t in tables:
        c = con.cursor()
        c.execute(f"SELECT * FROM [{t.replace(']', ']]')}]")   # تهريب ]
        cols = [d[0] for d in c.description]
        yield t, cols, c


def open_access(path):
    try:
        import pyodbc
    except ImportError:
        raise RuntimeError("دعم Access يتطلب مكتبة pyodbc.\n"
                           "نفّذ في موجّه الأوامر:  pip install pyodbc")
    drivers = [d for d in pyodbc.drivers() if "Microsoft Access Driver" in d]
    if not drivers:
        raise RuntimeError(
            "مُشغّل Access غير مثبّت على النظام.\n"
            "ثبّت «Microsoft Access Database Engine 2016 Redistributable» "
            "مجاناً من موقع مايكروسوفت ثم أعد المحاولة.")
    return pyodbc.connect(f"DRIVER={{{drivers[0]}}};DBQ={path};", autocommit=True)


# ════════════════════════════════════════════════════════════════
#  الواجهة
# ════════════════════════════════════════════════════════════════
class SearchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x760")
        self.configure(bg="#0f172a")

        self.engine = None
        self._gen = 0                    # لإسقاط نتائج البحث القديمة
        self._last_query = ("", "")      # (source, q) — لآخر بحث مكتمل
        self._queue = queue.Queue()
        self._search_job = None
        self._busy = False               # فهرسة/تصدير جارٍ
        self._indexing = False           # بناء فهرس جارٍ (يمنع فتح ملف الفهرس)

        self._build_style()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(60, self._poll)
        self.after(200, self._autoload)

    def _on_close(self):
        if self._busy and not messagebox.askokcancel(
                "عملية جارية",
                "توجد فهرسة/تصدير قيد التنفيذ.\nالإغلاق الآن قد يُنتج ملفاً ناقصاً. إغلاق؟"):
            return
        self.destroy()

    # ── جسر آمن بين خيوط العمل والواجهة ──────────────────────────
    def _poll(self):
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "results":
                    gen, name, q, cols, rows, total, capped = payload
                    if gen == self._gen:
                        self._last_query = (name, q)
                        self._render(cols, rows, total, q, name, capped)
                elif kind == "status":
                    self.status.config(text=payload)
                elif kind == "indexed":
                    self._busy = self._indexing = False
                    self._on_indexed(*payload)
                elif kind == "done":
                    self._busy = False
                elif kind == "info":
                    messagebox.showinfo(*payload)
                elif kind == "warn":
                    self._busy = self._indexing = False
                    messagebox.showwarning(*payload)
                elif kind == "error":
                    self._busy = self._indexing = False
                    messagebox.showerror(*payload)
        except queue.Empty:
            pass
        self.after(60, self._poll)

    # ── الشكل ────────────────────────────────────────────────────
    def _build_style(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure("Treeview", background="#1e293b", fieldbackground="#1e293b",
                    foreground="#e2e8f0", rowheight=26, borderwidth=0)
        s.configure("Treeview.Heading", background="#162033", foreground="#c7d2fe",
                    relief="flat", font=("Segoe UI", 10, "bold"))
        s.map("Treeview.Heading", background=[("active", "#243049")])
        s.map("Treeview", background=[("selected", "#3730a3")])

    def _build_ui(self):
        top = tk.Frame(self, bg="#1e293b", height=64)
        top.pack(fill="x", side="top")
        tk.Label(top, text="🔎 " + APP_TITLE, bg="#1e293b", fg="#e2e8f0",
                 font=("Segoe UI", 15, "bold")).pack(side="right", padx=18, pady=14)
        tk.Button(top, text="📂 فتح ملف (DB / Excel / CSV / Access)",
                  command=self.open_file, bg="#6366f1", fg="white", relief="flat",
                  padx=14, pady=8, font=("Segoe UI", 11),
                  cursor="hand2").pack(side="left", padx=(14, 6))
        tk.Button(top, text="📁 فتح مجلد (Excel / CSV)", command=self.open_folder,
                  bg="#0ea5e9", fg="white", relief="flat", padx=14, pady=8,
                  font=("Segoe UI", 11), cursor="hand2").pack(side="left", padx=6)
        tk.Label(top, text=CREDIT, bg="#1e293b", fg="#ffffff",
                 font=("Segoe UI", 13, "bold")).place(relx=0.56, rely=0.5, anchor="center")

        bar = tk.Frame(self, bg="#0f172a")
        bar.pack(fill="x", padx=16, pady=(14, 6))
        tk.Label(bar, text="المصدر:", bg="#0f172a", fg="#94a3b8").pack(side="right")
        self.table_var = tk.StringVar()
        self.table_cb = ttk.Combobox(bar, textvariable=self.table_var,
                                     state="readonly", width=26)
        self.table_cb.pack(side="right", padx=8)
        self.table_cb.bind("<<ComboboxSelected>>", lambda e: self.run_search())

        self.q_var = tk.StringVar()
        self.entry = tk.Entry(bar, textvariable=self.q_var, bg="#1e293b",
                              fg="#e2e8f0", insertbackground="#e2e8f0",
                              relief="flat", font=("Segoe UI", 13))
        self.entry.pack(side="right", fill="x", expand=True, ipady=7, padx=8)
        self._placeholder(self.entry,
                          "اكتب كلمة للبحث في كل الأعمدة… (3 أحرف فأكثر للبحث الفوري)")
        self.q_var.trace_add("write", lambda *a: self._debounced_search())
        self.entry.bind("<Return>", lambda e: self.run_search(force=True))

        tk.Button(bar, text="⬇ تصدير النتائج", command=self.export_results,
                  bg="#16a34a", fg="white", relief="flat", padx=12, pady=6,
                  cursor="hand2").pack(side="left", padx=(8, 0))

        self.status = tk.Label(self, text="افتح ملفاً أو مجلداً للبدء…",
                               bg="#0f172a", fg="#94a3b8", anchor="e")
        self.status.pack(fill="x", padx=18)

        wrap = tk.Frame(self, bg="#0f172a")
        wrap.pack(fill="both", expand=True, padx=16, pady=10)
        self.tree = ttk.Treeview(wrap, show="headings")
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        # علامة مائية خافتة (تظهر عند عدم وجود نتائج)
        self.watermark = tk.Label(self.tree, text=WATERMARK, fg="#22304a",
                                  bg="#1e293b", font=("Segoe UI", 34, "bold"))
        self.watermark.place(relx=0.5, rely=0.5, anchor="center")

    def _placeholder(self, entry, text):
        self._ph_active = True
        entry.config(fg="#64748b")
        entry.insert(0, text)

        def fin(_):
            if self._ph_active:
                entry.delete(0, "end")
                entry.config(fg="#e2e8f0")
                self._ph_active = False

        def fout(_):
            if not entry.get():
                entry.config(fg="#64748b")
                entry.insert(0, text)
                self._ph_active = True

        entry.bind("<FocusIn>", fin)
        entry.bind("<FocusOut>", fout)

    def _query_text(self):
        return "" if self._ph_active else self.q_var.get().strip()

    # ── التحميل التلقائي عند بدء التشغيل ─────────────────────────
    def _autoload(self):
        base = base_dir()
        candidates = [(name, os.path.join(base, name)) for name in DATA_DIRS]
        candidates.append(("مجلد البرنامج", base))
        for label, folder in candidates:
            if not os.path.isdir(folder):
                continue
            try:
                entries = os.listdir(folder)
            except OSError:
                continue
            if any(f.lower().endswith(FOLDER_EXTS) and not f.startswith("~$")
                   for f in entries):
                self._start_ingest(self._ingest_folder, folder,
                                   f"تحميل البيانات الثابتة من «{label}»…")
                return
            singles = sorted(f for f in entries
                             if f.lower().endswith(SQLITE_EXTS + ACCESS_EXTS))
            if singles:
                self._start_ingest(self._ingest_file,
                                   os.path.join(folder, singles[0]),
                                   f"تحميل البيانات الثابتة من «{label}»…")
                return
        last = self._load_config().get("last")
        if last and os.path.exists(last):
            fn = self._ingest_folder if os.path.isdir(last) else self._ingest_file
            self._start_ingest(fn, last, "فتح آخر مصدر مستخدم…")

    def _config_path(self):
        d = base_dir()
        return os.path.join(d if _dir_writable(d) else appdata_dir(), CONFIG_NAME)

    def _load_config(self):
        try:
            with open(self._config_path(), encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_config(self, path):
        try:
            with open(self._config_path(), "w", encoding="utf-8") as f:
                json.dump({"last": path}, f, ensure_ascii=False)
        except OSError:
            pass

    # ── فتح المصادر ──────────────────────────────────────────────
    def open_file(self):
        path = filedialog.askopenfilename(
            title="اختر ملف بيانات",
            filetypes=[("كل المدعوم",
                        "*.db *.sqlite *.sqlite3 *.xlsx *.xls *.xlsm *.csv *.mdb *.accdb"),
                       ("SQLite", "*.db *.sqlite *.sqlite3"),
                       ("Excel", "*.xlsx *.xls *.xlsm"),
                       ("CSV", "*.csv"),
                       ("Access", "*.mdb *.accdb"),
                       ("كل الملفات", "*.*")])
        if path:
            self._start_ingest(self._ingest_file, path, "جارٍ الفتح…")

    def open_folder(self):
        folder = filedialog.askdirectory(title="اختر مجلداً يحتوي ملفات Excel أو CSV")
        if folder:
            self._start_ingest(self._ingest_folder, folder, "جارٍ الفتح…")

    def _start_ingest(self, fn, path, msg):
        if self._busy:
            messagebox.showinfo("انتظر", "عملية أخرى قيد التنفيذ — انتظر انتهاءها.")
            return
        self._busy = self._indexing = True   # يمنع البحث من فتح الفهرس أثناء استبداله
        self._gen += 1                        # يُسقط أي بحث سابق قيد التنفيذ
        self.status.config(text=msg)
        threading.Thread(target=fn, args=(path,), daemon=True).start()

    def _progress(self, src, rid):
        self._queue.put(("status", f"جارٍ الفهرسة… {rid:,} صف  ({src})"))

    def _ingest_file(self, path):
        try:
            idx = index_path_for(path)
            fp = fingerprint(path)
            skipped = []
            if Engine.stored_fingerprint(idx) != fp:
                self._queue.put(("status", "جارٍ بناء الفهرس…"))
                ext = os.path.splitext(path)[1].lower()
                if ext in SQLITE_EXTS:
                    con = ro_connect(path)
                    try:
                        Engine.build(idx, read_sqlite_sources(con), fp, self._progress)
                    finally:
                        con.close()
                elif ext in ACCESS_EXTS:
                    con = open_access(path)
                    try:
                        Engine.build(idx, read_access_sources(con), fp, self._progress)
                    finally:
                        con.close()
                elif ext in CSV_EXTS:
                    Engine.build(idx, read_csv_source(path), fp, self._progress)
                elif ext in EXCEL_EXTS:
                    Engine.build(idx, read_excel_source(path), fp, self._progress)
                else:
                    raise RuntimeError(f"صيغة غير مدعومة: {ext or path}")
            self._queue.put(("indexed", (idx, path, skipped)))
        except Exception as e:
            self._queue.put(("error", ("خطأ في الفتح", str(e))))

    def _ingest_folder(self, folder):
        try:
            if not list_data_files(folder):
                self._queue.put(("warn",
                                 ("تنبيه", "لم يُعثر على أي ملف Excel أو CSV في هذا المجلد.")))
                return
            idx = index_path_for(folder)
            fp = fingerprint(folder)
            skipped = []
            if Engine.stored_fingerprint(idx) != fp:
                self._queue.put(("status", "جارٍ بناء الفهرس…"))
                Engine.build(idx, read_folder_sources(folder, skipped), fp,
                             self._progress, source_errors=skipped)
            self._queue.put(("indexed", (idx, folder, skipped)))
        except Exception as e:
            self._queue.put(("error", ("خطأ في فتح المجلد", str(e))))

    def _on_indexed(self, idx, src_path, skipped):
        engine = Engine(idx)
        engine.load_meta()
        if not engine.sources:
            # الفهرس الجديد فارغ — لا نُبقي محرّكاً قديماً يشير لملف استُبدل
            self.engine = None
            self.table_cb["values"] = []
            self.table_var.set("")
            self.tree.delete(*self.tree.get_children())
            self.tree["columns"] = ()
            self.watermark.place(relx=0.5, rely=0.5, anchor="center")
            self.status.config(text="لا توجد بيانات صالحة في هذا المصدر.")
            messagebox.showwarning("تنبيه", "لا توجد بيانات صالحة في هذا المصدر.")
            return
        self.engine = engine
        self._save_config(src_path)
        names = list(engine.sources.keys())
        values = ([ALL_KEY] + names) if len(names) > 1 else names
        self.table_cb["values"] = values
        self.table_cb.current(0)
        total = sum(s["nrows"] for s in engine.sources.values())
        self.title(f"{APP_TITLE} — {os.path.basename(src_path) or src_path}")
        msg = f"جاهز: {len(names)} مصدر، {total:,} صف مفهرس"
        if skipped:
            msg += f"  (تُخطّي {len(skipped)} ملف تالف)"
        self.status.config(text=msg)
        self.run_search()

    # ── البحث ────────────────────────────────────────────────────
    def _debounced_search(self):
        if self._ph_active:
            return
        if self._search_job:
            self.after_cancel(self._search_job)
        self._search_job = self.after(160, self.run_search)

    def run_search(self, force=False):
        # إلغاء أي بحث مؤجَّل معلّق (Enter لا يُنتج بحثاً مكرراً)
        if self._search_job:
            self.after_cancel(self._search_job)
            self._search_job = None
        self._gen += 1                    # أي بحث سابق قيد التنفيذ يُصبح ملغى
        if not self.engine or self._indexing:
            return
        q = self._query_text()
        if q and len(q) < MIN_FAST_QUERY and not force:
            # الاستعلام القصير يتطلب مسحاً كاملاً — لا يعمل تلقائياً مع كل حرف
            self.status.config(
                text=f"أكمل إلى {MIN_FAST_QUERY} أحرف للبحث الفوري — أو اضغط Enter للبحث الدقيق")
            return
        name = self.table_var.get()
        self.status.config(text="جارٍ البحث…")
        threading.Thread(target=self._do_search,
                         args=(self._gen, name, q), daemon=True).start()

    @staticmethod
    def _columns_for(engine, name):
        if name != ALL_KEY:
            return list(engine.sources.get(name, {}).get("columns", []))
        # اتحاد أعمدة كل المصادر، مع عمود مصدر فريد لا يتصادم مع أي عمود بيانات
        data_cols = []
        seen = set()
        for meta in engine.sources.values():
            for c in meta["columns"]:
                if c not in seen:
                    seen.add(c)
                    data_cols.append(c)
        header = SOURCE_COL
        while header in seen:
            header += " "
        return [header] + data_cols

    @staticmethod
    def _rows_to_table(engine, name, cols, rows):
        """يحوّل (source, vals) إلى صفوف قوائم بمحاذاة الأعمدة — بلا pandas."""
        out = []
        if name != ALL_KEY:
            ncols = len(cols)
            for _src, vals in rows:
                v = vals.split(SEP)
                v = v + [""] * (ncols - len(v)) if len(v) < ncols else v[:ncols]
                out.append(v)
            return out
        pos = {c: i for i, c in enumerate(cols)}
        src_maps = {
            src: [pos.get(c) for c in meta["columns"]]
            for src, meta in engine.sources.items()
        }
        for src, vals in rows:
            row = [""] * len(cols)
            row[0] = src
            mapping = src_maps.get(src)
            if mapping:
                for i, v in enumerate(vals.split(SEP)):
                    if i < len(mapping) and mapping[i] is not None:
                        row[mapping[i]] = v
            out.append(row)
        return out

    def _do_search(self, gen, name, q):
        engine = self.engine          # لقطة ثابتة تفادياً لاستبداله أثناء البحث
        if engine is None:
            return
        try:
            total, capped, rows = engine.search(name, q, PAGE)
            cols = self._columns_for(engine, name)
            table = self._rows_to_table(engine, name, cols, rows)
            if gen == self._gen:
                self._queue.put(("results", (gen, name, q, cols, table, total, capped)))
        except Exception as e:
            if gen == self._gen:
                self._queue.put(("status", "تعذّر البحث: " + str(e)))

    def _render(self, cols, rows, total, q, name, capped=False):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=160, minwidth=80, anchor="e", stretch=False)
        insert = self.tree.insert
        for row in rows:
            insert("", "end", values=[_clip(v) for v in row])
        if rows:
            self.watermark.place_forget()
        else:
            self.watermark.place(relx=0.5, rely=0.5, anchor="center")
        count_txt = f"أكثر من {total:,}" if capped else f"{total:,}"
        extra = f" (معروض أول {PAGE:,})" if total > PAGE else ""
        qtxt = f' للبحث عن "{q}"' if q else ""
        self.status.config(text=f"المصدر «{name}» — {count_txt} نتيجة{extra}{qtxt}")

    # ── التصدير ──────────────────────────────────────────────────
    def export_results(self):
        if not self.engine:
            return
        if self._busy:
            messagebox.showinfo("انتظر", "عملية أخرى قيد التنفيذ — انتظر انتهاءها.")
            return
        # مجلد افتراضي بعيد عن مجلد البيانات حتى لا تُفهرَس النتائج المصدَّرة لاحقاً
        home = os.path.expanduser("~")
        desktop = os.path.join(home, "Desktop")
        path = filedialog.asksaveasfilename(
            title="حفظ النتائج", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV (أسرع للنتائج الكبيرة)", "*.csv")],
            initialdir=desktop if os.path.isdir(desktop) else home,
            initialfile="نتائج_البحث.xlsx")
        if not path:
            return
        name, q = self.table_var.get(), self._query_text()
        if q and len(q) < MIN_FAST_QUERY:
            name, q = self._last_query   # صدّر آخر بحث مكتمل فعلاً
        self._busy = True
        self.status.config(text="جارٍ تجهيز ملف التصدير…")
        threading.Thread(target=self._do_export,
                         args=(path, name, q), daemon=True).start()

    def _do_export(self, path, name, q):
        engine = self.engine
        if engine is None:
            self._queue.put(("done", None))
            return
        try:
            cols = self._columns_for(engine, name)
            rows_iter = engine.iter_all(name, q, EXPORT_CAP)
            n = 0
            if path.lower().endswith(".csv"):
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f)
                    w.writerow(cols)
                    batch = []
                    for item in rows_iter:
                        batch.append(item)
                        if len(batch) >= 50_000:
                            w.writerows(self._rows_to_table(engine, name, cols, batch))
                            n += len(batch)
                            batch.clear()
                            self._queue.put(("status", f"جارٍ التصدير… {n:,} صف"))
                    w.writerows(self._rows_to_table(engine, name, cols, batch))
                    n += len(batch)
            else:
                rows = list(rows_iter)
                n = len(rows)
                table = self._rows_to_table(engine, name, cols, rows)
                pd.DataFrame(table, columns=cols).to_excel(path, index=False)
            if n == 0:
                os.remove(path)
                self._queue.put(("info", ("تصدير", "لا توجد نتائج للتصدير.")))
                self._queue.put(("done", None))
                return
            extra = (f"\n(بلغ الحد الأقصى {EXPORT_CAP:,} صف — ضيّق البحث لتصدير أقل)"
                     if n >= EXPORT_CAP else "")
            self._queue.put(("info", ("تم", f"تم حفظ {n:,} صف في:\n{path}{extra}")))
            self._queue.put(("status", f"تم تصدير {n:,} صف"))
            self._queue.put(("done", None))
        except Exception as e:
            self._queue.put(("error", ("خطأ في التصدير", str(e))))


def _clip(v, n=120):
    s = "" if v is None else str(v)
    if "\n" in s or "\r" in s:
        s = s.replace("\n", " ").replace("\r", " ")
    return s if len(s) <= n else s[:n] + "…"


def main():
    try:
        con = sqlite3.connect(":memory:")
        con.execute("CREATE VIRTUAL TABLE _t USING fts5(x, tokenize='trigram')")
        con.close()
    except sqlite3.OperationalError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "غير مدعوم",
            "نسخة SQLite في Python لديك لا تدعم FTS5/trigram.\n"
            "ثبّت Python 3.11 أو أحدث من python.org.")
        raise SystemExit(1)
    SearchApp().mainloop()


if __name__ == "__main__":
    main()
