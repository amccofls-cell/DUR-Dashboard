from __future__ import annotations

import json
import os
import re
import shutil
import sqlite3
import threading
import traceback
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import load_workbook

APP_NAME = "DUR 약물안전성 조회"
VERSION = "1.4.1"

CATEGORIES = [
    ("combo_paid", "병용금기 급여"),
    ("combo_unpaid", "병용금기 비급여"),
    ("age", "연령금기"),
    ("preg", "임부금기"),
    ("lact", "수유부주의"),
    ("dup", "효능군중복"),
    ("tele", "비대면진료 처방금지"),
    ("cost", "비용효과적인 함량"),
]

DUR_CARDS = [
    ("combo", "병용금기"),
    ("age", "연령금기"),
    ("preg", "임부금기"),
    ("lact", "수유부주의"),
    ("dup", "효능군중복"),
    ("tele", "비대면진료 금지"),
    ("cost", "비용효과적 함량"),
]

# Mockup-based compact enterprise palette
APP_BG = "#F4F5F2"
SURFACE = "#FFFFFF"
SURFACE_ALT = "#FAFBF8"
TEXT = "#183042"
TEXT_DARK = "#20313D"
MUTED = "#7D8991"
BORDER = "#DDE3DF"
BORDER_DARK = "#C9D1CC"
SAGE = "#5E9872"
SAGE_DEEP = "#3F7B59"
SAGE_SOFT = "#EEF8F1"
SAGE_ROW = "#EDF6F0"
TERRA = "#C95737"
TERRA_DEEP = "#A93B24"
TERRA_SOFT = "#FFF4F0"
GRAY_SOFT = "#F7F8F7"
GRAY_ICON = "#AAB1AE"
ORANGE = "#E88A2D"
ORANGE_SOFT = "#FFF6EA"
ERROR = "#B34E46"
ERROR_SOFT = "#FFF1F0"


def app_dir() -> Path:
    root = Path(os.getenv("LOCALAPPDATA") or Path.home()) / "DUR_Dashboard"
    root.mkdir(parents=True, exist_ok=True)
    (root / "files").mkdir(exist_ok=True)
    return root


ROOT_DIR = app_dir()
CFG_PATH = ROOT_DIR / "config.json"
DB_PATH = ROOT_DIR / "dur_index.sqlite3"


def norm(v):
    if v is None:
        return ""
    s = str(v).strip().lower()
    s = re.sub(r"\s+", "", s)
    s = s.replace("㎖", "ml").replace("밀리리터", "ml").replace("밀리그램", "mg")
    s = re.sub(r"[()\[\]{}·,._\-/]", "", s)
    return s


def product_core(v):
    """Cross-list matching key: keep product name/strength, ignore ingredient/export/package parentheses."""
    if v is None:
        return ""
    s = str(v).strip()
    s = re.sub(r"_\s*\([^)]*\)\s*$", "", s)
    s = re.sub(r"\([^)]*\)", "", s)
    return norm(s)


def split_top_level_aliases(v):
    """Split comma/semicolon separated aliases, but keep punctuation inside parentheses."""
    text = "" if v is None else str(v).strip()
    if not text:
        return []
    out, buf, depth = [], [], 0
    for ch in text:
        if ch in "([{":
            depth += 1
        elif ch in ")]}":
            depth = max(0, depth - 1)
        if ch in ",;" and depth == 0:
            part = "".join(buf).strip()
            if part:
                out.append(part)
            buf = []
        else:
            buf.append(ch)
    part = "".join(buf).strip()
    if part:
        out.append(part)
    return out or [text]


def sval(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def unique_headers(values):
    out, seen = [], {}
    for i, v in enumerate(values):
        h = sval(v) or f"col_{i+1}"
        seen[h] = seen.get(h, 0) + 1
        if seen[h] > 1:
            h = f"{h}_{seen[h]}"
        out.append(h)
    return out


def find_header_row(ws, max_scan=15):
    keys = ("제품명", "품목명", "약품명", "성분명", "성분명a", "저함량")
    best = (1, -1)
    for r in range(1, min(ws.max_row, max_scan) + 1):
        vals = [norm(c.value) for c in ws[r][: min(ws.max_column, 40)]]
        score = sum(any(k in x for k in keys) for x in vals)
        if score > best[1]:
            best = (r, score)
    return best[0]


def detect_name_cols(headers, category):
    hn = [norm(h) for h in headers]
    if category.startswith("combo"):
        a = [i for i, h in enumerate(hn) if h in ("제품명a", "품목명a") or ("제품명" in h and h.endswith("a"))]
        b = [i for i, h in enumerate(hn) if h in ("제품명b", "품목명b") or ("제품명" in h and h.endswith("b"))]
        return (a[:1], b[:1])
    cand = []
    for i, h in enumerate(hn):
        if h in ("제품명", "품목명", "약품명") or h.startswith("제품명") or h.startswith("품목명") or h.startswith("약품명"):
            cand.append(i)
    return (cand[:2], [])


def cost_rows(ws):
    rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 8), values_only=True))
    header_row = 4 if len(rows) >= 4 else 1
    for idx, row in enumerate(rows, 1):
        vals = [norm(x) for x in row]
        if vals.count("제품코드") >= 2 or ("제품코드" in vals and "제품명" in vals and "함량" in vals):
            header_row = idx
    group_row = max(1, header_row - 1)
    groups = list(rows[group_row - 1]) if len(rows) >= group_row else []
    base = list(rows[header_row - 1])
    headers, current = [], ""
    for i, h in enumerate(base):
        g = sval(groups[i]) if i < len(groups) else ""
        if g:
            current = g
        prefix = "저함량" if "저함량" in current else ("고함량" if "고함량" in current else "")
        name = sval(h) or f"col_{i+1}"
        headers.append((prefix + " " + name).strip())
    headers = unique_headers(headers)
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(v is not None and sval(v) != "" for v in row):
            continue
        yield headers, row


# Only these columns are indexed from the very large combination-prohibition workbooks.
COMBO_KEEP = {
    1: "성분명A",   # A
    4: "제품명A",   # D
    6: "급여여부A", # F
    7: "성분명B",   # G
    10: "제품명B",  # J
    13: "고시번호", # M
    14: "고시일자", # N
    15: "상세정보", # O
}


def combo_rows(ws):
    hr = find_header_row(ws)
    for row in ws.iter_rows(min_row=hr + 1, min_col=1, max_col=15, values_only=True):
        if not any(row[i - 1] is not None and sval(row[i - 1]) != "" for i in COMBO_KEEP):
            continue
        yield {name: sval(row[i - 1]) for i, name in COMBO_KEEP.items()}


def normal_rows(ws):
    hr = find_header_row(ws)
    headers = unique_headers([c.value for c in ws[hr]])
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        if not any(v is not None and sval(v) != "" for v in row):
            continue
        yield headers, row


def record_product_code(d, category, side=""):
    """Internal-only product code for reliable cross-DUR matching; never displayed."""
    if category.startswith("combo"):
        return ""
    if category == "cost":
        prefix = "저함량" if side == "저함량" else "고함량"
        for k, v in d.items():
            if prefix in k and "제품코드" in k and sval(v):
                return sval(v)
        return ""
    for key in ("제품코드", "약품코드"):
        if sval(d.get(key)):
            return sval(d.get(key))
    return ""


class Indexer:
    def __init__(self, db_path=DB_PATH):
        self.db_path = db_path

    def conn(self):
        c = sqlite3.connect(self.db_path)
        c.execute("PRAGMA journal_mode=WAL")
        return c

    def init(self):
        with self.conn() as c:
            c.execute(
                "CREATE TABLE IF NOT EXISTS records("
                "id INTEGER PRIMARY KEY, category TEXT, sheet TEXT, product TEXT, "
                "product_norm TEXT, core_norm TEXT, product_code TEXT, side TEXT, data TEXT)"
            )
            cols = [r[1] for r in c.execute("PRAGMA table_info(records)").fetchall()]
            if "core_norm" not in cols:
                c.execute("ALTER TABLE records ADD COLUMN core_norm TEXT DEFAULT ''")
            if "product_code" not in cols:
                c.execute("ALTER TABLE records ADD COLUMN product_code TEXT DEFAULT ''")
            c.execute("CREATE INDEX IF NOT EXISTS ix_records_cat_name ON records(category,product_norm)")
            c.execute("CREATE INDEX IF NOT EXISTS ix_records_cat_core ON records(category,core_norm)")
            c.execute("CREATE INDEX IF NOT EXISTS ix_records_cat_code ON records(category,product_code)")
            c.execute(
                "CREATE TABLE IF NOT EXISTS status("
                "category TEXT PRIMARY KEY, ok INTEGER, message TEXT, filename TEXT, "
                "indexed_at TEXT, sheet_count INTEGER, row_count INTEGER)"
            )

    def clear_category(self, c, cat):
        c.execute("DELETE FROM records WHERE category=?", (cat,))
        c.execute("DELETE FROM status WHERE category=?", (cat,))

    def index_file(self, category, path, progress=None):
        self.init()
        path = Path(path)
        total = 0
        sheets = 0
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            with self.conn() as c:
                self.clear_category(c, category)
                for ws in wb.worksheets:
                    sheets += 1
                    if progress:
                        progress(f"{ws.title} 시트 읽는 중…")
                    iterator = combo_rows(ws) if category.startswith("combo") else (cost_rows(ws) if category == "cost" else normal_rows(ws))
                    batch = []
                    for item in iterator:
                        if category.startswith("combo"):
                            d = item
                            for side, product_name in (("A", d.get("제품명A", "")), ("B", d.get("제품명B", ""))):
                                if product_name:
                                    batch.append((
                                        category, ws.title, product_name, norm(product_name),
                                        product_core(product_name), "", side, json.dumps(d, ensure_ascii=False)
                                    ))
                            if len(batch) >= 2500:
                                c.executemany(
                                    "INSERT INTO records(category,sheet,product,product_norm,core_norm,product_code,side,data) VALUES(?,?,?,?,?,?,?,?)",
                                    batch,
                                )
                                total += len(batch)
                                batch = []
                            continue

                        headers, row = item
                        d = {headers[i]: sval(row[i]) if i < len(row) else "" for i in range(len(headers))}
                        if category == "cost":
                            low = next((d[k] for k in d if "저함량" in k and "제품명" in k), "")
                            high = next((d[k] for k in d if "고함량" in k and "제품명" in k), "")
                            for side, product_name in (("저함량", low), ("고함량", high)):
                                if product_name:
                                    batch.append((
                                        category, ws.title, product_name, norm(product_name),
                                        product_core(product_name), record_product_code(d, category, side), side, json.dumps(d, ensure_ascii=False)
                                    ))
                        else:
                            a, _ = detect_name_cols(headers, category)
                            for i in a:
                                raw_name = sval(row[i]) if i < len(row) else ""
                                if not raw_name:
                                    continue
                                names = split_top_level_aliases(raw_name) if category == "tele" else [raw_name]
                                for product_name in names:
                                    batch.append((
                                        category, ws.title, product_name, norm(product_name),
                                        product_core(product_name), record_product_code(d, category, ""), "", json.dumps(d, ensure_ascii=False)
                                    ))
                        if len(batch) >= 2500:
                            c.executemany(
                                "INSERT INTO records(category,sheet,product,product_norm,core_norm,product_code,side,data) VALUES(?,?,?,?,?,?,?,?)",
                                batch,
                            )
                            total += len(batch)
                            batch = []
                    if batch:
                        c.executemany(
                            "INSERT INTO records(category,sheet,product,product_norm,core_norm,product_code,side,data) VALUES(?,?,?,?,?,?,?,?)",
                            batch,
                        )
                        total += len(batch)
                c.execute(
                    "INSERT OR REPLACE INTO status VALUES(?,?,?,?,?,?,?)",
                    (category, 1, "정상", path.name, datetime.now().isoformat(timespec="seconds"), sheets, total),
                )
            wb.close()
            return True, f"{sheets}개 시트 · {total:,}개 검색 엔트리"
        except Exception as e:
            with self.conn() as c:
                self.clear_category(c, category)
                c.execute(
                    "INSERT OR REPLACE INTO status VALUES(?,?,?,?,?,?,?)",
                    (category, 0, str(e), path.name, datetime.now().isoformat(timespec="seconds"), sheets, total),
                )
            return False, str(e)

    def statuses(self):
        self.init()
        with self.conn() as c:
            return {r[0]: r[1:] for r in c.execute(
                "SELECT category,ok,message,filename,indexed_at,sheet_count,row_count FROM status"
            )}

    def suggestions(self, q, limit=60):
        nq = norm(q)
        if not nq:
            return []
        with self.conn() as c:
            rows = c.execute(
                "SELECT product, COUNT(*) n FROM records WHERE product_norm LIKE ? "
                "GROUP BY product ORDER BY CASE WHEN product_norm LIKE ? THEN 0 ELSE 1 END, length(product), product LIMIT ?",
                (f"%{nq}%", f"{nq}%", limit),
            ).fetchall()
        return [x[0] for x in rows]

    def search_product(self, product):
        """Match exact name, core name, internal product code, then safe age/tele fallback."""
        np = norm(product)
        cp = product_core(product)
        out = {k: [] for k, _ in CATEGORIES}
        raw = {}
        codes = set()
        with self.conn() as c:
            for cat, _ in CATEGORIES:
                rows = c.execute(
                    "SELECT sheet,product,product_code,side,data FROM records WHERE category=? AND product_norm=?",
                    (cat, np),
                ).fetchall()
                if not rows and cp and len(cp) >= 4:
                    rows = c.execute(
                        "SELECT sheet,product,product_code,side,data FROM records WHERE category=? AND core_norm=?",
                        (cat, cp),
                    ).fetchall()
                raw[cat] = rows
                for r in rows:
                    if r[2]:
                        codes.add(str(r[2]).strip())

            if codes:
                ph = ",".join("?" for _ in codes)
                for cat, _ in CATEGORIES:
                    if raw.get(cat) or cat.startswith("combo"):
                        continue
                    raw[cat] = c.execute(
                        f"SELECT sheet,product,product_code,side,data FROM records WHERE category=? AND product_code IN ({ph})",
                        (cat, *sorted(codes)),
                    ).fetchall()

            for cat in ("age", "tele"):
                if raw.get(cat) or not cp or len(cp) < 5:
                    continue
                raw[cat] = c.execute(
                    "SELECT sheet,product,product_code,side,data FROM records WHERE category=? AND (product_norm LIKE ? OR core_norm LIKE ?)",
                    (cat, f"%{cp}%", f"%{cp}%"),
                ).fetchall()

            for cat, _ in CATEGORIES:
                rows = raw.get(cat, [])
                out[cat] = [
                    {"sheet": r[0], "product": r[1], "side": r[3], "data": json.loads(r[4])}
                    for r in rows
                ]
        return out



class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME}  {VERSION}")
        self.geometry("1536x910")
        self.minsize(1180, 740)
        self.configure(bg=APP_BG)

        self.idx = Indexer()
        self.idx.init()
        self.cfg = self.load_cfg()
        self.results = {}
        self.selected_drug = None
        self.selected_dur = "combo"
        self.current_page = "lookup"
        self.result_mode = "drug"
        self.detail_hits = []

        self.setup_style()
        self.build_shell()
        self.refresh_status()
        self.show_lookup_page()

    # ---------- configuration / fonts ----------
    def load_cfg(self):
        try:
            return json.loads(CFG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def save_cfg(self):
        CFG_PATH.write_text(json.dumps(self.cfg, ensure_ascii=False, indent=2), encoding="utf-8")

    def font(self, size=10, weight="normal"):
        # Windows chooses NanumGothic for Korean if installed; otherwise Malgun Gothic fallback.
        return ("NanumGothic", size, weight)

    def font_en(self, size=10, weight="normal"):
        return ("Arial", size, weight)

    def setup_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=APP_BG)
        style.configure("Surface.TFrame", background=SURFACE)
        style.configure("TLabel", background=APP_BG, foreground=TEXT_DARK, font=self.font(10))
        style.configure("Surface.TLabel", background=SURFACE, foreground=TEXT_DARK, font=self.font(10))
        style.configure("TSeparator", background=BORDER)
        style.configure(
            "Compact.Treeview",
            font=self.font(9),
            rowheight=30,
            background=SURFACE,
            fieldbackground=SURFACE,
            foreground=TEXT_DARK,
            borderwidth=0,
            relief="flat",
        )
        style.map("Compact.Treeview", background=[("selected", "#E9F2EE")], foreground=[("selected", TEXT_DARK)])
        style.configure(
            "Compact.Treeview.Heading",
            font=self.font(9, "bold"),
            background="#F8FAF8",
            foreground="#425563",
            relief="flat",
            borderwidth=0,
            padding=(7, 7),
        )
        style.map("Compact.Treeview.Heading", background=[("active", "#F3F6F4")])

    # ---------- root shell ----------
    def build_shell(self):
        self.build_header()
        self.page_host = tk.Frame(self, bg=APP_BG)
        self.page_host.pack(fill="both", expand=True)

    def build_header(self):
        h = tk.Frame(self, bg=SURFACE, height=74, highlightbackground=BORDER, highlightthickness=1)
        h.pack(fill="x")
        h.pack_propagate(False)

        brand = tk.Frame(h, bg=SURFACE)
        brand.pack(side="left", padx=(28, 34), pady=11)
        icon = tk.Label(brand, text="❦", bg=SURFACE, fg=SAGE_DEEP, font=("Segoe UI Symbol", 28))
        icon.pack(side="left", padx=(0, 12))
        btxt = tk.Frame(brand, bg=SURFACE)
        btxt.pack(side="left")
        tk.Label(btxt, text="DUR 약물안전성 조회", bg=SURFACE, fg=TEXT, font=self.font(15, "bold")).pack(anchor="w")
        tk.Label(btxt, text="의약품 DUR 통합 분석 시스템", bg=SURFACE, fg="#61717C", font=self.font(9)).pack(anchor="w", pady=(1, 0))

        nav = tk.Frame(h, bg=SURFACE)
        nav.pack(side="left", fill="y")
        self.nav_buttons = {}
        for key, label, symbol in [
            ("home", "홈", "⌂"),
            ("files", "기준파일 관리", "▣"),
            ("lookup", "DUR 조회", "◆"),
            ("settings", "설정", "⚙"),
        ]:
            btn = tk.Frame(nav, bg=SURFACE, width=138, height=73, cursor="hand2")
            btn.pack(side="left", fill="y")
            btn.pack_propagate(False)
            inner = tk.Frame(btn, bg=SURFACE)
            inner.pack(expand=True)
            s = tk.Label(inner, text=symbol, bg=SURFACE, fg="#687783", font=("Segoe UI Symbol", 14))
            s.pack(side="left", padx=(0, 9))
            t = tk.Label(inner, text=label, bg=SURFACE, fg=TEXT_DARK, font=self.font(10))
            t.pack(side="left")
            line = tk.Frame(btn, bg=SURFACE, height=3)
            line.pack(side="bottom", fill="x")
            for w in (btn, inner, s, t):
                w.bind("<Button-1>", lambda e, k=key: self.navigate(k))
            self.nav_buttons[key] = (btn, s, t, line)

        right = tk.Frame(h, bg=SURFACE)
        right.pack(side="right", padx=28, pady=11)
        tk.Label(right, text="기준파일 최종 업데이트", bg=SURFACE, fg="#526D60", font=self.font(9)).pack(anchor="w")
        self.header_update = tk.Label(right, text="● —", bg=SURFACE, fg=SAGE_DEEP, font=self.font(10))
        self.header_update.pack(anchor="w", pady=(2, 0))

    def navigate(self, page):
        self.current_page = page
        self.paint_nav()
        for w in self.page_host.winfo_children():
            w.destroy()
        if page == "lookup":
            self.show_lookup_page()
        elif page == "files":
            self.show_files_page()
        elif page == "home":
            self.show_home_page()
        else:
            self.show_settings_page()

    def paint_nav(self):
        for key, (btn, s, t, line) in self.nav_buttons.items():
            active = key == self.current_page
            fg = TERRA if active else TEXT_DARK
            s.config(fg=fg)
            t.config(fg=fg, font=self.font(10, "bold" if active else "normal"))
            line.config(bg=TERRA if active else SURFACE)

    # ---------- home ----------
    def show_home_page(self):
        self.paint_nav()
        wrap = tk.Frame(self.page_host, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=24, pady=20)
        tk.Label(wrap, text="DUR Dashboard", bg=APP_BG, fg=TEXT, font=self.font(20, "bold")).pack(anchor="w")
        tk.Label(wrap, text="기준파일을 등록하고 여러 품목을 한 번에 조회할 수 있습니다.", bg=APP_BG, fg=MUTED, font=self.font(10)).pack(anchor="w", pady=(4, 18))

        row = tk.Frame(wrap, bg=APP_BG)
        row.pack(fill="x")
        cards = [
            ("기준파일 관리", "8개 DUR Excel 자동 분류 · 인덱싱", self.show_files_page),
            ("DUR 조회", "여러 품목을 동시에 입력하고 통합 판정", self.show_lookup_page),
        ]
        for title, sub, cmd in cards:
            c = tk.Frame(row, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1, width=420, height=130, cursor="hand2")
            c.pack(side="left", padx=(0, 14), fill="x", expand=True)
            c.pack_propagate(False)
            tk.Label(c, text=title, bg=SURFACE, fg=TEXT, font=self.font(14, "bold")).pack(anchor="w", padx=20, pady=(22, 6))
            tk.Label(c, text=sub, bg=SURFACE, fg=MUTED, font=self.font(9)).pack(anchor="w", padx=20)
            c.bind("<Button-1>", lambda e, f=cmd: f())

    def show_settings_page(self):
        self.paint_nav()
        wrap = tk.Frame(self.page_host, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=24, pady=20)
        tk.Label(wrap, text="설정", bg=APP_BG, fg=TEXT, font=self.font(18, "bold")).pack(anchor="w")
        tk.Label(wrap, text="DUR Dashboard v1.3 · 데이터는 PC 내부에서만 처리됩니다.", bg=APP_BG, fg=MUTED, font=self.font(10)).pack(anchor="w", pady=(6, 0))

    # ---------- files page ----------
    def show_files_page(self):
        self.current_page = "files"
        self.paint_nav()
        wrap = tk.Frame(self.page_host, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=22, pady=18)

        head = tk.Frame(wrap, bg=APP_BG)
        head.pack(fill="x", pady=(0, 12))
        tk.Label(head, text="기준파일 관리", bg=APP_BG, fg=TEXT, font=self.font(17, "bold")).pack(side="left")
        self.make_button(head, "＋ 8개 파일 일괄 등록", self.choose_files, primary=True).pack(side="right")
        self.make_button(head, "다시 인덱싱", self.reindex_all).pack(side="right", padx=(0, 8))

        panel = tk.Frame(wrap, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="both", expand=True)
        tk.Label(panel, text="파일 종류", bg="#F8FAF8", fg="#52626D", font=self.font(9, "bold"), width=26, anchor="w").grid(row=0, column=0, sticky="nsew", padx=(18, 0), pady=10)
        tk.Label(panel, text="상태", bg="#F8FAF8", fg="#52626D", font=self.font(9, "bold"), width=14, anchor="w").grid(row=0, column=1, sticky="nsew", pady=10)
        tk.Label(panel, text="마지막 등록", bg="#F8FAF8", fg="#52626D", font=self.font(9, "bold"), width=22, anchor="w").grid(row=0, column=2, sticky="nsew", pady=10)
        tk.Label(panel, text="등록 파일", bg="#F8FAF8", fg="#52626D", font=self.font(9, "bold"), anchor="w").grid(row=0, column=3, sticky="nsew", padx=(0, 18), pady=10)
        panel.columnconfigure(3, weight=1)
        self.file_rows = {}
        for i, (key, label) in enumerate(CATEGORIES, start=1):
            bg = SURFACE if i % 2 else "#FBFCFB"
            a = tk.Label(panel, text=label, bg=bg, fg=TEXT_DARK, font=self.font(10), anchor="w")
            b = tk.Label(panel, text="미등록", bg=bg, fg=MUTED, font=self.font(9), anchor="w")
            c = tk.Label(panel, text="—", bg=bg, fg=MUTED, font=self.font(9), anchor="w")
            d = tk.Label(panel, text="—", bg=bg, fg=MUTED, font=self.font(9), anchor="w")
            for col, w in enumerate((a, b, c, d)):
                w.grid(row=i, column=col, sticky="nsew", padx=(18 if col == 0 else 0, 18 if col == 3 else 0), pady=10)
            self.file_rows[key] = (b, c, d)
        self.refresh_status()

    # ---------- lookup page ----------
    def show_lookup_page(self):
        self.current_page = "lookup"
        self.paint_nav()
        wrap = tk.Frame(self.page_host, bg=APP_BG)
        wrap.pack(fill="both", expand=True)

        # Top compact result tabs, modeled after mockup
        tabs = tk.Frame(wrap, bg=APP_BG)
        tabs.pack(fill="x", padx=16, pady=(12, 0))
        self.mode_buttons = {}
        for mode, label in (("drug", "약품별 결과"), ("dur", "DUR 종류별 결과")):
            box = tk.Frame(tabs, bg=SURFACE, width=170, height=44, cursor="hand2", highlightbackground=BORDER, highlightthickness=1)
            box.pack(side="left", padx=(0, 2))
            box.pack_propagate(False)
            lbl = tk.Label(box, text=label, bg=SURFACE, fg=TEXT_DARK, font=self.font(10, "bold"))
            lbl.pack(expand=True)
            line = tk.Frame(box, bg=SURFACE, height=3)
            line.pack(side="bottom", fill="x")
            for w in (box, lbl):
                w.bind("<Button-1>", lambda e, m=mode: self.switch_result_mode(m))
            self.mode_buttons[mode] = (box, lbl, line)

        self.lookup_body = tk.Frame(wrap, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        self.lookup_body.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self.paint_mode_tabs()
        self.build_lookup_body()

    def switch_result_mode(self, mode):
        self.result_mode = mode
        self.paint_mode_tabs()
        self.render_main_mode()

    def paint_mode_tabs(self):
        if not hasattr(self, "mode_buttons"):
            return
        for mode, (box, lbl, line) in self.mode_buttons.items():
            active = mode == self.result_mode
            lbl.config(fg=TERRA_DEEP if active else TEXT_DARK)
            line.config(bg=TERRA if active else SURFACE)

    def build_lookup_body(self):
        # Search strip
        strip = tk.Frame(self.lookup_body, bg="#FBFCFB", height=58)
        strip.pack(fill="x")
        strip.pack_propagate(False)
        tk.Label(strip, text="품목 일괄 검색", bg="#FBFCFB", fg=TEXT_DARK, font=self.font(10, "bold")).pack(side="left", padx=(18, 10))
        self.query_entry = tk.Entry(strip, bg=SURFACE, fg=TEXT_DARK, font=self.font(10), relief="flat", bd=0, highlightthickness=1, highlightbackground=BORDER, highlightcolor=SAGE)
        self.query_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=8)
        self.query_entry.insert(0, "품목 여러 개는 쉼표(,) 또는 세미콜론(;)으로 입력")
        self.query_entry.bind("<FocusIn>", self.clear_query_placeholder)
        self.query_entry.bind("<Return>", lambda e: self.batch_search())
        self.make_button(strip, "여러 줄 입력", self.open_multiline_dialog).pack(side="left", padx=(0, 8))
        self.make_button(strip, "조회", self.batch_search, primary=True).pack(side="left", padx=(0, 18))

        # main area
        self.main_area = tk.Frame(self.lookup_body, bg=SURFACE)
        self.main_area.pack(fill="both", expand=True)
        self.render_main_mode()

    def clear_query_placeholder(self, _=None):
        s = self.query_entry.get().strip()
        if s.startswith("품목 여러 개는"):
            self.query_entry.delete(0, "end")

    def open_multiline_dialog(self):
        d = tk.Toplevel(self)
        d.title("품목 여러 줄 입력")
        d.geometry("560x430")
        d.configure(bg=SURFACE)
        d.transient(self)
        d.grab_set()
        tk.Label(d, text="검색 품목 입력", bg=SURFACE, fg=TEXT, font=self.font(14, "bold")).pack(anchor="w", padx=20, pady=(20, 5))
        tk.Label(d, text="한 줄에 하나씩 입력할 수 있습니다.", bg=SURFACE, fg=MUTED, font=self.font(9)).pack(anchor="w", padx=20)
        t = tk.Text(d, bg=GRAY_SOFT, fg=TEXT_DARK, font=self.font(10), relief="flat", highlightthickness=1, highlightbackground=BORDER, padx=12, pady=10)
        t.pack(fill="both", expand=True, padx=20, pady=14)
        def apply():
            raw = t.get("1.0", "end").strip()
            self.query_entry.delete(0, "end")
            self.query_entry.insert(0, "; ".join([x.strip() for x in raw.splitlines() if x.strip()]))
            d.destroy()
            self.batch_search()
        self.make_button(d, "통합 조회", apply, primary=True).pack(anchor="e", padx=20, pady=(0, 18))

    def render_main_mode(self):
        if not hasattr(self, "main_area"):
            return
        for w in self.main_area.winfo_children():
            w.destroy()
        if self.result_mode == "drug":
            self.build_drug_mode()
        else:
            self.build_dur_mode()

    # ---------- drug mode ----------
    def build_drug_mode(self):
        left = tk.Frame(self.main_area, bg=SURFACE, width=388, highlightbackground=BORDER, highlightthickness=1)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        right = tk.Frame(self.main_area, bg=SURFACE)
        right.pack(side="left", fill="both", expand=True)

        # left drug list header
        h = tk.Frame(left, bg=SURFACE)
        h.pack(fill="x", padx=18, pady=(15, 10))
        self.drug_count_lbl = tk.Label(h, text="검색 품목 (0)", bg=SURFACE, fg=TEXT, font=self.font(12, "bold"))
        self.drug_count_lbl.pack(side="left")

        list_host = tk.Frame(left, bg=SURFACE)
        list_host.pack(fill="both", expand=True, padx=(10, 4), pady=(0, 8))
        self.drug_canvas = tk.Canvas(list_host, bg=SURFACE, highlightthickness=0, bd=0)
        sb = ttk.Scrollbar(list_host, orient="vertical", command=self.drug_canvas.yview)
        self.drug_canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.drug_canvas.pack(side="left", fill="both", expand=True)
        self.drug_list_frame = tk.Frame(self.drug_canvas, bg=SURFACE)
        self.drug_canvas_window = self.drug_canvas.create_window((0, 0), window=self.drug_list_frame, anchor="nw")
        self.drug_list_frame.bind("<Configure>", lambda e: self.drug_canvas.configure(scrollregion=self.drug_canvas.bbox("all")))
        self.drug_canvas.bind("<Configure>", lambda e: self.drug_canvas.itemconfig(self.drug_canvas_window, width=e.width))

        # right content
        self.right = right
        self.drug_title = tk.Label(right, text="품목을 검색해 주세요", bg=SURFACE, fg=TEXT, font=self.font(17, "bold"), anchor="w")
        self.drug_title.pack(fill="x", padx=22, pady=(14, 2))
        self.drug_subtitle = tk.Label(right, text="", bg=SURFACE, fg=TEXT_DARK, font=self.font(10), anchor="w")
        self.drug_subtitle.pack(fill="x", padx=22, pady=(0, 12))

        self.cards_frame = tk.Frame(right, bg=SURFACE)
        self.cards_frame.pack(fill="x", padx=22)
        self.cards = {}
        for i, (key, label) in enumerate(DUR_CARDS):
            card = tk.Frame(self.cards_frame, bg=GRAY_SOFT, height=92, highlightbackground=BORDER, highlightthickness=1, cursor="hand2")
            card.grid(row=i // 4, column=i % 4, sticky="nsew", padx=(0 if i % 4 == 0 else 8, 0), pady=(0, 8))
            card.grid_propagate(False)
            title = tk.Label(card, text=label, bg=GRAY_SOFT, fg="#667781", font=self.font(9), anchor="w")
            title.pack(fill="x", padx=16, pady=(12, 5))
            line = tk.Frame(card, bg=GRAY_SOFT)
            line.pack(fill="x", padx=16)
            dot = tk.Label(line, text="●", bg=GRAY_SOFT, fg=GRAY_ICON, font=self.font(8))
            dot.pack(side="left", padx=(0, 8))
            val = tk.Label(line, text="해당 없음", bg=GRAY_SOFT, fg="#5A6872", font=self.font(11), anchor="w")
            val.pack(side="left")
            count = tk.Label(line, text="", bg=GRAY_SOFT, fg=TEXT, font=self.font(16, "bold"))
            count.pack(side="right")
            arrow = tk.Label(card, text="›", bg=GRAY_SOFT, fg=TEXT_DARK, font=self.font(14), anchor="e")
            arrow.place(relx=0.96, rely=0.72, anchor="e")
            for w in (card, title, line, dot, val, count, arrow):
                w.bind("<Button-1>", lambda e, k=key: self.show_detail(k))
            self.cards[key] = (card, title, dot, val, count, arrow)
        for i in range(4):
            self.cards_frame.columnconfigure(i, weight=1)

        # detailed DUR tab strip
        self.detail_tabs = tk.Frame(right, bg=SURFACE)
        self.detail_tabs.pack(fill="x", padx=22, pady=(14, 0))
        self.detail_tab_buttons = {}

        self.detail_count = tk.Label(self.detail_tabs, text="", bg=SURFACE, fg=TEXT_DARK, font=self.font(11))
        self.detail_count.pack(side="right", padx=(8, 8))

        self.detail_host = tk.Frame(right, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        self.detail_host.pack(fill="both", expand=True, padx=22, pady=(0, 14))

        self.render_drug_list()
        if self.selected_drug and self.selected_drug in self.results:
            self.render_selected_drug()

    def render_drug_list(self):
        if not hasattr(self, "drug_list_frame"):
            return
        for w in self.drug_list_frame.winfo_children():
            w.destroy()
        self.drug_count_lbl.config(text=f"검색 품목 ({len(self.results)})")
        if not self.results:
            tk.Label(self.drug_list_frame, text="검색 결과가 없습니다.", bg=SURFACE, fg=MUTED, font=self.font(9)).pack(anchor="w", padx=12, pady=18)
            return
        latest = self.latest_upload_date()
        for q, r in self.results.items():
            selected = q == self.selected_drug
            bg = SAGE_ROW if selected else SURFACE
            row = tk.Frame(self.drug_list_frame, bg=bg, height=61, cursor="hand2")
            row.pack(fill="x", pady=0)
            row.pack_propagate(False)
            accent = tk.Frame(row, bg=SAGE if selected else bg, width=4)
            accent.pack(side="left", fill="y")
            status_text, status_color, sub = self.drug_list_status(r)
            icon = tk.Label(row, text=status_text, bg=bg, fg=status_color, font=self.font(12, "bold"), width=2)
            icon.pack(side="left", padx=(8, 5))
            text = tk.Frame(row, bg=bg)
            text.pack(side="left", fill="both", expand=True, pady=8)
            name = r.get("matched") or q
            tk.Label(text, text=self.ellipsize(name, 31), bg=bg, fg=TEXT_DARK, font=self.font(9, "bold"), anchor="w").pack(fill="x")
            tk.Label(text, text=sub, bg=bg, fg=(SAGE_DEEP if "DUR 해당" in sub else (ERROR if "확인" in sub else MUTED)), font=self.font(8), anchor="w").pack(fill="x", pady=(2, 0))
            date = tk.Label(row, text=latest, bg=bg, fg="#8B9498", font=self.font(8))
            date.pack(side="right", padx=10)
            for w in (row, accent, icon, text, date):
                w.bind("<Button-1>", lambda e, qq=q: self.choose_drug(qq))
            for child in text.winfo_children():
                child.bind("<Button-1>", lambda e, qq=q: self.choose_drug(qq))

    def drug_list_status(self, r):
        if not r.get("matched"):
            return "!", ORANGE, "품목 확인 필요"
        data = r.get("data") or {}
        dur_hits = sum(1 for k, _ in DUR_CARDS if self.has_hits(data, k))
        if dur_hits:
            return "●", SAGE, f"{dur_hits}개 DUR 해당"
        complete = self.all_required_complete()
        return ("●", "#9CA5A1", "DUR 해당 없음") if complete else ("!", ORANGE, "일부 DUR 확인불가")

    def choose_drug(self, q):
        self.selected_drug = q
        self.render_drug_list()
        self.render_selected_drug()

    def render_selected_drug(self):
        r = self.results.get(self.selected_drug)
        if not r:
            return
        if not r.get("matched"):
            self.drug_title.config(text=self.selected_drug)
            self.drug_subtitle.config(text="품목 확인 필요 · 등록된 DUR 품목명에서 후보를 찾지 못했습니다.", fg=ERROR)
            self.paint_empty_cards("확인불가")
            self.build_detail_tabs(None)
            self.render_empty_detail("등록된 DUR 품목명에서 후보를 찾지 못했습니다.")
            return

        self.drug_title.config(text=r["matched"])
        data = r["data"]
        hit_keys = [k for k, _ in DUR_CARDS if self.has_hits(data, k)]
        summary = self.drug_summary(data)
        self.drug_subtitle.config(text=summary, fg=TEXT_DARK)
        self.paint_cards(data)
        self.build_detail_tabs(data)
        if self.selected_dur not in [k for k, _ in DUR_CARDS]:
            self.selected_dur = "combo"
        if not self.has_hits(data, self.selected_dur):
            self.selected_dur = hit_keys[0] if hit_keys else "combo"
        self.render_detail(self.selected_dur)

    def drug_summary(self, data):
        parts = []
        hit_count = 0
        for key, label in DUR_CARDS:
            v = self.verdict(data, key)
            if v.startswith("○"):
                hit_count += 1
                if key == "combo":
                    n = sum(len(data[c]) for c in self.mapping()[key])
                    parts.append(f"병용금기 {n}건")
                elif key == "preg":
                    parts.append(v.replace("○ ", "임부금기 "))
                elif key == "dup":
                    n = sum(len(data[c]) for c in self.mapping()[key])
                    parts.append(f"효능군중복 {n}건")
        prefix = f"{hit_count}개 DUR 해당" if hit_count else "DUR 해당 없음"
        return "   |   ".join([prefix] + parts)

    def build_detail_tabs(self, data):
        for w in self.detail_tabs.winfo_children():
            if w is not self.detail_count:
                w.destroy()
        self.detail_tab_buttons = {}
        if data is None:
            self.detail_count.config(text="")
            return
        keys = [k for k, _ in DUR_CARDS if self.has_hits(data, k)]
        if not keys:
            keys = ["combo"]
        for key in keys:
            label = dict(DUR_CARDS)[key] + " 상세"
            box = tk.Frame(self.detail_tabs, bg=SURFACE, height=40, cursor="hand2")
            box.pack(side="left", padx=(0, 4))
            text = tk.Label(box, text=label, bg=SURFACE, fg=TEXT_DARK, font=self.font(9))
            text.pack(side="left", padx=14, pady=9)
            line = tk.Frame(box, bg=SURFACE, height=2)
            line.pack(side="bottom", fill="x")
            for w in (box, text):
                w.bind("<Button-1>", lambda e, k=key: self.show_detail(k))
            self.detail_tab_buttons[key] = (box, text, line)

    def paint_detail_tabs(self):
        for key, (box, text, line) in self.detail_tab_buttons.items():
            active = key == self.selected_dur
            text.config(fg=TERRA_DEEP if active else "#596872", font=self.font(9, "bold" if active else "normal"))
            line.config(bg=TERRA if active else SURFACE)

    def show_detail(self, key):
        self.selected_dur = key
        self.paint_detail_tabs()
        self.render_detail(key)

    def render_detail(self, key):
        for w in self.detail_host.winfo_children():
            w.destroy()
        r = self.results.get(self.selected_drug)
        if not r or not r.get("data"):
            self.render_empty_detail("상세정보가 없습니다.")
            return
        data = r["data"]
        self.paint_detail_tabs()
        if key == "combo":
            self.render_combo_detail(data)
        else:
            self.render_generic_detail(data, key)

    def render_empty_detail(self, text):
        for w in self.detail_host.winfo_children():
            w.destroy()
        tk.Label(self.detail_host, text=text, bg=SURFACE, fg=MUTED, font=self.font(10)).pack(anchor="w", padx=18, pady=20)

    def render_combo_detail(self, data):
        hits = []
        for cat in self.mapping()["combo"]:
            for h in data.get(cat, []):
                d = h["data"]
                other = "B" if h["side"] == "A" else "A"
                hits.append({
                    "상대 품목": self.find_value(d, [f"제품명{other}", f"품목명{other}"]) or "—",
                    "상대 성분": self.find_value(d, [f"성분명{other}"]) or "—",
                    "급여구분": self.find_value(d, [f"급여여부{other}", "급여여부A"]) or ("급여" if cat == "combo_paid" else "비급여"),
                    "고시번호": self.find_value(d, ["고시번호"]) or "—",
                    "고시일자": self.find_value(d, ["고시일자"]) or "—",
                    "상세내용": self.find_value(d, ["상세정보"]) or "파일에서 확인되지 않음",
                })
        self.detail_hits = hits
        self.detail_count.config(text=f"총 {len(hits)}건" if hits else "")
        if not hits:
            self.render_empty_detail(self.verdict(data, "combo"))
            return

        top = tk.Frame(self.detail_host, bg=SURFACE)
        top.pack(fill="both", expand=True)
        cols = ("No.", "상대 품목", "상대 성분", "급여구분", "고시번호", "고시일자")
        tree = ttk.Treeview(top, columns=cols, show="headings", style="Compact.Treeview", height=7)
        widths = {"No.": 45, "상대 품목": 330, "상대 성분": 240, "급여구분": 90, "고시번호": 120, "고시일자": 120}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths[c], anchor="center" if c in ("No.", "급여구분", "고시번호", "고시일자") else "w")
        sb = ttk.Scrollbar(top, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True)
        for i, h in enumerate(hits, 1):
            tree.insert("", "end", iid=str(i - 1), values=(i, h["상대 품목"], h["상대 성분"], h["급여구분"], h["고시번호"], h["고시일자"]))
        tree.bind("<<TreeviewSelect>>", lambda e: self.render_combo_selected(tree))

        self.combo_selected = tk.Frame(self.detail_host, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        self.combo_selected.pack(fill="x", padx=0, pady=(8, 0))
        if hits:
            tree.selection_set("0")
            tree.focus("0")
            self.render_combo_selected(tree)

    def render_combo_selected(self, tree):
        sel = tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        h = self.detail_hits[idx]
        for w in self.combo_selected.winfo_children():
            w.destroy()
        tk.Label(self.combo_selected, text="선택 항목 상세정보", bg=SURFACE, fg=TEXT, font=self.font(10, "bold")).pack(anchor="w", padx=14, pady=(10, 7))
        grid = tk.Frame(self.combo_selected, bg=SURFACE)
        grid.pack(fill="x", padx=14, pady=(0, 12))
        items = [
            ("상대 품목", h["상대 품목"]), ("고시번호", h["고시번호"]),
            ("상대 성분", h["상대 성분"]), ("고시일자", h["고시일자"]),
            ("급여구분", h["급여구분"]), ("상세내용", h["상세내용"]),
        ]
        for i, (k, v) in enumerate(items):
            r, c = divmod(i, 2)
            tk.Label(grid, text=k, bg="#F7F9F7", fg=TEXT_DARK, font=self.font(9), width=11, anchor="w", padx=8, pady=7).grid(row=r, column=c * 2, sticky="nsew", padx=(0, 0), pady=1)
            tk.Label(grid, text=v, bg=SURFACE, fg=TEXT_DARK, font=self.font(9), anchor="w", justify="left", wraplength=530, padx=10, pady=7).grid(row=r, column=c * 2 + 1, sticky="nsew", pady=1)
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(3, weight=1)

    def render_generic_detail(self, data, key):
        rows = []
        for cat in self.mapping()[key]:
            for h in data.get(cat, []):
                d = h["data"]
                clean = []
                for k, v in d.items():
                    if v and self.visible_field(k) and norm(k) not in (norm("원본시트"), norm("원본파일명"), norm("기준년월")):
                        clean.append((k, v))
                rows.append(clean)
        self.detail_count.config(text=f"총 {len(rows)}건" if rows else "")
        if not rows:
            self.render_empty_detail(self.verdict(data, key))
            return
        canvas = tk.Canvas(self.detail_host, bg=SURFACE, highlightthickness=0)
        sb = ttk.Scrollbar(self.detail_host, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)
        host = tk.Frame(canvas, bg=SURFACE)
        win = canvas.create_window((0, 0), window=host, anchor="nw")
        host.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
        for idx, group in enumerate(rows, 1):
            card = tk.Frame(host, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
            card.pack(fill="x", padx=10, pady=(10 if idx == 1 else 0, 8))
            tk.Label(card, text=f"{dict(DUR_CARDS)[key]} {idx}", bg=SURFACE, fg=TEXT, font=self.font(10, "bold")).pack(anchor="w", padx=12, pady=(10, 5))
            g = tk.Frame(card, bg=SURFACE)
            g.pack(fill="x", padx=12, pady=(0, 10))
            for r, (k, v) in enumerate(group):
                tk.Label(g, text=k, bg="#F7F9F7", fg="#54636D", font=self.font(9), width=18, anchor="w", padx=7, pady=6).grid(row=r, column=0, sticky="nsew", pady=1)
                tk.Label(g, text=v, bg=SURFACE, fg=TEXT_DARK, font=self.font(9), anchor="w", justify="left", wraplength=850, padx=10, pady=6).grid(row=r, column=1, sticky="nsew", pady=1)
            g.columnconfigure(1, weight=1)

    # ---------- DUR mode ----------
    def build_dur_mode(self):
        host = tk.Frame(self.main_area, bg=SURFACE)
        host.pack(fill="both", expand=True)
        head = tk.Frame(host, bg=SURFACE)
        head.pack(fill="x", padx=18, pady=(15, 10))
        tk.Label(head, text="DUR 종류별 해당 품목", bg=SURFACE, fg=TEXT, font=self.font(13, "bold")).pack(side="left")

        self.dur_filter_bar = tk.Frame(host, bg=SURFACE)
        self.dur_filter_bar.pack(fill="x", padx=18, pady=(0, 10))
        counts = self.dur_counts()
        self.dur_filter_buttons = {}
        for key, label in DUR_CARDS:
            cnt = counts.get(key, 0)
            b = tk.Label(self.dur_filter_bar, text=f"{label} {cnt}", bg="#F7F9F7", fg=TEXT_DARK, font=self.font(9), padx=12, pady=7, cursor="hand2", highlightbackground=BORDER, highlightthickness=1)
            b.pack(side="left", padx=(0, 6))
            b.bind("<Button-1>", lambda e, k=key: self.render_dur_section(k))
            self.dur_filter_buttons[key] = b

        self.dur_section = tk.Frame(host, bg=SURFACE)
        self.dur_section.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        first = next((k for k, _ in DUR_CARDS if counts.get(k, 0)), "combo")
        self.render_dur_section(first)

    def dur_counts(self):
        out = {k: 0 for k, _ in DUR_CARDS}
        for _, r in self.results.items():
            if not r.get("data"):
                continue
            for k, _ in DUR_CARDS:
                if self.has_hits(r["data"], k):
                    out[k] += 1
        return out

    def render_dur_section(self, key):
        if not hasattr(self, "dur_section"):
            return
        for k, b in self.dur_filter_buttons.items():
            active = k == key
            b.config(bg=TERRA_SOFT if active else "#F7F9F7", fg=TERRA_DEEP if active else TEXT_DARK)
        for w in self.dur_section.winfo_children():
            w.destroy()
        label = dict(DUR_CARDS)[key]
        rows = []
        for q, r in self.results.items():
            if not r.get("data"):
                continue
            if self.has_hits(r["data"], key):
                rows.append((r.get("matched") or q, self.verdict(r["data"], key), self.short_detail(r["data"], key)))
        top = tk.Frame(self.dur_section, bg=SURFACE)
        top.pack(fill="x", pady=(2, 8))
        tk.Label(top, text=label, bg=SURFACE, fg=TEXT, font=self.font(12, "bold")).pack(side="left")
        tk.Label(top, text=f"{len(rows)}개 품목", bg=SURFACE, fg=MUTED, font=self.font(9)).pack(side="right")
        tree = ttk.Treeview(self.dur_section, columns=("품목", "판정", "핵심정보"), show="headings", style="Compact.Treeview")
        for c, w in (("품목", 420), ("판정", 180), ("핵심정보", 650)):
            tree.heading(c, text=c)
            tree.column(c, width=w, anchor="w")
        tree.pack(fill="both", expand=True)
        for name, verdict, detail in rows:
            tree.insert("", "end", values=(name, verdict, detail))
        if not rows:
            tree.insert("", "end", values=("해당 품목 없음", "—", "현재 검색 품목 중 해당 결과가 없습니다."))

    # ---------- result calculations ----------
    def mapping(self):
        return {
            "combo": ["combo_paid", "combo_unpaid"],
            "age": ["age"],
            "preg": ["preg"],
            "lact": ["lact"],
            "dup": ["dup"],
            "tele": ["tele"],
            "cost": ["cost"],
        }

    def all_required_complete(self):
        sts = self.idx.statuses()
        return all(sts.get(k) and sts[k][0] for k, _ in CATEGORIES)

    def has_hits(self, data, key):
        return any(data.get(c) for c in self.mapping()[key])

    def find_value(self, d, keys):
        for target in keys:
            for k, v in d.items():
                if norm(target) == norm(k) and v:
                    return v
        return ""

    def verdict(self, data, key):
        sts = self.idx.statuses()
        cats = self.mapping()[key]
        complete = all(sts.get(c) and sts[c][0] for c in cats)
        hits = sum(len(data.get(c, [])) for c in cats)
        if not complete:
            return "확인불가"
        if not hits:
            return "— 해당 없음"
        if key == "preg":
            grades = sorted({self.find_value(x["data"], ["금기등급", "등급"]) for x in data["preg"] if self.find_value(x["data"], ["금기등급", "등급"])})
            return "○ " + ("/".join(g + "등급" for g in grades) if grades else "해당")
        if key == "age":
            x = data["age"][0]["data"]
            age = self.find_value(x, ["특정연령"])
            unit = self.find_value(x, ["특정연령단위", "특정연령단위코드", "연령단위"])
            cond = self.find_value(x, ["연령처리조건"])
            return "○ " + (" ".join(z for z in [age + unit if age else "", cond] if z) or "해당")
        return f"○ 해당 ({hits})" if hits > 1 else "○ 해당"

    def paint_empty_cards(self, text):
        for key, _ in DUR_CARDS:
            self.set_card(key, text, None, "warn")

    def set_card(self, key, text, count, state):
        if key not in self.cards:
            return
        card, title, dot, val, cnt, arrow = self.cards[key]
        if state == "yes":
            bg, fg, dotc = SAGE_SOFT, SAGE_DEEP, SAGE
        elif state == "preg":
            bg, fg, dotc = TERRA_SOFT, TERRA_DEEP, TERRA
        elif state == "warn":
            bg, fg, dotc = ORANGE_SOFT, ORANGE, ORANGE
        else:
            bg, fg, dotc = GRAY_SOFT, "#556671", GRAY_ICON
        card.config(bg=bg, highlightbackground=TERRA if key == self.selected_dur and state != "no" else BORDER)
        for w in (title, dot, val, cnt, arrow):
            w.config(bg=bg)
        dot.config(fg=dotc)
        val.config(text=text, fg=fg)
        cnt.config(text=count or "", fg=fg)

    def paint_cards(self, data):
        for key, _ in DUR_CARDS:
            v = self.verdict(data, key)
            hits = sum(len(data.get(c, [])) for c in self.mapping()[key])
            if v == "확인불가":
                self.set_card(key, "확인불가", None, "warn")
            elif not v.startswith("○"):
                self.set_card(key, "해당 없음", None, "no")
            elif key == "preg":
                self.set_card(key, v.replace("○ ", ""), None, "preg")
            elif key == "age":
                self.set_card(key, v.replace("○ ", ""), None, "yes")
            else:
                self.set_card(key, "해당", f"{hits}건" if hits > 1 else "", "yes")

    def visible_field(self, key):
        n = norm(key)
        hidden = ("성분코드", "주성분코드", "일반명코드", "제품코드", "약품코드", "업체명")
        return not any(norm(x) in n for x in hidden)

    def short_detail(self, data, key):
        if key == "combo":
            vals = []
            for cat in self.mapping()[key]:
                for h in data.get(cat, [])[:3]:
                    d = h["data"]
                    other = "B" if h["side"] == "A" else "A"
                    p = self.find_value(d, [f"제품명{other}", f"품목명{other}"])
                    s = self.find_value(d, [f"성분명{other}"])
                    vals.append(" / ".join(x for x in [p, s] if x))
            return "; ".join(vals)
        if key == "preg":
            return self.verdict(data, key).replace("○ ", "")
        if key == "age":
            return self.verdict(data, key).replace("○ ", "")
        if key == "dup":
            vals = []
            for h in data.get("dup", [])[:3]:
                d = h["data"]
                eff = self.find_value(d, ["효능군"])
                grp = self.find_value(d, ["Group"])
                vals.append(" / ".join(x for x in [eff, grp] if x))
            return "; ".join(vals)
        if key == "tele":
            vals = []
            for h in data.get("tele", [])[:3]:
                d = h["data"]
                vals.append(self.find_value(d, ["약품구분", "비고"]))
            return "; ".join(x for x in vals if x)
        return self.verdict(data, key).replace("○ ", "")

    # ---------- searching ----------
    def batch_search(self):
        raw = self.query_entry.get().strip() if hasattr(self, "query_entry") else ""
        if not raw or raw.startswith("품목 여러 개는"):
            messagebox.showinfo("입력 필요", "검색할 품목명을 입력해 주세요.")
            return
        queries = [x.strip() for x in re.split(r"[,;\n]", raw) if x.strip()]
        queries = list(dict.fromkeys(queries))
        if not queries:
            return
        self.open_candidate_dialog(queries)

    def open_candidate_dialog(self, queries):
        """Resolve every partial query to an explicit user-selected product before DUR lookup."""
        candidate_map = {q: self.idx.suggestions(q, 80) for q in queries}

        d = tk.Toplevel(self)
        d.title("검색 품목 후보 선택")
        d.geometry("980x620")
        d.minsize(760, 430)
        d.configure(bg=SURFACE)
        d.transient(self)
        d.grab_set()

        tk.Label(d, text="검색 품목 후보 선택", bg=SURFACE, fg=TEXT, font=self.font(15, "bold")).pack(anchor="w", padx=24, pady=(20, 4))
        tk.Label(
            d,
            text="입력한 검색어마다 실제 조회할 품목을 선택하세요. 후보가 여러 개인 경우 함량·제형까지 확인할 수 있습니다.",
            bg=SURFACE, fg=MUTED, font=self.font(9)
        ).pack(anchor="w", padx=24, pady=(0, 14))

        host = tk.Frame(d, bg=SURFACE)
        host.pack(fill="both", expand=True, padx=24)
        canvas = tk.Canvas(host, bg=SURFACE, highlightthickness=0)
        sb = ttk.Scrollbar(host, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        rows_frame = tk.Frame(canvas, bg=SURFACE)
        win = canvas.create_window((0, 0), window=rows_frame, anchor="nw")
        rows_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))

        tk.Label(rows_frame, text="입력 검색어", bg=GRAY_SOFT, fg=TEXT_DARK, font=self.font(9, "bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=(0, 1), ipady=8)
        tk.Label(rows_frame, text="선택 품목", bg=GRAY_SOFT, fg=TEXT_DARK, font=self.font(9, "bold"), anchor="w").grid(row=0, column=1, sticky="ew", ipady=8)
        rows_frame.columnconfigure(0, weight=0, minsize=230)
        rows_frame.columnconfigure(1, weight=1)

        combo_vars = {}
        for ridx, q in enumerate(queries, start=1):
            vals = candidate_map[q]
            tk.Label(rows_frame, text=q, bg=SURFACE, fg=TEXT_DARK, font=self.font(9, "bold"), anchor="w").grid(row=ridx, column=0, sticky="ew", padx=(4, 14), pady=5)
            display_vals = vals if vals else ["후보 없음"]
            var = tk.StringVar(value=display_vals[0])
            cb = ttk.Combobox(rows_frame, textvariable=var, values=display_vals, state="readonly", font=self.font(9))
            cb.grid(row=ridx, column=1, sticky="ew", padx=(0, 4), pady=5, ipady=3)
            combo_vars[q] = (var, vals)

        foot = tk.Frame(d, bg=SURFACE)
        foot.pack(fill="x", padx=24, pady=18)

        def run_selected():
            self.results = {}
            for q in queries:
                var, vals = combo_vars[q]
                matched = var.get().strip() if vals else None
                if not matched or matched == "후보 없음":
                    self.results[q] = {"query": q, "matched": None, "data": None, "candidates": vals}
                else:
                    self.results[q] = {
                        "query": q,
                        "matched": matched,
                        "data": self.idx.search_product(matched),
                        "candidates": vals,
                    }
            d.destroy()
            self.selected_drug = next(iter(self.results.keys()), None)
            self.result_mode = "drug"
            self.paint_mode_tabs()
            self.render_main_mode()

        self.make_button(foot, "취소", d.destroy).pack(side="right")
        self.make_button(foot, "선택 품목으로 조회", run_selected, primary=True).pack(side="right", padx=(0, 8))

    # ---------- file handling ----------
    def classify_file(self, path):
        n = norm(Path(path).name)
        if "병용금기" in n:
            if "비급여" in n:
                return "combo_unpaid"
            if "급여" in n:
                return "combo_paid"
        if "연령금기" in n:
            return "age"
        if "임부금기" in n:
            return "preg"
        if "수유부주의" in n:
            return "lact"
        if "효능군중복" in n:
            return "dup"
        if "비대면진료" in n:
            return "tele"
        if "비용효과" in n:
            return "cost"
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            sig = norm(" ".join(wb.sheetnames))
            ws = wb[wb.sheetnames[0]]
            vals = []
            for row in ws.iter_rows(min_row=1, max_row=min(8, ws.max_row), values_only=True):
                vals += [sval(x) for x in row if x is not None]
            sig += norm(" ".join(vals[:120]))
            wb.close()
            if "저함량" in sig and "고함량" in sig:
                return "cost"
            if "효능군중복점검코드" in sig:
                return "dup"
            if "특정연령" in sig:
                return "age"
            if "금기등급" in sig:
                return "preg"
            if "주성분코드" in sig and "공고번호" in sig:
                return "lact"
            if "성분명a" in sig and "성분명b" in sig:
                return "combo_unpaid" if "비급여" in sig else "combo_paid"
            if "약품구분" in sig and "비고" in sig:
                return "tele"
        except Exception:
            pass
        return None

    def choose_files(self):
        paths = filedialog.askopenfilenames(title="DUR 품목리스트 Excel 선택", filetypes=[("Excel", "*.xlsx")])
        if not paths:
            return
        classified, unknown, duplicates = {}, [], []
        for p in paths:
            k = self.classify_file(p)
            if not k:
                unknown.append(Path(p).name)
            elif k in classified:
                duplicates.append(Path(p).name)
            else:
                classified[k] = p
        if unknown or duplicates:
            msg = []
            if unknown:
                msg.append("자동 분류 실패:\n- " + "\n- ".join(unknown))
            if duplicates:
                msg.append("같은 종류로 중복 분류:\n- " + "\n- ".join(duplicates))
            messagebox.showwarning("일부 파일 확인 필요", "\n\n".join(msg) + "\n\n분류된 파일은 계속 등록합니다.")
        if not classified:
            return
        for k, p in classified.items():
            dest = ROOT_DIR / "files" / (k + "__" + Path(p).name)
            for old in (ROOT_DIR / "files").glob(k + "__*"):
                try:
                    old.unlink()
                except Exception:
                    pass
            shutil.copy2(p, dest)
            self.cfg[k] = str(dest)
            self.cfg.setdefault("_uploaded_at", {})[k] = datetime.now().isoformat(timespec="minutes")
        self.save_cfg()
        self.index_many([(k, Path(self.cfg[k])) for k in classified])

    def index_many(self, items):
        def worker():
            errors = []
            for k, p in items:
                ok, msg = self.idx.index_file(k, p)
                if not ok:
                    errors.append(f"{dict(CATEGORIES)[k]}: {msg}")
                self.after(0, self.refresh_status)
            self.after(0, lambda: self.index_many_done(errors, len(items)))
        threading.Thread(target=worker, daemon=True).start()

    def index_many_done(self, errors, n):
        self.refresh_status()
        if errors:
            messagebox.showwarning("일부 파일 확인 필요", "\n".join(errors))
        else:
            messagebox.showinfo("완료", f"{n}개 기준파일 등록/인덱싱이 완료되었습니다.")

    def reindex_all(self):
        items = [(k, Path(v)) for k, v in self.cfg.items() if k in dict(CATEGORIES) and isinstance(v, str) and Path(v).exists()]
        if not items:
            messagebox.showwarning("기준파일 없음", "먼저 Excel 파일을 등록해 주세요.")
            return
        self.index_many(items)

    def refresh_status(self):
        sts = self.idx.statuses()
        times = self.cfg.get("_uploaded_at", {})
        latest = max(times.values()) if times else ""
        self.header_update.config(text=("● " + latest.replace("T", " ")[:16]) if latest else "● —")
        if hasattr(self, "file_rows"):
            for key, _ in CATEGORIES:
                status_lbl, date_lbl, file_lbl = self.file_rows[key]
                v = sts.get(key)
                stamp = times.get(key)
                date_lbl.config(text=stamp.replace("T", " ")[:16] if stamp else "—")
                path = self.cfg.get(key)
                file_lbl.config(text=Path(path).name.split("__", 1)[-1] if path else "—")
                if v and v[0]:
                    status_lbl.config(text=f"● 정상 · {v[4]}시트", fg=SAGE_DEEP)
                elif v:
                    status_lbl.config(text="● 오류", fg=ERROR)
                else:
                    status_lbl.config(text="● 미등록", fg=MUTED)

    def latest_upload_date(self):
        times = self.cfg.get("_uploaded_at", {})
        latest = max(times.values()) if times else ""
        return latest[:10] if latest else "—"

    # ---------- helpers ----------
    def make_button(self, parent, text, command, primary=False):
        bg = TERRA if primary else SURFACE
        fg = "white" if primary else TEXT_DARK
        active = TERRA_DEEP if primary else "#F0F3F1"
        b = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active,
            activeforeground=fg,
            relief="flat",
            bd=0,
            font=self.font(9, "bold" if primary else "normal"),
            padx=14,
            pady=8,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=(TERRA if primary else BORDER),
        )
        return b

    def ellipsize(self, text, max_chars):
        text = str(text)
        return text if len(text) <= max_chars else text[: max_chars - 1] + "…"


if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception:
        err = traceback.format_exc()
        try:
            messagebox.showerror("DUR Dashboard 오류", err)
        except Exception:
            print(err)
