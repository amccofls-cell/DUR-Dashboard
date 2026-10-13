from __future__ import annotations

import json
import os
import re
import shutil
import sqlite3
import threading
import time
import traceback
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import load_workbook

APP_NAME = "DUR 품목리스트 조회 프로그램"
VERSION = "1.5.2"

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

APP_BG = "#F4F5F2"
SURFACE = "#FFFFFF"
TEXT = "#183042"
TEXT_DARK = "#20313D"
MUTED = "#7D8991"
BORDER = "#DDE3DF"
SAGE = "#5E9872"
SAGE_DEEP = "#3F7B59"
SAGE_SOFT = "#EEF8F1"
TERRA = "#C95737"
TERRA_DEEP = "#A93B24"
TERRA_SOFT = "#FFF4F0"
GRAY_SOFT = "#F7F8F7"
ORANGE = "#E88A2D"
ERROR = "#B34E46"

PLACEHOLDER = "품목명의 일부를 입력하세요"


def app_dir() -> Path:
    root = Path(os.getenv("LOCALAPPDATA") or Path.home()) / "DUR_Dashboard"
    root.mkdir(parents=True, exist_ok=True)
    (root / "files").mkdir(exist_ok=True)
    return root


ROOT_DIR = app_dir()
CFG_PATH = ROOT_DIR / "config.json"
DB_PATH = ROOT_DIR / "dur_index.sqlite3"


def sval(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def norm(v):
    s = sval(v).lower()
    s = re.sub(r"\s+", "", s)
    s = s.replace("㎖", "ml").replace("밀리리터", "ml").replace("밀리그램", "mg")
    s = re.sub(r"[()\[\]{}·,._\-/]", "", s)
    return s


def product_core(v):
    s = sval(v)
    s = re.sub(r"_\s*\([^)]*\)\s*$", "", s)
    s = re.sub(r"\([^)]*\)", "", s)
    return norm(s)


def split_top_level_aliases(v):
    text = sval(v)
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


def unique_headers(values):
    out, seen = [], {}
    for i, v in enumerate(values):
        h = sval(v) or f"col_{i+1}"
        seen[h] = seen.get(h, 0) + 1
        if seen[h] > 1:
            h = f"{h}_{seen[h]}"
        out.append(h)
    return out


def find_header_row(ws, max_scan=18):
    keys = ("제품명", "품목명", "약품명", "성분명", "성분명a", "저함량", "특정연령", "약품구분")
    best = (1, -1)
    for r in range(1, min(ws.max_row, max_scan) + 1):
        vals = [norm(c.value) for c in ws[r][: min(ws.max_column, 40)]]
        score = sum(any(k in x for k in keys) for x in vals)
        if score > best[1]:
            best = (r, score)
    return best[0]


COMBO_KEEP = {
    1: "성분명A",
    4: "제품명A",
    6: "급여여부A",
    7: "성분명B",
    10: "제품명B",
    13: "고시번호",
    14: "고시일자",
    15: "상세정보",
}


def combo_rows(ws):
    hr = find_header_row(ws)
    for row in ws.iter_rows(min_row=hr + 1, min_col=1, max_col=15, values_only=True):
        if not any(row[i - 1] is not None and sval(row[i - 1]) for i in COMBO_KEEP):
            continue
        yield {name: sval(row[i - 1]) for i, name in COMBO_KEEP.items()}


def normal_rows(ws):
    hr = find_header_row(ws)
    headers = unique_headers([c.value for c in ws[hr]])
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        if not any(v is not None and sval(v) for v in row):
            continue
        yield headers, row


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
        headers.append((prefix + " " + (sval(h) or f"col_{i+1}")).strip())
    headers = unique_headers(headers)
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(v is not None and sval(v) for v in row):
            continue
        yield headers, row


def detect_name_cols(headers):
    hn = [norm(h) for h in headers]
    out = []
    for i, h in enumerate(hn):
        if h in ("제품명", "품목명", "약품명") or h.startswith("제품명") or h.startswith("품목명") or h.startswith("약품명"):
            out.append(i)
    return out[:3]


def record_product_code(d, category, side=""):
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
        self.db_path = Path(db_path)
        self.lock = threading.RLock()

    def conn(self):
        c = sqlite3.connect(self.db_path, timeout=60)
        c.execute("PRAGMA busy_timeout=60000")
        return c

    def with_retry(self, fn, tries=8):
        last = None
        for i in range(tries):
            try:
                return fn()
            except sqlite3.OperationalError as e:
                last = e
                if "locked" not in str(e).lower():
                    raise
                time.sleep(0.4 + i * 0.4)
        raise last

    def init(self):
        def _work():
            with self.lock, self.conn() as c:
                try:
                    c.execute("PRAGMA journal_mode=WAL")
                except sqlite3.OperationalError:
                    pass
                c.execute(
                    "CREATE TABLE IF NOT EXISTS records("
                    "id INTEGER PRIMARY KEY, category TEXT, sheet TEXT, product TEXT, "
                    "product_norm TEXT, core_norm TEXT, product_code TEXT, side TEXT, data TEXT)"
                )
                c.execute("CREATE INDEX IF NOT EXISTS ix_records_cat_name ON records(category,product_norm)")
                c.execute("CREATE INDEX IF NOT EXISTS ix_records_cat_core ON records(category,core_norm)")
                c.execute("CREATE INDEX IF NOT EXISTS ix_records_cat_code ON records(category,product_code)")
                c.execute(
                    "CREATE TABLE IF NOT EXISTS status("
                    "category TEXT PRIMARY KEY, ok INTEGER, message TEXT, filename TEXT, "
                    "indexed_at TEXT, sheet_count INTEGER, row_count INTEGER)"
                )
        return self.with_retry(_work)

    def clear_category(self, c, category):
        c.execute("DELETE FROM records WHERE category=?", (category,))
        c.execute("DELETE FROM status WHERE category=?", (category,))

    def index_file(self, category, path, progress=None):
        path = Path(path)
        self.init()

        def _work():
            total = 0
            sheets = 0
            wb = load_workbook(path, read_only=True, data_only=True)
            try:
                with self.lock, self.conn() as c:
                    self.clear_category(c, category)
                    for ws in wb.worksheets:
                        sheets += 1
                        if progress:
                            progress(f"{ws.title} 시트 인덱싱 중…")
                        batch = []

                        if category.startswith("combo"):
                            for d in combo_rows(ws):
                                for side, product_name in (("A", d.get("제품명A", "")), ("B", d.get("제품명B", ""))):
                                    if product_name:
                                        batch.append((category, ws.title, product_name, norm(product_name),
                                                      product_core(product_name), "", side,
                                                      json.dumps(d, ensure_ascii=False)))
                                    if len(batch) >= 3000:
                                        c.executemany(
                                            "INSERT INTO records(category,sheet,product,product_norm,core_norm,product_code,side,data) VALUES(?,?,?,?,?,?,?,?)",
                                            batch
                                        )
                                        total += len(batch)
                                        batch = []
                        else:
                            iterator = cost_rows(ws) if category == "cost" else normal_rows(ws)
                            for headers, row in iterator:
                                d = {headers[i]: sval(row[i]) if i < len(row) else "" for i in range(len(headers))}
                                if category == "cost":
                                    low = next((d[k] for k in d if "저함량" in k and "제품명" in k), "")
                                    high = next((d[k] for k in d if "고함량" in k and "제품명" in k), "")
                                    pairs = [("저함량", low), ("고함량", high)]
                                else:
                                    pairs = []
                                    for i in detect_name_cols(headers):
                                        raw_name = sval(row[i]) if i < len(row) else ""
                                        if not raw_name:
                                            continue
                                        names = split_top_level_aliases(raw_name) if category == "tele" else [raw_name]
                                        pairs.extend([("", n) for n in names])

                                for side, product_name in pairs:
                                    if not product_name:
                                        continue
                                    batch.append((category, ws.title, product_name, norm(product_name),
                                                  product_core(product_name),
                                                  record_product_code(d, category, side),
                                                  side, json.dumps(d, ensure_ascii=False)))
                                if len(batch) >= 3000:
                                    c.executemany(
                                        "INSERT INTO records(category,sheet,product,product_norm,core_norm,product_code,side,data) VALUES(?,?,?,?,?,?,?,?)",
                                        batch
                                    )
                                    total += len(batch)
                                    batch = []

                        if batch:
                            c.executemany(
                                "INSERT INTO records(category,sheet,product,product_norm,core_norm,product_code,side,data) VALUES(?,?,?,?,?,?,?,?)",
                                batch
                            )
                            total += len(batch)

                    c.execute(
                        "INSERT OR REPLACE INTO status VALUES(?,?,?,?,?,?,?)",
                        (category, 1, "정상", path.name, datetime.now().isoformat(timespec="seconds"), sheets, total),
                    )
                return True, f"{sheets}개 시트 · {total:,}개 엔트리"
            finally:
                wb.close()

        try:
            return self.with_retry(_work)
        except Exception as e:
            try:
                with self.conn() as c:
                    self.clear_category(c, category)
                    c.execute("INSERT OR REPLACE INTO status VALUES(?,?,?,?,?,?,?)",
                              (category, 0, str(e), path.name, datetime.now().isoformat(timespec="seconds"), 0, 0))
            except Exception:
                pass
            return False, str(e)

    def statuses(self):
        self.init()
        def _work():
            with self.conn() as c:
                return {r[0]: r[1:] for r in c.execute(
                    "SELECT category,ok,message,filename,indexed_at,sheet_count,row_count FROM status"
                )}
        return self.with_retry(_work)

    def suggestions(self, q, limit=40):
        nq = norm(q)
        if not nq:
            return []
        def _work():
            with self.conn() as c:
                rows = c.execute(
                    "SELECT product, COUNT(*) n FROM records "
                    "WHERE product_norm LIKE ? "
                    "GROUP BY product "
                    "ORDER BY CASE WHEN product_norm LIKE ? THEN 0 ELSE 1 END, length(product), product "
                    "LIMIT ?",
                    (f"%{nq}%", f"{nq}%", limit),
                ).fetchall()
            return [r[0] for r in rows]
        return self.with_retry(_work)

    def search_product(self, product):
        np = norm(product)
        cp = product_core(product)
        raw = {}
        codes = set()

        def _work():
            with self.conn() as c:
                for cat, _ in CATEGORIES:
                    rows = c.execute(
                        "SELECT sheet,product,product_code,side,data FROM records "
                        "WHERE category=? AND product_norm=?",
                        (cat, np)
                    ).fetchall()
                    if not rows and cp and len(cp) >= 4:
                        rows = c.execute(
                            "SELECT sheet,product,product_code,side,data FROM records "
                            "WHERE category=? AND core_norm=?",
                            (cat, cp)
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
                            f"SELECT sheet,product,product_code,side,data FROM records "
                            f"WHERE category=? AND product_code IN ({ph})",
                            (cat, *sorted(codes)),
                        ).fetchall()

                for cat in ("age", "tele"):
                    if raw.get(cat) or not cp or len(cp) < 5:
                        continue
                    raw[cat] = c.execute(
                        "SELECT sheet,product,product_code,side,data FROM records "
                        "WHERE category=? AND (product_norm LIKE ? OR core_norm LIKE ?)",
                        (cat, f"%{cp}%", f"%{cp}%"),
                    ).fetchall()

            out = {}
            for cat, _ in CATEGORIES:
                out[cat] = [
                    {"sheet": r[0], "product": r[1], "product_code": r[2], "side": r[3], "data": json.loads(r[4])}
                    for r in raw.get(cat, [])
                ]
            return out
        return self.with_retry(_work)


def live_small_file_match(category, path, product, known_codes=None):
    """Direct fallback for small age/tele workbooks when local SQLite index is stale/incomplete."""
    known_codes = {str(x).strip() for x in (known_codes or set()) if str(x).strip()}
    np = norm(product)
    cp = product_core(product)
    hits = []
    p = Path(path)
    if not p.exists():
        return hits

    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            for headers, row in normal_rows(ws):
                d = {headers[i]: sval(row[i]) if i < len(row) else "" for i in range(len(headers))}
                name_cols = detect_name_cols(headers)
                if not name_cols:
                    continue

                raw_names = []
                for i in name_cols:
                    raw = sval(row[i]) if i < len(row) else ""
                    if not raw:
                        continue
                    raw_names.extend(split_top_level_aliases(raw) if category == "tele" else [raw])

                row_code = record_product_code(d, category, "")
                matched = False
                matched_name = ""

                for candidate in raw_names:
                    if norm(candidate) == np:
                        matched, matched_name = True, candidate
                        break
                    ccore = product_core(candidate)
                    if cp and ccore == cp:
                        matched, matched_name = True, candidate
                        break
                    if cp and len(cp) >= 5 and (cp in ccore or ccore in cp):
                        matched, matched_name = True, candidate
                        break

                if not matched and row_code and row_code in known_codes:
                    matched = True
                    matched_name = raw_names[0] if raw_names else product

                if matched:
                    hits.append({
                        "sheet": ws.title,
                        "product": matched_name or product,
                        "product_code": row_code,
                        "side": "",
                        "data": d,
                    })
    finally:
        wb.close()
    return hits


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME}  {VERSION}")
        self.geometry("1500x900")
        self.minsize(1180, 760)
        self.configure(bg=APP_BG)

        self.idx = Indexer()
        try:
            self.idx.init()
        except sqlite3.OperationalError as e:
            if "locked" in str(e).lower():
                messagebox.showerror("데이터베이스 사용 중",
                                     "검색 인덱스가 다른 DUR 프로그램에서 사용 중입니다.\n"
                                     "실행 중인 DUR Dashboard를 모두 종료한 뒤 다시 실행해 주세요.")
                self.destroy()
                return
            raise

        self.cfg = self.load_cfg()
        self.selected_products = []
        self.results = {}
        self.current_product = None
        self.current_dur = "combo"

        self.setup_style()
        self.build_header()
        self.page_host = tk.Frame(self, bg=APP_BG)
        self.page_host.pack(fill="both", expand=True)

        self.footer = tk.Label(
            self, text="Created by Chaereen Jeong",
            bg=APP_BG, fg="#8A948F", font=self.font_en(8)
        )
        self.footer.place(relx=0.992, rely=0.992, anchor="se")

        self.show_lookup_page()
        self.after(500, self.auto_reindex_v151)

    def load_cfg(self):
        try:
            return json.loads(CFG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def save_cfg(self):
        CFG_PATH.write_text(json.dumps(self.cfg, ensure_ascii=False, indent=2), encoding="utf-8")

    def font(self, size=10, weight="normal"):
        return ("NanumGothic", size, weight)

    def font_en(self, size=10, weight="normal"):
        return ("Arial", size, weight)

    def setup_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Compact.Treeview", font=self.font(9), rowheight=30,
                        background=SURFACE, fieldbackground=SURFACE, foreground=TEXT_DARK,
                        borderwidth=0, relief="flat")
        style.map("Compact.Treeview", background=[("selected", "#E9F2EE")], foreground=[("selected", TEXT_DARK)])
        style.configure("Compact.Treeview.Heading", font=self.font(9, "bold"),
                        background="#F8FAF8", foreground="#425563", relief="flat",
                        borderwidth=0, padding=(7, 7))

    def build_header(self):
        h = tk.Frame(self, bg=SURFACE, height=70, highlightbackground=BORDER, highlightthickness=1)
        h.pack(fill="x")
        h.pack_propagate(False)

        brand = tk.Frame(h, bg=SURFACE)
        brand.pack(side="left", padx=(26, 36), pady=12)
        tk.Label(brand, text="❦", bg=SURFACE, fg=SAGE_DEEP,
                 font=("Segoe UI Symbol", 26)).pack(side="left", padx=(0, 10))
        tk.Label(brand, text=APP_NAME, bg=SURFACE, fg=TEXT,
                 font=self.font(15, "bold")).pack(side="left")

        nav = tk.Frame(h, bg=SURFACE)
        nav.pack(side="left", fill="y")
        self.nav = {}
        for key, label in (("lookup", "DUR 조회"), ("files", "기준파일 관리")):
            box = tk.Frame(nav, bg=SURFACE, width=150, height=69, cursor="hand2")
            box.pack(side="left")
            box.pack_propagate(False)
            lbl = tk.Label(box, text=label, bg=SURFACE, fg=TEXT_DARK, font=self.font(10))
            lbl.pack(expand=True)
            line = tk.Frame(box, bg=SURFACE, height=3)
            line.pack(side="bottom", fill="x")
            for w in (box, lbl):
                w.bind("<Button-1>", lambda e, k=key: self.navigate(k))
            self.nav[key] = (lbl, line)

        right = tk.Frame(h, bg=SURFACE)
        right.pack(side="right", padx=28, pady=10)
        tk.Label(right, text="기준파일 최종 업데이트", bg=SURFACE, fg="#526D60",
                 font=self.font(9)).pack(anchor="w")
        self.header_update = tk.Label(right, text="● —", bg=SURFACE, fg=SAGE_DEEP,
                                      font=self.font(10))
        self.header_update.pack(anchor="w", pady=(2, 0))

    def navigate(self, page):
        for w in self.page_host.winfo_children():
            w.destroy()
        for key, (lbl, line) in self.nav.items():
            active = key == page
            lbl.config(fg=TERRA if active else TEXT_DARK,
                       font=self.font(10, "bold" if active else "normal"))
            line.config(bg=TERRA if active else SURFACE)
        if page == "files":
            self.show_files_page()
        else:
            self.show_lookup_page()

    def make_button(self, parent, text, command, primary=False):
        return tk.Button(
            parent, text=text, command=command,
            bg=TERRA if primary else SURFACE,
            fg="white" if primary else TEXT_DARK,
            activebackground=TERRA_DEEP if primary else "#F0F3F1",
            activeforeground="white" if primary else TEXT_DARK,
            relief="flat", bd=0, padx=14, pady=8,
            font=self.font(9, "bold" if primary else "normal"),
            cursor="hand2", highlightthickness=1,
            highlightbackground=TERRA if primary else BORDER
        )

    def show_lookup_page(self):
        self.navigate_state("lookup")
        wrap = tk.Frame(self.page_host, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=16, pady=12)

        top = tk.Frame(wrap, bg=APP_BG)
        top.pack(fill="x", pady=(0, 8))
        tk.Label(top, text="품목 일괄 검색", bg=APP_BG, fg=TEXT,
                 font=self.font(13, "bold")).pack(side="left")

        search_panel = tk.Frame(wrap, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        search_panel.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(search_panel, bg=SURFACE)
        inner.pack(fill="x", padx=14, pady=12)

        self.query_var = tk.StringVar()
        self.query = tk.Entry(inner, textvariable=self.query_var, bg=SURFACE, fg=TEXT_DARK,
                              font=self.font(10), relief="flat", bd=0,
                              highlightthickness=1, highlightbackground=BORDER,
                              highlightcolor=SAGE)
        self.query.pack(side="left", fill="x", expand=True, ipady=8)
        self.query.insert(0, PLACEHOLDER)
        self.query.bind("<FocusIn>", self.clear_placeholder)
        self.query.bind("<KeyRelease>", self.on_query_change)
        self.query.bind("<Down>", lambda e: self.focus_suggestions())
        self.query.bind("<Return>", lambda e: self.choose_first_suggestion())

        self.make_button(inner, "선택 품목 조회", self.run_selected, primary=True).pack(side="left", padx=(10, 0))
        self.make_button(inner, "선택 초기화", self.clear_selected).pack(side="left", padx=(8, 0))

        self.suggest_frame = tk.Frame(search_panel, bg=SURFACE)
        self.suggest_frame.pack(fill="x", padx=14, pady=(0, 8))
        self.suggest_list = tk.Listbox(self.suggest_frame, height=0, bg=SURFACE, fg=TEXT_DARK,
                                       font=self.font(9), relief="flat", bd=0,
                                       highlightthickness=1, highlightbackground=BORDER,
                                       selectbackground="#E8F1EC", selectforeground=TEXT_DARK,
                                       activestyle="none")
        self.suggest_list.pack(fill="x")
        self.suggest_list.bind("<Double-Button-1>", lambda e: self.select_suggestion())
        self.suggest_list.bind("<Return>", lambda e: self.select_suggestion())
        self.suggest_list.pack_forget()

        selected_panel = tk.Frame(wrap, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        selected_panel.pack(fill="x", pady=(0, 10))
        tk.Label(selected_panel, text="선택 품목", bg=SURFACE, fg=TEXT_DARK,
                 font=self.font(9, "bold")).pack(side="left", padx=12, pady=9)
        self.selected_text = tk.Label(selected_panel, text="선택된 품목 없음", bg=SURFACE,
                                      fg=MUTED, font=self.font(9), anchor="w")
        self.selected_text.pack(side="left", fill="x", expand=True, padx=(0, 12))

        body = tk.Frame(wrap, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        body.pack(fill="both", expand=True)
        self.result_body = body
        self.render_empty()

    def navigate_state(self, active):
        for key, (lbl, line) in self.nav.items():
            is_active = key == active
            lbl.config(fg=TERRA if is_active else TEXT_DARK,
                       font=self.font(10, "bold" if is_active else "normal"))
            line.config(bg=TERRA if is_active else SURFACE)

    def clear_placeholder(self, _=None):
        if self.query.get() == PLACEHOLDER:
            self.query.delete(0, "end")

    def on_query_change(self, event=None):
        if event and event.keysym in ("Up", "Down", "Return", "Escape"):
            return
        q = self.query.get().strip()
        if not q or q == PLACEHOLDER:
            self.hide_suggestions()
            return
        items = self.idx.suggestions(q, 30)
        self.suggest_list.delete(0, "end")
        for p in items:
            self.suggest_list.insert("end", p)
        if items:
            self.suggest_list.config(height=min(8, len(items)))
            self.suggest_list.pack(fill="x")
        else:
            self.hide_suggestions()

    def hide_suggestions(self):
        if hasattr(self, "suggest_list"):
            self.suggest_list.pack_forget()

    def focus_suggestions(self):
        if self.suggest_list.winfo_ismapped() and self.suggest_list.size():
            self.suggest_list.focus_set()
            self.suggest_list.selection_clear(0, "end")
            self.suggest_list.selection_set(0)
            return "break"

    def choose_first_suggestion(self):
        if self.suggest_list.winfo_ismapped() and self.suggest_list.size():
            self.suggest_list.selection_clear(0, "end")
            self.suggest_list.selection_set(0)
            self.select_suggestion()
            return "break"
        return None

    def select_suggestion(self):
        sel = self.suggest_list.curselection()
        if not sel:
            return
        product = self.suggest_list.get(sel[0])
        if product not in self.selected_products:
            self.selected_products.append(product)
        self.refresh_selected_label()
        self.query.delete(0, "end")
        self.query.focus_set()
        self.hide_suggestions()

    def refresh_selected_label(self):
        if not self.selected_products:
            self.selected_text.config(text="선택된 품목 없음", fg=MUTED)
            return
        preview = " · ".join(self.selected_products[:5])
        if len(self.selected_products) > 5:
            preview += f"  외 {len(self.selected_products)-5}개"
        self.selected_text.config(text=preview, fg=TEXT_DARK)

    def clear_selected(self):
        self.selected_products.clear()
        self.results.clear()
        self.current_product = None
        self.refresh_selected_label()
        if hasattr(self, "result_body"):
            self.render_empty()

    def run_selected(self):
        if not self.selected_products:
            messagebox.showwarning("선택 품목 없음", "검색어를 입력한 뒤 자동으로 나타나는 후보에서 품목을 선택해 주세요.")
            return

        self.results = {}
        for p in self.selected_products:
            result = self.idx.search_product(p)

            # Collect any product codes already resolved from other DUR lists.
            known_codes = set()
            for hits in result.values():
                for h in hits:
                    code = sval(h.get("product_code"))
                    if code:
                        known_codes.add(code)

            # Age and telemedicine workbooks are small enough to verify directly on a DB miss.
            # This protects against stale/incomplete local SQLite indexes from older versions.
            for cat in ("age", "tele"):
                if result.get(cat):
                    continue
                source = self.cfg.get(cat)
                if isinstance(source, str) and Path(source).exists():
                    direct_hits = live_small_file_match(cat, source, p, known_codes)
                    if direct_hits:
                        result[cat] = direct_hits
                        for h in direct_hits:
                            code = sval(h.get("product_code"))
                            if code:
                                known_codes.add(code)

            self.results[p] = result

        self.current_product = self.selected_products[0]
        self.render_results()

    def render_empty(self):
        for w in self.result_body.winfo_children():
            w.destroy()
        tk.Label(self.result_body, text="검색어를 입력하면 후보 품목이 검색창 바로 아래에 자동으로 표시됩니다.",
                 bg=SURFACE, fg=MUTED, font=self.font(11)).pack(expand=True)

    def card_state(self, product, key):
        r = self.results.get(product, {})
        if key == "combo":
            n = len(r.get("combo_paid", [])) + len(r.get("combo_unpaid", []))
            return ("해당", f"{n}건", n > 0)
        hits = r.get(key, [])
        if not hits:
            return ("해당 없음", "", False)
        d = hits[0]["data"]
        if key == "preg":
            grade = next((sval(v) for k, v in d.items() if "등급" in k and sval(v)), "")
            return ("해당", f"{grade}등급" if grade else "해당", True)
        if key == "age":
            age = next((sval(v) for k, v in d.items() if "특정연령" in k and "단위" not in k and sval(v)), "")
            unit = next((sval(v) for k, v in d.items() if "연령단위" in k and "코드" not in k and sval(v)), "")
            return ("해당", (age + unit).strip() if age else "해당", True)
        return ("해당", f"{len(hits)}건" if len(hits) > 1 else "해당", True)

    def render_results(self):
        for w in self.result_body.winfo_children():
            w.destroy()

        left = tk.Frame(self.result_body, bg=SURFACE, width=300)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        right = tk.Frame(self.result_body, bg=SURFACE)
        right.pack(side="left", fill="both", expand=True)

        tk.Label(left, text=f"검색 품목 ({len(self.selected_products)})", bg=SURFACE,
                 fg=TEXT_DARK, font=self.font(10, "bold")).pack(anchor="w", padx=14, pady=(14, 8))

        for p in self.selected_products:
            row = tk.Frame(left, bg="#EFF7F2" if p == self.current_product else SURFACE,
                           cursor="hand2", height=52)
            row.pack(fill="x")
            row.pack_propagate(False)
            tk.Label(row, text="●" if p == self.current_product else "○",
                     bg=row["bg"], fg=SAGE_DEEP if p == self.current_product else "#AAB1AE",
                     font=self.font(9)).pack(side="left", padx=(12, 8))
            name = tk.Label(row, text=p, bg=row["bg"], fg=TEXT_DARK,
                            font=self.font(9, "bold" if p == self.current_product else "normal"),
                            anchor="w")
            name.pack(side="left", fill="x", expand=True)
            for w in (row, name):
                w.bind("<Button-1>", lambda e, prod=p: self.pick_product(prod))

        p = self.current_product
        head = tk.Frame(right, bg=SURFACE)
        head.pack(fill="x", padx=18, pady=(16, 8))
        tk.Label(head, text=p, bg=SURFACE, fg=TEXT, font=self.font(16, "bold")).pack(anchor="w")

        cards = tk.Frame(right, bg=SURFACE)
        cards.pack(fill="x", padx=18)
        for i, (key, label) in enumerate(DUR_CARDS):
            state, sub, active = self.card_state(p, key)
            bg = SAGE_SOFT if active else "#FAFBFA"
            border = SAGE if active else BORDER
            c = tk.Frame(cards, bg=bg, highlightbackground=border, highlightthickness=1,
                         width=205, height=82, cursor="hand2")
            c.grid(row=i//4, column=i%4, sticky="nsew", padx=4, pady=4)
            c.grid_propagate(False)
            tk.Label(c, text=label, bg=bg, fg=TEXT_DARK, font=self.font(9, "bold")).pack(anchor="w", padx=12, pady=(10, 4))
            text = f"● {sub or state}" if active else "○ 해당 없음"
            color = TERRA_DEEP if key == "preg" and active else (SAGE_DEEP if active else MUTED)
            tk.Label(c, text=text, bg=bg, fg=color,
                     font=self.font(11, "bold" if active else "normal")).pack(anchor="w", padx=12)
            for w in (c,):
                w.bind("<Button-1>", lambda e, k=key: self.render_detail(k))
        for col in range(4):
            cards.columnconfigure(col, weight=1)

        self.detail_host = tk.Frame(right, bg=SURFACE)
        self.detail_host.pack(fill="both", expand=True, padx=18, pady=(10, 16))
        self.render_detail("combo")

    def pick_product(self, p):
        self.current_product = p
        self.render_results()

    def render_detail(self, key):
        for w in self.detail_host.winfo_children():
            w.destroy()

        p = self.current_product
        r = self.results.get(p, {})
        hits = (r.get("combo_paid", []) + r.get("combo_unpaid", [])) if key == "combo" else r.get(key, [])

        title = dict(DUR_CARDS)[key]
        tk.Label(self.detail_host, text=f"{title} 상세정보", bg=SURFACE, fg=TEXT_DARK,
                 font=self.font(10, "bold")).pack(anchor="w", pady=(2, 8))

        if not hits:
            tk.Label(self.detail_host, text="해당 없음", bg=SURFACE, fg=MUTED,
                     font=self.font(10)).pack(anchor="w")
            return

        if key == "combo":
            cols = ("상대 품목", "상대 성분", "급여구분", "고시번호", "고시일자")
            tree = ttk.Treeview(self.detail_host, columns=cols, show="headings",
                                style="Compact.Treeview", height=8)
            for c in cols:
                tree.heading(c, text=c)
            tree.column("상대 품목", width=320)
            tree.column("상대 성분", width=220)
            tree.column("급여구분", width=90, anchor="center")
            tree.column("고시번호", width=120, anchor="center")
            tree.column("고시일자", width=120, anchor="center")
            rows_data = []
            for h in hits:
                d = h["data"]
                side = h["side"]
                if side == "A":
                    other_p, other_i, pay = d.get("제품명B",""), d.get("성분명B",""), d.get("급여여부A","")
                else:
                    other_p, other_i, pay = d.get("제품명A",""), d.get("성분명A",""), d.get("급여여부A","")
                vals = (other_p, other_i, pay, d.get("고시번호",""), d.get("고시일자",""))
                tree.insert("", "end", values=vals)
                rows_data.append(d)
            tree.pack(fill="x")

            detail = tk.Text(self.detail_host, height=5, bg=GRAY_SOFT, fg=TEXT_DARK,
                             font=self.font(9), relief="flat", bd=0, padx=10, pady=8,
                             wrap="word")
            detail.pack(fill="x", pady=(8,0))
            def show_sel(_=None):
                sel = tree.selection()
                if not sel:
                    return
                idx = tree.index(sel[0])
                d = rows_data[idx]
                detail.config(state="normal")
                detail.delete("1.0","end")
                detail.insert("1.0", d.get("상세정보","") or "상세정보 없음")
                detail.config(state="disabled")
            tree.bind("<<TreeviewSelect>>", show_sel)
            if tree.get_children():
                tree.selection_set(tree.get_children()[0])
                show_sel()
            return

        data = hits[0]["data"]
        hidden = ("성분코드","제품코드","약품코드","일반명코드","주성분코드","업체명","원본","시트")
        box = tk.Frame(self.detail_host, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        box.pack(fill="both", expand=True)
        rr = 0
        for k, v in data.items():
            if not sval(v) or any(h in str(k) for h in hidden):
                continue
            tk.Label(box, text=str(k), bg=SURFACE, fg=MUTED,
                     font=self.font(9), anchor="w").grid(row=rr, column=0, sticky="nw", padx=12, pady=6)
            tk.Label(box, text=sval(v), bg=SURFACE, fg=TEXT_DARK,
                     font=self.font(9), anchor="w", justify="left",
                     wraplength=760).grid(row=rr, column=1, sticky="nw", padx=(8,12), pady=6)
            rr += 1
        box.columnconfigure(1, weight=1)

    def show_files_page(self):
        self.navigate_state("files")
        wrap = tk.Frame(self.page_host, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=20, pady=16)

        head = tk.Frame(wrap, bg=APP_BG)
        head.pack(fill="x", pady=(0,10))
        tk.Label(head, text="기준파일 관리", bg=APP_BG, fg=TEXT,
                 font=self.font(14, "bold")).pack(side="left")
        self.make_button(head, "8개 파일 일괄 등록", self.choose_files, primary=True).pack(side="right")
        self.make_button(head, "다시 인덱싱", self.reindex_all).pack(side="right", padx=(0,8))

        panel = tk.Frame(wrap, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="both", expand=True)
        for col, label in enumerate(("파일 종류","상태","마지막 등록","등록 파일")):
            tk.Label(panel, text=label, bg="#F8FAF8", fg="#52626D",
                     font=self.font(9,"bold"), anchor="w").grid(row=0,column=col,sticky="nsew",padx=12,pady=9)
        panel.columnconfigure(3, weight=1)
        self.file_rows = {}
        for i,(key,label) in enumerate(CATEGORIES,1):
            bg = SURFACE if i%2 else "#FBFCFB"
            a=tk.Label(panel,text=label,bg=bg,fg=TEXT_DARK,font=self.font(9),anchor="w")
            b=tk.Label(panel,text="미등록",bg=bg,fg=MUTED,font=self.font(9),anchor="w")
            c=tk.Label(panel,text="—",bg=bg,fg=MUTED,font=self.font(9),anchor="w")
            d=tk.Label(panel,text="—",bg=bg,fg=MUTED,font=self.font(9),anchor="w")
            for j,w in enumerate((a,b,c,d)):
                w.grid(row=i,column=j,sticky="nsew",padx=12,pady=9)
            self.file_rows[key]=(b,c,d)
        self.refresh_status()

    def classify_file(self, path):
        name = Path(path).name
        n = norm(name)
        if "병용금기" in name and "비급여" in name:
            return "combo_unpaid"
        if "병용금기" in name and ("급여" in name and "비급여" not in name):
            return "combo_paid"
        if "연령금기" in name:
            return "age"
        if "임부금기" in name:
            return "preg"
        if "수유부주의" in name:
            return "lact"
        if "효능군중복" in name:
            return "dup"
        if "비대면진료" in name:
            return "tele"
        if "비용효과" in name:
            return "cost"
        return None

    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="DUR 기준 Excel 파일 선택",
            filetypes=[("Excel files","*.xlsx *.xlsm")]
        )
        if not paths:
            return
        mapped = {}
        unknown = []
        for p in paths:
            cat = self.classify_file(p)
            if cat:
                mapped[cat]=Path(p)
            else:
                unknown.append(Path(p).name)
        if not mapped:
            messagebox.showwarning("분류 실패","DUR 파일 종류를 판별하지 못했습니다.")
            return

        files_dir = ROOT_DIR / "files"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        items = []
        self.cfg.setdefault("_uploaded_at", {})
        for cat,p in mapped.items():
            dest = files_dir / f"{cat}__{p.name}"
            shutil.copy2(p,dest)
            self.cfg[cat]=str(dest)
            self.cfg["_uploaded_at"][cat]=datetime.now().isoformat(timespec="minutes")
            items.append((cat,dest))
        self.save_cfg()
        self.index_many(items)
        if unknown:
            messagebox.showwarning("일부 파일 미분류","\n".join(unknown))

    def index_many(self, items):
        win = tk.Toplevel(self)
        win.title("인덱싱")
        win.geometry("520x180")
        win.configure(bg=SURFACE)
        win.transient(self)
        label = tk.Label(win,text="인덱싱 준비 중…",bg=SURFACE,fg=TEXT_DARK,font=self.font(10))
        label.pack(pady=(34,10))
        bar = ttk.Progressbar(win,mode="determinate",maximum=max(1,len(items)),length=430)
        bar.pack(pady=10)

        def worker():
            errors=[]
            for i,(cat,path) in enumerate(items,1):
                def prog(msg, cat=cat):
                    self.after(0, lambda: label.config(text=f"{dict(CATEGORIES)[cat]} · {msg}"))
                ok,msg=self.idx.index_file(cat,path,prog)
                if not ok:
                    errors.append(f"{dict(CATEGORIES)[cat]}: {msg}")
                self.after(0, lambda i=i: bar.config(value=i))
            def done():
                win.destroy()
                self.refresh_status()
                if hasattr(self,"file_rows"):
                    self.refresh_status()
                if errors:
                    messagebox.showwarning("일부 파일 확인 필요","\n".join(errors))
                else:
                    messagebox.showinfo("완료","기준파일 인덱싱이 완료되었습니다.")
            self.after(0,done)
        threading.Thread(target=worker,daemon=True).start()

    def reindex_all(self):
        items=[]
        for k,_ in CATEGORIES:
            v=self.cfg.get(k)
            if isinstance(v,str) and Path(v).exists():
                items.append((k,Path(v)))
        if not items:
            messagebox.showwarning("기준파일 없음","먼저 Excel 파일을 등록해 주세요.")
            return
        self.index_many(items)

    def refresh_status(self):
        try:
            sts=self.idx.statuses()
        except Exception:
            sts={}
        times=self.cfg.get("_uploaded_at",{})
        latest=max(times.values()) if times else ""
        self.header_update.config(text=("● "+latest.replace("T"," ")[:16]) if latest else "● —")
        if hasattr(self,"file_rows"):
            for key,_ in CATEGORIES:
                status_lbl,date_lbl,file_lbl=self.file_rows[key]
                v=sts.get(key)
                stamp=times.get(key)
                date_lbl.config(text=stamp.replace("T"," ")[:16] if stamp else "—")
                path=self.cfg.get(key)
                file_lbl.config(text=Path(path).name.split("__",1)[-1] if path else "—")
                if v and v[0]:
                    status_lbl.config(text=f"● 정상 · {v[4]}시트",fg=SAGE_DEEP)
                elif v:
                    status_lbl.config(text="● 오류",fg=ERROR)
                else:
                    status_lbl.config(text="● 미등록",fg=MUTED)

    def auto_reindex_v151(self):
        # v1.5.2 uses a new migration key so an old v1.5.1 completion flag cannot suppress repair.
        migration_key = "_index_migration_v152_age_tele"
        if self.cfg.get(migration_key):
            return

        items=[]
        for cat in ("age","tele"):
            p=self.cfg.get(cat)
            if isinstance(p,str) and Path(p).exists():
                items.append((cat,Path(p)))

        if not items:
            return

        def worker():
            all_ok = True
            for cat,path in items:
                ok, _ = self.idx.index_file(cat,path)
                all_ok = all_ok and ok
            if all_ok:
                self.cfg[migration_key] = datetime.now().isoformat(timespec="seconds")
                self.save_cfg()
            self.after(0,self.refresh_status)

        threading.Thread(target=worker,daemon=True).start()


if __name__ == "__main__":
    try:
        app = App()
        if app.winfo_exists():
            app.mainloop()
    except Exception:
        err=traceback.format_exc()
        try:
            messagebox.showerror("DUR 프로그램 오류",err)
        except Exception:
            print(err)
