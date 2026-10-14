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
from tkinter import font as tkfont

from openpyxl import load_workbook

APP_NAME = "DUR 품목리스트 조회 프로그램"
VERSION = "1.6"

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

APP_BG = "#F6F8FB"
SURFACE = "#FFFFFF"
TEXT = "#14233B"
TEXT_DARK = "#26354D"
MUTED = "#718096"
BORDER = "#E1E7EF"
SAGE = "#14A58C"
SAGE_DEEP = "#0F816E"
SAGE_SOFT = "#EAF8F5"
TERRA = "#2F80ED"
TERRA_DEEP = "#1769C2"
TERRA_SOFT = "#EAF3FF"
GRAY_SOFT = "#F8FAFC"
ORANGE = "#F59E0B"
ERROR = "#E14B4B"
BLUE = "#1976E9"
BLUE_SOFT = "#EAF4FF"
RED = "#E64C4C"
RED_SOFT = "#FFF0F0"
ORANGE_SOFT = "#FFF6E7"
PURPLE = "#7B61D1"
PURPLE_SOFT = "#F4F0FF"
GREEN = "#0F9D78"
GREEN_SOFT = "#ECFAF4"

PLACEHOLDER = "품목명의 일부를 입력하세요"
DETAIL_FONT = 12  # 상세정보 본문 글자 크기


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
    s = s.replace("그람", "그램")  # 일부 목록의 구표기(밀리그람) 통일
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
                c.execute("CREATE TABLE IF NOT EXISTS names(category TEXT, product TEXT, product_norm TEXT)")
                c.execute("CREATE INDEX IF NOT EXISTS ix_names_cat ON names(category)")
                if c.execute("SELECT 1 FROM names LIMIT 1").fetchone() is None \
                        and c.execute("SELECT 1 FROM records LIMIT 1").fetchone() is not None:
                    c.execute("INSERT INTO names SELECT DISTINCT category,product,product_norm FROM records")
                c.execute(
                    "CREATE TABLE IF NOT EXISTS status("
                    "category TEXT PRIMARY KEY, ok INTEGER, message TEXT, filename TEXT, "
                    "indexed_at TEXT, sheet_count INTEGER, row_count INTEGER)"
                )
        return self.with_retry(_work)

    def clear_category(self, c, category):
        c.execute("DELETE FROM records WHERE category=?", (category,))
        c.execute("DELETE FROM names WHERE category=?", (category,))
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
                        "INSERT INTO names SELECT DISTINCT category,product,product_norm FROM records WHERE category=?",
                        (category,),
                    )
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

    def migrate_norm_v161(self):
        """그람→그램 표기 통일을 기존 DB에 직접 반영하고 후보 테이블을 다시 만든다(엑셀 재인덱싱 불필요)."""
        self.init()
        def _work():
            with self.lock, self.conn() as c:
                for col in ("product_norm", "core_norm"):
                    c.execute(
                        f"UPDATE records SET {col}=replace(replace({col},'밀리그람','mg'),'그람','그램') "
                        f"WHERE {col} LIKE '%그람%'"
                    )
                c.execute("DELETE FROM names")
                c.execute("INSERT INTO names SELECT DISTINCT category,product,product_norm FROM records")
        return self.with_retry(_work)

    def delete_category(self, category):
        """Remove one category from the local search index."""
        self.init()
        def _work():
            with self.lock, self.conn() as c:
                self.clear_category(c, category)
        return self.with_retry(_work)

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
                    "SELECT product, MIN(CASE WHEN product_norm LIKE ? THEN 0 ELSE 1 END) rk FROM names "
                    "WHERE product_norm LIKE ? "
                    "GROUP BY product "
                    "ORDER BY rk, length(product), product "
                    "LIMIT ?",
                    (f"{nq}%", f"%{nq}%", limit),
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
        self.ui_font = self.resolve_font_family()
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
            bg=APP_BG, fg="#8A948F", font=self.font_en(9)
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

    def resolve_font_family(self):
        """Use Pretendard when installed, with safe Windows fallbacks."""
        try:
            fams = set(tkfont.families(self))
            for name in ("Pretendard", "Pretendard Variable", "맑은 고딕", "Malgun Gothic", "Arial"):
                if name in fams:
                    return name
        except Exception:
            pass
        return "Malgun Gothic"

    def font(self, size=11, weight="normal"):
        return (self.ui_font, size, weight)

    def font_en(self, size=10, weight="normal"):
        return ("Arial", size, weight)

    def setup_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Compact.Treeview",
            font=self.font(11),
            rowheight=36,
            background=SURFACE,
            fieldbackground=SURFACE,
            foreground=TEXT_DARK,
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "Compact.Treeview",
            background=[("selected", "#E8F3FF")],
            foreground=[("selected", TEXT_DARK)],
        )
        style.configure(
            "Compact.Treeview.Heading",
            font=self.font(10, "bold"),
            background="#F3F6FA",
            foreground="#42526B",
            relief="flat",
            borderwidth=0,
            padding=(8, 10),
        )
        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor="#EDF1F6",
            background=BLUE,
            bordercolor="#EDF1F6",
            lightcolor=BLUE,
            darkcolor=BLUE,
        )

    def build_header(self):
        h = tk.Frame(self, bg=SURFACE, height=82, highlightbackground=BORDER, highlightthickness=1)
        h.pack(fill="x")
        h.pack_propagate(False)

        brand = tk.Frame(h, bg=SURFACE)
        brand.pack(side="left", padx=(28, 34), pady=14)
        icon = tk.Label(
            brand, text="◆", bg=SURFACE, fg=BLUE,
            font=("Segoe UI Symbol", 26, "bold")
        )
        icon.pack(side="left", padx=(0, 12))
        tk.Label(
            brand, text=APP_NAME, bg=SURFACE, fg=TEXT,
            font=self.font(19, "bold")
        ).pack(side="left")

        nav = tk.Frame(h, bg="#F8FAFC", highlightbackground=BORDER, highlightthickness=1)
        nav.pack(side="left", pady=14)
        self.nav = {}
        for key, label, symbol in (
            ("lookup", "DUR 조회", "⌕"),
            ("files", "기준파일 관리", "▤"),
        ):
            box = tk.Frame(nav, bg="#F8FAFC", width=156, height=52, cursor="hand2")
            box.pack(side="left")
            box.pack_propagate(False)
            inner = tk.Frame(box, bg="#F8FAFC")
            inner.pack(expand=True)
            s = tk.Label(inner, text=symbol, bg="#F8FAFC", fg="#65758B",
                         font=("Segoe UI Symbol", 16))
            s.pack(side="left", padx=(0, 8))
            lbl = tk.Label(inner, text=label, bg="#F8FAFC", fg=TEXT_DARK,
                           font=self.font(11, "bold"))
            lbl.pack(side="left")
            line = tk.Frame(box, bg="#F8FAFC", height=3)
            line.pack(side="bottom", fill="x")
            for w in (box, inner, s, lbl):
                w.bind("<Button-1>", lambda e, k=key: self.navigate(k))
            self.nav[key] = (lbl, line)

        right = tk.Frame(h, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        right.pack(side="right", padx=26, pady=12)
        tk.Label(
            right, text="●  기준파일 최종 업데이트",
            bg=SURFACE, fg=SAGE_DEEP, font=self.font(9, "bold")
        ).pack(anchor="w", padx=14, pady=(8, 1))
        self.header_update = tk.Label(
            right, text="—", bg=SURFACE, fg=TEXT_DARK, font=self.font(11, "bold")
        )
        self.header_update.pack(anchor="w", padx=14, pady=(0, 8))

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

    def make_button(self, parent, text, command, primary=False, danger=False):
        if danger:
            bg, fg, active, border = SURFACE, ERROR, "#FFF1F1", "#F3B4B4"
        elif primary:
            bg, fg, active, border = BLUE, "white", TERRA_DEEP, BLUE
        else:
            bg, fg, active, border = SURFACE, TEXT_DARK, "#F0F4F8", BORDER
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active,
            activeforeground=fg,
            relief="flat",
            bd=0,
            padx=16,
            pady=10,
            font=self.font(10, "bold" if (primary or danger) else "normal"),
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=border,
        )

    def build_side_nav(self, parent, active):
        side = tk.Frame(parent, bg=SURFACE, width=184, highlightbackground=BORDER, highlightthickness=1)
        side.pack(side="left", fill="y", padx=(0, 12))
        side.pack_propagate(False)

        tk.Label(side, text="메뉴", bg=SURFACE, fg=MUTED, font=self.font(9, "bold")).pack(
            anchor="w", padx=18, pady=(20, 8)
        )

        for key, label, symbol in (
            ("lookup", "DUR 조회", "⌕"),
            ("files", "기준파일 관리", "▤"),
        ):
            is_active = key == active
            bg = BLUE_SOFT if is_active else SURFACE
            fg = BLUE if is_active else TEXT_DARK
            row = tk.Frame(side, bg=bg, height=54, cursor="hand2")
            row.pack(fill="x", padx=10, pady=3)
            row.pack_propagate(False)
            tk.Label(row, text=symbol, bg=bg, fg=fg, font=("Segoe UI Symbol", 16)).pack(
                side="left", padx=(12, 10)
            )
            lbl = tk.Label(row, text=label, bg=bg, fg=fg, font=self.font(11, "bold" if is_active else "normal"))
            lbl.pack(side="left")
            for w in (row, lbl):
                w.bind("<Button-1>", lambda e, k=key: self.navigate(k))

        tk.Frame(side, bg=BORDER, height=1).pack(fill="x", padx=14, pady=14)
        tk.Label(
            side, text="로컬 전용\\nExcel 원본은 PC 내부에서만 처리됩니다.",
            bg=SURFACE, fg=MUTED, justify="left", font=self.font(9)
        ).pack(anchor="w", padx=18)

        return side

    def show_lookup_page(self):
        self.navigate_state("lookup")
        wrap = tk.Frame(self.page_host, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=16, pady=14)

        self.build_side_nav(wrap, "lookup")

        work = tk.Frame(wrap, bg=APP_BG)
        work.pack(side="left", fill="both", expand=True)

        search_col = tk.Frame(work, bg=SURFACE, width=365, highlightbackground=BORDER, highlightthickness=1)
        search_col.pack(side="left", fill="y", padx=(0, 12))
        search_col.pack_propagate(False)

        tk.Label(
            search_col, text="품목 검색 (일괄 조회)", bg=SURFACE, fg=TEXT,
            font=self.font(14, "bold")
        ).pack(anchor="w", padx=18, pady=(20, 12))

        search_box = tk.Frame(search_col, bg=SURFACE)
        search_box.pack(fill="x", padx=16)
        self.query_var = tk.StringVar()
        self.query = tk.Entry(
            search_box, textvariable=self.query_var, bg=SURFACE, fg=TEXT_DARK,
            font=self.font(12), relief="flat", bd=0,
            highlightthickness=1, highlightbackground="#AFC8EE", highlightcolor=BLUE
        )
        self.query.pack(fill="x", ipady=11)
        self.query.insert(0, PLACEHOLDER)
        self.query.bind("<FocusIn>", self.clear_placeholder)
        self.query.bind("<KeyRelease>", self.on_query_change)
        self.query.bind("<Down>", lambda e: self.focus_suggestions())
        self.query.bind("<Return>", lambda e: self.choose_first_suggestion())

        self.suggest_frame = tk.Frame(search_col, bg=SURFACE)
        self.suggest_frame.pack(fill="x", padx=16, pady=(4, 0))
        self.suggest_list = tk.Listbox(
            self.suggest_frame, height=0, bg=SURFACE, fg=TEXT_DARK,
            font=self.font(10), relief="flat", bd=0,
            highlightthickness=1, highlightbackground=BORDER,
            selectbackground="#DCEBFF", selectforeground=BLUE,
            activestyle="none"
        )
        self.suggest_list.pack(fill="x")
        self.suggest_list.bind("<Double-Button-1>", lambda e: self.select_suggestion())
        self.suggest_list.bind("<Return>", lambda e: self.select_suggestion())
        self.suggest_list.pack_forget()

        sel_head = tk.Frame(search_col, bg=SURFACE)
        sel_head.pack(fill="x", padx=16, pady=(18, 8))
        tk.Label(sel_head, text="선택한 품목", bg=SURFACE, fg=TEXT,
                 font=self.font(13, "bold")).pack(side="left")
        self.selected_count = tk.Label(sel_head, text="(0)", bg=SURFACE, fg=MUTED,
                                       font=self.font(11, "bold"))
        self.selected_count.pack(side="left", padx=(4, 0))
        self.make_button(sel_head, "전체 삭제", self.clear_selected, danger=True).pack(side="right")

        self.selected_list_host = tk.Frame(search_col, bg=SURFACE)
        self.selected_list_host.pack(fill="both", expand=True, padx=16)

        actions = tk.Frame(search_col, bg=SURFACE)
        actions.pack(fill="x", padx=16, pady=16)
        self.make_button(
            actions, "⌕  선택 품목으로 조회", self.run_selected, primary=True
        ).pack(fill="x")
        self.make_button(
            actions, "선택 품목 모두 삭제", self.clear_selected
        ).pack(fill="x", pady=(8, 0))

        result = tk.Frame(work, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        result.pack(side="left", fill="both", expand=True)
        self.result_body = result

        self.refresh_selected_panel()
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
        job = getattr(self, "_sugg_job", None)
        if job:
            self.after_cancel(job)
        self._sugg_job = self.after(150, self.update_suggestions)

    def update_suggestions(self):
        self._sugg_job = None
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
        self.refresh_selected_panel()
        self.query.delete(0, "end")
        self.query.focus_set()
        self.hide_suggestions()

    def refresh_selected_panel(self):
        if not hasattr(self, "selected_list_host"):
            return
        for w in self.selected_list_host.winfo_children():
            w.destroy()
        self.selected_count.config(text=f"({len(self.selected_products)})")
        if not self.selected_products:
            tk.Label(
                self.selected_list_host,
                text="검색 결과에서 품목을 선택하면\n이곳에 누적됩니다.",
                bg=SURFACE, fg=MUTED, justify="center", font=self.font(10)
            ).pack(expand=True, pady=28)
            return

        for product in self.selected_products:
            active = product == self.current_product
            bg = "#EAF4FF" if active else "#F8FAFC"
            row = tk.Frame(
                self.selected_list_host, bg=bg,
                highlightbackground="#C8DCF7" if active else BORDER,
                highlightthickness=1, height=58, cursor="hand2"
            )
            row.pack(fill="x", pady=4)
            row.pack_propagate(False)
            name = tk.Label(
                row, text=product, bg=bg, fg=BLUE if active else TEXT_DARK,
                font=self.font(10, "bold" if active else "normal"),
                anchor="w"
            )
            name.pack(side="left", fill="x", expand=True, padx=(12, 6))
            close = tk.Label(row, text="×", bg=bg, fg=MUTED, font=self.font(14, "bold"), cursor="hand2")
            close.pack(side="right", padx=10)
            name.bind("<Button-1>", lambda e, p=product: self.pick_product(p))
            row.bind("<Button-1>", lambda e, p=product: self.pick_product(p))
            close.bind("<Button-1>", lambda e, p=product: self.remove_selected_product(p))

    def refresh_selected_label(self):
        self.refresh_selected_panel()

    def remove_selected_product(self, product):
        if product in self.selected_products:
            self.selected_products.remove(product)
        self.results.pop(product, None)
        if self.current_product == product:
            self.current_product = self.selected_products[0] if self.selected_products else None
        self.refresh_selected_panel()
        if not self.selected_products:
            self.render_empty()
        elif self.results and self.current_product in self.results:
            self.render_results()

    def clear_selected(self):
        self.selected_products.clear()
        self.results.clear()
        self.current_product = None
        self.refresh_selected_panel()
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
        box = tk.Frame(self.result_body, bg=SURFACE)
        box.pack(expand=True)
        tk.Label(box, text="⌕", bg=SURFACE, fg="#A8B7CB",
                 font=("Segoe UI Symbol", 34)).pack(pady=(0, 8))
        tk.Label(
            box, text="조회할 품목을 선택해 주세요",
            bg=SURFACE, fg=TEXT, font=self.font(16, "bold")
        ).pack()
        tk.Label(
            box,
            text="품목명의 일부를 입력하면 검색창 아래에 후보가 자동으로 표시됩니다.",
            bg=SURFACE, fg=MUTED, font=self.font(10)
        ).pack(pady=(6, 0))

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

        p = self.current_product
        if not p:
            self.render_empty()
            return

        self.refresh_selected_panel()

        right = tk.Frame(self.result_body, bg=SURFACE)
        right.pack(fill="both", expand=True, padx=16, pady=14)

        head = tk.Frame(right, bg=SURFACE)
        head.pack(fill="x", pady=(0, 10))
        tk.Label(head, text="◆", bg=SURFACE, fg=SAGE,
                 font=("Segoe UI Symbol", 20)).pack(side="left", padx=(0, 10))
        tk.Label(
            head, text=p, bg=SURFACE, fg=TEXT,
            font=self.font(18, "bold")
        ).pack(side="left")

        active_parts = []
        for key, label in DUR_CARDS:
            state, sub, active = self.card_state(p, key)
            if active:
                active_parts.append(f"{label} {sub or ''}".strip())
        tk.Label(
            right,
            text=(f"{len(active_parts)}개 DUR 해당  ·  " + "  ·  ".join(active_parts[:4]))
                 if active_parts else "7개 관리항목 모두 해당 없음",
            bg=SURFACE, fg=TEXT_DARK, font=self.font(11, "bold")
        ).pack(anchor="w", pady=(0, 12))

        cards = tk.Frame(right, bg=SURFACE)
        cards.pack(fill="x")

        palette = {
            "combo": (RED, RED_SOFT),
            "age": (ORANGE, ORANGE_SOFT),
            "preg": (RED, "#FFF2F4"),
            "lact": (PURPLE, PURPLE_SOFT),
            "dup": (GREEN, GREEN_SOFT),
            "tele": (BLUE, BLUE_SOFT),
            "cost": (SAGE_DEEP, SAGE_SOFT),
        }

        for i, (key, label) in enumerate(DUR_CARDS):
            state, sub, active = self.card_state(p, key)
            accent, soft = palette[key]
            bg = soft if active else "#FBFCFE"
            border = accent if active else BORDER
            c = tk.Frame(
                cards, bg=bg, highlightbackground=border, highlightthickness=1,
                height=106, cursor="hand2"
            )
            c.grid(row=i // 3, column=i % 3, sticky="nsew", padx=5, pady=5)
            c.grid_propagate(False)

            top = tk.Frame(c, bg=bg)
            top.pack(fill="x", padx=14, pady=(12, 5))
            tk.Label(top, text=label, bg=bg, fg=accent if active else TEXT_DARK,
                     font=self.font(11, "bold")).pack(side="left")
            tk.Label(top, text="›" if active else "–", bg=bg,
                     fg=accent if active else MUTED, font=self.font(16, "bold")).pack(side="right")

            main_text = sub or state if active else "해당 없음"
            tk.Label(
                c,
                text=("●  " + main_text) if active else "○  해당 없음",
                bg=bg, fg=accent if active else MUTED,
                font=self.font(13, "bold" if active else "normal")
            ).pack(anchor="w", padx=14, pady=(4, 0))

            c.bind("<Button-1>", lambda e, k=key: self.render_detail(k))

        for col in range(3):
            cards.columnconfigure(col, weight=1)

        self.detail_tabs = tk.Frame(right, bg=SURFACE)
        self.detail_tabs.pack(fill="x", pady=(12, 0))

        self.detail_host = tk.Frame(right, bg=SURFACE)
        self.detail_host.pack(fill="both", expand=True, pady=(0, 4))

        self.render_detail(next((k for k, _ in DUR_CARDS if self.card_state(p, k)[2]), "combo"))

    def pick_product(self, p):
        self.current_product = p
        self.refresh_selected_panel()
        if p in self.results:
            self.render_results()

    # ---- 상세정보 선택·복사 지원 ----
    def copy_to_clipboard(self, text, btn=None):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update_idletasks()
        if btn is not None:
            old = btn.cget("text")
            btn.config(text="복사됨 ✓")
            self.after(1200, lambda: btn.winfo_exists() and btn.config(text=old))

    def make_readonly_text(self, parent, height=8, bg=SURFACE):
        """드래그로 블럭 선택·Ctrl+C 복사는 되고 편집은 막은 Text. (holder, text) 반환."""
        holder = tk.Frame(parent, bg=bg, highlightbackground=BORDER, highlightthickness=1)
        t = tk.Text(
            holder, height=height, bg=bg, fg=TEXT_DARK, font=self.font(DETAIL_FONT),
            relief="flat", bd=0, padx=14, pady=10, wrap="word", cursor="xterm",
            selectbackground="#BBD8FF", selectforeground=TEXT_DARK,
            inactiveselectbackground="#BBD8FF", insertwidth=0, highlightthickness=0,
        )
        sb = ttk.Scrollbar(holder, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        t.pack(side="left", fill="both", expand=True)

        move_keys = {"Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next",
                     "Shift_L", "Shift_R", "Control_L", "Control_R"}

        def on_key(e):
            if e.keysym in move_keys:
                return None
            return "break"  # 입력·삭제·붙여넣기 차단

        def select_all(_=None):
            t.tag_add("sel", "1.0", "end-1c")
            return "break"

        def copy_sel(_=None):
            t.event_generate("<<Copy>>")
            return "break"

        t.bind("<Key>", on_key)
        for k in ("a", "A"):
            t.bind(f"<Control-{k}>", select_all)
        for k in ("c", "C"):
            t.bind(f"<Control-{k}>", copy_sel)
        t.bind("<<Paste>>", lambda e: "break")
        t.bind("<<Cut>>", lambda e: "break")
        t.bind("<Button-1>", lambda e: t.focus_set(), add="+")

        menu = tk.Menu(t, tearoff=0)
        menu.add_command(label="복사", command=lambda: t.event_generate("<<Copy>>"))
        menu.add_command(label="전체 선택", command=select_all)

        def popup(e):
            t.focus_set()
            try:
                menu.tk_popup(e.x_root, e.y_root)
            finally:
                menu.grab_release()

        t.bind("<Button-3>", popup)
        return holder, t

    def render_detail(self, key):
        for w in self.detail_host.winfo_children():
            w.destroy()

        p = self.current_product
        r = self.results.get(p, {})
        hits = (r.get("combo_paid", []) + r.get("combo_unpaid", [])) if key == "combo" else r.get(key, [])

        title = dict(DUR_CARDS)[key]

        if hasattr(self, "detail_tabs"):
            for w in self.detail_tabs.winfo_children():
                w.destroy()
            for tab_key, tab_label in DUR_CARDS:
                st, sub, active = self.card_state(p, tab_key)
                if not active:
                    continue
                selected = tab_key == key
                btn = tk.Button(
                    self.detail_tabs,
                    text=f"{tab_label} 상세" + (f" ({sub})" if sub and sub.endswith("건") else ""),
                    command=lambda k=tab_key: self.render_detail(k),
                    bg=SURFACE,
                    fg=(BLUE if selected else TEXT_DARK),
                    activebackground=SURFACE,
                    activeforeground=BLUE,
                    relief="flat",
                    bd=0,
                    font=self.font(10, "bold" if selected else "normal"),
                    padx=14,
                    pady=9,
                    cursor="hand2",
                    highlightthickness=0,
                )
                btn.pack(side="left")
            tk.Frame(self.detail_tabs, bg=BORDER, height=1).pack(fill="x", side="bottom")

        head = tk.Frame(self.detail_host, bg=SURFACE)
        head.pack(fill="x", pady=(12, 8))
        tk.Label(head, text=f"{title} 상세정보", bg=SURFACE, fg=TEXT_DARK,
                 font=self.font(13, "bold")).pack(side="left")

        def small_button(parent, text, command):
            b = self.make_button(parent, text, command)
            b.config(padx=12, pady=5)
            return b

        if not hits:
            tk.Label(self.detail_host, text="해당 없음", bg=SURFACE, fg=MUTED,
                     font=self.font(DETAIL_FONT)).pack(anchor="w")
            return

        if key == "combo":
            cols = ("상대 품목", "상대 성분", "급여구분", "고시번호", "고시일자")
            pw = ttk.PanedWindow(self.detail_host, orient="vertical")
            pw.pack(fill="both", expand=True)
            top = tk.Frame(pw, bg=SURFACE)
            bottom = tk.Frame(pw, bg=SURFACE)
            pw.add(top, weight=3)
            pw.add(bottom, weight=2)

            tree = ttk.Treeview(top, columns=cols, show="headings",
                                style="Compact.Treeview", height=4, selectmode="extended")
            tsb = ttk.Scrollbar(top, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=tsb.set)
            for c in cols:
                tree.heading(c, text=c, command=lambda c=c: sort_by(c))
            tree.column("상대 품목", width=300, minwidth=180)
            tree.column("상대 성분", width=200, minwidth=120)
            tree.column("급여구분", width=80, minwidth=70, anchor="center")
            tree.column("고시번호", width=110, minwidth=90, anchor="center")
            tree.column("고시일자", width=110, minwidth=90, anchor="center")
            rows_data = []
            for h in hits:
                d = h["data"]
                side = h["side"]
                if side == "A":
                    other_p, other_i, pay = d.get("제품명B",""), d.get("성분명B",""), d.get("급여여부A","")
                else:
                    other_p, other_i, pay = d.get("제품명A",""), d.get("성분명A",""), d.get("급여여부A","")
                vals = (other_p, other_i, pay, d.get("고시번호",""), d.get("고시일자",""))
                tree.insert("", "end", iid=str(len(rows_data)), values=vals)
                rows_data.append(d)
            tsb.pack(side="right", fill="y")
            tree.pack(side="left", fill="both", expand=True)

            # 열 제목 클릭: 오름차순 → 내림차순 → 원래 순서
            sort_state = {"col": None, "dir": 0}

            def sort_key(v):
                v = str(v).strip()
                if not v:
                    return None  # 빈 값은 항상 맨 아래
                return (0, int(v)) if v.isdigit() else (1, v.casefold())

            def sort_by(col):
                if sort_state["col"] == col:
                    sort_state["dir"] = {1: -1, -1: 0, 0: 1}[sort_state["dir"]]
                else:
                    sort_state["col"], sort_state["dir"] = col, 1
                direction = sort_state["dir"]
                base = sorted(tree.get_children(), key=int)  # 동률일 때 원래 순서 유지
                if direction == 0:
                    ordered = base
                else:
                    keyed = [(sort_key(tree.set(i, col)), i) for i in base]
                    full = sorted([x for x in keyed if x[0] is not None],
                                  key=lambda x: x[0], reverse=(direction == -1))
                    ordered = [i for _, i in full] + [i for k, i in keyed if k is None]
                for pos, i in enumerate(ordered):
                    tree.move(i, "", pos)
                for c in cols:
                    mark = (" ▲" if direction == 1 else " ▼") if (c == col and direction != 0) else ""
                    tree.heading(c, text=c + mark)
                sel = tree.selection()
                if sel:
                    tree.see(sel[0])

            def rows_text(items, with_header=False):
                lines = ["\t".join(cols)] if with_header else []
                for i in items:
                    lines.append("\t".join(str(x) for x in tree.item(i, "values")))
                return "\n".join(lines)

            def copy_selected(_=None):
                items = sorted(tree.selection(), key=tree.index) or list(tree.get_children())
                self.copy_to_clipboard(rows_text(items))
                return "break"

            def copy_all(btn):
                self.copy_to_clipboard(rows_text(tree.get_children(), with_header=True), btn)

            all_btn = small_button(head, "표 전체 복사", lambda: copy_all(all_btn))
            all_btn.pack(side="right")
            tk.Label(head, text="  열 제목을 클릭하면 정렬됩니다", bg=SURFACE, fg=MUTED,
                     font=self.font(10)).pack(side="left")
            tree.bind("<Control-c>", copy_selected)
            tree.bind("<Control-C>", copy_selected)
            tree.bind("<Control-a>", lambda e: (tree.selection_set(tree.get_children()), "break")[1])
            tmenu = tk.Menu(tree, tearoff=0)
            tmenu.add_command(label="선택 행 복사", command=copy_selected)
            tmenu.add_command(label="전체 행 복사 (머리글 포함)",
                              command=lambda: self.copy_to_clipboard(rows_text(tree.get_children(), True)))

            def tpopup(e):
                row = tree.identify_row(e.y)
                if row and row not in tree.selection():
                    tree.selection_set(row)
                try:
                    tmenu.tk_popup(e.x_root, e.y_root)
                finally:
                    tmenu.grab_release()

            tree.bind("<Button-3>", tpopup)

            sub_head = tk.Frame(bottom, bg=SURFACE)
            sub_head.pack(fill="x", pady=(10, 4))
            tk.Label(sub_head, text="상세내용 (드래그해서 선택 · Ctrl+C로 복사 · 위 표와의 경계선을 끌어 크기 조절)",
                     bg=SURFACE, fg=MUTED, font=self.font(10)).pack(side="left")

            holder, detail = self.make_readonly_text(bottom, height=4, bg=GRAY_SOFT)
            holder.pack(fill="both", expand=True)
            cp_btn = small_button(sub_head, "상세내용 복사", lambda: self.copy_to_clipboard(
                detail.get("1.0", "end-1c"), cp_btn))
            cp_btn.pack(side="right")

            def show_sel(_=None):
                sel = sorted(tree.selection(), key=tree.index)
                if not sel:
                    return
                d = rows_data[int(sel[0])]
                detail.delete("1.0", "end")
                detail.insert("1.0", d.get("상세정보", "") or "상세정보 없음")
            tree.bind("<<TreeviewSelect>>", show_sel)
            if tree.get_children():
                tree.selection_set(tree.get_children()[0])
                show_sel()
            return

        data = hits[0]["data"]
        hidden = ("성분코드","제품코드","약품코드","일반명코드","주성분코드","업체명","원본","시트")
        fields = []
        for k, v in data.items():
            kk = re.sub(r"\s+", "", str(k))  # 머리글 안의 줄바꿈/공백 제거 (예: "일반명\n코드")
            if sval(v) and not any(h in kk for h in hidden):
                fields.append((kk, sval(v)))

        holder, txt = self.make_readonly_text(self.detail_host, height=8, bg=SURFACE)
        holder.pack(fill="both", expand=True)
        key_font = tkfont.Font(font=self.font(DETAIL_FONT, "bold"))
        tab = (max((key_font.measure(k) for k, _ in fields), default=80)) + 40
        txt.configure(tabs=(tab,))
        txt.tag_configure("row", lmargin1=0, lmargin2=tab, spacing1=6, spacing3=6)
        txt.tag_configure("k", font=self.font(DETAIL_FONT, "bold"), foreground=MUTED)
        for k, v in fields:
            txt.insert("end", k, ("row", "k"))
            txt.insert("end", "\t", ("row",))
            txt.insert("end", v + "\n", ("row",))
        txt.delete("end-2c", "end-1c") if fields else None  # 마지막 빈 줄 제거

        copy_btn = small_button(head, "내용 전체 복사", lambda: self.copy_to_clipboard(
            "\n".join(f"{k}: {v}" for k, v in fields), copy_btn))
        copy_btn.pack(side="right")

    def show_files_page(self):
        self.navigate_state("files")
        wrap = tk.Frame(self.page_host, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=16, pady=14)

        self.build_side_nav(wrap, "files")

        main = tk.Frame(wrap, bg=APP_BG)
        main.pack(side="left", fill="both", expand=True)

        head = tk.Frame(main, bg=APP_BG)
        head.pack(fill="x", pady=(0, 12))
        title_box = tk.Frame(head, bg=APP_BG)
        title_box.pack(side="left")
        tk.Label(title_box, text="기준파일 관리", bg=APP_BG, fg=TEXT,
                 font=self.font(18, "bold")).pack(anchor="w")
        tk.Label(
            title_box,
            text="새 파일을 등록하거나 기존 인덱스를 완전히 삭제할 수 있습니다.",
            bg=APP_BG, fg=MUTED, font=self.font(10)
        ).pack(anchor="w", pady=(4, 0))

        self.make_button(head, "전체 기준파일 삭제", self.delete_all_files, danger=True).pack(side="right")
        self.make_button(head, "파일 일괄 등록", self.choose_files, primary=True).pack(side="right", padx=(0, 8))
        self.make_button(head, "다시 인덱싱", self.reindex_all).pack(side="right", padx=(0, 8))

        panel = tk.Frame(main, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="both", expand=True)

        headers = ("파일 종류", "상태", "마지막 등록", "등록 파일", "")
        widths = (22, 15, 20, 1, 12)
        for col, label in enumerate(headers):
            tk.Label(
                panel, text=label, bg="#F4F7FB", fg="#52626D",
                font=self.font(10, "bold"), anchor="w"
            ).grid(row=0, column=col, sticky="nsew", padx=14, pady=12)
        panel.columnconfigure(3, weight=1)

        self.file_rows = {}
        for i, (key, label) in enumerate(CATEGORIES, 1):
            bg = SURFACE if i % 2 else "#FBFCFE"
            a = tk.Label(panel, text=label, bg=bg, fg=TEXT_DARK, font=self.font(11, "bold"), anchor="w")
            b = tk.Label(panel, text="미등록", bg=bg, fg=MUTED, font=self.font(10), anchor="w")
            c = tk.Label(panel, text="—", bg=bg, fg=MUTED, font=self.font(10), anchor="w")
            d = tk.Label(panel, text="—", bg=bg, fg=MUTED, font=self.font(10), anchor="w")
            e = self.make_button(panel, "삭제", lambda k=key: self.delete_category_file(k), danger=True)

            for j, w in enumerate((a, b, c, d)):
                w.grid(row=i, column=j, sticky="nsew", padx=14, pady=12)
            e.grid(row=i, column=4, sticky="e", padx=12, pady=7)
            self.file_rows[key] = (b, c, d, e)

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
            old_path = self.cfg.get(cat)
            dest = files_dir / f"{cat}__{p.name}"
            shutil.copy2(p,dest)
            if old_path and old_path != str(dest):
                try:
                    old = Path(old_path)
                    if old.exists() and old.parent == files_dir:
                        old.unlink()
                except Exception:
                    pass
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

    def delete_category_file(self, category, ask=True):
        label = dict(CATEGORIES)[category]
        path = self.cfg.get(category)

        if ask and not messagebox.askyesno(
            "기준파일 삭제",
            f"'{label}' 기준파일과 검색 인덱스를 완전히 삭제할까요?\n\n"
            "삭제 후 다른 Excel 파일을 다시 등록할 수 있습니다."
        ):
            return

        try:
            self.idx.delete_category(category)
        except Exception as e:
            messagebox.showerror("인덱스 삭제 실패", str(e))
            return

        if isinstance(path, str):
            try:
                p = Path(path)
                if p.exists() and p.parent == (ROOT_DIR / "files"):
                    p.unlink()
            except Exception:
                pass

        self.cfg.pop(category, None)
        times = self.cfg.get("_uploaded_at", {})
        times.pop(category, None)

        if category in ("age", "tele"):
            self.cfg.pop("_index_migration_v161_norm_all", None)
            self.cfg.pop("_v151_auto_reindex_done", None)

        self.save_cfg()
        self.refresh_status()

    def delete_all_files(self):
        existing = [k for k, _ in CATEGORIES if self.cfg.get(k)]
        if not existing:
            messagebox.showinfo("삭제할 파일 없음", "등록된 기준파일이 없습니다.")
            return
        if not messagebox.askyesno(
            "전체 기준파일 삭제",
            "등록된 모든 기준파일과 모든 검색 인덱스를 완전히 삭제할까요?\n\n"
            "원본으로 등록했던 외부 Excel 파일은 삭제하지 않고, 프로그램 내부 복사본과 인덱스만 삭제합니다."
        ):
            return
        for category in existing:
            self.delete_category_file(category, ask=False)
        messagebox.showinfo("삭제 완료", "등록된 기준파일과 검색 인덱스를 모두 삭제했습니다.")

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
        self.header_update.config(text=(latest.replace("T"," ")[:16]) if latest else "—")
        if hasattr(self,"file_rows"):
            for key,_ in CATEGORIES:
                row_items=self.file_rows[key]
                status_lbl,date_lbl,file_lbl=row_items[:3]
                delete_btn=row_items[3] if len(row_items) > 3 else None
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
                if delete_btn is not None:
                    delete_btn.config(state=("normal" if path else "disabled"))

    def auto_reindex_v151(self):
        # v1.6.1: 표기 통일(그람→그램)과 후보 전용 테이블을 DB 안에서 직접 처리한다.
        migration_key = "_index_migration_v161_norm_all"
        if self.cfg.get(migration_key):
            return

        def worker():
            try:
                self.idx.migrate_norm_v161()
                self.cfg[migration_key] = datetime.now().isoformat(timespec="seconds")
                self.save_cfg()
            except Exception:
                pass
            self.after(0, self.refresh_status)

        threading.Thread(target=worker, daemon=True).start()


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
