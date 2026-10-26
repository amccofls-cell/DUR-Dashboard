from __future__ import annotations

import json
import os
import re
import shutil
import sqlite3
import sys
import threading
import time
import traceback
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter import font as tkfont

from openpyxl import load_workbook

APP_NAME = "DUR 품목리스트 조회 프로그램"
VERSION = "1.8"

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

APP_BG = "#E9E9E9"
SURFACE = "#FFFFFF"
TEXT = "#14233B"
TEXT_DARK = "#26354D"
MUTED = "#718096"
BORDER = "#ADB1B7"
SAGE = "#14A58C"
SAGE_DEEP = "#0F816E"
SAGE_SOFT = "#EAF8F5"
TERRA = "#2F80ED"
TERRA_DEEP = "#1769C2"
TERRA_SOFT = "#EAF3FF"
GRAY_SOFT = "#F3F3F3"
RAIL_BG = "#16283F"
RAIL_FG = "#9DB0CB"
RAIL_ACTIVE = "#223A57"
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

TITLE_BG = "#F3DFB8"      # 베이지 타이틀 바
TITLE_LINE = "#E08A1E"    # 주황 하단선
TITLE_FG = "#4A3210"
PANEL = "#EFEFEF"         # 회색 조건 패널
GRID_LINE = "#C4C8CE"

PLACEHOLDER = "품목명의 일부를 입력하세요"
DETAIL_FONT = 12  # 상세정보 본문 글자 크기


def resource_dir() -> Path:
    """PyInstaller 실행 시 임시 해제 폴더, 아니면 스크립트 폴더."""
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def register_bundled_fonts():
    """동봉된 Pretendard를 이 프로그램에서만 쓰도록 등록한다(시스템에 설치하지 않음)."""
    folder = resource_dir() / "fonts"
    if not folder.is_dir():
        return
    try:
        if sys.platform.startswith("win"):
            import ctypes
            for f in sorted(folder.glob("Pretendard-*.*tf")):
                ctypes.windll.gdi32.AddFontResourceExW(str(f), 0x10, 0)  # FR_PRIVATE
        elif sys.platform.startswith("linux"):
            dst = Path.home() / ".fonts"
            dst.mkdir(exist_ok=True)
            for f in folder.glob("Pretendard-*.*tf"):
                if not (dst / f.name).exists():
                    shutil.copy2(f, dst / f.name)
    except Exception:
        pass


register_bundled_fonts()


def app_dir() -> Path:
    root = Path(os.getenv("LOCALAPPDATA") or Path.home()) / "DUR_Dashboard"
    root.mkdir(parents=True, exist_ok=True)
    (root / "files").mkdir(exist_ok=True)
    return root


ROOT_DIR = app_dir()
CFG_PATH = ROOT_DIR / "config.json"
DB_PATH = ROOT_DIR / "dur_index.sqlite3"
SETS_PATH = ROOT_DIR / "product_sets.json"  # 저장된 품목 묶음


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

    def has_product(self, product):
        """현재 인덱스(어느 DUR 목록이든)에 이 품목이 있는지 확인한다."""
        np = norm(product)
        cp = product_core(product)

        def _work():
            with self.conn() as c:
                if c.execute("SELECT 1 FROM names WHERE product_norm=? LIMIT 1", (np,)).fetchone():
                    return True
                if cp and len(cp) >= 4:
                    for cat, _ in CATEGORIES:
                        if c.execute("SELECT 1 FROM records WHERE category=? AND core_norm=? LIMIT 1",
                                     (cat, cp)).fetchone():
                            return True
            return False
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
    RAIL_W = 224

    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME}  {VERSION}")
        self.geometry("1560x920")
        self.minsize(1220, 780)
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
        self.product_sets = self.load_sets()
        self.ui_font = self.resolve_font_family()
        self.selected_products = []
        self.results = {}
        self.current_product = None
        self.current_dur = "combo"
        self.view_mode = "product"  # "product": 약품별 보기, "dur": DUR별 보기

        self.setup_style()
        self.build_rail()
        self.build_header()
        self.page_host = tk.Frame(self, bg=APP_BG)
        self.page_host.pack(fill="both", expand=True)

        self.show_lookup_page()
        self.navigate_state("lookup")
        self.after(500, self.auto_reindex_v151)

    # ------------------------------------------------------------------ cfg
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

    def title_bar(self, parent, text, right_text=None):
        """베이지 타이틀 바(주황 하단선)."""
        bar = tk.Frame(parent, bg=TITLE_BG)
        bar.pack(fill="x")
        tk.Label(bar, text="■", bg=TITLE_BG, fg=TITLE_LINE, font=self.font(8)).pack(side="left", padx=(10, 4), pady=5)
        tk.Label(bar, text=text, bg=TITLE_BG, fg=TITLE_FG, font=self.font(11, "bold")).pack(side="left", pady=5)
        if right_text:
            tk.Label(bar, text=right_text, bg=TITLE_BG, fg=TITLE_FG, font=self.font(9)).pack(side="right", padx=10)
        tk.Frame(parent, bg=TITLE_LINE, height=2).pack(fill="x")
        return bar

    def font_en(self, size=10, weight="normal"):
        return ("Arial", size, weight)

    def setup_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Compact.Treeview",
            font=self.font(10),
            rowheight=28,
            background=SURFACE,
            fieldbackground=SURFACE,
            foreground=TEXT_DARK,
            bordercolor=BORDER,
            lightcolor=GRID_LINE,
            darkcolor=GRID_LINE,
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "Compact.Treeview",
            background=[("selected", "#FBE3B5")],
            foreground=[("selected", TITLE_FG)],
        )
        style.configure(
            "Compact.Treeview.Heading",
            font=self.font(10, "bold"),
            background=TITLE_BG,
            foreground=TITLE_FG,
            relief="solid",
            bordercolor=BORDER,
            lightcolor=TITLE_BG,
            darkcolor=TITLE_BG,
            borderwidth=1,
            padding=(8, 6),
        )
        style.map("Compact.Treeview.Heading", background=[("active", "#EBCB8E")])
        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor="#EDF1F6",
            background=BLUE,
            bordercolor="#EDF1F6",
            lightcolor=BLUE,
            darkcolor=BLUE,
        )

    # ------------------------------------------------------------------ rail
    def build_rail(self):
        rail = tk.Frame(self, bg=RAIL_BG, width=self.RAIL_W,
                        highlightbackground=RAIL_BG, highlightthickness=1)
        rail.pack(side="left", fill="y")
        rail.pack_propagate(False)

        brand = tk.Frame(rail, bg=RAIL_BG)
        brand.pack(fill="x", padx=20, pady=(26, 4))
        tk.Label(brand, text="◆", bg=RAIL_BG, fg=SAGE,
                 font=("Segoe UI Symbol", 22, "bold")).pack(side="left", padx=(0, 10))
        tk.Label(brand, text="DUR", bg=RAIL_BG, fg="#FFFFFF",
                 font=self.font(19, "bold")).pack(side="left")
        tk.Label(rail, text=APP_NAME, bg=RAIL_BG, fg="#6E82A0",
                 font=self.font(9), wraplength=180, justify="left"
                 ).pack(anchor="w", padx=22, pady=(2, 22))

        tk.Frame(rail, bg="#243B58", height=1).pack(fill="x", padx=18)

        self.nav = {}
        for key, label, symbol in (
            ("lookup", "DUR 조회", "⌕"),
            ("files", "기준파일 관리", "▤"),
        ):
            row = tk.Frame(rail, bg=RAIL_BG, height=52, cursor="hand2")
            row.pack(fill="x", padx=12, pady=4)
            row.pack_propagate(False)
            dot = tk.Frame(row, bg=RAIL_BG, width=4, height=26)
            dot.pack(side="left", padx=(6, 8), pady=0)
            s = tk.Label(row, text=symbol, bg=RAIL_BG, fg=RAIL_FG,
                         font=("Segoe UI Symbol", 15))
            s.pack(side="left", padx=(0, 10))
            lbl = tk.Label(row, text=label, bg=RAIL_BG, fg=RAIL_FG, font=self.font(11))
            lbl.pack(side="left")
            for w in (row, dot, s, lbl):
                w.bind("<Button-1>", lambda e, k=key: self.navigate(k))
            self.nav[key] = (row, dot, s, lbl)

        rail_bottom = tk.Frame(rail, bg=RAIL_BG)
        rail_bottom.pack(side="bottom", fill="x", padx=20, pady=18)

        upd = tk.Frame(rail_bottom, bg="#16283F", highlightbackground="#243B58", highlightthickness=1)
        upd.pack(fill="x", pady=(0, 14), ipady=8)
        tk.Label(upd, text="기준파일 최종 업데이트", bg="#16283F", fg="#7E93B2",
                 font=self.font(8, "bold")).pack(anchor="w", padx=12, pady=(2, 1))
        self.header_update = tk.Label(upd, text="—", bg="#16283F", fg="#E8EFF9",
                                      font=self.font(10, "bold"))
        self.header_update.pack(anchor="w", padx=12, pady=(0, 2))

        tk.Label(rail_bottom, text="로컬 전용 · Excel 원본은\nPC 내부에서만 처리됩니다.",
                 bg=RAIL_BG, fg="#5E7290", justify="left", font=self.font(8)).pack(anchor="w")
        self.footer = tk.Label(rail_bottom, text="Created by Chaereen Jeong",
                               bg=RAIL_BG, fg="#46587A", font=self.font_en(8))
        self.footer.pack(anchor="w", pady=(10, 0))

    def build_header(self):
        h = tk.Frame(self, bg=TITLE_BG, height=46)
        h.pack(side="top", fill="x")
        h.pack_propagate(False)
        tk.Frame(self, bg=TITLE_LINE, height=2).pack(side="top", fill="x")

        self.page_title = tk.Label(h, text="DUR 조회", bg=TITLE_BG, fg=TITLE_FG,
                                   font=self.font(14, "bold"))
        self.page_title.pack(side="left", padx=(18, 10))

        self.page_subtitle = tk.Label(h, text="품목별 DUR 등재 여부를 일괄 조회합니다",
                                      bg=TITLE_BG, fg="#7A5A24", font=self.font(10))
        self.page_subtitle.pack(side="left", padx=(0, 10), pady=(3, 0))

        ver = tk.Label(h, text=f"v{VERSION}", bg=SURFACE, fg=TITLE_FG,
                       font=self.font_en(9, "bold"), padx=8, pady=2)
        ver.pack(side="right", padx=16)
        ver.config(highlightthickness=1, highlightbackground=TITLE_LINE)

    # ------------------------------------------------------------------ nav
    def navigate(self, page):
        for w in self.page_host.winfo_children():
            w.destroy()
        titles = {
            "lookup": ("DUR 조회", "품목별 DUR 등재 여부를 일괄 조회합니다"),
            "files": ("기준파일 관리", "DUR 기준 Excel 파일을 등록하고 인덱스를 관리합니다"),
        }
        title, sub = titles.get(page, ("DUR 조회", ""))
        self.page_title.config(text=title)
        self.page_subtitle.config(text=sub)
        self.navigate_state(page)
        if page == "files":
            self.show_files_page()
        else:
            self.show_lookup_page()

    def navigate_state(self, active):
        for key, (row, dot, s, lbl) in self.nav.items():
            is_active = key == active
            row_bg = RAIL_ACTIVE if is_active else RAIL_BG
            row.config(bg=row_bg)
            dot.config(bg=BLUE if is_active else RAIL_BG)
            s.config(bg=row_bg, fg="#FFFFFF" if is_active else RAIL_FG)
            lbl.config(bg=row_bg, fg="#FFFFFF" if is_active else RAIL_FG,
                       font=self.font(11, "bold" if is_active else "normal"))

    def make_button(self, parent, text, command, primary=False, danger=False):
        if danger:
            bg, fg, active, border = SURFACE, ERROR, "#FFF1F1", "#F3B4B4"
        elif primary:
            bg, fg, active, border = BLUE, "white", "#1B6FD8", BLUE
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

    # ------------------------------------------------------------------ lookup
    def show_lookup_page(self):
        wrap = tk.Frame(self.page_host, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=18, pady=16)

        search_col = tk.Frame(wrap, bg=PANEL, width=372,
                              highlightbackground=BORDER, highlightthickness=1)
        search_col.pack(side="left", fill="y", padx=(0, 14))
        search_col.pack_propagate(False)

        self.title_bar(search_col, "품목 검색", "일괄 조회")

        tk.Label(search_col, text="품목명의 일부를 입력하면 후보가 표시됩니다.",
                 bg=PANEL, fg=MUTED, font=self.font(9)).pack(anchor="w", padx=18, pady=(4, 8))

        search_box = tk.Frame(search_col, bg=PANEL)
        search_box.pack(fill="x", padx=16)
        self.search_box = search_box
        icon = tk.Label(search_box, text="⌕", bg=SURFACE, fg=MUTED,
                        font=("Segoe UI Symbol", 13), padx=10)
        icon.pack(side="left", fill="y")
        self.query_var = tk.StringVar()
        self.query = tk.Entry(
            search_box, textvariable=self.query_var, bg=SURFACE, fg=TEXT_DARK,
            font=self.font(12), relief="flat", bd=0,
            highlightthickness=2, highlightbackground="#D4DEEA", highlightcolor=BLUE
        )
        self.query.pack(side="left", fill="x", expand=True, ipady=11)
        self.query.insert(0, PLACEHOLDER)
        self.query.bind("<FocusIn>", self.clear_placeholder)
        self.query.bind("<KeyRelease>", self.on_query_change)
        self.query.bind("<Down>", lambda e: self.focus_suggestions())
        self.query.bind("<Return>", lambda e: "break")  # 검색어 직접 입력 금지: 후보 목록에서만 선택

        # 후보 목록 틀은 평소에는 배치하지 않는다(빈 공간이 남지 않도록 보일 때만 pack).
        self.suggest_frame = tk.Frame(search_col, bg=PANEL)
        self.suggest_list = tk.Listbox(
            self.suggest_frame, height=0, bg=PANEL, fg=TEXT_DARK,
            font=self.font(10), relief="flat", bd=0,
            highlightthickness=1, highlightbackground=BORDER,
            selectbackground=BLUE_SOFT, selectforeground=BLUE,
            activestyle="none"
        )
        self.suggest_list.pack(fill="x")
        self.suggest_list.bind("<Double-Button-1>", lambda e: self.select_suggestion())
        self.suggest_list.bind("<Return>", lambda e: self.select_suggestion())

        sel_head = tk.Frame(search_col, bg=PANEL)
        sel_head.pack(fill="x", padx=18, pady=(16, 8))
        tk.Label(sel_head, text="선택한 품목", bg=PANEL, fg=TITLE_FG,
                 font=self.font(11, "bold")).pack(side="left")
        self.selected_count = tk.Label(sel_head, text="0", bg=BLUE_SOFT, fg=BLUE,
                                       font=self.font(9, "bold"), padx=8, pady=2)
        self.selected_count.pack(side="left", padx=(8, 0))
        del_btn = self.make_button(sel_head, "전체 삭제", self.clear_selected, danger=True)
        del_btn.config(padx=10, pady=4)
        del_btn.pack(side="right")

        # 조회 버튼은 먼저 아래쪽에 고정해 두고(목록이 아무리 길어져도 가려지지 않음),
        # 남는 공간에 스크롤 가능한 선택 목록을 배치한다.
        actions = tk.Frame(search_col, bg=PANEL)
        actions.pack(side="bottom", fill="x", padx=16, pady=14)
        set_row = tk.Frame(actions, bg=PANEL)
        set_row.pack(fill="x", pady=(0, 8))
        self.save_set_btn = self.make_button(set_row, "묶음으로 저장", self.save_current_set)
        self.save_set_btn.config(pady=7)
        self.save_set_btn.pack(side="left", fill="x", expand=True, padx=(0, 4))
        open_set_btn = self.make_button(set_row, "저장된 묶음", self.open_sets_dialog)
        open_set_btn.config(pady=7)
        open_set_btn.pack(side="left", fill="x", expand=True, padx=(4, 0))
        self.make_button(
            actions, "⌕  선택 품목으로 조회", self.run_selected, primary=True
        ).pack(fill="x", ipady=2)

        self.selected_list_host = tk.Frame(search_col, bg=PANEL)
        self.selected_list_host.pack(fill="both", expand=True, padx=16)
        self.sel_canvas = tk.Canvas(self.selected_list_host, bg=PANEL,
                                    highlightthickness=0, bd=0)
        self.sel_scroll = ttk.Scrollbar(self.selected_list_host, orient="vertical",
                                        command=self.sel_canvas.yview)
        self.sel_canvas.configure(yscrollcommand=self._sel_yset)
        self.sel_canvas.pack(side="left", fill="both", expand=True)
        self.sel_inner = tk.Frame(self.sel_canvas, bg=PANEL)
        self.sel_window = self.sel_canvas.create_window((0, 0), window=self.sel_inner, anchor="nw")
        self.sel_inner.bind("<Configure>", lambda e: self._sel_fit())
        self.sel_canvas.bind("<Configure>", lambda e: self._sel_fit())
        self._bind_sel_wheel(self.sel_canvas)

        result = tk.Frame(wrap, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        result.pack(side="left", fill="both", expand=True)
        self.result_body = result

        self.refresh_selected_panel()
        if self.current_product and self.current_product in self.results:
            self.render_results()
        else:
            self.render_empty()

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
            self.suggest_frame.pack(fill="x", padx=16, pady=(4, 0), after=self.search_box)
        else:
            self.hide_suggestions()

    def hide_suggestions(self):
        if hasattr(self, "suggest_frame"):
            self.suggest_frame.pack_forget()

    def focus_suggestions(self):
        if self.suggest_list.winfo_ismapped() and self.suggest_list.size():
            self.suggest_list.focus_set()
            self.suggest_list.selection_clear(0, "end")
            self.suggest_list.selection_set(0)
            return "break"

    def select_suggestion(self):
        sel = self.suggest_list.curselection()
        if not sel:
            return
        product = self.suggest_list.get(sel[0])
        if product not in self.selected_products:
            self.selected_products.append(product)
        self.refresh_selected_panel(to_end=True)
        self.query.delete(0, "end")
        self.query.focus_set()
        self.hide_suggestions()

    # -------------------------------------------- 선택 품목 목록(스크롤 영역)
    def _sel_yset(self, first, last):
        """내용이 영역보다 길 때만 스크롤바를 보여 준다."""
        self.sel_scroll.set(first, last)
        needed = not (float(first) <= 0.0 and float(last) >= 1.0)
        shown = self.sel_scroll.winfo_manager() != ""
        if needed and not shown:
            self.sel_scroll.pack(side="right", fill="y", before=self.sel_canvas)
        elif not needed and shown:
            self.sel_scroll.pack_forget()

    def _sel_fit(self):
        """내부 프레임 너비를 영역에 맞추고, 짧을 때는 영역 높이만큼 채운다."""
        if not hasattr(self, "sel_canvas"):
            return
        cw = self.sel_canvas.winfo_width()
        ch = self.sel_canvas.winfo_height()
        need = self.sel_inner.winfo_reqheight()
        self.sel_canvas.itemconfigure(self.sel_window, width=cw, height=max(ch, need))
        self.sel_canvas.configure(scrollregion=(0, 0, cw, max(ch, need)))

    def _bind_sel_wheel(self, widget):
        def on_wheel(e):
            if self.sel_inner.winfo_reqheight() <= self.sel_canvas.winfo_height():
                return "break"
            step = -1 if (getattr(e, "num", 0) == 4 or getattr(e, "delta", 0) > 0) else 1
            self.sel_canvas.yview_scroll(step * 2, "units")
            return "break"
        widget.bind("<MouseWheel>", on_wheel, add="+")
        widget.bind("<Button-4>", on_wheel, add="+")
        widget.bind("<Button-5>", on_wheel, add="+")
        for ch in widget.winfo_children():
            self._bind_sel_wheel(ch)

    def refresh_selected_panel(self, to_end=False):
        if not hasattr(self, "sel_inner"):
            return
        first = self.sel_canvas.yview()[0]
        for w in self.sel_inner.winfo_children():
            w.destroy()
        self.selected_count.config(text=str(len(self.selected_products)))
        if not self.selected_products:
            empty = tk.Frame(self.sel_inner, bg=GRAY_SOFT,
                             highlightbackground=BORDER, highlightthickness=1)
            empty.pack(fill="both", expand=True, pady=(0, 2))
            tk.Label(empty, text="＋", bg=GRAY_SOFT, fg="#A8B7CB",
                     font=("Segoe UI Symbol", 20)).pack(pady=(26, 4))
            tk.Label(empty, text="검색 후보에서 품목을 선택하면\n이곳에 누적됩니다.",
                     bg=GRAY_SOFT, fg=MUTED, justify="center", font=self.font(9)).pack(pady=(0, 22))
            self._bind_sel_wheel(self.sel_inner)
            self.sel_canvas.update_idletasks()
            self._sel_fit()
            self.sel_canvas.yview_moveto(0)
            return

        active_row = None
        for i, product in enumerate(self.selected_products, 1):
            active = product == self.current_product
            bg = BLUE_SOFT if active else "#F8FAFC"
            border = "#B9D4F5" if active else BORDER
            row = tk.Frame(self.sel_inner, bg=bg, highlightbackground=border,
                           highlightthickness=1, height=52, cursor="hand2")
            row.pack(fill="x", pady=3)
            row.pack_propagate(False)
            if active:
                active_row = row
                bar = tk.Frame(row, bg=BLUE, width=3)
                bar.pack(side="left", fill="y")
            num = tk.Label(row, text=str(i), bg=bg, fg=MUTED,
                           font=self.font_en(9, "bold"), width=3)
            num.pack(side="left", padx=(8, 0))
            close = tk.Label(row, text="✕", bg=bg, fg="#9AA8BC",
                             font=self.font(10, "bold"), cursor="hand2", padx=8)
            close.pack(side="right")
            name = tk.Label(row, text=product, bg=bg, fg=BLUE if active else TEXT_DARK,
                            font=self.font(10, "bold" if active else "normal"),
                            anchor="w", justify="left", wraplength=220)
            name.pack(side="left", fill="x", expand=True, padx=(2, 6))
            widgets = (row, num, name)
            for w in widgets:
                w.bind("<Button-1>", lambda e, p=product: self.pick_product(p))
            close.bind("<Button-1>", lambda e, p=product: self.remove_selected_product(p))
            close.bind("<Enter>", lambda e, c=close: c.config(fg=ERROR))
            close.bind("<Leave>", lambda e, c=close: c.config(fg="#9AA8BC"))

        self._bind_sel_wheel(self.sel_inner)
        self.sel_canvas.update_idletasks()
        self._sel_fit()
        self.sel_canvas.update_idletasks()
        if to_end:
            self.sel_canvas.yview_moveto(1.0)
        else:
            self.sel_canvas.yview_moveto(first)
            if active_row is not None:  # 활성 행이 화면 밖이면 보이는 위치로 이동
                total = max(1, self.sel_inner.winfo_reqheight())
                view_h = self.sel_canvas.winfo_height()
                top = self.sel_canvas.canvasy(0)
                y, h = active_row.winfo_y(), active_row.winfo_height() + 6
                if y < top:
                    self.sel_canvas.yview_moveto(y / total)
                elif y + h > top + view_h:
                    self.sel_canvas.yview_moveto((y + h - view_h) / total)

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

    # ------------------------------------------------------------ 품목 묶음
    def load_sets(self):
        try:
            data = json.loads(SETS_PATH.read_text(encoding="utf-8"))
        except Exception:
            return []
        out = []
        for d in data if isinstance(data, list) else []:
            if not (isinstance(d, dict) and isinstance(d.get("name"), str)
                    and isinstance(d.get("products"), list)):
                continue
            prods = [str(p) for p in d["products"] if str(p).strip()]
            if d["name"].strip() and prods:
                out.append({"name": d["name"].strip(), "products": prods,
                            "saved_at": str(d.get("saved_at", ""))})
        return out

    def write_sets(self):
        tmp = SETS_PATH.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(self.product_sets, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, SETS_PATH)

    def save_current_set(self, name=None):
        """지금 '선택한 품목' 전체를 이름을 붙여 묶음으로 저장한다."""
        if not self.selected_products:
            messagebox.showwarning("저장할 품목 없음", "먼저 품목을 선택한 뒤 저장해 주세요.")
            return
        if name is None:
            name = simpledialog.askstring(
                "묶음으로 저장",
                f"선택한 {len(self.selected_products)}개 품목을 저장할 이름을 입력하세요.",
                initialvalue=f"묶음 {len(self.product_sets) + 1}", parent=self)
        if name is None:
            return
        name = name.strip()[:40]
        if not name:
            messagebox.showwarning("이름 필요", "묶음 이름을 입력해 주세요.")
            return
        existing = next((s for s in self.product_sets if s["name"] == name), None)
        if existing and not messagebox.askyesno(
                "같은 이름의 묶음",
                f"'{name}' 묶음이 이미 있습니다.\n현재 선택한 품목으로 덮어쓸까요?"):
            return
        entry = {"name": name, "products": list(self.selected_products),
                 "saved_at": datetime.now().isoformat(timespec="seconds")}
        if existing:
            existing.update(entry)
        else:
            self.product_sets.append(entry)
        try:
            self.write_sets()
        except Exception as e:
            messagebox.showerror("묶음 저장 실패", str(e))
            return
        btn = getattr(self, "save_set_btn", None)
        if btn is not None and btn.winfo_exists():
            old = "묶음으로 저장"
            btn.config(text="저장됨 ✓")
            self.after(1400, lambda: btn.winfo_exists() and btn.config(text=old))

    def missing_products(self, products):
        try:
            return [p for p in products if not self.idx.has_product(p)]
        except Exception:
            return []

    def apply_set(self, products, replace=True, run=True):
        """묶음의 품목을 선택 목록에 넣고(replace=True면 교체), 필요하면 바로 조회한다."""
        products = list(dict.fromkeys(products))
        if replace:
            saved = [set(s["products"]) for s in self.product_sets]
            cur = set(self.selected_products)
            unsaved = bool(cur) and not cur <= set(self.results) and cur not in saved
            if unsaved and not messagebox.askyesno(
                    "선택 목록 교체",
                    f"현재 선택한 {len(self.selected_products)}개 품목은 저장·조회되지 않았습니다.\n"
                    "이 묶음으로 교체할까요?"):
                return False
            self.selected_products = products
            self.results = {}
            self.current_product = None
            self.refresh_selected_panel()
        else:
            for p in products:
                if p not in self.selected_products:
                    self.selected_products.append(p)
            self.refresh_selected_panel(to_end=True)
        if run:
            missing = self.missing_products(products)
            self.run_selected()
            if missing:
                messagebox.showwarning(
                    "일부 품목 확인 필요",
                    f"다음 {len(missing)}개 품목은 현재 등록된 기준파일에서 찾을 수 없습니다.\n"
                    "제품명이 바뀌었거나 기준파일이 등록되지 않았을 수 있습니다.\n"
                    "해당 품목은 다시 검색해서 선택해 주세요.\n\n"
                    + "\n".join(f"- {m}" for m in missing[:15])
                    + (f"\n… 외 {len(missing) - 15}개" if len(missing) > 15 else ""))
        return True

    def open_sets_dialog(self):
        win = tk.Toplevel(self)
        win.title("저장된 묶음")
        win.configure(bg=SURFACE)
        win.transient(self)
        w, h = 600, 560
        win.geometry(f"{w}x{h}")
        win.minsize(520, 380)
        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - w) // 2
        y = self.winfo_y() + (self.winfo_height() - h) // 2
        win.geometry(f"+{max(0, x)}+{max(0, y)}")

        head = tk.Frame(win, bg=SURFACE)
        head.pack(fill="x", padx=22, pady=(18, 6))
        tk.Label(head, text="저장된 묶음", bg=SURFACE, fg=TEXT,
                 font=self.font(15, "bold")).pack(side="left")
        tk.Label(win, text="묶음을 클릭하면 품목 검색 없이 바로 DUR 해당 여부를 조회합니다.",
                 bg=SURFACE, fg=MUTED, font=self.font(10)).pack(anchor="w", padx=22, pady=(0, 10))

        foot = tk.Frame(win, bg=SURFACE)
        foot.pack(side="bottom", fill="x", padx=22, pady=14)
        self.make_button(foot, "닫기", win.destroy).pack(side="right")

        body = tk.Frame(win, bg=SURFACE)
        body.pack(fill="both", expand=True, padx=22)
        canvas = tk.Canvas(body, bg=SURFACE, highlightthickness=0, bd=0)
        sb = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)

        def yset(first, last):  # 목록이 길 때만 스크롤바 표시
            sb.set(first, last)
            needed = not (float(first) <= 0.0 and float(last) >= 1.0)
            shown = sb.winfo_manager() != ""
            if needed and not shown:
                sb.pack(side="right", fill="y", before=canvas)
            elif not needed and shown:
                sb.pack_forget()

        canvas.configure(yscrollcommand=yset)
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg=SURFACE)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))

        def on_wheel(e):
            if inner.winfo_reqheight() <= canvas.winfo_height():
                return
            step = -1 if (getattr(e, "num", 0) == 4 or getattr(e, "delta", 0) > 0) else 1
            canvas.yview_scroll(step * 2, "units")
        win.bind("<MouseWheel>", on_wheel)
        win.bind("<Button-4>", on_wheel)
        win.bind("<Button-5>", on_wheel)

        def short(p):
            t = re.sub(r"[_(].*$", "", p).strip()
            return t or p

        def do_run(entry):
            win.destroy()
            self.apply_set(entry["products"], replace=True, run=True)

        def do_add(entry):
            win.destroy()
            self.apply_set(entry["products"], replace=False, run=False)

        def do_delete(entry):
            if not messagebox.askyesno("묶음 삭제", f"'{entry['name']}' 묶음을 삭제할까요?", parent=win):
                return
            if entry in self.product_sets:
                self.product_sets.remove(entry)
            try:
                self.write_sets()
            except Exception as e:
                messagebox.showerror("묶음 삭제 실패", str(e), parent=win)
            render()

        def bind_click(widget, entry):
            if isinstance(widget, tk.Button):
                return
            widget.bind("<Button-1>", lambda e, en=entry: do_run(en))
            for ch in widget.winfo_children():
                bind_click(ch, entry)

        def render():
            for c in inner.winfo_children():
                c.destroy()
            # 최근 저장한 묶음이 위로 (저장 시각이 같으면 나중에 만든 것이 위)
            entries = [e for _, e in sorted(enumerate(self.product_sets),
                                            key=lambda t: (t[1].get("saved_at", ""), t[0]),
                                            reverse=True)]
            if not entries:
                tk.Label(inner, text="저장된 묶음이 없습니다.\n품목을 선택한 뒤 '묶음으로 저장'을 눌러 보세요.",
                         bg=SURFACE, fg=MUTED, font=self.font(11), justify="center").pack(pady=60)
                return
            for entry in entries:
                bg = GRAY_SOFT
                card = tk.Frame(inner, bg=bg, highlightbackground=BORDER, highlightthickness=1,
                                cursor="hand2")
                card.pack(fill="x", pady=5, padx=(0, 6))
                top = tk.Frame(card, bg=bg)
                top.pack(fill="x", padx=14, pady=(12, 2))
                tk.Label(top, text=entry["name"], bg=bg, fg=TEXT, font=self.font(12, "bold"),
                         anchor="w").pack(side="left")
                tk.Label(top, text=f"{len(entry['products'])}개 품목", bg=BLUE_SOFT, fg=BLUE,
                         font=self.font(9, "bold"), padx=8, pady=2).pack(side="left", padx=(8, 0))
                saved = entry.get("saved_at", "").replace("T", " ")[:16]
                if saved:
                    tk.Label(top, text=f"저장 {saved}", bg=bg, fg=MUTED,
                             font=self.font(9)).pack(side="right")
                names = [short(p) for p in entry["products"]]
                preview = " · ".join(names[:4]) + (f"  외 {len(names) - 4}개" if len(names) > 4 else "")
                tk.Label(card, text=preview, bg=bg, fg=TEXT_DARK, font=self.font(10),
                         anchor="w", justify="left", wraplength=470).pack(fill="x", padx=14, pady=(0, 8))
                btns = tk.Frame(card, bg=bg)
                btns.pack(fill="x", padx=14, pady=(0, 12))
                run_b = self.make_button(btns, "⌕  바로 조회", lambda en=entry: do_run(en), primary=True)
                run_b.config(padx=14, pady=5)
                run_b.pack(side="left")
                add_b = self.make_button(btns, "목록에 추가", lambda en=entry: do_add(en))
                add_b.config(padx=12, pady=5)
                add_b.pack(side="left", padx=(8, 0))
                del_b = self.make_button(btns, "삭제", lambda en=entry: do_delete(en), danger=True)
                del_b.config(padx=12, pady=5)
                del_b.pack(side="right")
                bind_click(card, entry)

        render()
        try:
            win.grab_set()
        except tk.TclError:
            pass
        win.focus_set()

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

    # ------------------------------------------------------------------ results
    def render_empty(self):
        for w in self.result_body.winfo_children():
            w.destroy()
        self.title_bar(self.result_body, "조회 결과")
        box = tk.Frame(self.result_body, bg=SURFACE)
        box.pack(expand=True)
        badge = tk.Frame(box, bg=GRAY_SOFT, highlightbackground=BORDER, highlightthickness=1)
        badge.pack(pady=(0, 14), ipadx=18, ipady=14)
        tk.Label(badge, text="⌕", bg=GRAY_SOFT, fg="#9FB0C6",
                 font=("Segoe UI Symbol", 30)).pack()
        tk.Label(box, text="조회할 품목을 선택해 주세요",
                 bg=SURFACE, fg=TEXT, font=self.font(16, "bold")).pack()
        tk.Label(
            box,
            text="왼쪽 검색창에 품목명의 일부를 입력하면 후보가 자동으로 표시됩니다.\n후보를 선택해 누적한 뒤 [선택 품목으로 조회]를 누르세요.",
            bg=SURFACE, fg=MUTED, justify="center", font=self.font(10)
        ).pack(pady=(8, 0))

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

        self.title_bar(self.result_body, "조회 결과", re.sub(r"[_(].*$", "", p).strip() or p)
        right = tk.Frame(self.result_body, bg=SURFACE)
        right.pack(fill="both", expand=True, padx=14, pady=12)

        self.build_view_toggle(right)
        if self.view_mode == "dur":
            self.render_by_dur(right)
            return

        head = tk.Frame(right, bg=SURFACE)
        head.pack(fill="x", pady=(0, 6))
        bar = tk.Frame(head, bg=SAGE, width=5, height=30)
        bar.pack(side="left", padx=(0, 12))
        tk.Label(
            head, text=p, bg=SURFACE, fg=TEXT,
            font=self.font(19, "bold")
        ).pack(side="left")

        active_parts = []
        for key, label in DUR_CARDS:
            state, sub, active = self.card_state(p, key)
            if active:
                active_parts.append(f"{label} {sub or ''}".strip())
        summary = tk.Frame(right, bg=SURFACE)
        summary.pack(fill="x", pady=(2, 12))
        if active_parts:
            count_chip = tk.Label(summary, text=f"{len(active_parts)}개 DUR 해당",
                                  bg=RED_SOFT, fg=RED, font=self.font(10, "bold"),
                                  padx=10, pady=4)
            count_chip.pack(side="left", padx=(0, 8))
            palette = self.dur_palette()
            for key, label in DUR_CARDS:
                state, sub, active = self.card_state(p, key)
                if not active:
                    continue
                accent, soft = palette[key]
                chip = tk.Label(summary, text=f"{label} {sub or ''}".strip(),
                                bg=soft, fg=accent, font=self.font(9, "bold"),
                                padx=9, pady=4, cursor="hand2")
                chip.pack(side="left", padx=(0, 6))
                chip.bind("<Button-1>", lambda e, k=key: self.render_detail(k))
        else:
            tk.Label(summary, text="7개 관리항목 모두 해당 없음",
                     bg=GRAY_SOFT, fg=MUTED, font=self.font(10, "bold"),
                     padx=10, pady=4).pack(side="left")

        cards = tk.Frame(right, bg=SURFACE)
        cards.pack(fill="x")

        palette = self.dur_palette()

        for i, (key, label) in enumerate(DUR_CARDS):
            state, sub, active = self.card_state(p, key)
            accent, soft = palette[key]
            bg = soft if active else "#FBFCFE"
            border = accent if active else BORDER
            c = tk.Frame(
                cards, bg=bg, highlightbackground=border, highlightthickness=1,
                height=112, cursor="hand2"
            )
            c.grid(row=i // 4, column=i % 4, sticky="nsew", padx=5, pady=5)
            c.grid_propagate(False)

            edge = tk.Frame(c, bg=accent if active else "#E4EAF2", width=4)
            edge.pack(side="left", fill="y")

            inner = tk.Frame(c, bg=bg)
            inner.pack(fill="both", expand=True)

            top = tk.Frame(inner, bg=bg)
            top.pack(fill="x", padx=12, pady=(12, 4))
            tk.Label(top, text=label, bg=bg, fg=accent if active else TEXT_DARK,
                     font=self.font(11, "bold")).pack(side="left")
            tk.Label(top, text="›" if active else "–", bg=bg,
                     fg=accent if active else "#B6C2D2",
                     font=self.font(15, "bold")).pack(side="right")

            main_text = sub or state if active else "해당 없음"
            tk.Label(
                inner,
                text=("●  " + main_text) if active else "○  해당 없음",
                bg=bg, fg=accent if active else MUTED,
                font=self.font(13, "bold" if active else "normal")
            ).pack(anchor="w", padx=12, pady=(4, 0))


            def descendants(w):
                out = []
                for ch in w.winfo_children():
                    out.append(ch)
                    out.extend(descendants(ch))
                return out

            hover_bg = "#F1F5FA" if not active else bg
            card_widgets = [c] + descendants(c)
            state_box = {"on": False}

            def paint(color, widgets=card_widgets, base=bg, hover=hover_bg):
                for w in widgets:
                    try:
                        if w.cget("bg") in (base, hover):
                            w.config(bg=color)
                    except tk.TclError:
                        pass

            def pointer_inside(card=c):
                try:
                    px, py = card.winfo_pointerxy()
                    x0, y0 = card.winfo_rootx(), card.winfo_rooty()
                    return x0 <= px < x0 + card.winfo_width() and y0 <= py < y0 + card.winfo_height()
                except tk.TclError:
                    return False

            def sync_hover(box=state_box, base=bg, hv=hover_bg, paint=paint, inside=pointer_inside):
                # 자식 위젯 사이를 오갈 때 생기는 Leave/Enter 연타를 무시하고, 상태가 바뀔 때만 한 번 칠한다.
                now = inside()
                if now != box["on"]:
                    box["on"] = now
                    paint(hv if now else base)

            for w in card_widgets:
                try:
                    w.bind("<Button-1>", lambda e, k=key: self.render_detail(k))
                    if not active:
                        w.bind("<Enter>", lambda e, f=sync_hover: f(), add="+")
                        w.bind("<Leave>", lambda e, f=sync_hover, a=self.after: a(15, f), add="+")
                except tk.TclError:
                    pass

        for col in range(4):
            cards.columnconfigure(col, weight=1, uniform="duecard")  # 열 폭 고정(내용 폭에 따라 흔들리지 않게)
        for row in range((len(DUR_CARDS) + 3) // 4):
            cards.rowconfigure(row, minsize=122, uniform="duerow")

        self.detail_tabs = tk.Frame(right, bg=SURFACE)
        self.detail_tabs.pack(fill="x", pady=(14, 0))

        self.detail_host = tk.Frame(right, bg=SURFACE)
        self.detail_host.pack(fill="both", expand=True, pady=(0, 4))

        self.render_detail(next((k for k, _ in DUR_CARDS if self.card_state(p, k)[2]), "combo"))

    def dur_palette(self):
        return {
            "combo": (RED, RED_SOFT),
            "age": (ORANGE, ORANGE_SOFT),
            "preg": (RED, "#FFF2F4"),
            "lact": (PURPLE, PURPLE_SOFT),
            "dup": (GREEN, GREEN_SOFT),
            "tele": (BLUE, BLUE_SOFT),
            "cost": (SAGE_DEEP, SAGE_SOFT),
        }

    def set_view_mode(self, mode):
        self.view_mode = mode
        self.render_results()

    def build_view_toggle(self, parent):
        bar = tk.Frame(parent, bg=SURFACE)
        bar.pack(fill="x", pady=(0, 12))
        seg = tk.Frame(bar, bg="#EDF1F6", highlightbackground=BORDER, highlightthickness=1)
        seg.pack(side="left")
        for mode, label in (("product", "약품별 보기"), ("dur", "DUR별 보기")):
            on = self.view_mode == mode
            tk.Button(
                seg, text=label, command=lambda m=mode: self.set_view_mode(m),
                bg=SURFACE if on else "#EDF1F6", fg=BLUE if on else MUTED,
                activebackground=SURFACE if on else "#E3E9F1",
                activeforeground=BLUE if on else MUTED,
                relief="flat", bd=0, font=self.font(10, "bold"),
                padx=20, pady=8, cursor="hand2",
                highlightthickness=1,
                highlightbackground=BLUE if on else "#EDF1F6",
            ).pack(side="left")

    def dur_groups(self):
        """DUR 항목별로 해당 품목을 모은다. ([(key, label, [(품목, 요약)])], 조회된 품목 수)"""
        prods = [p for p in self.selected_products if p in self.results]
        groups = []
        for key, label in DUR_CARDS:
            items = []
            for p in prods:
                state, sub, active = self.card_state(p, key)
                if active:
                    items.append((p, sub or "해당"))
            groups.append((key, label, items))
        return groups, len(prods)

    def dur_groups_text(self):
        groups, total = self.dur_groups()
        lines = []
        for key, label, items in groups:
            if items:
                lines.append(f"[{label}] {len(items)}/{total}개 품목 해당")
                lines += [f"- {p}: {sub}" for p, sub in items]
            else:
                lines.append(f"[{label}] 해당 품목 없음")
        return "\n".join(lines)

    def render_by_dur(self, right):
        groups, total = self.dur_groups()
        palette = self.dur_palette()
        hit_durs = sum(1 for _, _, items in groups if items)

        head = tk.Frame(right, bg=SURFACE)
        head.pack(fill="x", pady=(0, 6))
        bar = tk.Frame(head, bg=BLUE, width=5, height=28)
        bar.pack(side="left", padx=(0, 12))
        tk.Label(head, text="DUR 항목별 해당 품목", bg=SURFACE, fg=TEXT,
                 font=self.font(16, "bold")).pack(side="left")
        copy_btn = self.make_button(head, "전체 복사", lambda: self.copy_to_clipboard(
            self.dur_groups_text(), copy_btn))
        copy_btn.config(padx=12, pady=5)
        copy_btn.pack(side="right")

        info = tk.Frame(right, bg=SURFACE)
        info.pack(fill="x", pady=(0, 8))
        tk.Label(info, text=f"조회 품목 {total}개", bg=GRAY_SOFT, fg=TEXT_DARK,
                 font=self.font(9, "bold"), padx=9, pady=4).pack(side="left")
        tk.Label(info, text=f"해당 DUR {hit_durs}개", bg=BLUE_SOFT, fg=BLUE,
                 font=self.font(9, "bold"), padx=9, pady=4).pack(side="left", padx=(6, 0))
        tk.Label(info, text="품목을 더블클릭하면 해당 약품의 상세정보로 이동합니다",
                 bg=SURFACE, fg=MUTED, font=self.font(9)).pack(side="left", padx=(10, 0))

        box = tk.Frame(right, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        box.pack(fill="both", expand=True)
        tree = ttk.Treeview(box, columns=("sub",), show="tree headings",
                            style="Compact.Treeview", selectmode="extended")
        sb = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.heading("#0", text="DUR 항목 / 해당 품목")
        tree.heading("sub", text="내용")
        tree.column("#0", width=520, minwidth=260)
        tree.column("sub", width=140, minwidth=90, anchor="center", stretch=False)
        sb.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)

        for key, _ in DUR_CARDS:
            accent, soft = palette[key]
            tree.tag_configure(f"g_{key}", foreground=accent, background=soft,
                               font=self.font(11, "bold"))
        tree.tag_configure("g_empty", foreground=MUTED, background=GRAY_SOFT)
        tree.tag_configure("child", foreground=TEXT_DARK)

        label_of = dict(DUR_CARDS)
        child_map = {}
        group_children = {}
        for key, label, items in groups:
            if items:
                tree.insert("", "end", iid=key, open=True, tags=(f"g_{key}",),
                            text=f"{label}   {len(items)}/{total}개 품목 해당")
            else:
                tree.insert("", "end", iid=key, open=True, tags=("g_empty",),
                            text=f"{label}   해당 품목 없음")
            group_children[key] = []
            for n, (p, sub) in enumerate(items):
                iid = f"{key}::{n}"
                tree.insert(key, "end", iid=iid, text="      " + p, values=(sub,), tags=("child",))
                child_map[iid] = (key, p)
                group_children[key].append(iid)

        def copy_selected(_=None):
            lines, seen = [], set()
            for iid in tree.selection():
                if iid in group_children:
                    if not group_children[iid]:
                        lines.append(f"{label_of[iid]}\t(해당 품목 없음)\t")
                    targets = group_children[iid]
                else:
                    targets = [iid]
                for t in targets:
                    if t in seen or t not in child_map:
                        continue
                    seen.add(t)
                    key, p = child_map[t]
                    lines.append(f"{label_of[key]}\t{p}\t{tree.set(t, 'sub')}")
            if lines:
                self.copy_to_clipboard("\n".join(lines))
            return "break"

        def go_detail(iid):
            if iid not in child_map:
                return
            key, p = child_map[iid]
            self.view_mode = "product"
            self.current_product = p
            self.refresh_selected_panel()
            self.render_results()
            self.render_detail(key)

        def on_double(e):
            go_detail(tree.identify_row(e.y))

        def on_return(_=None):
            sel = tree.selection()
            if sel:
                go_detail(sel[0])

        tree.bind("<Double-1>", on_double)
        tree.bind("<Return>", on_return)
        tree.bind("<Control-c>", copy_selected)
        tree.bind("<Control-C>", copy_selected)
        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label="선택 행 복사", command=copy_selected)
        menu.add_command(label="전체 복사", command=lambda: self.copy_to_clipboard(self.dur_groups_text()))

        def popup(e):
            row = tree.identify_row(e.y)
            if row and row not in tree.selection():
                tree.selection_set(row)
            try:
                menu.tk_popup(e.x_root, e.y_root)
            finally:
                menu.grab_release()

        tree.bind("<Button-3>", popup)

    def pick_product(self, p):
        self.view_mode = "product"
        self.current_product = p
        self.refresh_selected_panel()
        if p in self.results:
            self.render_results()

    # ------------------------------------------------- 상세정보 선택·복사 지원
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
        palette = self.dur_palette()
        accent, soft = palette[key]

        if hasattr(self, "detail_tabs"):
            for w in self.detail_tabs.winfo_children():
                w.destroy()
            for tab_key, tab_label in DUR_CARDS:
                st, sub, active = self.card_state(p, tab_key)
                if not active:
                    continue
                selected = tab_key == key
                t_accent, t_soft = palette[tab_key]
                btn = tk.Button(
                    self.detail_tabs,
                    text=f"{tab_label} 상세" + (f" ({sub})" if sub and sub.endswith("건") else ""),
                    command=lambda k=tab_key: self.render_detail(k),
                    bg=t_soft if selected else SURFACE,
                    fg=(t_accent if selected else TEXT_DARK),
                    activebackground=t_soft,
                    activeforeground=t_accent,
                    relief="flat",
                    bd=0,
                    font=self.font(10, "bold" if selected else "normal"),
                    padx=14,
                    pady=9,
                    cursor="hand2",
                    highlightthickness=1,
                    highlightbackground=t_accent if selected else BORDER,
                )
                btn.pack(side="left", padx=(0, 6))

        head = tk.Frame(self.detail_host, bg=TITLE_BG, highlightbackground=TITLE_LINE, highlightthickness=1)
        head.pack(fill="x", pady=(12, 8))
        dot = tk.Frame(head, bg=accent, width=6, height=20)
        dot.pack(side="left", padx=(0, 10), pady=4)
        count_note = f"  ·  {len(hits)}건" if key != "combo" and len(hits) > 1 else ""
        tk.Label(head, text=f"{title} 상세정보{count_note}", bg=TITLE_BG, fg=TITLE_FG,
                 font=self.font(12, "bold")).pack(side="left")

        def small_button(parent, text, command):
            b = self.make_button(parent, text, command)
            b.config(padx=12, pady=5)
            return b

        if not hits:
            empty = tk.Frame(self.detail_host, bg=GRAY_SOFT,
                             highlightbackground=BORDER, highlightthickness=1)
            empty.pack(fill="both", expand=True)
            tk.Label(empty, text="해당 없음", bg=GRAY_SOFT, fg=MUTED,
                     font=self.font(DETAIL_FONT)).pack(expand=True)
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
                    other_p, other_i, pay = d.get("제품명B", ""), d.get("성분명B", ""), d.get("급여여부A", "")
                else:
                    other_p, other_i, pay = d.get("제품명A", ""), d.get("성분명A", ""), d.get("급여여부A", "")
                vals = (other_p, other_i, pay, d.get("고시번호", ""), d.get("고시일자", ""))
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
            tk.Label(head, text="  열 제목을 클릭하면 정렬됩니다", bg=TITLE_BG, fg="#7A5A24",
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

        hidden = ("성분코드", "제품코드", "약품코드", "일반명코드", "주성분코드", "업체명", "원본", "시트")
        blocks = []  # 해당 건마다 (항목, 값) 목록 하나씩
        for h in hits:
            fl = []
            for k, v in h["data"].items():
                kk = re.sub(r"\s+", "", str(k))  # 머리글 안의 줄바꿈/공백 제거 (예: "일반명\n코드")
                if sval(v) and not any(x in kk for x in hidden):
                    fl.append((kk, sval(v)))
            blocks.append(fl)
        multi = len(blocks) > 1

        holder, txt = self.make_readonly_text(self.detail_host, height=8, bg=GRAY_SOFT)
        holder.pack(fill="both", expand=True)
        key_font = tkfont.Font(font=self.font(DETAIL_FONT, "bold"))
        all_keys = [k for fl in blocks for k, _ in fl]
        tab = (max((key_font.measure(k) for k in all_keys), default=80)) + 40
        txt.configure(tabs=(tab,))
        txt.tag_configure("row", lmargin1=0, lmargin2=tab, spacing1=6, spacing3=6)
        txt.tag_configure("k", font=self.font(DETAIL_FONT, "bold"), foreground=MUTED)
        txt.tag_configure("hdr", font=self.font(DETAIL_FONT, "bold"), foreground=accent,
                          background=soft, spacing1=14, spacing3=4)
        for n, fl in enumerate(blocks, 1):
            if multi:  # 여러 건이면 건마다 구분 머리글을 붙여 모두 표시
                txt.insert("end", f"  {n} / {len(blocks)}\n", ("hdr",))
            for k, v in fl:
                txt.insert("end", k, ("row", "k"))
                txt.insert("end", "\t", ("row",))
                txt.insert("end", v + "\n", ("row",))
        if all_keys or multi:
            txt.delete("end-2c", "end-1c")  # 마지막 빈 줄 제거

        def all_text():
            parts = []
            for n, fl in enumerate(blocks, 1):
                head_line = f"[{n}/{len(blocks)}]\n" if multi else ""
                parts.append(head_line + "\n".join(f"{k}: {v}" for k, v in fl))
            return "\n\n".join(parts)

        copy_btn = small_button(head, "내용 전체 복사", lambda: self.copy_to_clipboard(all_text(), copy_btn))
        copy_btn.pack(side="right")

    # ------------------------------------------------------------------ files
    def show_files_page(self):
        self.navigate_state("files")
        wrap = tk.Frame(self.page_host, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=18, pady=16)

        main = tk.Frame(wrap, bg=APP_BG)
        main.pack(fill="both", expand=True)

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

        btns = tk.Frame(head, bg=APP_BG)
        btns.pack(side="right")
        self.make_button(btns, "전체 기준파일 삭제", self.delete_all_files, danger=True
                         ).pack(side="right", padx=(8, 0))
        self.make_button(btns, "파일 일괄 등록", self.choose_files, primary=True
                         ).pack(side="right", padx=(8, 0), ipady=2)
        self.make_button(btns, "다시 인덱싱", self.reindex_all).pack(side="right")

        panel = tk.Frame(main, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="both", expand=True)

        headers = ("파일 종류", "상태", "마지막 등록", "등록 파일", "")
        for col, label in enumerate(headers):
            tk.Label(
                panel, text=label, bg="#F1F5FA", fg="#52626D",
                font=self.font(10, "bold"), anchor="w", padx=14, pady=12
            ).grid(row=0, column=col, sticky="nsew")
        panel.columnconfigure(3, weight=1)

        self.file_rows = {}
        for i, (key, label) in enumerate(CATEGORIES, 1):
            bg = SURFACE if i % 2 else "#FBFCFE"
            a = tk.Label(panel, text=label, bg=bg, fg=TEXT_DARK, font=self.font(11, "bold"),
                         anchor="w", padx=14, pady=13)
            b = tk.Label(panel, text="미등록", bg=bg, fg=MUTED, font=self.font(10),
                         anchor="w", padx=14, pady=13)
            c = tk.Label(panel, text="—", bg=bg, fg=MUTED, font=self.font(10),
                         anchor="w", padx=14, pady=13)
            d = tk.Label(panel, text="—", bg=bg, fg=MUTED, font=self.font(10),
                         anchor="w", padx=14, pady=13)
            cell = tk.Frame(panel, bg=bg)
            e = self.make_button(cell, "삭제", lambda k=key: self.delete_category_file(k), danger=True)
            e.config(padx=12, pady=5)
            e.pack(padx=12, pady=8)

            for j, w in enumerate((a, b, c, d)):
                w.grid(row=i, column=j, sticky="nsew")
            cell.grid(row=i, column=4, sticky="nsew")
            self.file_rows[key] = (b, c, d, e)

        self.refresh_status()

    def classify_file(self, path):
        name = Path(path).name
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
            filetypes=[("Excel files", "*.xlsx *.xlsm")]
        )
        if not paths:
            return
        mapped = {}
        unknown = []
        for p in paths:
            cat = self.classify_file(p)
            if cat:
                mapped[cat] = Path(p)
            else:
                unknown.append(Path(p).name)
        if not mapped:
            messagebox.showwarning("분류 실패", "DUR 파일 종류를 판별하지 못했습니다.")
            return

        files_dir = ROOT_DIR / "files"
        items = []
        self.cfg.setdefault("_uploaded_at", {})
        for cat, p in mapped.items():
            old_path = self.cfg.get(cat)
            dest = files_dir / f"{cat}__{p.name}"
            shutil.copy2(p, dest)
            if old_path and old_path != str(dest):
                try:
                    old = Path(old_path)
                    if old.exists() and old.parent == files_dir:
                        old.unlink()
                except Exception:
                    pass
            self.cfg[cat] = str(dest)
            self.cfg["_uploaded_at"][cat] = datetime.now().isoformat(timespec="minutes")
            items.append((cat, dest))
        self.save_cfg()
        self.index_many(items)
        if unknown:
            messagebox.showwarning("일부 파일 미분류", "\n".join(unknown))

    def index_many(self, items):
        win = tk.Toplevel(self)
        win.title("인덱싱")
        win.geometry("560x220")
        win.configure(bg=SURFACE)
        win.transient(self)
        win.resizable(False, False)
        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 560) // 2
        y = self.winfo_y() + (self.winfo_height() - 220) // 2
        win.geometry(f"+{x}+{y}")

        inner = tk.Frame(win, bg=SURFACE)
        inner.pack(expand=True)
        tk.Label(inner, text="기준파일 인덱싱 중", bg=SURFACE, fg=TEXT,
                 font=self.font(13, "bold")).pack(pady=(10, 2))
        label = tk.Label(inner, text="준비 중…", bg=SURFACE, fg=MUTED, font=self.font(10))
        label.pack(pady=(4, 12))
        bar = ttk.Progressbar(inner, mode="determinate", maximum=max(1, len(items)), length=460)
        bar.pack(pady=6, padx=40)

        def worker():
            errors = []
            for i, (cat, path) in enumerate(items, 1):
                def prog(msg, cat=cat):
                    self.after(0, lambda: label.config(text=f"{dict(CATEGORIES)[cat]} · {msg}"))
                ok, msg = self.idx.index_file(cat, path, prog)
                if not ok:
                    errors.append(f"{dict(CATEGORIES)[cat]}: {msg}")
                self.after(0, lambda i=i: bar.config(value=i))
            def done():
                win.destroy()
                self.refresh_status()
                if errors:
                    messagebox.showwarning("일부 파일 확인 필요", "\n".join(errors))
                else:
                    messagebox.showinfo("완료", "기준파일 인덱싱이 완료되었습니다.")
            self.after(0, done)
        threading.Thread(target=worker, daemon=True).start()

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
        items = []
        for k, _ in CATEGORIES:
            v = self.cfg.get(k)
            if isinstance(v, str) and Path(v).exists():
                items.append((k, Path(v)))
        if not items:
            messagebox.showwarning("기준파일 없음", "먼저 Excel 파일을 등록해 주세요.")
            return
        self.index_many(items)

    def refresh_status(self):
        try:
            sts = self.idx.statuses()
        except Exception:
            sts = {}
        times = self.cfg.get("_uploaded_at", {})
        latest = max(times.values()) if times else ""
        self.header_update.config(text=(latest.replace("T", " ")[:16]) if latest else "—")
        if hasattr(self, "file_rows"):
            for key, _ in CATEGORIES:
                row_items = self.file_rows[key]
                status_lbl, date_lbl, file_lbl = row_items[:3]
                delete_btn = row_items[3] if len(row_items) > 3 else None
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
        err = traceback.format_exc()
        try:
            messagebox.showerror("DUR 프로그램 오류", err)
        except Exception:
            print(err)
