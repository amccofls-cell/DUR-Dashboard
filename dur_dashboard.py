from __future__ import annotations

import json, os, re, shutil, sqlite3, sys, threading, traceback
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import load_workbook

APP_NAME = "DUR Dashboard"
VERSION = "1.2.0"

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
    ("combo", "병용금기"), ("age", "연령금기"), ("preg", "임부금기"),
    ("lact", "수유부주의"), ("dup", "효능군중복"),
    ("tele", "비대면진료 금지"), ("cost", "비용효과적 함량"),
]

BG="#F6F3ED"; PANEL="#FFFFFF"; TEXT="#30352F"; MUTED="#74786F"; BORDER="#E4E0D8"
SAGE="#7C9278"; DEEP_SAGE="#526653"; SAGE_SOFT="#EEF3EC"
TERRA="#C9785D"; TERRA_SOFT="#F7E8E1"; NO="#777B74"; NO_BG="#F6F6F3"
WARN="#A86635"; WARN_BG="#FFF4E8"; ERROR="#A84F46"


def app_dir() -> Path:
    root = Path(os.getenv("LOCALAPPDATA") or Path.home()) / "DUR_Dashboard"
    root.mkdir(parents=True, exist_ok=True)
    (root/"files").mkdir(exist_ok=True)
    return root

ROOT_DIR=app_dir(); CFG_PATH=ROOT_DIR/"config.json"; DB_PATH=ROOT_DIR/"dur_index.sqlite3"

def norm(v):
    if v is None: return ""
    s=str(v).strip().lower()
    s=re.sub(r"\s+", "", s)
    s=s.replace("㎖","ml").replace("밀리리터","ml").replace("밀리그램","mg")
    s=re.sub(r"[()\[\]{}·,._\-/]", "", s)
    return s

def sval(v):
    if v is None: return ""
    if isinstance(v, datetime): return v.strftime("%Y-%m-%d")
    return str(v).strip()

def unique_headers(values):
    out=[]; seen={}
    for i,v in enumerate(values):
        h=sval(v) or f"col_{i+1}"
        seen[h]=seen.get(h,0)+1
        if seen[h]>1: h=f"{h}_{seen[h]}"
        out.append(h)
    return out

def find_header_row(ws, max_scan=15):
    keys=("제품명","품목명","약품명","성분명","성분명a","저함량")
    best=(1,-1)
    for r in range(1,min(ws.max_row,max_scan)+1):
        vals=[norm(c.value) for c in ws[r][:min(ws.max_column,40)]]
        score=sum(any(k in x for k in keys) for x in vals)
        if score>best[1]: best=(r,score)
    return best[0]

def detect_name_cols(headers, category):
    hn=[norm(h) for h in headers]
    if category.startswith("combo"):
        a=[i for i,h in enumerate(hn) if h in ("제품명a","품목명a") or ("제품명" in h and h.endswith("a"))]
        b=[i for i,h in enumerate(hn) if h in ("제품명b","품목명b") or ("제품명" in h and h.endswith("b"))]
        return (a[:1],b[:1])
    cand=[]
    for i,h in enumerate(hn):
        if h in ("제품명","품목명","약품명") or h.startswith("제품명") or h.startswith("품목명") or h.startswith("약품명"):
            cand.append(i)
    return (cand[:2],[])

def cost_rows(ws):
    # 비용효과 파일은 보통 3~4행 다단 헤더. 실제 데이터 시작점을 찾고 저/고함량 그룹을 보존한다.
    rows=list(ws.iter_rows(min_row=1,max_row=min(ws.max_row,8),values_only=True))
    header_row=4 if len(rows)>=4 else 1
    for idx,row in enumerate(rows,1):
        vals=[norm(x) for x in row]
        if vals.count("제품코드")>=2 or ("제품코드" in vals and "제품명" in vals and "함량" in vals): header_row=idx
    group_row=max(1,header_row-1)
    groups=list(rows[group_row-1]) if len(rows)>=group_row else []
    base=list(rows[header_row-1])
    headers=[]; current=""
    for i,h in enumerate(base):
        g=sval(groups[i]) if i<len(groups) else ""
        if g: current=g
        prefix="저함량" if "저함량" in current else ("고함량" if "고함량" in current else "")
        name=sval(h) or f"col_{i+1}"
        headers.append((prefix+" "+name).strip())
    headers=unique_headers(headers)
    for row in ws.iter_rows(min_row=header_row+1, values_only=True):
        if not any(v is not None and sval(v)!="" for v in row): continue
        yield headers,row

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
    """병용금기 대용량 파일은 A,D,F,G,J,M,N,O 열만 인덱스에 저장한다."""
    hr=find_header_row(ws)
    for row in ws.iter_rows(min_row=hr+1, min_col=1, max_col=15, values_only=True):
        if not any(row[i-1] is not None and sval(row[i-1])!="" for i in COMBO_KEEP):
            continue
        d={name:sval(row[i-1]) for i,name in COMBO_KEEP.items()}
        yield d

def normal_rows(ws):
    hr=find_header_row(ws)
    headers=unique_headers([c.value for c in ws[hr]])
    for row in ws.iter_rows(min_row=hr+1,values_only=True):
        if not any(v is not None and sval(v)!="" for v in row): continue
        yield headers,row

class Indexer:
    def __init__(self, db_path=DB_PATH): self.db_path=db_path
    def conn(self):
        c=sqlite3.connect(self.db_path); c.execute("PRAGMA journal_mode=WAL"); return c
    def init(self):
        with self.conn() as c:
            c.execute("CREATE TABLE IF NOT EXISTS records(id INTEGER PRIMARY KEY, category TEXT, sheet TEXT, product TEXT, product_norm TEXT, side TEXT, data TEXT)")
            c.execute("CREATE INDEX IF NOT EXISTS ix_records_cat_name ON records(category,product_norm)")
            c.execute("CREATE TABLE IF NOT EXISTS status(category TEXT PRIMARY KEY, ok INTEGER, message TEXT, filename TEXT, indexed_at TEXT, sheet_count INTEGER, row_count INTEGER)")
    def clear_category(self,c,cat): c.execute("DELETE FROM records WHERE category=?",(cat,)); c.execute("DELETE FROM status WHERE category=?",(cat,))
    def index_file(self, category, path, progress=None):
        self.init(); path=Path(path); total=0; sheets=0
        try:
            wb=load_workbook(path,read_only=True,data_only=True)
            with self.conn() as c:
                self.clear_category(c,category)
                for ws in wb.worksheets:
                    sheets+=1
                    if progress: progress(f"{ws.title} 시트 읽는 중…")
                    iterator=combo_rows(ws) if category.startswith("combo") else (cost_rows(ws) if category=="cost" else normal_rows(ws))
                    batch=[]
                    for item in iterator:
                        if category.startswith("combo"):
                            d=item
                            for side,p in (("A",d.get("제품명A","")),("B",d.get("제품명B",""))):
                                if p: batch.append((category,ws.title,p,norm(p),side,json.dumps(d,ensure_ascii=False)))
                            if len(batch)>=2000:
                                c.executemany("INSERT INTO records(category,sheet,product,product_norm,side,data) VALUES(?,?,?,?,?,?)",batch); total+=len(batch); batch=[]
                            continue
                        headers,row=item
                        d={headers[i]:sval(row[i]) if i<len(row) else "" for i in range(len(headers))}
                        if category=="cost":
                            # 저/고함량 제품명 각각 검색 가능하게 2개의 검색 엔트리 생성
                            low=next((d[k] for k in d if "저함량" in k and "제품명" in k),"")
                            high=next((d[k] for k in d if "고함량" in k and "제품명" in k),"")
                            for side,p in (("저함량",low),("고함량",high)):
                                if p: batch.append((category,ws.title,p,norm(p),side,json.dumps(d,ensure_ascii=False)))
                        else:
                            a,_=detect_name_cols(headers,category)
                            for i in a:
                                p=sval(row[i]) if i<len(row) else ""
                                if p: batch.append((category,ws.title,p,norm(p),"",json.dumps(d,ensure_ascii=False)))
                        if len(batch)>=2000:
                            c.executemany("INSERT INTO records(category,sheet,product,product_norm,side,data) VALUES(?,?,?,?,?,?)",batch); total+=len(batch); batch=[]
                    if batch:
                        c.executemany("INSERT INTO records(category,sheet,product,product_norm,side,data) VALUES(?,?,?,?,?,?)",batch); total+=len(batch)
                c.execute("INSERT OR REPLACE INTO status VALUES(?,?,?,?,?,?,?)",(category,1,"정상",path.name,datetime.now().isoformat(timespec="seconds"),sheets,total))
            wb.close(); return True,f"{sheets}개 시트 · {total:,}개 검색 엔트리"
        except Exception as e:
            with self.conn() as c:
                self.clear_category(c,category)
                c.execute("INSERT OR REPLACE INTO status VALUES(?,?,?,?,?,?,?)",(category,0,str(e),path.name,datetime.now().isoformat(timespec="seconds"),sheets,total))
            return False,str(e)
    def statuses(self):
        self.init()
        with self.conn() as c: return {r[0]:r[1:] for r in c.execute("SELECT category,ok,message,filename,indexed_at,sheet_count,row_count FROM status")}
    def suggestions(self,q,limit=30):
        nq=norm(q)
        if not nq:return []
        with self.conn() as c:
            rows=c.execute("SELECT product, COUNT(*) n FROM records WHERE product_norm LIKE ? GROUP BY product ORDER BY CASE WHEN product_norm LIKE ? THEN 0 ELSE 1 END, length(product), product LIMIT ?",(f"%{nq}%",f"{nq}%",limit)).fetchall()
        return [x[0] for x in rows]
    def search_product(self,product):
        np=norm(product); out={k:[] for k,_ in CATEGORIES}
        with self.conn() as c:
            for cat,_ in CATEGORIES:
                rows=c.execute("SELECT sheet,product,side,data FROM records WHERE category=? AND product_norm=?",(cat,np)).fetchall()
                out[cat]=[{"sheet":r[0],"product":r[1],"side":r[2],"data":json.loads(r[3])} for r in rows]
        return out

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(f"{APP_NAME}  {VERSION}"); self.geometry("1480x900"); self.minsize(1180,760); self.configure(bg=BG)
        self.idx=Indexer(); self.idx.init(); self.cfg=self.load_cfg(); self.results={}; self.selected_drug=None
        self.setup_style(); self.build(); self.refresh_status()
    def load_cfg(self):
        try:return json.loads(CFG_PATH.read_text(encoding="utf-8"))
        except:return {}
    def save_cfg(self): CFG_PATH.write_text(json.dumps(self.cfg,ensure_ascii=False,indent=2),encoding="utf-8")
    def ui_font(self, size=10, weight="normal"):
        # Arial 우선, 한글은 Windows font fallback으로 NanumGothic/Malgun Gothic 사용
        return ("Arial", size, weight)
    def setup_style(self):
        s=ttk.Style(self); s.theme_use("clam")
        s.configure("TFrame",background=BG); s.configure("Surface.TFrame",background=PANEL)
        s.configure("TLabel",background=BG,foreground=TEXT,font=self.ui_font(10))
        s.configure("Surface.TLabel",background=PANEL,foreground=TEXT,font=self.ui_font(10))
        s.configure("Title.TLabel",background=BG,foreground=DEEP_SAGE,font=self.ui_font(22,"bold"))
        s.configure("Muted.TLabel",background=BG,foreground=MUTED,font=self.ui_font(9))
        s.configure("TButton",font=self.ui_font(9,"bold"),padding=(13,9),background="#F1EFEA",foreground=TEXT,borderwidth=0)
        s.map("TButton",background=[("active","#E9E5DE")])
        s.configure("Primary.TButton",background=DEEP_SAGE,foreground="white",borderwidth=0)
        s.map("Primary.TButton",background=[("active",SAGE)])
        s.configure("Treeview",font=self.ui_font(10),rowheight=38,background=PANEL,fieldbackground=PANEL,foreground=TEXT,borderwidth=0)
        s.map("Treeview",background=[("selected",SAGE_SOFT)],foreground=[("selected",DEEP_SAGE)])
        s.configure("Treeview.Heading",font=self.ui_font(9,"bold"),background="#F0F2EC",foreground=DEEP_SAGE,relief="flat",padding=(8,8))
        s.configure("TNotebook",background=BG,borderwidth=0); s.configure("TNotebook.Tab",font=self.ui_font(10,"bold"),padding=(20,11),background="#ECE9E2",foreground=MUTED)
        s.map("TNotebook.Tab",background=[("selected",PANEL)],foreground=[("selected",TERRA)])
    def build(self):
        root=ttk.Frame(self); root.pack(fill="both",expand=True,padx=26,pady=22)
        top=ttk.Frame(root); top.pack(fill="x",pady=(0,16))
        ttk.Label(top,text="DUR Dashboard",style="Title.TLabel").pack(side="left")
        ttk.Label(top,text="의약품 DUR 통합 조회 · 로컬 전용",style="Muted.TLabel").pack(side="left",padx=16,pady=(10,0))
        self.global_status=ttk.Label(top,text="",style="Muted.TLabel"); self.global_status.pack(side="right",pady=(10,0))
        body=ttk.Frame(root); body.pack(fill="both",expand=True)
        self.left=ttk.Frame(body,style="Surface.TFrame",width=330); self.left.pack(side="left",fill="y",padx=(0,16)); self.left.pack_propagate(False)
        self.main=ttk.Frame(body); self.main.pack(side="left",fill="both",expand=True)
        self.build_left(); self.build_main()
    def build_left(self):
        tk.Label(self.left,text="기준 파일",bg=PANEL,fg=DEEP_SAGE,font=self.ui_font(14,"bold")).pack(anchor="w",padx=20,pady=(20,4))
        tk.Label(self.left,text="8개 Excel을 한 번에 선택하면 자동 분류 후 인덱싱합니다.",bg=PANEL,fg=MUTED,font=self.ui_font(9),wraplength=285,justify="left").pack(anchor="w",padx=20,pady=(0,14))
        ttk.Button(self.left,text="＋  DUR 파일 일괄 등록",style="Primary.TButton",command=self.choose_files).pack(fill="x",padx=18,pady=(0,14))
        self.slot_widgets={}
        for key,label in CATEGORIES:
            f=tk.Frame(self.left,bg=PANEL); f.pack(fill="x",padx=18,pady=4)
            dot=tk.Label(f,text="●",bg=PANEL,fg="#C8CBC4",font=self.ui_font(8)); dot.pack(side="left",padx=(0,8),anchor="n")
            mid=tk.Frame(f,bg=PANEL); mid.pack(side="left",fill="x",expand=True)
            tk.Label(mid,text=label,bg=PANEL,fg=TEXT,font=self.ui_font(9,"bold"),anchor="w").pack(fill="x")
            date=tk.Label(mid,text="마지막 등록 —",bg=PANEL,fg="#9A9D96",font=self.ui_font(8),anchor="w"); date.pack(fill="x",pady=(1,0))
            st=tk.Label(f,text="미등록",bg=PANEL,fg="#9A9D96",font=self.ui_font(8),anchor="e"); st.pack(side="right",padx=(6,0))
            self.slot_widgets[key]=(dot,st,date)
        ttk.Separator(self.left).pack(fill="x",padx=18,pady=14)
        self.upload_note=tk.Label(self.left,text="파일은 PC 내부에서만 처리됩니다.",bg=PANEL,fg=MUTED,font=self.ui_font(8),justify="left"); self.upload_note.pack(anchor="w",padx=20)
        ttk.Button(self.left,text="등록 파일 다시 인덱싱",command=self.reindex_all).pack(fill="x",padx=18,pady=16)
    def build_main(self):
        search=tk.Frame(self.main,bg=PANEL,highlightbackground=BORDER,highlightthickness=1); search.pack(fill="x")
        tk.Label(search,text="품목 일괄 검색",bg=PANEL,fg=TEXT,font=self.ui_font(13,"bold")).pack(anchor="w",padx=20,pady=(16,4))
        tk.Label(search,text="품목명을 한 줄에 하나씩 입력하세요. 일부 이름도 검색할 수 있습니다.",bg=PANEL,fg=MUTED,font=self.ui_font(9)).pack(anchor="w",padx=20)
        row=tk.Frame(search,bg=PANEL); row.pack(fill="x",padx=20,pady=(10,16))
        self.multi=tk.Text(row,height=4,font=self.ui_font(11),relief="flat",bg=NO_BG,fg=TEXT,insertbackground="#101828",padx=12,pady=10,highlightthickness=1,highlightbackground=BORDER); self.multi.pack(side="left",fill="x",expand=True)
        self.multi.insert("1.0","예)\n마운자로\n쎄레빅스\n암로젯")
        self.multi.bind("<FocusIn>",self.clear_example)
        ttk.Button(row,text="통합 조회",style="Primary.TButton",command=self.batch_search).pack(side="left",padx=(12,0),fill="y")
        self.summary=tk.Label(self.main,text="검색할 품목을 입력해 주세요.",bg=BG,fg=MUTED,font=self.ui_font(9),anchor="w"); self.summary.pack(fill="x",pady=(12,8))
        self.tabs=ttk.Notebook(self.main); self.tabs.pack(fill="both",expand=True)
        self.by_drug=ttk.Frame(self.tabs,style="Surface.TFrame"); self.by_dur=ttk.Frame(self.tabs,style="Surface.TFrame")
        self.tabs.add(self.by_drug,text="약품별 결과"); self.tabs.add(self.by_dur,text="DUR 종류별 결과")
        self.build_drug_tab(); self.build_dur_tab()
    def clear_example(self,_=None):
        if self.multi.get("1.0","end").strip().startswith("예)"): self.multi.delete("1.0","end")
    def build_drug_tab(self):
        pane=tk.PanedWindow(self.by_drug,orient="horizontal",bg="#EAECF0",sashwidth=1,bd=0); pane.pack(fill="both",expand=True)
        left=tk.Frame(pane,bg=PANEL,width=280); right=tk.Frame(pane,bg=PANEL); pane.add(left,minsize=230); pane.add(right,minsize=600)
        tk.Label(left,text="검색 품목",bg=PANEL,fg=TEXT,font=self.ui_font(10,"bold")).pack(anchor="w",padx=16,pady=(14,8))
        self.drug_list=tk.Listbox(left,font=self.ui_font(9),relief="flat",bd=0,highlightthickness=0,selectbackground=SAGE_SOFT,selectforeground=DEEP_SAGE,activestyle="none"); self.drug_list.pack(fill="both",expand=True,padx=8,pady=(0,8)); self.drug_list.bind("<<ListboxSelect>>",self.select_drug)
        self.drug_title=tk.Label(right,text="품목을 선택해 주세요",bg=PANEL,fg=TEXT,font=self.ui_font(17,"bold")); self.drug_title.pack(anchor="w",padx=18,pady=(16,10))
        self.cards_frame=tk.Frame(right,bg=PANEL); self.cards_frame.pack(fill="x",padx=18); self.cards={}
        for i,(key,label) in enumerate(DUR_CARDS):
            f=tk.Frame(self.cards_frame,bg=NO_BG,highlightbackground=BORDER,highlightthickness=1,height=74); f.grid(row=i//4,column=i%4,sticky="nsew",padx=(0 if i%4==0 else 7,0),pady=(0,7)); f.grid_propagate(False)
            tk.Label(f,text=label,bg=NO_BG,fg=MUTED,font=self.ui_font(8)).pack(anchor="w",padx=11,pady=(10,2))
            val=tk.Label(f,text="—",bg=NO_BG,fg=MUTED,font=self.ui_font(10,"bold")); val.pack(anchor="w",padx=11)
            f.bind("<Button-1>",lambda e,k=key:self.show_detail(k)); val.bind("<Button-1>",lambda e,k=key:self.show_detail(k)); self.cards[key]=(f,val)
        for i in range(4): self.cards_frame.columnconfigure(i,weight=1)
        self.detail_title=tk.Label(right,text="상세정보",bg=PANEL,fg=TEXT,font=self.ui_font(10,"bold")); self.detail_title.pack(anchor="w",padx=18,pady=(8,6))
        self.detail=ttk.Treeview(right,columns=("항목","내용"),show="headings"); self.detail.heading("항목",text="항목"); self.detail.heading("내용",text="내용"); self.detail.column("항목",width=175,anchor="w"); self.detail.column("내용",width=650,anchor="w"); self.detail.pack(fill="both",expand=True,padx=18,pady=(0,16))
    def build_dur_tab(self):
        tk.Label(self.by_dur,text="DUR 종류별 해당 품목",bg=PANEL,fg=TEXT,font=self.ui_font(13,"bold")).pack(anchor="w",padx=18,pady=(16,8))
        self.dur_tree=ttk.Treeview(self.by_dur,columns=("DUR","품목","판정","상세 요약"),show="headings")
        for c,w in (("DUR",180),("품목",300),("판정",130),("상세 요약",600)):
            self.dur_tree.heading(c,text=c); self.dur_tree.column(c,width=w,anchor="w")
        self.dur_tree.pack(fill="both",expand=True,padx=18,pady=(0,18))
    def classify_file(self,path):
        n=norm(Path(path).name)
        # 파일명 우선. 급여/비급여는 서로 독립적으로 분류.
        if "병용금기" in n:
            if "비급여" in n:return "combo_unpaid"
            if "급여" in n:return "combo_paid"
        if "연령금기" in n:return "age"
        if "임부금기" in n:return "preg"
        if "수유부주의" in n:return "lact"
        if "효능군중복" in n:return "dup"
        if "비대면진료" in n:return "tele"
        if "비용효과" in n:return "cost"
        # 이름이 불명확하면 시트명/초기 헤더로 보조 판별.
        try:
            wb=load_workbook(path,read_only=True,data_only=True)
            sig=norm(" ".join(wb.sheetnames))
            ws=wb[wb.sheetnames[0]]; vals=[]
            for row in ws.iter_rows(min_row=1,max_row=min(8,ws.max_row),values_only=True): vals += [sval(x) for x in row if x is not None]
            sig += norm(" ".join(vals[:120])); wb.close()
            if "저함량" in sig and "고함량" in sig:return "cost"
            if "효능군중복점검코드" in sig:return "dup"
            if "특정연령" in sig:return "age"
            if "금기등급" in sig:return "preg"
            if "주성분코드" in sig and "공고번호" in sig:return "lact"
            if "성분명a" in sig and "성분명b" in sig:
                return "combo_unpaid" if "비급여" in sig else "combo_paid"
            if "약품구분" in sig and "비고" in sig:return "tele"
        except: pass
        return None
    def choose_files(self):
        paths=filedialog.askopenfilenames(title="DUR 품목리스트 Excel 선택 (최대 8개)",filetypes=[("Excel","*.xlsx")])
        if not paths:return
        classified={}; unknown=[]; dup=[]
        for p in paths:
            k=self.classify_file(p)
            if not k: unknown.append(Path(p).name)
            elif k in classified: dup.append(Path(p).name)
            else: classified[k]=p
        if unknown or dup:
            msg=[]
            if unknown:msg.append("자동 분류 실패:\n- "+"\n- ".join(unknown))
            if dup:msg.append("같은 종류로 중복 분류:\n- "+"\n- ".join(dup))
            messagebox.showwarning("일부 파일 확인 필요","\n\n".join(msg)+"\n\n분류된 파일은 계속 등록합니다.")
        if not classified:return
        for k,p in classified.items():
            dest=ROOT_DIR/"files"/(k+"__"+Path(p).name)
            for old in (ROOT_DIR/"files").glob(k+"__*"):
                try:old.unlink()
                except:pass
            shutil.copy2(p,dest); self.cfg[k]=str(dest); self.cfg.setdefault("_uploaded_at",{})[k]=datetime.now().isoformat(timespec="minutes")
        self.save_cfg(); self.index_many([(k,Path(self.cfg[k])) for k in classified])
    def index_many(self,items):
        self.upload_note.config(text=f"{len(items)}개 파일 자동 분류 완료 · 인덱싱 중…",fg=DEEP_SAGE)
        def worker():
            errors=[]
            for k,p in items:
                self.after(0,lambda kk=k:self.set_slot_working(kk))
                ok,msg=self.idx.index_file(k,p)
                if not ok: errors.append(f"{dict(CATEGORIES)[k]}: {msg}")
                self.after(0,self.refresh_status)
            self.after(0,lambda:self.index_many_done(errors,len(items)))
        threading.Thread(target=worker,daemon=True).start()
    def set_slot_working(self,k):
        dot,st,date=self.slot_widgets[k]; dot.config(fg=DEEP_SAGE); st.config(text="인덱싱 중",fg=DEEP_SAGE)
    def index_many_done(self,errors,n):
        self.refresh_status(); self.upload_note.config(text=f"{n}개 파일 등록/인덱싱 완료" if not errors else "일부 파일 인덱싱 오류",fg=DEEP_SAGE if not errors else "#B42318")
        messagebox.showinfo("완료" if not errors else "확인 필요", "DUR 기준파일 인덱싱이 완료되었습니다." if not errors else "\n".join(errors))
    def refresh_status(self):
        sts=self.idx.statuses(); good=0; upload_times=self.cfg.get("_uploaded_at",{}); recent=[]
        for key,label in CATEGORIES:
            dot,st,date=self.slot_widgets[key]; v=sts.get(key); stamp=upload_times.get(key)
            date.config(text="마지막 등록 "+(stamp.replace("T"," ")[:16] if stamp else "—"))
            if stamp: recent.append(stamp)
            if v and v[0]:
                good+=1; dot.config(fg=SAGE); st.config(text=f"정상 · {v[4]}시트",fg=DEEP_SAGE)
            elif v: dot.config(fg=ERROR); st.config(text="오류",fg=ERROR)
            else: dot.config(fg="#C8CBC4"); st.config(text="미등록",fg="#9A9D96")
        latest=max(recent).replace("T"," ")[:16] if recent else "—"
        self.global_status.config(text=f"기준파일 {good}/8 정상   ·   최종 등록 {latest}")
    def reindex_all(self):
        items=[(k,Path(v)) for k,v in self.cfg.items() if k in dict(CATEGORIES) and isinstance(v,str) and Path(v).exists()]
        if not items:return messagebox.showwarning("기준파일 없음","먼저 Excel 파일을 등록해 주세요.")
        self.index_many(items)
    def batch_search(self):
        raw=self.multi.get("1.0","end").strip(); queries=[x.strip() for x in raw.splitlines() if x.strip() and x.strip()!="예)"]
        # comma-separated input also accepted
        q2=[]
        for q in queries:q2 += [z.strip() for z in re.split(r"[,;]",q) if z.strip()]
        queries=list(dict.fromkeys(q2))
        if not queries:return messagebox.showinfo("입력 필요","검색할 품목명을 한 줄에 하나씩 입력해 주세요.")
        self.results={}; unresolved=[]
        for q in queries:
            vals=self.idx.suggestions(q,30)
            if not vals: self.results[q]={"query":q,"matched":None,"data":None}; unresolved.append(q); continue
            nq=norm(q); exact=[v for v in vals if norm(v)==nq]
            matched=exact[0] if exact else vals[0]
            self.results[q]={"query":q,"matched":matched,"data":self.idx.search_product(matched),"candidates":vals}
        self.render_batch(); self.summary.config(text=f"{len(queries)}개 입력 · {len(queries)-len(unresolved)}개 품목 후보 확인 · {len(unresolved)}개 미확인")
    def render_batch(self):
        self.drug_list.delete(0,"end")
        for q,r in self.results.items():
            label=("✓ " if r["matched"] else "? ")+(r["matched"] or q)
            if r["matched"] and norm(r["matched"])!=norm(q): label += f"   ← {q}"
            self.drug_list.insert("end",label)
        self.render_dur_view()
        if self.results:self.drug_list.selection_set(0); self.select_drug()
    def select_drug(self,_=None):
        sel=self.drug_list.curselection()
        if not sel:return
        q=list(self.results.keys())[sel[0]]; r=self.results[q]; self.selected_drug=q
        if not r["matched"]:
            self.drug_title.config(text=q+" · 검색 결과 없음"); self.paint_empty_cards("확인불가"); self.detail.delete(*self.detail.get_children()); self.detail.insert("","end",values=("결과","등록된 DUR 품목명에서 후보를 찾지 못했습니다.")); return
        self.drug_title.config(text=r["matched"]); self.paint_cards(r["data"]); self.show_detail(next((k for k,_ in DUR_CARDS if self.has_hits(r["data"],k)),"combo"))
    def mapping(self):return {"combo":["combo_paid","combo_unpaid"],"age":["age"],"preg":["preg"],"lact":["lact"],"dup":["dup"],"tele":["tele"],"cost":["cost"]}
    def has_hits(self,data,key):return any(data.get(c) for c in self.mapping()[key])
    def paint_empty_cards(self,text):
        for key,_ in DUR_CARDS:self.set_card(key,text,WARN_BG,WARN)
    def set_card(self,key,text,bg,fg):
        f,val=self.cards[key]; f.config(bg=bg); val.config(text=text,bg=bg,fg=fg)
        for w in f.winfo_children():
            if isinstance(w,tk.Label):w.config(bg=bg)
    def find_value(self,d,keys):
        for target in keys:
            for k,v in d.items():
                if norm(target)==norm(k) and v:return v
        return ""
    def verdict(self,data,key):
        sts=self.idx.statuses(); cats=self.mapping()[key]; complete=all(sts.get(c) and sts[c][0] for c in cats); hits=sum(len(data[c]) for c in cats)
        if not complete:return "확인불가"
        if not hits:return "— 해당 없음"
        if key=="preg":
            grades=sorted({self.find_value(x["data"],["금기등급","등급"]) for x in data["preg"] if self.find_value(x["data"],["금기등급","등급"])})
            return "○ "+("/".join(g+"등급" for g in grades) if grades else "해당")
        if key=="age":
            x=data["age"][0]["data"]; age=self.find_value(x,["특정연령"]); unit=self.find_value(x,["연령단위"]); cond=self.find_value(x,["연령처리조건"])
            return "○ "+(" ".join(z for z in [age+unit if age else "",cond] if z) or "해당")
        return f"○ 해당 ({hits})" if hits>1 else "○ 해당"
    def paint_cards(self,data):
        for key,_ in DUR_CARDS:
            v=self.verdict(data,key)
            if v.startswith("○"): self.set_card(key,v,SAGE_SOFT,DEEP_SAGE)
            elif v=="확인불가": self.set_card(key,v,WARN_BG,WARN)
            else:self.set_card(key,v,NO_BG,NO)
    def visible_field(self,key):
        """조회 화면에서는 코드류와 업체명은 숨긴다. 검색/식별용 내부 데이터는 유지 가능."""
        n=norm(key)
        hidden=("성분코드","주성분코드","일반명코드","제품코드","약품코드","업체명")
        return not any(norm(x) in n for x in hidden)
    def show_detail(self,key):
        if self.selected_drug is None:return
        r=self.results.get(self.selected_drug); self.detail.delete(*self.detail.get_children()); self.detail_title.config(text=dict(DUR_CARDS)[key]+" 상세정보")
        if not r or not r["data"]:return
        rows=[]
        for cat in self.mapping()[key]:
            for hit in r["data"][cat]:
                d=hit["data"]
                if key=="combo":
                    other="B" if hit["side"]=="A" else "A"
                    rows.append(("상대 품목",self.find_value(d,[f"제품명{other}",f"품목명{other}"]) or "—"))
                    rows.append(("상대 성분",self.find_value(d,[f"성분명{other}"]) or "—"))
                    rows.append(("급여 구분",self.find_value(d,["급여여부A"]) or dict(CATEGORIES)[cat].replace("병용금기 ","")))
                    rows.append(("고시번호",self.find_value(d,["고시번호"]) or "—"))
                    rows.append(("고시일자",self.find_value(d,["고시일자"]) or "—"))
                    rows.append(("상세내용",self.find_value(d,["상세정보"]) or "파일에서 확인되지 않음"))
                else:
                    for k,v in d.items():
                        if v and self.visible_field(k) and norm(k) not in (norm("원본시트"),norm("원본파일명"),norm("기준년월")):
                            rows.append((k,v))
                rows.append(("",""))
        if not rows:rows=[("결과",self.verdict(r["data"],key))]
        for a,b in rows:self.detail.insert("","end",values=(a,b))
    def render_dur_view(self):
        self.dur_tree.delete(*self.dur_tree.get_children())
        for key,label in DUR_CARDS:
            for q,r in self.results.items():
                if not r["data"]:continue
                v=self.verdict(r["data"],key)
                if v.startswith("○"):
                    details=[]
                    for cat in self.mapping()[key]:
                        for h in r["data"][cat]:
                            d=h["data"]
                            if key=="combo":
                                side=h["side"]; other="B" if side=="A" else "A"; otherp=self.find_value(d,[f"제품명{other}",f"품목명{other}"]); others=self.find_value(d,[f"성분명{other}"])
                                if otherp or others:details.append(" / ".join(x for x in [otherp,others] if x))
                            elif key=="preg":details.append("등급 "+self.find_value(d,["금기등급","등급"]))
                            elif key=="age":details.append(" ".join(x for x in [self.find_value(d,["특정연령"])+self.find_value(d,["연령단위"]),self.find_value(d,["연령처리조건"])] if x))
                            else:details.append(h["sheet"])
                    summary="; ".join(dict.fromkeys(x for x in details if x))[:500]
                    self.dur_tree.insert("","end",values=(label,r["matched"],v,summary))
        if not self.dur_tree.get_children(): self.dur_tree.insert("","end",values=("—","해당 품목 없음","—","현재 검색 품목 중 DUR 해당 결과가 없습니다."))

if __name__=="__main__":
    try: App().mainloop()
    except Exception:
        err=traceback.format_exc()
        try: messagebox.showerror("DUR Dashboard 오류",err)
        except: print(err)
