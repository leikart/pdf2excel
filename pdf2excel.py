"""
PDF 轉 Excel 通用工具 v4
- 雙模式：自動表格偵測 / 文字座標精準解析
- 座標模式：自動從標題列計算欄界、合併同列多欄內容
- 浮水印過濾、重複標題刪除、錨點欄續行合併
- 合併儲存格（出貨日期/星期同天合併）
- 預覽 + 欄位名稱調整 + 另存選項
- 批次轉換
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, os, re

try:
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable,"-m","pip","install","pdfplumber","openpyxl","-q"])
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell

BG="#12121F"; BG2="#1A1A2E"; CARD="#0D2137"; ACCENT="#E94560"
ACCENT2="#6C3FC4"; TEXT="#F0F0F0"; TEXT2="#8FA8C8"
SUCCESS="#3EC97A"; ERROR="#E94560"; WARNING="#F5A623"
ACCENT_HOVER="#FF6B81"; CARD2="#162840"

DEFAULT_SETTINGS = {
    "mode": "auto",
    "watermark_keywords": "",
    "anchor_col": "0",
    "skip_duplicate_header": True,
    "col_bounds": "",
    "merge_col_range": "",
}

def clean(v):
    if v is None: return ""
    return re.sub(r"\s+"," ",str(v)).strip()

def rows_equal(a, b):
    ca=[clean(x) for x in (a or [])]
    cb=[clean(x) for x in (b or [])]
    return ca==cb and any(ca)

def normalize(table):
    return [[clean(v) for v in row] for row in table if any(clean(v) for v in row)]

# ══════════════════════════════════════════
#  模式一：自動表格偵測
# ══════════════════════════════════════════
T_LINES={"vertical_strategy":"lines","horizontal_strategy":"lines",
         "snap_tolerance":3,"join_tolerance":3,"edge_min_length":3,
         "min_words_vertical":1,"min_words_horizontal":1,
         "intersection_tolerance":3,"text_tolerance":3}
T_TEXT={**T_LINES,"vertical_strategy":"text","horizontal_strategy":"text"}

def parse_auto(pdf_path, settings, log_cb=None):
    wm=[k.strip() for k in settings["watermark_keywords"].split(",") if k.strip()]
    skip_dup=settings.get("skip_duplicate_header",True)
    strategy="lines"; all_rows=[]; header=None

    def _flat(row): return " ".join(v.replace("\n"," ") if v else "" for v in row)
    def _is_wm(row): return any(k in _flat(row) for k in wm) if wm else False
    def _is_hdr(row):
        t=_flat(row)
        return ("出貨單編號" in t and "出貨日期" in t) or ("出貨日期" in t and "星期" in t and "位置" in t)

    with pdfplumber.open(pdf_path) as pdf:
        if log_cb: log_cb(f"  共 {len(pdf.pages)} 頁（自動模式）")
        for pn,page in enumerate(pdf.pages,1):
            if log_cb: log_cb(f"  解析第 {pn} 頁...")
            tbl_objs=page.find_tables(T_LINES)
            if not tbl_objs:
                tbl_objs=page.find_tables(T_TEXT); strategy="text"
            else: strategy="lines"

            if tbl_objs:
                words=page.extract_words(x_tolerance=3,y_tolerance=3)
                line_map={}
                for w in words:
                    y=round(w["top"]/3)*3
                    line_map.setdefault(y,[]).append(w)
                for tbl in tbl_objs:
                    for tbl_row in tbl.rows:
                        row_data=[]
                        for cell in tbl_row.cells:
                            if cell is None: row_data.append(""); continue
                            x0,y0,x1,y1=cell
                            cell_lines={}
                            for y,ws in line_map.items():
                                for w in ws:
                                    if y0<=w["top"]<=y1 and x0<=w["x0"]<=x1:
                                        cell_lines.setdefault(round(w["top"],1),[]).append(w["text"])
                            text="\n".join(" ".join(cell_lines[y]) for y in sorted(cell_lines))
                            row_data.append(text)
                        if not any(row_data): continue
                        if _is_wm(row_data): continue
                        if _is_hdr(row_data):
                            if header is None:
                                header=[clean(v) for v in row_data]; all_rows.append(header)
                            elif not skip_dup: all_rows.append([clean(v) for v in row_data])
                            continue
                        all_rows.append(row_data)
            else:
                tables=page.extract_tables(T_TEXT)
                for t in (tables or []):
                    for row in t:
                        r=[clean(v) for v in row]
                        if not any(r): continue
                        if _is_wm(r): continue
                        if _is_hdr(r):
                            if header is None: header=r; all_rows.append(r)
                            elif not skip_dup: all_rows.append(r)
                            continue
                        all_rows.append(r)
    return all_rows, strategy

# ══════════════════════════════════════════
#  模式二：文字座標精準解析
# ══════════════════════════════════════════
def _calc_bounds_from_header(header_words):
    """從標題列文字位置自動計算欄界（取相鄰欄的中間點）"""
    sorted_h = sorted(header_words, key=lambda w: w["x0"])
    bounds = [0.0]
    for i in range(len(sorted_h)-1):
        mid = (sorted_h[i]["x1"] + sorted_h[i+1]["x0"]) / 2
        bounds.append(round(mid, 1))
    bounds.append(9999.0)
    return bounds

def _get_col(x, bounds):
    for i in range(len(bounds)-1):
        if bounds[i] <= x < bounds[i+1]: return i
    return len(bounds)-2

def _is_header(line_text, hdr_keys):
    return sum(1 for k in hdr_keys if k in line_text) >= 2

def _is_watermark(line_text, wm_keys):
    return any(k in line_text for k in wm_keys)

def parse_coords(pdf_path, settings, log_cb=None):
    wm  = [k.strip() for k in settings["watermark_keywords"].split(",") if k.strip()]
    skip_dup = settings.get("skip_duplicate_header", True)
    anchor   = int(settings.get("anchor_col","0") or 0)

    # 手動欄界
    cb_str = settings.get("col_bounds","").strip()
    manual_bounds = None
    if cb_str:
        try:
            nums = [float(x) for x in cb_str.split(",") if x.strip()]
            manual_bounds = sorted(nums) + [9999.0]
        except: pass

    all_rows=[]; bounds=manual_bounds; hdr_keys=[]; col_count=0; header_row=None

    with pdfplumber.open(pdf_path) as pdf:
        if log_cb: log_cb(f"  共 {len(pdf.pages)} 頁（座標模式）")
        for pn,page in enumerate(pdf.pages,1):
            if log_cb: log_cb(f"  解析第 {pn} 頁...")
            words = page.extract_words(x_tolerance=3,y_tolerance=3,
                                       keep_blank_chars=False,use_text_flow=False)
            if not words: continue

            # 依 y 分組
            lines={}
            for w in words:
                y=round(w["top"]/4)*4
                lines.setdefault(y,[]).append(w)

            for y in sorted(lines.keys()):
                ws=sorted(lines[y],key=lambda w:w["x0"])
                lt=" ".join(w["text"] for w in ws)

                # 過濾浮水印
                if wm and _is_watermark(lt, wm): continue

                # 偵測標題列：自動計算欄界
                if bounds is None or (_is_header(lt, hdr_keys if hdr_keys else ["出貨","日期","位置","編號"])):
                    if bounds is None:
                        bounds = _calc_bounds_from_header(ws)
                        col_count = len(bounds)-1
                        hdr_keys  = [w["text"] for w in ws]
                        if log_cb: log_cb(f"  偵測到 {col_count} 欄，欄界：{[round(b) for b in bounds[:-1]]}")
                    # 標題列：第一次加入，之後跳過（skip_dup）
                    if header_row is None:
                        header_row = [w["text"] for w in ws]
                        # 展開成欄位陣列
                        row=[""]*(col_count or len(ws))
                        for w in ws:
                            ci=_get_col(w["x0"],bounds)
                            if ci<len(row): row[ci]=(row[ci]+" "+w["text"]).strip()
                        all_rows.append(row)
                    elif not skip_dup:
                        row=[""]*(col_count)
                        for w in ws:
                            ci=_get_col(w["x0"],bounds)
                            if ci<len(row): row[ci]=(row[ci]+" "+w["text"]).strip()
                        all_rows.append(row)
                    continue

                if bounds is None: continue

                # 一般資料列
                row=[""]*(col_count or 8)
                for w in ws:
                    ci=_get_col(w["x0"],bounds)
                    if ci<len(row): row[ci]=(row[ci]+" "+w["text"]).strip()
                if any(row): all_rows.append(row)

    # 合併指定欄（如把「星期欄」和「位置欄」分開時位置被切兩半）
    merge_range = settings.get("merge_col_range","").strip()
    if merge_range:
        try:
            parts=merge_range.split("-")
            mc_start=int(parts[0]); mc_end=int(parts[1])
            merged_rows=[]
            for row in all_rows:
                new_row=list(row[:mc_start])
                combined=" ".join(row[ci] for ci in range(mc_start,min(mc_end+1,len(row))) if row[ci])
                new_row.append(combined)
                new_row.extend(row[mc_end+1:])
                merged_rows.append(new_row)
            all_rows=merged_rows
        except: pass

    # 處理錨點欄續行
    result=_merge_continuation(all_rows, anchor)
    return result, "coords"

def _merge_continuation(rows, anchor_col):
    """錨點欄空白的列 → 合併到上一筆的備註欄"""
    if not rows: return []
    result=[]
    for row in rows:
        av=row[anchor_col] if anchor_col<len(row) else ""
        # 判斷是否為標題列
        is_hdr=any(k in " ".join(row) for k in ["出貨單編號","出貨日期","位置","數量","進場時間"])
        if is_hdr and not result:
            result.append(row); continue
        if not av and result and not is_hdr:
            extra=" ".join(v for v in row if v)
            if extra:
                last=result[-1]
                note=last[-1] if last else ""
                result[-1]=last[:-1]+[((note+"\n"+extra).strip() if note else extra)]
        else:
            result.append(row)
    return result

def _merge_tables(all_tables, settings):
    skip=settings.get("skip_duplicate_header",True)
    merged=[]; header=None
    for t in all_tables:
        if not t: continue
        if header is None:
            header=t[0]; merged.extend(t)
        else:
            start=1 if (skip and rows_equal(t[0],header)) else 0
            merged.extend(t[start:])
    return merged

# ══════════════════════════════════════════
#  合併儲存格偵測
# ══════════════════════════════════════════
def detect_merges(table, merge_cols=None):
    """
    偵測需要合併的欄位範圍。
    merge_cols: 指定欄位索引清單，只合併這些欄（預設只合併出貨日期=1和星期=2）
    若為 None 且無法判斷，預設只合併索引1和2。
    """
    if len(table)<2: return {}
    data=table[1:]; info={}
    # 預設只合併出貨日期（欄1）和星期（欄2）
    if merge_cols is None:
        # 嘗試從標題判斷日期/星期欄位
        header=table[0]
        date_cols=[]
        for ci,h in enumerate(header):
            if any(k in str(h) for k in ["日期","星期","週","week","date","Date"]):
                date_cols.append(ci)
        merge_cols=date_cols if date_cols else [1,2]
    for ci in merge_cols:
        ranges=[]; i=0
        while i<len(data):
            val=data[i][ci] if ci<len(data[i]) else ""
            if val:
                j=i+1
                while j<len(data):
                    nv=data[j][ci] if ci<len(data[j]) else ""
                    if nv=="": j+=1   # 空白代表同一天，繼續合併
                    else: break        # 有新值代表換天，停止
                if j-i>1: ranges.append((i,j-1))
                i=j
            else: i+=1
        if ranges: info[ci]=ranges
    return info

# ══════════════════════════════════════════
#  寫入 Excel
# ══════════════════════════════════════════
def write_excel(table, merge_info, out_path, col_map=None):
    wb=Workbook(); ws=wb.active; ws.title="轉換結果"
    hfill=PatternFill(start_color="0F3460",end_color="0F3460",fill_type="solid")
    hfont=Font(color="FFFFFF",bold=True,size=11)
    mfill=PatternFill(start_color="D6E4F0",end_color="D6E4F0",fill_type="solid")
    o_fill=PatternFill(start_color="F5F9FF",end_color="F5F9FF",fill_type="solid")
    e_fill=PatternFill(start_color="FFFFFF",end_color="FFFFFF",fill_type="solid")
    thin=Side(style="thin")
    bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    ctr=Alignment(horizontal="center",vertical="top",wrap_text=True)
    lwrap=Alignment(horizontal="left",vertical="top",wrap_text=True)
    if not table: wb.save(out_path); return
    ncols=max(len(r) for r in table)

    # 欄寬：取各欄最長內容（換行只算最長的一段），上限50
    col_max=[12]*ncols
    for row in table:
        for ci,v in enumerate(row):
            if v:
                # 換行內容取最長一行來計算欄寬
                max_line=max((len(line) for line in str(v).split("\n")),default=0)
                col_max[ci]=min(max(col_max[ci], max_line), 50)
    for ci in range(ncols):
        ws.column_dimensions[get_column_letter(ci+1)].width=col_max[ci]+3

    drc=0
    for ri,row in enumerate(table):
        er=ri+1; is_hdr=(ri==0)
        max_lines=1  # 這一列最多幾行
        for ci in range(ncols):
            val=row[ci] if ci<len(row) else ""
            if is_hdr and col_map and ci in col_map: val=col_map[ci]
            # 保留換行符（\n → Excel 換行）
            cell=ws.cell(row=er,column=ci+1,value=val)
            cell.border=bdr
            cell.alignment=ctr if is_hdr else lwrap
            if is_hdr:
                cell.fill=hfill; cell.font=hfont
            else:
                cell.fill=o_fill if drc%2==0 else e_fill
                cell.font=Font(size=10)
                if val: max_lines=max(max_lines, str(val).count("\n")+1)
        if not is_hdr:
            # 列高：依各欄換行數計算，每行15pt，最少22，最多300
            ws.row_dimensions[er].height=min(max(22, 15*max_lines), 300)
            drc+=1
        else:
            ws.row_dimensions[er].height=24
    ws.freeze_panes="A2"
    for ci,ranges in merge_info.items():
        for (si,ei) in ranges:
            sr=si+2; er2=ei+2
            if sr==er2: continue
            try:
                ws.merge_cells(start_row=sr,start_column=ci+1,end_row=er2,end_column=ci+1)
                cell=ws.cell(row=sr,column=ci+1)
                cell.alignment=ctr; cell.fill=mfill
                cell.font=Font(size=10,bold=True,color="1F3864"); cell.border=bdr
            except: pass
    wb.save(out_path)

def convert(pdf_path, out_path, settings, log_cb=None):
    mode=settings.get("mode","auto")
    if mode=="coords":
        table,strategy=parse_coords(pdf_path,settings,log_cb)
    elif mode=="table":
        table,strategy=parse_auto(pdf_path,settings,log_cb)
    else:
        table,strategy=parse_auto(pdf_path,settings,log_cb)
        if not table or len(table)<3:
            if log_cb: log_cb("  表格偵測結果少，切換座標模式...")
            table,strategy=parse_coords(pdf_path,settings,log_cb)
    if log_cb: log_cb(f"  策略：{strategy}，{len(table)} 列（含標題）")
    mi=detect_merges(table)
    if out_path: write_excel(table,mi,out_path)
    return table, strategy

# ══════════════════════════════════════════
#  GUI
# ══════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 轉 Excel 通用工具 v4")
        self.geometry("1100x780"); self.minsize(960,660)
        self.configure(bg=BG)
        self.pdf_files=[]; self.running=False
        self.settings=dict(DEFAULT_SETTINGS)
        self.out_dir=tk.StringVar(value=os.path.expanduser("~\\Desktop"))
        self._build()

    def _build(self):
        hdr=tk.Frame(self,bg=CARD,height=68); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr,text="PDF  →  Excel",font=("Segoe UI",22,"bold"),bg=CARD,fg="white").pack(side="left",padx=24,pady=12)
        tk.Label(hdr,text="通用版 v4・雙模式・進階設定・合併儲存格・換行保留",font=("Segoe UI",11),bg=CARD,fg=TEXT2).pack(side="left")

        main=tk.Frame(self,bg=BG); main.pack(fill="both",expand=True,padx=16,pady=10)

        # 左：檔案清單
        left=tk.Frame(main,bg=BG); left.pack(side="left",fill="both",expand=True)
        r1=tk.Frame(left,bg=BG); r1.pack(fill="x",pady=(0,4))
        tk.Label(r1,text="PDF 檔案清單",font=("Segoe UI",12,"bold"),bg=BG,fg=TEXT).pack(side="left")
        tk.Button(r1,text="＋ 新增",font=("Segoe UI",11),bg=ACCENT,fg="white",relief="flat",cursor="hand2",padx=14,pady=4,command=self.add_files).pack(side="right",padx=(4,0))
        tk.Button(r1,text="清除全部",font=("Segoe UI",11),bg=BG2,fg=TEXT2,relief="flat",cursor="hand2",padx=14,pady=4,command=self.clear_files).pack(side="right")
        lf=tk.Frame(left,bg=CARD); lf.pack(fill="both",expand=True)
        sb=tk.Scrollbar(lf); sb.pack(side="right",fill="y")
        self.listbox=tk.Listbox(lf,yscrollcommand=sb.set,bg=CARD,fg=TEXT,selectbackground=ACCENT2,font=("Segoe UI",11),relief="flat",bd=0,highlightthickness=0,activestyle="none")
        self.listbox.pack(fill="both",expand=True,padx=2,pady=2)
        sb.config(command=self.listbox.yview); self.listbox.bind("<Delete>",self.remove_sel)
        tk.Label(left,text="選中後按 Delete 可移除",font=("Segoe UI",8),bg=BG,fg=TEXT2).pack(anchor="w",pady=(3,0))

        # 右：設定面板
        right=tk.Frame(main,bg=BG,width=320); right.pack(side="right",fill="y",padx=(16,0)); right.pack_propagate(False)

        # 輸出資料夾
        tk.Label(right,text="輸出資料夾",font=("Segoe UI",11,"bold"),bg=BG,fg=TEXT).pack(anchor="w")
        dr=tk.Frame(right,bg=BG); dr.pack(fill="x",pady=(3,8))
        tk.Entry(dr,textvariable=self.out_dir,bg=CARD,fg=TEXT,insertbackground="white",relief="flat",font=("Segoe UI",10)).pack(side="left",fill="x",expand=True,ipady=5,padx=(0,4))
        tk.Button(dr,text="瀏覽",font=("Segoe UI",10),bg=BG2,fg=TEXT2,relief="flat",cursor="hand2",padx=10,pady=3,command=self.browse_dir).pack(side="right")

        # 進階設定
        adv=tk.LabelFrame(right,text=" ⚙  進階設定 ",font=("Segoe UI",11,"bold"),bg=BG,fg=TEXT2,bd=1,relief="groove")
        adv.pack(fill="x",pady=(0,8))

        # 解析模式
        tk.Label(adv,text="解析模式",font=("Segoe UI",9),bg=BG,fg=TEXT2).pack(anchor="w",padx=8,pady=(6,2))
        self.mode_var=tk.StringVar(value="auto")
        mf=tk.Frame(adv,bg=BG); mf.pack(fill="x",padx=8,pady=(0,6))
        for val,lbl,tip in [("auto","自動","先試框線，失敗改座標"),
                            ("table","表格偵測","有框線的PDF"),
                            ("coords","文字座標","無框線/特殊版面")]:
            col=tk.Frame(mf,bg=BG); col.pack(side="left",padx=(0,8))
            tk.Radiobutton(col,text=lbl,variable=self.mode_var,value=val,
                           bg=BG,fg=TEXT,selectcolor=BG2,activebackground=BG,
                           font=("Segoe UI",11),command=self._on_mode).pack()
            tk.Label(col,text=tip,font=("Segoe UI",8),bg=BG,fg=TEXT2).pack()

        def _row(parent, label, var, tip="", wide=False):
            tk.Label(parent,text=label,font=("Segoe UI",8),bg=BG,fg=TEXT2).pack(anchor="w",padx=8,pady=(4,1))
            tk.Entry(parent,textvariable=var,bg=CARD,fg=TEXT,insertbackground="white",
                     relief="flat",font=("Segoe UI",8)).pack(fill="x",padx=8,ipady=3,pady=(0,1))
            if tip: tk.Label(parent,text=tip,font=("Segoe UI",7),bg=BG,fg=TEXT2).pack(anchor="w",padx=8,pady=(0,3))

        self.bounds_var=tk.StringVar()
        self.bounds_frame=tk.Frame(adv,bg=BG); self.bounds_frame.pack(fill="x")
        _row(self.bounds_frame,"欄界 x 座標（逗號分隔，留空=自動偵測）",
             self.bounds_var,"例：0,116,182,250,361,443,513")
        self.bounds_frame.pack_forget()

        self.wm_var=tk.StringVar()
        _row(adv,"浮水印關鍵字（逗號分隔）",self.wm_var,"例：B242,允將,CONFIDENTIAL")

        self.anchor_var=tk.StringVar(value="0")
        _row(adv,"錨點欄（第幾欄為主鍵，0開始）",self.anchor_var,"空白該列→合併備註到上一筆")

        self.merge_range_var=tk.StringVar()
        self.merge_frame=tk.Frame(adv,bg=BG); self.merge_frame.pack(fill="x")
        _row(self.merge_frame,"合併欄位範圍（起-迄，如 3-4）",
             self.merge_range_var,"將指定欄合併為一欄（修正錯位）")
        self.merge_frame.pack_forget()

        self.skip_dup=tk.BooleanVar(value=True)
        tk.Checkbutton(adv,text="自動刪除重複標題列",variable=self.skip_dup,
                       bg=BG,fg=TEXT,selectcolor=BG2,activebackground=BG,
                       font=("Segoe UI",9)).pack(anchor="w",padx=8,pady=(2,6))

        # 進度條
        tk.Label(right,text="轉換進度",font=("Segoe UI",9,"bold"),bg=BG,fg=TEXT).pack(anchor="w")
        s=ttk.Style(); s.theme_use("clam")
        s.configure("P.Horizontal.TProgressbar",troughcolor=CARD,background=ACCENT,
                    darkcolor=ACCENT,lightcolor=ACCENT,bordercolor=BG,thickness=14)
        self.pbar=ttk.Progressbar(right,style="P.Horizontal.TProgressbar",orient="horizontal",mode="determinate")
        self.pbar.pack(fill="x",pady=(3,2))
        self.pbar_lbl=tk.Label(right,text="",font=("Segoe UI",8),bg=BG,fg=TEXT2)
        self.pbar_lbl.pack(anchor="w",pady=(0,8))

        bc=dict(relief="flat",cursor="hand2")
        tk.Button(right,text="🔍  預覽 / 調整欄位",font=("Segoe UI",12),bg=CARD,fg=TEXT,pady=9,command=self.open_preview,**bc).pack(fill="x",pady=(0,5))
        self.btn_start=tk.Button(right,text="▶  開始批次轉換",font=("Segoe UI",14,"bold"),bg=ACCENT,fg="white",pady=11,command=self.start_convert,**bc)
        self.btn_start.pack(fill="x",pady=(0,5))
        tk.Button(right,text="📂  開啟輸出資料夾",font=("Segoe UI",11),bg=BG2,fg=TEXT2,pady=7,command=self.open_out,**bc).pack(fill="x")

        # Log
        logf=tk.Frame(self,bg=BG2,height=148); logf.pack(fill="x",padx=16,pady=(0,10)); logf.pack_propagate(False)
        tk.Label(logf,text="執行記錄",font=("Segoe UI",11,"bold"),bg=BG2,fg=TEXT2).pack(anchor="w",padx=10,pady=(6,0))
        lsb=tk.Scrollbar(logf); lsb.pack(side="right",fill="y")
        self.log_box=tk.Text(logf,yscrollcommand=lsb.set,bg=BG2,fg=TEXT2,font=("Consolas",10),
                             relief="flat",bd=0,state="disabled",wrap="word",highlightthickness=0)
        self.log_box.pack(fill="both",expand=True,padx=10,pady=(0,6))
        lsb.config(command=self.log_box.yview)
        self.log_box.tag_config("ok",foreground=SUCCESS)
        self.log_box.tag_config("err",foreground=ERROR)
        self.log_box.tag_config("warn",foreground=WARNING)

    def _on_mode(self):
        m=self.mode_var.get()
        if m=="coords":
            self.bounds_frame.pack(fill="x")
            self.merge_frame.pack(fill="x")
        else:
            self.bounds_frame.pack_forget()
            self.merge_frame.pack_forget()

    def _get_settings(self):
        return {
            "mode": self.mode_var.get(),
            "watermark_keywords": self.wm_var.get(),
            "anchor_col": self.anchor_var.get().strip() or "0",
            "skip_duplicate_header": self.skip_dup.get(),
            "col_bounds": self.bounds_var.get(),
            "merge_col_range": self.merge_range_var.get(),
        }

    def add_files(self):
        fs=filedialog.askopenfilenames(title="選擇 PDF",filetypes=[("PDF","*.pdf"),("所有","*.*")])
        for f in fs:
            if f not in self.pdf_files: self.pdf_files.append(f); self.listbox.insert("end",os.path.basename(f))

    def remove_sel(self,e=None):
        for i in reversed(self.listbox.curselection()): self.listbox.delete(i); self.pdf_files.pop(i)

    def clear_files(self): self.pdf_files.clear(); self.listbox.delete(0,"end")
    def browse_dir(self):
        d=filedialog.askdirectory()
        if d: self.out_dir.set(d)
    def open_out(self):
        d=self.out_dir.get()
        if os.path.isdir(d): os.startfile(d)

    def log(self,msg,tag="info"):
        self.log_box.config(state="normal"); self.log_box.insert("end",msg+"\n",tag)
        self.log_box.see("end"); self.log_box.config(state="disabled")

    def open_preview(self):
        if not self.pdf_files: messagebox.showwarning("提示","請先新增 PDF！"); return
        sel=self.listbox.curselection()
        pdf=self.pdf_files[sel[0] if sel else 0]
        settings=self._get_settings()
        self.log(f"載入預覽：{os.path.basename(pdf)}")
        def _load():
            try:
                table,strategy=convert(pdf,None,settings,log_cb=lambda m:self.after(0,lambda m=m:self.log(m)))
                self.after(0,lambda:PreviewWin(self,pdf,table,strategy,settings))
            except Exception as e:
                self.after(0,lambda:self.log(f"預覽失敗：{e}","err"))
        threading.Thread(target=_load,daemon=True).start()

    def start_convert(self):
        if self.running: return
        if not self.pdf_files: messagebox.showwarning("提示","請先新增 PDF！"); return
        if not os.path.isdir(self.out_dir.get()): messagebox.showerror("錯誤",f"資料夾不存在：{self.out_dir.get()}"); return
        self.running=True; self.btn_start.config(state="disabled",text="⏳  轉換中...")
        self.pbar["value"]=0; threading.Thread(target=self._run,daemon=True).start()

    def _run(self):
        files=list(self.pdf_files); out_dir=self.out_dir.get()
        settings=self._get_settings(); ok=fail=0
        self.log("═"*44); self.log(f"開始轉換 {len(files)} 個（模式：{settings['mode']}）")
        for i,pdf in enumerate(files):
            fname=os.path.basename(pdf); self.log(f"\n[{i+1}/{len(files)}] {fname}")
            self.after(0,lambda v=i/len(files)*100:self._set_pbar(v))
            out=os.path.join(out_dir,os.path.splitext(fname)[0]+".xlsx")
            if os.path.exists(out):
                ans=messagebox.askyesno("檔案已存在",
                    f"檔案已存在，是否覆蓋？\n{os.path.basename(out)}")
                if not ans:
                    self.log(f"  ⏭  跳過（使用者取消）","warn"); continue
            try:
                table,strategy=convert(pdf,out,settings,log_cb=lambda m:self.after(0,lambda m=m:self.log(m)))
                ok+=1; self.log(f"  ✅ {max(0,len(table)-1)} 筆 [{strategy}] → {os.path.basename(out)}","ok")
            except Exception as e:
                fail+=1; self.log(f"  ❌ {e}","err")
        self.after(0,lambda:self._set_pbar(100))
        self.log("\n"+"═"*44)
        self.log(f"完成：成功 {ok} / 失敗 {fail}","ok" if fail==0 else "err")
        self.after(0,self._done)

    def _set_pbar(self,v): self.pbar["value"]=v; self.pbar_lbl.config(text=f"{v:.0f}%")
    def _done(self): self.running=False; self.btn_start.config(state="normal",text="▶  開始批次轉換")


class PreviewWin(tk.Toplevel):
    def __init__(self,master,pdf_path,table,strategy,settings):
        super().__init__(master)
        self.pdf_path=pdf_path; self.table=table
        self.strategy=strategy; self.settings=settings; self.col_vars=[]
        self.title(f"預覽 — {os.path.basename(pdf_path)}")
        self.geometry("1300x740"); self.configure(bg=BG); self._build()

    def _build(self):
        info=tk.Frame(self,bg=BG2,height=32); info.pack(fill="x"); info.pack_propagate(False)
        nrows=len(self.table); ncols=max((len(r) for r in self.table),default=0)
        tk.Label(info,text=f"  策略：{self.strategy}    {nrows} 列 × {ncols} 欄    （前 50 列預覽）",
                 font=("Segoe UI",9),bg=BG2,fg=TEXT2).pack(side="left",padx=8)

        if self.table:
            hf=tk.Frame(self,bg=BG); hf.pack(fill="x",padx=10,pady=(6,2))
            tk.Label(hf,text="欄位名稱（可修改後點「套用」更新預覽）：",font=("Segoe UI",9,"bold"),bg=BG,fg=TEXT).pack(side="left")
            ce=tk.Frame(self,bg=BG); ce.pack(fill="x",padx=10,pady=(0,4))
            for ci,hval in enumerate(self.table[0]):
                v=tk.StringVar(value=clean(hval)); self.col_vars.append(v)
                f=tk.Frame(ce,bg=CARD,padx=4,pady=2); f.pack(side="left",padx=(0,3))
                tk.Label(f,text=f"欄{ci+1}",font=("Segoe UI",7),bg=CARD,fg=TEXT2).pack()
                tk.Entry(f,textvariable=v,width=12,bg=BG2,fg=TEXT,insertbackground="white",relief="flat",font=("Segoe UI",10)).pack(ipady=3)

        tf=tk.Frame(self,bg=BG); tf.pack(fill="both",expand=True,padx=10,pady=(0,4))
        preview=self.table[:50]
        ncols2=max((len(r) for r in preview),default=1)
        cols=[f"C{i+1}" for i in range(ncols2)]
        self.tree=ttk.Treeview(tf,columns=cols,show="headings",height=17)
        cw=max(70,min(160,900//max(len(cols),1)))
        for c in cols: self.tree.heading(c,text=c); self.tree.column(c,width=cw,anchor="w")
        vsb=ttk.Scrollbar(tf,orient="vertical",command=self.tree.yview)
        hsb=ttk.Scrollbar(tf,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        tf.rowconfigure(0,weight=1); tf.columnconfigure(0,weight=1)
        s=ttk.Style()
        s.configure("Treeview",background=CARD,foreground=TEXT,fieldbackground=CARD,font=("Segoe UI",11),rowheight=26)
        s.configure("Treeview.Heading",background=BG2,foreground=TEXT,font=("Segoe UI",11,"bold"))
        self._fill(preview)
        self.tree.bind("<Double-1>",self._on_double_click)
        tk.Label(tf,text="💡 雙擊儲存格可直接編輯　Enter 換行　Ctrl+Enter / Tab 儲存　Esc 取消",font=("Segoe UI",10),bg=BG,fg=TEXT2).grid(row=2,column=0,sticky="w",pady=(4,0))

        bf=tk.Frame(self,bg=BG); bf.pack(fill="x",padx=10,pady=(0,8))
        tk.Button(bf,text="🔄  套用欄位名稱",font=("Segoe UI",11),bg=CARD,fg=TEXT,relief="flat",cursor="hand2",padx=14,pady=6,command=self._apply).pack(side="left",padx=(0,8))
        tk.Button(bf,text="💾  另存為 Excel",font=("Segoe UI",12,"bold"),bg=ACCENT,fg="white",relief="flat",cursor="hand2",padx=18,pady=7,command=self.save).pack(side="right")
        tk.Button(bf,text="關閉",font=("Segoe UI",11),bg=BG2,fg=TEXT2,relief="flat",cursor="hand2",padx=14,pady=6,command=self.destroy).pack(side="right",padx=(0,6))

    def _fill(self,data):
        self.tree.delete(*self.tree.get_children())
        if not data: return
        header=[v.get() if i<len(self.col_vars) else f"欄{i+1}" for i,v in enumerate(self.col_vars)] if self.col_vars else [clean(v) or f"欄{i+1}" for i,v in enumerate(data[0])]
        ncols=max(len(r) for r in data)
        cols=[f"C{i+1}" for i in range(ncols)]
        for c,h in zip(cols,header): self.tree.heading(c,text=h)
        for row in data[1:]:
            vals=[row[i] if i<len(row) else "" for i in range(ncols)]
            self.tree.insert("","end",values=[str(v).replace("\n"," / ") for v in vals])

    def _apply(self): self._fill(self.table[:50])

    def _on_double_click(self, event):
        """雙擊儲存格 → inline Text 編輯
        Enter = 換行   Ctrl+Enter / Tab = 儲存   Esc = 取消
        """
        item=self.tree.focus()
        if not item: return
        col_id=self.tree.identify_column(event.x)
        ci=int(col_id.replace("#",""))-1
        items=list(self.tree.get_children())
        if item not in items: return
        ri=items.index(item)+1
        if ri>=len(self.table): return
        cur_val=self.table[ri][ci] if ci<len(self.table[ri]) else ""
        bbox=self.tree.bbox(item, col_id)
        if not bbox: return
        x,y,w,h=bbox
        # 統一用 Text 框，高度依內容自動設定（最少2行，最多8行）
        n_lines=max(cur_val.count("\n")+1, 2)
        edit_h=min(n_lines, 8)
        widget=tk.Text(self.tree, bg="#FFFDE7", fg="black",
                       relief="solid", bd=1,
                       font=("Segoe UI",10), wrap="word",
                       height=edit_h, insertbackground="black",
                       selectbackground=ACCENT2)
        widget.insert("1.0", cur_val)
        widget.place(x=x, y=y, width=max(w,220), height=edit_h*18+4)
        widget.focus_set()
        widget.mark_set("insert","end")

        orig_val=cur_val  # 用於 Esc 取消

        def _commit(e=None):
            new_val=widget.get("1.0","end-1c")
            while len(self.table[ri])<ci+1: self.table[ri].append("")
            self.table[ri][ci]=new_val
            widget.destroy()
            self._fill(self.table[:50])

        def _cancel(e=None):
            widget.destroy()

        def _key(e):
            # Ctrl+Enter 或 Tab → 儲存
            if e.keysym in ("Tab",) or (e.keysym=="Return" and (e.state&0x4)):
                _commit(); return "break"
            # Esc → 取消
            if e.keysym=="Escape":
                _cancel(); return "break"
            # 其他鍵正常處理（Enter 換行）

        widget.bind("<KeyPress>", _key)
        widget.bind("<FocusOut>", _commit)

    def save(self):
        out=filedialog.asksaveasfilename(title="另存 Excel",defaultextension=".xlsx",
            initialfile=os.path.splitext(os.path.basename(self.pdf_path))[0]+".xlsx",
            filetypes=[("Excel","*.xlsx")])
        if not out: return
        try:
            col_map={i:v.get() for i,v in enumerate(self.col_vars) if v.get()} if self.col_vars else None
            mi=detect_merges(self.table)
            write_excel(self.table,mi,out,col_map=col_map)
            messagebox.showinfo("完成",f"已儲存：\n{out}")
            os.startfile(os.path.dirname(out))
        except Exception as e:
            messagebox.showerror("錯誤",str(e))

if __name__=="__main__":
    App().mainloop()
