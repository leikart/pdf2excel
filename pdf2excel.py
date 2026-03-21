"""
PDF 轉 Excel 通用工具
- 自動偵測表格（框線 / 文字對齊 雙策略）
- 自動辨識標題列、去除重複標題
- 合併儲存格（跟 PDF 一樣）
- 多頁合併成單一工作表
- 預覽 + 欄位名稱調整介面
- 批次轉換 + 另存選項
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
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "pdfplumber", "openpyxl", "-q"])
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell

BG=="#1A1A2E"; BG2="#16213E"; CARD="#0F3460"; ACCENT="#E94560"
ACCENT2="#533483"; TEXT="#EAEAEA"; TEXT2="#A0A8C0"
SUCCESS="#4CAF82"; ERROR="#E94560"; WARNING="#F0A500"
BG      = "#1A1A2E"
BG2     = "#16213E"
CARD    = "#0F3460"
ACCENT  = "#E94560"
ACCENT2 = "#533483"
TEXT    = "#EAEAEA"
TEXT2   = "#A0A8C0"
SUCCESS = "#4CAF82"
ERROR   = "#E94560"
WARNING = "#F0A500"

TABLE_LINES = {
    "vertical_strategy":"lines","horizontal_strategy":"lines",
    "snap_tolerance":3,"join_tolerance":3,"edge_min_length":3,
    "min_words_vertical":1,"min_words_horizontal":1,
    "intersection_tolerance":3,"text_tolerance":3,
}
TABLE_TEXT = {**TABLE_LINES,"vertical_strategy":"text","horizontal_strategy":"text"}

def clean(v):
    if v is None: return ""
    return re.sub(r"\s+"," ",str(v)).strip()

def rows_equal(a,b):
    ca=[clean(x) for x in (a or [])]
    cb=[clean(x) for x in (b or [])]
    return ca==cb and any(ca)

def normalize(table):
    out=[]
    for row in table:
        r=[clean(v) for v in row]
        if any(r): out.append(r)
    return out

def parse_pdf(pdf_path, log_cb=None):
    all_tables=[]; strategy_used="lines"
    with pdfplumber.open(pdf_path) as pdf:
        total=len(pdf.pages)
        if log_cb: log_cb(f"  共 {total} 頁")
        for pn,page in enumerate(pdf.pages,1):
            if log_cb: log_cb(f"  解析第 {pn} 頁...")
            tables=page.extract_tables(TABLE_LINES)
            if not tables or not any(len(t)>1 for t in tables):
                tables=page.extract_tables(TABLE_TEXT); strategy_used="text"
            else:
                strategy_used="lines"
            for t in (tables or []):
                n=normalize(t)
                if n: all_tables.append(n)
    merged=[]; header=None
    for table in all_tables:
        if not table: continue
        if header is None:
            header=table[0]; merged.extend(table)
        else:
            start=1 if rows_equal(table[0],header) else 0
            merged.extend(table[start:])
    return merged, strategy_used

def detect_merges(merged_table):
    if len(merged_table)<2: return {}
    ncols=max(len(r) for r in merged_table)
    data=merged_table[1:]
    info={}
    for ci in range(ncols):
        ranges=[]; i=0
        while i<len(data):
            val=data[i][ci] if ci<len(data[i]) else ""
            if val:
                j=i+1
                while j<len(data):
                    nv=data[j][ci] if ci<len(data[j]) else ""
                    if nv=="" or nv==val: j+=1
                    else: break
                if j-i>1: ranges.append((i,j-1))
                i=j
            else: i+=1
        if ranges: info[ci]=ranges
    return info

def write_excel(merged_table, merge_info, out_path, col_map=None):
    wb=Workbook(); ws=wb.active; ws.title="轉換結果"
    hfill=PatternFill(start_color="0F3460",end_color="0F3460",fill_type="solid")
    hfont=Font(color="FFFFFF",bold=True,size=11)
    mfill=PatternFill(start_color="D6E4F0",end_color="D6E4F0",fill_type="solid")
    o_fill=PatternFill(start_color="F5F9FF",end_color="F5F9FF",fill_type="solid")
    e_fill=PatternFill(start_color="FFFFFF",end_color="FFFFFF",fill_type="solid")
    thin=Side(style="thin")
    bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    ctr=Alignment(horizontal="center",vertical="center",wrap_text=True)
    lwrap=Alignment(horizontal="left",vertical="center",wrap_text=True)
    if not merged_table: wb.save(out_path); return
    ncols=max(len(r) for r in merged_table)
    col_max=[8]*ncols
    for row in merged_table:
        for ci,val in enumerate(row): col_max[ci]=min(max(col_max[ci],len(str(val or ""))),40)
    for ci in range(ncols):
        ws.column_dimensions[get_column_letter(ci+1)].width=col_max[ci]+4
    drc=0
    for ri,row in enumerate(merged_table):
        er=ri+1; is_hdr=(ri==0)
        ws.row_dimensions[er].height=20 if is_hdr else 16
        for ci in range(ncols):
            val=row[ci] if ci<len(row) else ""
            if is_hdr and col_map and ci in col_map: val=col_map[ci]
            cell=ws.cell(row=er,column=ci+1,value=val)
            cell.border=bdr
            cell.alignment=ctr if is_hdr else lwrap
            if is_hdr: cell.fill=hfill; cell.font=hfont
            else: cell.fill=o_fill if drc%2==0 else e_fill; cell.font=Font(size=10)
        if not is_hdr: drc+=1
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
            except Exception: pass
    wb.save(out_path)

def convert_pdf_to_excel(pdf_path, out_path, col_map=None, log_cb=None):
    merged,strategy=parse_pdf(pdf_path,log_cb)
    if log_cb: log_cb(f"  策略：{strategy}，{len(merged)} 列（含標題）")
    mi=detect_merges(merged)
    write_excel(merged,mi,out_path,col_map=col_map)
    return max(0,len(merged)-1)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 轉 Excel 通用工具")
        self.geometry("940x680"); self.minsize(800,580)
        self.configure(bg=BG)
        self.pdf_files=[]; self.out_dir=tk.StringVar(value=os.path.expanduser("~\\Desktop"))
        self.running=False; self._build_ui()

    def _build_ui(self):
        hdr=tk.Frame(self,bg=CARD,height=58); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr,text="PDF  →  Excel",font=("Segoe UI",18,"bold"),bg=CARD,fg="white").pack(side="left",padx=22,pady=8)
        tk.Label(hdr,text="通用版・自動偵測・合併儲存格・多頁整合",font=("Segoe UI",10),bg=CARD,fg=TEXT2).pack(side="left")

        body=tk.Frame(self,bg=BG); body.pack(fill="both",expand=True,padx=18,pady=12)
        left=tk.Frame(body,bg=BG); left.pack(side="left",fill="both",expand=True)

        r1=tk.Frame(left,bg=BG); r1.pack(fill="x",pady=(0,5))
        tk.Label(r1,text="PDF 檔案清單",font=("Segoe UI",10,"bold"),bg=BG,fg=TEXT).pack(side="left")
        tk.Button(r1,text="＋ 新增",font=("Segoe UI",9),bg=ACCENT,fg="white",relief="flat",cursor="hand2",padx=10,pady=2,command=self.add_files).pack(side="right",padx=(3,0))
        tk.Button(r1,text="清除全部",font=("Segoe UI",9),bg=BG2,fg=TEXT2,relief="flat",cursor="hand2",padx=10,pady=2,command=self.clear_files).pack(side="right")

        lf=tk.Frame(left,bg=CARD); lf.pack(fill="both",expand=True)
        sb=tk.Scrollbar(lf); sb.pack(side="right",fill="y")
        self.listbox=tk.Listbox(lf,yscrollcommand=sb.set,bg=CARD,fg=TEXT,selectbackground=ACCENT2,font=("Segoe UI",9),relief="flat",bd=0,highlightthickness=0,activestyle="none")
        self.listbox.pack(fill="both",expand=True,padx=2,pady=2)
        sb.config(command=self.listbox.yview)
        self.listbox.bind("<Delete>",self.remove_selected)
        tk.Label(left,text="選中後按 Delete 可移除",font=("Segoe UI",8),bg=BG,fg=TEXT2).pack(anchor="w",pady=(3,0))

        right=tk.Frame(body,bg=BG,width=238); right.pack(side="right",fill="y",padx=(14,0)); right.pack_propagate(False)
        tk.Label(right,text="輸出資料夾",font=("Segoe UI",10,"bold"),bg=BG,fg=TEXT).pack(anchor="w")
        dr=tk.Frame(right,bg=BG); dr.pack(fill="x",pady=(3,10))
        tk.Entry(dr,textvariable=self.out_dir,bg=CARD,fg=TEXT,insertbackground="white",relief="flat",font=("Segoe UI",8)).pack(side="left",fill="x",expand=True,ipady=5,padx=(0,3))
        tk.Button(dr,text="瀏覽",font=("Segoe UI",9),bg=BG2,fg=TEXT2,relief="flat",cursor="hand2",padx=8,pady=2,command=self.browse_dir).pack(side="right")

        self.count_lbl=tk.Label(right,text="尚未選擇檔案",font=("Segoe UI",9),bg=BG,fg=TEXT2)
        self.count_lbl.pack(anchor="w",pady=(0,10))

        tk.Label(right,text="轉換進度",font=("Segoe UI",10,"bold"),bg=BG,fg=TEXT).pack(anchor="w")
        s=ttk.Style(); s.theme_use("clam")
        s.configure("P.Horizontal.TProgressbar",troughcolor=CARD,background=ACCENT,darkcolor=ACCENT,lightcolor=ACCENT,bordercolor=BG,thickness=12)
        self.pbar=ttk.Progressbar(right,style="P.Horizontal.TProgressbar",orient="horizontal",mode="determinate")
        self.pbar.pack(fill="x",pady=(3,2))
        self.pbar_lbl=tk.Label(right,text="",font=("Segoe UI",8),bg=BG,fg=TEXT2); self.pbar_lbl.pack(anchor="w",pady=(0,10))

        bc=dict(relief="flat",cursor="hand2")
        tk.Button(right,text="🔍  預覽 / 調整欄位",font=("Segoe UI",10),bg=CARD,fg=TEXT,pady=8,command=self.open_preview,**bc).pack(fill="x",pady=(0,5))
        self.btn_start=tk.Button(right,text="▶  開始批次轉換",font=("Segoe UI",13,"bold"),bg=ACCENT,fg="white",pady=10,command=self.start_convert,**bc)
        self.btn_start.pack(fill="x",pady=(0,5))
        tk.Button(right,text="📂  開啟輸出資料夾",font=("Segoe UI",9),bg=BG2,fg=TEXT2,pady=6,command=self.open_out,**bc).pack(fill="x")

        logf=tk.Frame(self,bg=BG2,height=155); logf.pack(fill="x",padx=18,pady=(0,12)); logf.pack_propagate(False)
        tk.Label(logf,text="執行記錄",font=("Segoe UI",9,"bold"),bg=BG2,fg=TEXT2).pack(anchor="w",padx=10,pady=(5,0))
        lsb=tk.Scrollbar(logf); lsb.pack(side="right",fill="y")
        self.log_box=tk.Text(logf,yscrollcommand=lsb.set,bg=BG2,fg=TEXT2,font=("Consolas",8),relief="flat",bd=0,state="disabled",wrap="word",highlightthickness=0)
        self.log_box.pack(fill="both",expand=True,padx=10,pady=(0,6))
        lsb.config(command=self.log_box.yview)
        self.log_box.tag_config("ok",foreground=SUCCESS)
        self.log_box.tag_config("err",foreground=ERROR)
        self.log_box.tag_config("warn",foreground=WARNING)

    def add_files(self):
        fs=filedialog.askopenfilenames(title="選擇 PDF 檔案",filetypes=[("PDF","*.pdf"),("所有","*.*")])
        for f in fs:
            if f not in self.pdf_files: self.pdf_files.append(f); self.listbox.insert("end",os.path.basename(f))
        self._upd_count()

    def remove_selected(self,e=None):
        for i in reversed(self.listbox.curselection()): self.listbox.delete(i); self.pdf_files.pop(i)
        self._upd_count()

    def clear_files(self): self.pdf_files.clear(); self.listbox.delete(0,"end"); self._upd_count()
    def browse_dir(self):
        d=filedialog.askdirectory(title="選擇輸出資料夾")
        if d: self.out_dir.set(d)
    def open_out(self):
        d=self.out_dir.get()
        if os.path.isdir(d): os.startfile(d)
    def _upd_count(self):
        n=len(self.pdf_files); self.count_lbl.config(text=f"已選 {n} 個 PDF" if n else "尚未選擇檔案")

    def log(self,msg,tag="info"):
        self.log_box.config(state="normal"); self.log_box.insert("end",msg+"\n",tag)
        self.log_box.see("end"); self.log_box.config(state="disabled")

    def open_preview(self):
        if not self.pdf_files: messagebox.showwarning("提示","請先新增 PDF 檔案！"); return
        sel=self.listbox.curselection()
        pdf=self.pdf_files[sel[0] if sel else 0]
        self.log(f"載入預覽：{os.path.basename(pdf)}")
        def _load():
            try:
                merged,strategy=parse_pdf(pdf,log_cb=lambda m:self.after(0,lambda m=m:self.log(m)))
                self.after(0,lambda:PreviewWindow(self,pdf,merged,strategy))
            except Exception as e:
                self.after(0,lambda:self.log(f"預覽失敗：{e}","err"))
        threading.Thread(target=_load,daemon=True).start()

    def start_convert(self):
        if self.running: return
        if not self.pdf_files: messagebox.showwarning("提示","請先新增 PDF 檔案！"); return
        if not os.path.isdir(self.out_dir.get()): messagebox.showerror("錯誤",f"資料夾不存在：\n{self.out_dir.get()}"); return
        self.running=True; self.btn_start.config(state="disabled",text="⏳  轉換中...")
        self.pbar["value"]=0; threading.Thread(target=self._run,daemon=True).start()

    def _run(self):
        files=list(self.pdf_files); out_dir=self.out_dir.get(); ok=fail=0
        self.log("═"*46); self.log(f"開始批次轉換，共 {len(files)} 個")
        for i,pdf in enumerate(files):
            fname=os.path.basename(pdf); self.log(f"\n[{i+1}/{len(files)}] {fname}")
            self.after(0,lambda v=i/len(files)*100:self._set_pbar(v))
            out_path=os.path.join(out_dir,os.path.splitext(fname)[0]+".xlsx")
            try:
                n=convert_pdf_to_excel(pdf,out_path,log_cb=lambda m:self.after(0,lambda m=m:self.log(m)))
                ok+=1; self.log(f"  ✅ {n} 筆 → {os.path.basename(out_path)}","ok")
            except Exception as e:
                fail+=1; self.log(f"  ❌ 失敗：{e}","err")
        self.after(0,lambda:self._set_pbar(100))
        self.log("\n"+"═"*46)
        self.log(f"完成：成功 {ok} / 失敗 {fail}","ok" if fail==0 else "err")
        self.after(0,self._done)

    def _set_pbar(self,v): self.pbar["value"]=v; self.pbar_lbl.config(text=f"{v:.0f}%")
    def _done(self): self.running=False; self.btn_start.config(state="normal",text="▶  開始批次轉換")


class PreviewWindow(tk.Toplevel):
    def __init__(self,master,pdf_path,merged_table,strategy):
        super().__init__(master)
        self.pdf_path=pdf_path; self.merged_table=merged_table
        self.strategy=strategy; self.col_vars=[]
        self.title(f"預覽 / 欄位調整 — {os.path.basename(pdf_path)}")
        self.geometry("1160x640"); self.configure(bg=BG); self._build()

    def _build(self):
        info=tk.Frame(self,bg=BG2,height=34); info.pack(fill="x"); info.pack_propagate(False)
        ncols=max((len(r) for r in self.merged_table),default=0)
        nrows=len(self.merged_table)
        tk.Label(info,text=f"  策略：{self.strategy}    共 {nrows} 列 × {ncols} 欄    （前 50 列預覽）",
                 font=("Segoe UI",9),bg=BG2,fg=TEXT2).pack(side="left",padx=6)

        if self.merged_table:
            hf=tk.Frame(self,bg=BG); hf.pack(fill="x",padx=12,pady=(8,3))
            tk.Label(hf,text="欄位名稱（可直接修改後點「套用」更新預覽）：",font=("Segoe UI",9,"bold"),bg=BG,fg=TEXT).pack(side="left")
            ce=tk.Frame(self,bg=BG); ce.pack(fill="x",padx=12,pady=(0,4))
            for ci,hval in enumerate(self.merged_table[0]):
                v=tk.StringVar(value=clean(hval)); self.col_vars.append(v)
                f=tk.Frame(ce,bg=CARD,padx=4,pady=3); f.pack(side="left",padx=(0,4))
                tk.Label(f,text=f"欄{ci+1}",font=("Segoe UI",7),bg=CARD,fg=TEXT2).pack()
                tk.Entry(f,textvariable=v,width=12,bg=BG2,fg=TEXT,insertbackground="white",relief="flat",font=("Segoe UI",9)).pack(ipady=3)

        tf=tk.Frame(self,bg=BG); tf.pack(fill="both",expand=True,padx=12,pady=(0,6))
        preview=self.merged_table[:50]
        ncols2=max((len(r) for r in preview),default=1)
        cols=[f"欄{i+1}" for i in range(ncols2)]
        self.tree=ttk.Treeview(tf,columns=cols,show="headings",height=18)
        cw=max(80,min(180,900//max(len(cols),1)))
        for c in cols: self.tree.heading(c,text=c); self.tree.column(c,width=cw,anchor="w")
        vsb=ttk.Scrollbar(tf,orient="vertical",command=self.tree.yview)
        hsb=ttk.Scrollbar(tf,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        tf.rowconfigure(0,weight=1); tf.columnconfigure(0,weight=1)
        s=ttk.Style()
        s.configure("Treeview",background=CARD,foreground=TEXT,fieldbackground=CARD,font=("Segoe UI",9),rowheight=20)
        s.configure("Treeview.Heading",background=BG2,foreground=TEXT,font=("Segoe UI",9,"bold"))
        self._fill_tree(preview)

        bf=tk.Frame(self,bg=BG); bf.pack(fill="x",padx=12,pady=(0,10))
        tk.Button(bf,text="🔄  套用欄位名稱",font=("Segoe UI",9),bg=CARD,fg=TEXT,relief="flat",cursor="hand2",padx=12,pady=5,command=self._apply).pack(side="left",padx=(0,8))
        tk.Button(bf,text="💾  另存為 Excel",font=("Segoe UI",10,"bold"),bg=ACCENT,fg="white",relief="flat",cursor="hand2",padx=16,pady=6,command=self.save).pack(side="right")
        tk.Button(bf,text="關閉",font=("Segoe UI",9),bg=BG2,fg=TEXT2,relief="flat",cursor="hand2",padx=12,pady=6,command=self.destroy).pack(side="right",padx=(0,6))

    def _fill_tree(self,data):
        self.tree.delete(*self.tree.get_children())
        if not data: return
        header=[v.get() if i<len(self.col_vars) else f"欄{i+1}" for i,v in enumerate(self.col_vars)] if self.col_vars else [clean(v) or f"欄{i+1}" for i,v in enumerate(data[0])]
        ncols=max(len(r) for r in data)
        cols=[f"欄{i+1}" for i in range(ncols)]
        for c,h in zip(cols,header): self.tree.heading(c,text=h)
        for row in data[1:]:
            vals=[row[i] if i<len(row) else "" for i in range(ncols)]
            self.tree.insert("","end",values=vals)

    def _apply(self): self._fill_tree(self.merged_table[:50])

    def save(self):
        out=filedialog.asksaveasfilename(title="另存 Excel",defaultextension=".xlsx",
            initialfile=os.path.splitext(os.path.basename(self.pdf_path))[0]+".xlsx",
            filetypes=[("Excel","*.xlsx")])
        if not out: return
        try:
            col_map={i:v.get() for i,v in enumerate(self.col_vars) if v.get()} if self.col_vars else None
            mi=detect_merges(self.merged_table)
            write_excel(self.merged_table,mi,out,col_map=col_map)
            messagebox.showinfo("完成",f"已儲存：\n{out}")
            os.startfile(os.path.dirname(out))
        except Exception as e:
            messagebox.showerror("錯誤",str(e))

if __name__=="__main__":
    App().mainloop()
