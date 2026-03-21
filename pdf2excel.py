import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import re

try:
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell
except ImportError:
    import sys, subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "pdfplumber", "openpyxl", "-q"])
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell

# ── 顏色 ────────────────────────────────────────────
BG      = "#1A1A2E"
BG2     = "#16213E"
CARD    = "#0F3460"
ACCENT  = "#E94560"
TEXT    = "#EAEAEA"
TEXT2   = "#A0A8C0"
SUCCESS = "#4CAF82"
ERROR   = "#E94560"

# ── 欄位定義 ────────────────────────────────────────
COLUMNS    = ["出貨單編號", "出貨日期", "星期", "位置", "數量", "進場時間", "備註"]
COL_BOUNDS = [0, 95, 155, 195, 355, 415, 460, 9999]

def get_col_index(x):
    for i in range(len(COL_BOUNDS) - 1):
        if COL_BOUNDS[i] <= x < COL_BOUNDS[i + 1]:
            return i
    return len(COLUMNS) - 1

def is_header_row(row):
    text = "".join(str(v) for v in row if v)
    return "出貨單編號" in text and "出貨日期" in text

def is_page_title(text):
    return "B242" in text and "允將出貨順序" in text

def parse_pdf_to_rows(pdf_path, log_cb=None):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        if log_cb:
            log_cb(f"  📑 共 {total} 頁")
        for page_num, page in enumerate(pdf.pages, 1):
            if log_cb:
                log_cb(f"  解析第 {page_num} 頁...")
            words = page.extract_words(
                x_tolerance=3, y_tolerance=3,
                keep_blank_chars=False, use_text_flow=False)
            if not words:
                continue
            lines = {}
            for w in words:
                y = round(w["top"] / 4) * 4
                lines.setdefault(y, []).append(w)
            for y in sorted(lines.keys()):
                ws_line = sorted(lines[y], key=lambda w: w["x0"])
                line_text = " ".join(w["text"] for w in ws_line)
                if is_page_title(line_text):
                    continue
                row_data = [""] * len(COLUMNS)
                for w in ws_line:
                    ci = get_col_index(w["x0"])
                    if ci < len(COLUMNS):
                        row_data[ci] = (row_data[ci] + " " + w["text"]).strip()
                if not any(row_data):
                    continue
                if is_header_row(row_data):
                    continue
                all_rows.append(row_data)
    return all_rows

def merge_rows(raw_rows):
    records = []
    last_date = ""
    last_weekday = ""
    for row in raw_rows:
        no      = row[0].strip()
        date    = row[1].strip()
        weekday = row[2].strip()
        loc     = row[3].strip()
        qty     = row[4].strip()
        time_   = row[5].strip()
        note    = row[6].strip()
        if date:
            last_date = date
        if weekday:
            last_weekday = weekday
        if no:
            records.append({
                "出貨單編號": no,
                "出貨日期":   last_date,
                "星期":       last_weekday,
                "位置":       loc,
                "數量":       qty,
                "進場時間":   time_,
                "備註":       note,
            })
        else:
            if records:
                extra = " ".join(filter(None, [loc, qty, time_, note]))
                if extra:
                    if records[-1]["備註"]:
                        records[-1]["備註"] += "\n" + extra
                    else:
                        records[-1]["備註"] = extra
    # 計算需合併的日期範圍
    date_merge = []
    i = 0
    while i < len(records):
        j = i
        while j + 1 < len(records) and records[j+1]["出貨日期"] == records[i]["出貨日期"]:
            j += 1
        if j > i:
            date_merge.append((i, j))
        i = j + 1
    return records, date_merge

def records_to_excel(records, date_merge, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "出貨順序"

    header_fill = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    date_fill   = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    odd_fill    = PatternFill(start_color="F5F9FF", end_color="F5F9FF", fill_type="solid")
    even_fill   = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    thin        = Side(style="thin")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    col_widths = [16, 12, 6, 22, 8, 10, 30]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[1].height = 22
    for ci, col_name in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=ci, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = center
    ws.freeze_panes = "A2"

    for ri, rec in enumerate(records):
        row_num = ri + 2
        fill = odd_fill if ri % 2 == 0 else even_fill
        for ci, col_name in enumerate(COLUMNS, 1):
            val  = rec.get(col_name, "")
            cell = ws.cell(row=row_num, column=ci, value=val)
            cell.border    = border
            cell.alignment = center if ci <= 3 or ci in (5, 6) else left_wrap
            cell.fill      = fill
            cell.font      = Font(size=10)
        note_lines = rec.get("備註", "").count("\n") + 1
        ws.row_dimensions[row_num].height = max(15, 15 * note_lines)

    for (start_i, end_i) in date_merge:
        start_row = start_i + 2
        end_row   = end_i   + 2
        for col_idx in (2, 3):
            ws.merge_cells(start_row=start_row, start_column=col_idx,
                           end_row=end_row,     end_column=col_idx)
            cell = ws.cell(row=start_row, column=col_idx)
            cell.alignment = center
            cell.fill = date_fill
            cell.font = Font(size=10, bold=True, color="1F3864")
            cell.border = border

    wb.save(out_path)

def convert_pdf(pdf_path, out_path, log_cb=None):
    raw = parse_pdf_to_rows(pdf_path, log_cb)
    recs, merges = merge_rows(raw)
    records_to_excel(recs, merges, out_path)
    return len(recs)

def get_preview_data(pdf_path):
    raw = parse_pdf_to_rows(pdf_path)
    recs, _ = merge_rows(raw)
    return recs[:30], len(recs)


# ═══════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 轉 Excel 工具")
        self.geometry("900x660")
        self.minsize(780, 560)
        self.configure(bg=BG)
        self.pdf_files = []
        self.out_dir   = tk.StringVar(value=os.path.expanduser("~\\Desktop"))
        self.running   = False
        self._build_ui()

    def _build_ui(self):
        hdr = tk.Frame(self, bg=CARD, height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="PDF  →  Excel", font=("Segoe UI", 18, "bold"),
                 bg=CARD, fg="white").pack(side="left", padx=24, pady=10)
        tk.Label(hdr, text="出貨單專用・合併儲存格・多頁整合",
                 font=("Segoe UI", 10), bg=CARD, fg=TEXT2).pack(side="left")

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=14)

        left = tk.Frame(body, bg=BG)
        left.pack(side="left", fill="both", expand=True)

        r1 = tk.Frame(left, bg=BG)
        r1.pack(fill="x", pady=(0, 6))
        tk.Label(r1, text="PDF 檔案清單", font=("Segoe UI", 10, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")
        tk.Button(r1, text="＋ 新增", font=("Segoe UI", 9),
                  bg=ACCENT, fg="white", relief="flat", cursor="hand2",
                  padx=10, pady=2, command=self.add_files).pack(side="right", padx=(4,0))
        tk.Button(r1, text="清除", font=("Segoe UI", 9),
                  bg=BG2, fg=TEXT2, relief="flat", cursor="hand2",
                  padx=10, pady=2, command=self.clear_files).pack(side="right")

        lf = tk.Frame(left, bg=CARD)
        lf.pack(fill="both", expand=True)
        sb = tk.Scrollbar(lf)
        sb.pack(side="right", fill="y")
        self.listbox = tk.Listbox(lf, yscrollcommand=sb.set,
                                  bg=CARD, fg=TEXT, selectbackground="#533483",
                                  font=("Segoe UI", 9), relief="flat", bd=0,
                                  highlightthickness=0, activestyle="none")
        self.listbox.pack(fill="both", expand=True, padx=2, pady=2)
        sb.config(command=self.listbox.yview)
        self.listbox.bind("<Delete>", self.remove_selected)
        tk.Label(left, text="選中後按 Delete 移除", font=("Segoe UI", 8),
                 bg=BG, fg=TEXT2).pack(anchor="w", pady=(3,0))

        right = tk.Frame(body, bg=BG, width=230)
        right.pack(side="right", fill="y", padx=(16,0))
        right.pack_propagate(False)

        tk.Label(right, text="輸出資料夾", font=("Segoe UI", 10, "bold"),
                 bg=BG, fg=TEXT).pack(anchor="w")
        dr = tk.Frame(right, bg=BG)
        dr.pack(fill="x", pady=(4,10))
        tk.Entry(dr, textvariable=self.out_dir, bg=CARD, fg=TEXT,
                 insertbackground="white", relief="flat",
                 font=("Segoe UI", 8)).pack(side="left", fill="x", expand=True, ipady=5, padx=(0,4))
        tk.Button(dr, text="瀏覽", font=("Segoe UI", 9),
                  bg=BG2, fg=TEXT2, relief="flat", cursor="hand2",
                  padx=8, pady=2, command=self.browse_dir).pack(side="right")

        self.count_lbl = tk.Label(right, text="尚未選擇檔案",
                                  font=("Segoe UI", 9), bg=BG, fg=TEXT2)
        self.count_lbl.pack(anchor="w", pady=(0,12))

        tk.Label(right, text="轉換進度", font=("Segoe UI", 10, "bold"),
                 bg=BG, fg=TEXT).pack(anchor="w")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("P.Horizontal.TProgressbar",
                        troughcolor=CARD, background=ACCENT,
                        darkcolor=ACCENT, lightcolor=ACCENT,
                        bordercolor=BG, thickness=12)
        self.pbar = ttk.Progressbar(right, style="P.Horizontal.TProgressbar",
                                    orient="horizontal", mode="determinate")
        self.pbar.pack(fill="x", pady=(4,2))
        self.pbar_lbl = tk.Label(right, text="", font=("Segoe UI", 8),
                                 bg=BG, fg=TEXT2)
        self.pbar_lbl.pack(anchor="w", pady=(0,12))

        tk.Button(right, text="🔍  預覽第一個檔案",
                  font=("Segoe UI", 10), bg=CARD, fg=TEXT,
                  relief="flat", cursor="hand2", pady=8,
                  command=self.preview).pack(fill="x", pady=(0,6))

        self.btn = tk.Button(right, text="▶  開始轉換",
                             font=("Segoe UI", 13, "bold"),
                             bg=ACCENT, fg="white", relief="flat",
                             cursor="hand2", pady=10,
                             command=self.start_convert)
        self.btn.pack(fill="x", pady=(0,6))

        tk.Button(right, text="📂  開啟輸出資料夾",
                  font=("Segoe UI", 9), bg=BG2, fg=TEXT2,
                  relief="flat", cursor="hand2", pady=6,
                  command=self.open_out).pack(fill="x")

        logf = tk.Frame(self, bg=BG2, height=160)
        logf.pack(fill="x", padx=20, pady=(0,14))
        logf.pack_propagate(False)
        tk.Label(logf, text="執行記錄", font=("Segoe UI", 9, "bold"),
                 bg=BG2, fg=TEXT2).pack(anchor="w", padx=10, pady=(5,0))
        lsb = tk.Scrollbar(logf)
        lsb.pack(side="right", fill="y")
        self.log_box = tk.Text(logf, yscrollcommand=lsb.set,
                               bg=BG2, fg=TEXT2, font=("Consolas", 8),
                               relief="flat", bd=0, state="disabled",
                               wrap="word", highlightthickness=0)
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(0,6))
        lsb.config(command=self.log_box.yview)
        self.log_box.tag_config("ok",  foreground=SUCCESS)
        self.log_box.tag_config("err", foreground=ERROR)

    def add_files(self):
        fs = filedialog.askopenfilenames(
            title="選擇 PDF 檔案",
            filetypes=[("PDF", "*.pdf"), ("所有檔案", "*.*")])
        for f in fs:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
                self.listbox.insert("end", os.path.basename(f))
        self._upd_count()

    def remove_selected(self, e=None):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            self.pdf_files.pop(i)
        self._upd_count()

    def clear_files(self):
        self.pdf_files.clear()
        self.listbox.delete(0, "end")
        self._upd_count()

    def browse_dir(self):
        d = filedialog.askdirectory(title="選擇輸出資料夾")
        if d: self.out_dir.set(d)

    def open_out(self):
        d = self.out_dir.get()
        if os.path.isdir(d): os.startfile(d)

    def _upd_count(self):
        n = len(self.pdf_files)
        self.count_lbl.config(text=f"已選 {n} 個 PDF" if n else "尚未選擇檔案")

    def log(self, msg, tag="info"):
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n", tag)
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    def preview(self):
        if not self.pdf_files:
            messagebox.showwarning("提示", "請先新增 PDF 檔案！")
            return
        pdf = self.pdf_files[0]
        self.log(f"載入預覽：{os.path.basename(pdf)}")

        def _load():
            try:
                recs, total = get_preview_data(pdf)
                self.after(0, lambda: self._show_preview(recs, total, pdf))
            except Exception as e:
                self.after(0, lambda: self.log(f"預覽失敗：{e}", "err"))

        threading.Thread(target=_load, daemon=True).start()

    def _show_preview(self, recs, total, pdf_path):
        win = tk.Toplevel(self)
        win.title(f"預覽 - {os.path.basename(pdf_path)}（前{len(recs)}筆，共{total}筆）")
        win.geometry("1100x520")
        win.configure(bg=BG)

        tk.Label(win, text=f"預覽前 {len(recs)} 筆（共 {total} 筆）",
                 font=("Segoe UI", 10), bg=BG, fg=TEXT2).pack(anchor="w", padx=12, pady=(8,4))

        frame = tk.Frame(win, bg=BG)
        frame.pack(fill="both", expand=True, padx=12, pady=(0,12))

        cols = COLUMNS
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=20)
        col_ws = [14, 10, 6, 24, 6, 8, 36]
        for c, w in zip(cols, col_ws):
            tree.heading(c, text=c)
            tree.column(c, width=w*8, anchor="center" if w < 15 else "w")

        vsb = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Treeview", background=CARD, foreground=TEXT,
                        fieldbackground=CARD, font=("Segoe UI", 9), rowheight=22)
        style.configure("Treeview.Heading", background=BG2, foreground=TEXT,
                        font=("Segoe UI", 9, "bold"))

        for rec in recs:
            vals = [rec.get(c, "").replace("\n", " / ") for c in cols]
            tree.insert("", "end", values=vals)

        btnf = tk.Frame(win, bg=BG)
        btnf.pack(fill="x", padx=12, pady=(0,10))

        def save_now():
            out = filedialog.asksaveasfilename(
                title="另存 Excel",
                defaultextension=".xlsx",
                initialfile=os.path.splitext(os.path.basename(pdf_path))[0] + ".xlsx",
                filetypes=[("Excel", "*.xlsx")])
            if not out: return
            try:
                raw = parse_pdf_to_rows(pdf_path)
                recs2, merges = merge_rows(raw)
                records_to_excel(recs2, merges, out)
                messagebox.showinfo("完成", f"已儲存：\n{out}")
                os.startfile(os.path.dirname(out))
            except Exception as e:
                messagebox.showerror("錯誤", str(e))

        tk.Button(btnf, text="💾  另存為 Excel",
                  font=("Segoe UI", 10, "bold"), bg=ACCENT, fg="white",
                  relief="flat", cursor="hand2", padx=16, pady=6,
                  command=save_now).pack(side="right")
        tk.Button(btnf, text="關閉", font=("Segoe UI", 9),
                  bg=BG2, fg=TEXT2, relief="flat", cursor="hand2",
                  padx=12, pady=6, command=win.destroy).pack(side="right", padx=(0,8))

    def start_convert(self):
        if self.running: return
        if not self.pdf_files:
            messagebox.showwarning("提示", "請先新增 PDF 檔案！")
            return
        if not os.path.isdir(self.out_dir.get()):
            messagebox.showerror("錯誤", f"輸出資料夾不存在：\n{self.out_dir.get()}")
            return
        self.running = True
        self.btn.config(state="disabled", text="⏳  轉換中...")
        self.pbar["value"] = 0
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        files   = list(self.pdf_files)
        out_dir = self.out_dir.get()
        ok = fail = 0
        self.log("═" * 46)
        self.log(f"開始批次轉換，共 {len(files)} 個")
        for i, pdf in enumerate(files):
            fname = os.path.basename(pdf)
            self.log(f"\n[{i+1}/{len(files)}] {fname}")
            self.after(0, lambda v=(i/len(files)*100): self._set_pbar(v))
            out_path = os.path.join(out_dir,
                       os.path.splitext(fname)[0] + ".xlsx")
            try:
                n = convert_pdf(pdf, out_path,
                                log_cb=lambda m: self.after(0, lambda m=m: self.log(m)))
                ok += 1
                self.log(f"  ✅ 完成！{n} 筆記錄 → {os.path.basename(out_path)}", "ok")
            except Exception as e:
                fail += 1
                self.log(f"  ❌ 失敗：{e}", "err")
        self.after(0, lambda: self._set_pbar(100))
        self.log("\n" + "═" * 46)
        self.log(f"完成：成功 {ok} / 失敗 {fail}",
                 "ok" if fail == 0 else "err")
        self.after(0, self._done)

    def _set_pbar(self, v):
        self.pbar["value"] = v
        self.pbar_lbl.config(text=f"{v:.0f}%")

    def _done(self):
        self.running = False
        self.btn.config(state="normal", text="▶  開始轉換")


if __name__ == "__main__":
    App().mainloop()
