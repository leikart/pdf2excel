import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import re

# ── 依賴套件（打包時會一起帶入）──────────────────────
try:
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell
except ImportError as e:
    import sys
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "pdfplumber", "openpyxl", "-q"])
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell

# ── 顏色常數 ─────────────────────────────────────────
BG        = "#1A1A2E"
BG2       = "#16213E"
CARD      = "#0F3460"
ACCENT    = "#E94560"
ACCENT2   = "#533483"
TEXT      = "#EAEAEA"
TEXT2     = "#A0A8C0"
SUCCESS   = "#4CAF82"
WARNING   = "#F0A500"
ERROR     = "#E94560"
WHITE     = "#FFFFFF"

# ── 表格擷取設定 ─────────────────────────────────────
TABLE_SETTINGS_LINES = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 3,
    "join_tolerance": 3,
    "edge_min_length": 3,
    "min_words_vertical": 1,
    "min_words_horizontal": 1,
    "intersection_tolerance": 3,
    "text_tolerance": 3,
}
TABLE_SETTINGS_TEXT = {
    **TABLE_SETTINGS_LINES,
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
}

HEADER_FILL  = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
HEADER_FONT  = Font(color="FFFFFF", bold=True, size=11)
ROW_ODD      = PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid")
ROW_EVEN     = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
THIN_BORDER  = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

def clean_cell(v):
    if v is None: return ""
    return re.sub(r"\s+", " ", str(v).strip())

def auto_fit(ws):
    for col in ws.columns:
        mx = 0
        letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    mx = max(mx, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[letter].width = min(max(mx + 4, 10), 50)

def extract_tables(page):
    tables = page.extract_tables(TABLE_SETTINGS_LINES)
    if tables and any(len(t) > 0 for t in tables):
        return tables, "lines"
    tables = page.extract_tables(TABLE_SETTINGS_TEXT)
    return tables, "text"

def write_table(ws, data, t_idx, page_num):
    if not data: return 0
    start = ws.max_row + 2 if ws.max_row > 1 else 1
    if t_idx > 0:
        c = ws.cell(row=start, column=1,
                    value=f"▶ 第 {page_num} 頁 - 表格 {t_idx + 1}")
        c.font = Font(bold=True, color="C55A11", size=10)
        start += 1
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = ws.cell(row=start + ri, column=ci + 1)
            if isinstance(cell, MergedCell): continue
            cell.value = clean_cell(val)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            if ri == 0:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.fill = ROW_ODD if ri % 2 == 1 else ROW_EVEN
                cell.font = Font(size=10)
    ws.row_dimensions[start].height = 20
    return len(data)

def convert_pdf(pdf_path, out_dir, log_cb, prog_cb):
    base    = os.path.splitext(os.path.basename(pdf_path))[0]
    out     = os.path.join(out_dir, base + ".xlsx")
    wb      = Workbook()
    wb.remove(wb.active)
    log     = []
    total_t = 0

    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages
        total = len(pages)
        log_cb(f"  📑 共 {total} 頁")
        for i, page in enumerate(pages, 1):
            prog_cb(i / total)
            tables, strategy = extract_tables(page)
            valid = [t for t in (tables or []) if t and len(t) > 0]
            if not valid:
                log.append(f"第{i}頁：無表格")
                continue
            ws = wb.create_sheet(title=f"第{i}頁")
            c  = ws.cell(row=1, column=1,
                         value=f"{os.path.basename(pdf_path)}  |  第{i}/{total}頁  |  {strategy}")
            c.font  = Font(bold=True, size=11, color="1F3864")
            c.fill  = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
            ws.row_dimensions[1].height = 22
            for ti, t in enumerate(valid):
                write_table(ws, t, ti, i)
                total_t += 1
                log.append(f"第{i}頁 表格{ti+1}：{len(t)}列×{len(t[0]) if t else 0}欄 [{strategy}]")
            auto_fit(ws)
            ws.freeze_panes = "A3"

    ws_log = wb.create_sheet(title="轉換摘要", index=0)
    ws_log.column_dimensions["A"].width = 55
    for r in [("📄 來源", os.path.basename(pdf_path)),
              ("📊 擷取表格", total_t), ("", ""),
              ("── 各頁紀錄 ──", "")]:
        ws_log.append(r)
    for line in log:
        ws_log.append([line])
    for row in ws_log.iter_rows(min_row=1, max_row=3):
        for cell in row:
            cell.font = Font(bold=True, size=11)

    wb.save(out)
    return out, total_t


# ═══════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 轉 Excel 工具")
        self.geometry("780x620")
        self.minsize(680, 520)
        self.configure(bg=BG)
        self.resizable(True, True)

        self.pdf_files = []
        self.out_dir   = tk.StringVar(value=os.path.expanduser("~\\Desktop"))
        self.running   = False

        self._build_ui()

    # ── 建立介面 ──────────────────────────────────────
    def _build_ui(self):
        # 頂部標題列
        hdr = tk.Frame(self, bg=CARD, height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="PDF  →  Excel", font=("Segoe UI", 20, "bold"),
                 bg=CARD, fg=WHITE).pack(side="left", padx=24, pady=14)
        tk.Label(hdr, text="支援分隔線偵測・多檔批次轉換",
                 font=("Segoe UI", 10), bg=CARD, fg=TEXT2).pack(side="left", padx=0, pady=20)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        # ── 左側：檔案清單區 ──────────────────────────
        left = tk.Frame(body, bg=BG)
        left.pack(side="left", fill="both", expand=True)

        row1 = tk.Frame(left, bg=BG)
        row1.pack(fill="x", pady=(0, 8))
        tk.Label(row1, text="PDF 檔案清單", font=("Segoe UI", 11, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")
        tk.Button(row1, text="＋ 新增", font=("Segoe UI", 9),
                  bg=ACCENT, fg=WHITE, relief="flat", cursor="hand2",
                  padx=10, pady=3,
                  command=self.add_files).pack(side="right", padx=(4, 0))
        tk.Button(row1, text="清除全部", font=("Segoe UI", 9),
                  bg=BG2, fg=TEXT2, relief="flat", cursor="hand2",
                  padx=10, pady=3,
                  command=self.clear_files).pack(side="right")

        # 檔案清單（Listbox + Scrollbar）
        list_frame = tk.Frame(left, bg=CARD, bd=0)
        list_frame.pack(fill="both", expand=True)
        sb = tk.Scrollbar(list_frame)
        sb.pack(side="right", fill="y")
        self.listbox = tk.Listbox(
            list_frame, yscrollcommand=sb.set,
            bg=CARD, fg=TEXT, selectbackground=ACCENT2,
            selectforeground=WHITE, font=("Segoe UI", 9),
            relief="flat", bd=0, highlightthickness=0,
            activestyle="none",
        )
        self.listbox.pack(fill="both", expand=True, padx=2, pady=2)
        sb.config(command=self.listbox.yview)
        self.listbox.bind("<Delete>", self.remove_selected)

        tk.Label(left, text="選中後按 Delete 可移除單一檔案",
                 font=("Segoe UI", 8), bg=BG, fg=TEXT2).pack(anchor="w", pady=(4, 0))

        # ── 右側：設定與動作 ──────────────────────────
        right = tk.Frame(body, bg=BG, width=240)
        right.pack(side="right", fill="y", padx=(16, 0))
        right.pack_propagate(False)

        # 輸出資料夾
        tk.Label(right, text="輸出資料夾", font=("Segoe UI", 10, "bold"),
                 bg=BG, fg=TEXT).pack(anchor="w")
        dir_row = tk.Frame(right, bg=BG)
        dir_row.pack(fill="x", pady=(4, 12))
        self.dir_entry = tk.Entry(dir_row, textvariable=self.out_dir,
                                  bg=CARD, fg=TEXT, insertbackground=WHITE,
                                  relief="flat", font=("Segoe UI", 8))
        self.dir_entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 4))
        tk.Button(dir_row, text="瀏覽", font=("Segoe UI", 9),
                  bg=BG2, fg=TEXT, relief="flat", cursor="hand2",
                  padx=8, pady=3, command=self.browse_dir).pack(side="right")

        # 統計標籤
        self.count_label = tk.Label(right, text="尚未選擇檔案",
                                    font=("Segoe UI", 9), bg=BG, fg=TEXT2)
        self.count_label.pack(anchor="w", pady=(0, 16))

        # 進度條
        tk.Label(right, text="轉換進度", font=("Segoe UI", 10, "bold"),
                 bg=BG, fg=TEXT).pack(anchor="w")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=CARD, background=ACCENT,
                        darkcolor=ACCENT, lightcolor=ACCENT,
                        bordercolor=BG, thickness=14)
        self.pbar = ttk.Progressbar(right, style="Custom.Horizontal.TProgressbar",
                                    orient="horizontal", length=220, mode="determinate")
        self.pbar.pack(fill="x", pady=(4, 4))
        self.pbar_label = tk.Label(right, text="", font=("Segoe UI", 8),
                                   bg=BG, fg=TEXT2)
        self.pbar_label.pack(anchor="w", pady=(0, 16))

        # 開始按鈕
        self.btn_start = tk.Button(
            right, text="▶  開始轉換",
            font=("Segoe UI", 13, "bold"),
            bg=ACCENT, fg=WHITE, relief="flat", cursor="hand2",
            pady=12, command=self.start_convert,
        )
        self.btn_start.pack(fill="x", pady=(0, 8))

        self.btn_open = tk.Button(
            right, text="📂  開啟輸出資料夾",
            font=("Segoe UI", 9),
            bg=BG2, fg=TEXT2, relief="flat", cursor="hand2",
            pady=6, command=self.open_out_dir,
        )
        self.btn_open.pack(fill="x")

        # ── 底部 Log ──────────────────────────────────
        log_frame = tk.Frame(self, bg=BG2, height=180)
        log_frame.pack(fill="x", padx=20, pady=(0, 16))
        log_frame.pack_propagate(False)
        tk.Label(log_frame, text="執行記錄", font=("Segoe UI", 9, "bold"),
                 bg=BG2, fg=TEXT2).pack(anchor="w", padx=10, pady=(6, 0))
        log_sb = tk.Scrollbar(log_frame)
        log_sb.pack(side="right", fill="y")
        self.log_box = tk.Text(
            log_frame, yscrollcommand=log_sb.set,
            bg=BG2, fg=TEXT2, font=("Consolas", 8),
            relief="flat", bd=0, state="disabled",
            wrap="word", highlightthickness=0,
        )
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        log_sb.config(command=self.log_box.yview)
        self.log_box.tag_config("ok",   foreground=SUCCESS)
        self.log_box.tag_config("err",  foreground=ERROR)
        self.log_box.tag_config("info", foreground=TEXT2)

    # ── 動作函式 ─────────────────────────────────────
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="選擇 PDF 檔案",
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")]
        )
        for f in files:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
                self.listbox.insert("end", os.path.basename(f))
        self._update_count()

    def remove_selected(self, event=None):
        sel = list(self.listbox.curselection())
        for i in reversed(sel):
            self.listbox.delete(i)
            self.pdf_files.pop(i)
        self._update_count()

    def clear_files(self):
        self.pdf_files.clear()
        self.listbox.delete(0, "end")
        self._update_count()

    def browse_dir(self):
        d = filedialog.askdirectory(title="選擇輸出資料夾")
        if d:
            self.out_dir.set(d)

    def open_out_dir(self):
        d = self.out_dir.get()
        if os.path.isdir(d):
            os.startfile(d)

    def _update_count(self):
        n = len(self.pdf_files)
        self.count_label.config(
            text=f"已選 {n} 個 PDF 檔案" if n else "尚未選擇檔案"
        )

    def log(self, msg, tag="info"):
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n", tag)
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    def start_convert(self):
        if self.running: return
        if not self.pdf_files:
            messagebox.showwarning("提示", "請先新增 PDF 檔案！")
            return
        out = self.out_dir.get()
        if not os.path.isdir(out):
            messagebox.showerror("錯誤", f"輸出資料夾不存在：\n{out}")
            return
        self.running = True
        self.btn_start.config(state="disabled", text="⏳  轉換中...")
        self.pbar["value"] = 0
        threading.Thread(target=self._run_convert, daemon=True).start()

    def _run_convert(self):
        files   = list(self.pdf_files)
        out_dir = self.out_dir.get()
        ok = fail = 0
        self.log("═" * 48, "info")
        self.log(f"開始批次轉換，共 {len(files)} 個檔案", "info")

        for idx, pdf in enumerate(files):
            fname = os.path.basename(pdf)
            self.log(f"\n[{idx+1}/{len(files)}] {fname}", "info")

            # 整體進度
            self.pbar["value"] = idx / len(files) * 100

            def log_cb(msg):
                self.log(msg, "info")

            def prog_cb(frac, _idx=idx, _total=len(files)):
                overall = (_idx + frac) / _total * 100
                self.after(0, lambda v=overall: self._set_pbar(v))

            try:
                out_path, n_tables = convert_pdf(pdf, out_dir, log_cb, prog_cb)
                ok += 1
                self.log(f"  ✅ 完成！{n_tables} 個表格 → {os.path.basename(out_path)}", "ok")
            except Exception as e:
                fail += 1
                self.log(f"  ❌ 失敗：{e}", "err")

        self.after(0, lambda: self._set_pbar(100))
        self.log("\n" + "═" * 48, "info")
        self.log(f"完成：成功 {ok} / 失敗 {fail}", "ok" if fail == 0 else "err")
        self.after(0, self._done)

    def _set_pbar(self, v):
        self.pbar["value"] = v
        self.pbar_label.config(text=f"{v:.0f}%")

    def _done(self):
        self.running = False
        self.btn_start.config(state="normal", text="▶  開始轉換")


if __name__ == "__main__":
    app = App()
    app.mainloop()
