import os
import re
import shutil
import threading
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

SCHOOLS = {"RDXB": "Repton DXB", "RAB": "RAB", "FRY": "Fry", "ROSE": "Rose"}


def norm(value): return re.sub(r"\s+", " ", str(value or "").strip()).upper()
def as_number(value):
    if value in (None, ""): return 0.0
    if isinstance(value, (int, float)): return float(value)
    try: return float(str(value).replace(",", "").strip())
    except ValueError: return 0.0

def as_date(value):
    if isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value
    if isinstance(value, str):
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try: return datetime.strptime(value.strip(), fmt).date()
            except ValueError: pass
    return None

def find_header_row(ws):
    for r in range(1, min(ws.max_row, 20) + 1):
        values = [norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if "DATE" in values and "TOTAL" in values: return r
    raise ValueError(f"Could not find the sales header row in sheet '{ws.title}'.")

def parse_source(path):
    wb = load_workbook(path, data_only=True, read_only=True); ws = wb.active
    header_row = find_header_row(ws)
    headers = {norm(ws.cell(header_row, c).value): c for c in range(1, ws.max_column + 1)}
    date_col, total_col = headers.get("DATE"), headers.get("TOTAL")
    exchange_col = next((c for h, c in headers.items() if "EXCHANGE" in h), None); campus_col = headers.get("CAMPUS")
    if not date_col or not total_col or not exchange_col:
        wb.close(); raise ValueError(f"{Path(path).name}: Date, TOTAL or Exchange column was not found.")
    transaction_cols = list(range(4, total_col)); rows, campuses = [], set()
    fallback = next((s for s in SCHOOLS if s in norm(Path(path).name)), None)
    if not fallback and norm(ws.title) in SCHOOLS: fallback = norm(ws.title)
    for values in ws.iter_rows(min_row=header_row + 1, values_only=True):
        d = as_date(values[date_col - 1] if date_col <= len(values) else None)
        if not d: continue
        campus = norm(values[campus_col - 1]) if campus_col and campus_col <= len(values) else fallback
        campus = campus if campus in SCHOOLS else fallback
        if campus not in SCHOOLS: continue
        campuses.add(campus)
        total = sum(as_number(values[c - 1] if c <= len(values) else None) for c in transaction_cols)
        exchange = as_number(values[exchange_col - 1] if exchange_col <= len(values) else None)
        rows.append((d, campus, total, exchange))
    wb.close()
    if not rows: raise ValueError(f"{Path(path).name}: no usable dated sales rows were found.")
    return rows, campuses

def choose_rows(parsed):
    combined = [(p, rows, campuses) for p, rows, campuses in parsed if set(SCHOOLS).issubset(campuses)]
    if combined: return max(combined, key=lambda item: len(item[1]))[1]
    by_school = {}
    for _, rows, campuses in parsed:
        if len(campuses) == 1: by_school[next(iter(campuses))] = rows
    missing = [s for s in SCHOOLS if s not in by_school]
    if missing: raise ValueError("Missing source report(s): " + ", ".join(missing))
    return [row for school in SCHOOLS for row in by_school[school]]

def aggregate(rows):
    sales, exchanges = {}, {}
    for d, campus, total, exchange in rows:
        sales[(d, campus)] = sales.get((d, campus), 0.0) + total; exchanges[d] = exchanges.get(d, 0.0) + exchange
    return sales, exchanges, sorted({r[0] for r in rows})

def excel_date(value):
    if isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value
    return None

def update_target_native(target_path, rows, output_path, progress):
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Native Excel support is not installed. Run the updated desktop requirements install command.") from exc
    sales, exchanges, dates = aggregate(rows); progress("Preparing updated workbook…"); shutil.copy2(target_path, output_path)
    pythoncom.CoInitialize(); excel = None; book = None
    try:
        progress("Opening tracker in Excel…")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False; excel.DisplayAlerts = False; excel.ScreenUpdating = False; excel.EnableEvents = False
        # Do not change Application.Calculation before a workbook is open. Some managed Office builds reject it.
        book = excel.Workbooks.Open(os.path.abspath(output_path), UpdateLinks=0, ReadOnly=False)
        for d in dates:
            sheet_name = d.strftime("%b-%y"); progress(f"Updating {sheet_name} — {d:%d %b %Y}…")
            try: ws = book.Worksheets(sheet_name)
            except Exception: raise ValueError(f"Target workbook has no sheet named '{sheet_name}'.")
            used_last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1; target_row = None
            # Excel COM returns midnight dates as datetime objects.
            for r in range(1, used_last + 1):
                if excel_date(ws.Cells(r, 1).Value) == d: target_row = r; break
            if not target_row: raise ValueError(f"Could not find {d:%d %b %Y} in {sheet_name}.")
            ws.Cells(target_row, 6).Value = exchanges.get(d, 0.0)
            for campus, col in {"RDXB": 8, "RAB": 9, "FRY": 10, "ROSE": 11}.items(): ws.Cells(target_row, col).Value = sales.get((d, campus), 0.0)
        progress("Saving updated tracker…"); book.Save(); book.Close(SaveChanges=True); book = None
    finally:
        if book is not None:
            try: book.Close(SaveChanges=False)
            except Exception: pass
        if excel is not None:
            try: excel.Quit()
            except Exception: pass
        pythoncom.CoUninitialize()

class AutoTaskGUI(tk.Tk):
    def __init__(self):
        super().__init__(); self.title("AutoTask — Zona Daily Sales"); self.geometry("760x520"); self.minsize(700,480); self.configure(bg="#f5f7fa")
        self.target=tk.StringVar(); self.sources=[]; self.source_text=tk.StringVar(value="No files selected"); self.status=tk.StringVar(value="Ready"); self._build()
    def _build(self):
        style=ttk.Style(self); style.theme_use("vista" if "vista" in style.theme_names() else "clam")
        style.configure("Title.TLabel",font=("Segoe UI",19,"bold"),background="#f5f7fa",foreground="#172033"); style.configure("Sub.TLabel",font=("Segoe UI",9),background="#f5f7fa",foreground="#6b7280")
        style.configure("Card.TFrame",background="white"); style.configure("Card.TLabel",background="white",foreground="#172033",font=("Segoe UI",10,"bold")); style.configure("Hint.TLabel",background="white",foreground="#7b8493",font=("Segoe UI",9))
        outer=ttk.Frame(self,padding=(42,30,42,30)); outer.pack(fill="both",expand=True); ttk.Label(outer,text="AutoTask",style="Sub.TLabel").pack(anchor="w"); ttk.Label(outer,text="Zona Daily Sales",style="Title.TLabel").pack(anchor="w",pady=(3,2)); ttk.Label(outer,text="Update the Uniform Daily Sales Tracker from Zona supplier reports.",style="Sub.TLabel").pack(anchor="w",pady=(0,20))
        card=ttk.Frame(outer,style="Card.TFrame",padding=24); card.pack(fill="both",expand=True); ttk.Label(card,text="1   Daily Sales Tracker",style="Card.TLabel").pack(anchor="w"); ttk.Label(card,text="Select the current master workbook.",style="Hint.TLabel").pack(anchor="w",pady=(3,8))
        row1=ttk.Frame(card,style="Card.TFrame"); row1.pack(fill="x"); ttk.Entry(row1,textvariable=self.target,state="readonly").pack(side="left",fill="x",expand=True); ttk.Button(row1,text="Browse…",command=self.pick_target).pack(side="left",padx=(8,0)); ttk.Separator(card).pack(fill="x",pady=22)
        ttk.Label(card,text="2   Zona report(s)",style="Card.TLabel").pack(anchor="w"); ttk.Label(card,text="Choose the combined report or all four school reports.",style="Hint.TLabel").pack(anchor="w",pady=(3,8)); row2=ttk.Frame(card,style="Card.TFrame"); row2.pack(fill="x"); ttk.Entry(row2,textvariable=self.source_text,state="readonly").pack(side="left",fill="x",expand=True); ttk.Button(row2,text="Browse…",command=self.pick_sources).pack(side="left",padx=(8,0)); ttk.Separator(card).pack(fill="x",pady=22)
        actions=ttk.Frame(card,style="Card.TFrame"); actions.pack(fill="x"); self.run_btn=ttk.Button(actions,text="Process & Save",command=self.start_processing); self.run_btn.pack(side="left"); self.progress=ttk.Progressbar(actions,mode="indeterminate",length=105); self.progress.pack(side="left",padx=12); ttk.Label(actions,textvariable=self.status,style="Hint.TLabel").pack(side="left",fill="x",expand=True)
    def pick_target(self):
        p=filedialog.askopenfilename(title="Select Uniform Daily Sales Tracker",filetypes=[("Excel workbook","*.xlsx")]);
        if p:self.target.set(p)
    def pick_sources(self):
        paths=filedialog.askopenfilenames(title="Select Zona report(s)",filetypes=[("Excel workbook","*.xlsx")]);
        if paths:self.sources=list(paths);self.source_text.set(f"{len(paths)} file(s): "+", ".join(Path(p).name for p in paths))
    def set_stage(self,text): self.after(0,lambda:self.status.set(text))
    def start_processing(self):
        if not self.target.get() or not self.sources: messagebox.showwarning("Missing files","Select the tracker and Zona report(s) first.");return
        default=Path(self.target.get()).with_name(Path(self.target.get()).stem+" - Updated.xlsx"); output=filedialog.asksaveasfilename(title="Save updated tracker",initialdir=default.parent,initialfile=default.name,defaultextension=".xlsx",filetypes=[("Excel workbook","*.xlsx")])
        if not output:return
        self.run_btn.state(["disabled"]);self.progress.start(10);self.status.set("Reading Zona reports…");threading.Thread(target=self._process,args=(output,),daemon=True).start()
    def _process(self,output):
        try:
            parsed=[]
            for i,p in enumerate(self.sources,1): self.set_stage(f"Reading Zona report {i} of {len(self.sources)}…");parsed.append((p,*parse_source(p)))
            rows=choose_rows(parsed);dates=sorted({r[0] for r in rows});self.set_stage(f"Sales loaded: {dates[0]:%d %b}–{dates[-1]:%d %b}. Opening tracker…");update_target_native(self.target.get(),rows,output,self.set_stage);self.after(0,lambda:self._finish(output,None))
        except Exception as exc:self.after(0,lambda e=str(exc):self._finish(None,e))
    def _finish(self,output,error):
        self.progress.stop();self.run_btn.state(["!disabled"])
        if error:self.status.set("Failed");messagebox.showerror("Processing failed",error)
        else:
            self.status.set("Completed")
            if messagebox.askyesno("Completed",f"Updated tracker saved successfully.\n\n{output}\n\nOpen the file now?"):os.startfile(output)

if __name__=="__main__":AutoTaskGUI().mainloop()
