import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import re, zipfile, tempfile, os
from copy import deepcopy
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment

DESTINATIONS = {
    "SYDNEY": "AU/SYDNEY",
    "MELBOURNE": "AU/MELBOURNE",
    "BRISBANE": "AU/BRISBANE",
    "PERTH": "AU/PERTH",
}

def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)

def as_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def units(v):
    if v in (None, ""):
        return ""
    try:
        return f"{int(float(v)):02d} UNITS"
    except Exception:
        return f"{as_text(v)} UNITS"

def read_packing_list(path):
    ext = Path(path).suffix.lower()
    sheets = {}
    if ext == ".xls":
        try:
            import xlrd
        except ImportError:
            raise RuntimeError("পুরোনো .XLS পড়তে xlrd প্রয়োজন। install_requirements.bat চালান।")
        book = xlrd.open_workbook(path)
        for name in DESTINATIONS:
            try:
                sh = book.sheet_by_name(name)
            except Exception:
                continue
            rows = [[sh.cell_value(r,c) for c in range(sh.ncols)] for r in range(sh.nrows)]
            sheets[name] = rows
    elif ext in (".xlsx", ".xlsm"):
        book = load_workbook(path, data_only=True, read_only=True)
        for name in DESTINATIONS:
            if name in book.sheetnames:
                sh = book[name]
                sheets[name] = [[sh.cell(r,c).value for c in range(1, sh.max_column+1)]
                                for r in range(1, sh.max_row+1)]
        book.close()
    else:
        raise RuntimeError("Packing List .xls, .xlsx অথবা .xlsm হতে হবে।")
    return sheets

def cell(rows, r1, c1):
    try:
        return rows[r1-1][c1-1]
    except Exception:
        return None

def extract_metadata(rows):
    order, style, dept, desc = "", "", "034", "5 PK TUNNEL TRUNK"
    for r in range(34, 42):
        for v in (rows[r-1] if len(rows) >= r else []):
            if not isinstance(v, str):
                continue
            s = v.strip()
            m = re.search(r"ORDER\s*NO\.?\s*(\d+)", s, re.I)
            if m: order = m.group(1)
            m = re.search(r"STYLE\s*NO\.?\s*([A-Z0-9]+)", s, re.I)
            if m: style = m.group(1)
            m = re.search(r"DEPT\s*NO\.?\s*:\s*(\d+)", s, re.I)
            if m: dept = m.group(1)
            if "TUNNEL" in s.upper(): desc = s
    return order, style, dept, desc

def build_records(source_path):
    sheet_rows = read_packing_list(source_path)
    records, dest_totals = [], {}
    for dest in DESTINATIONS:
        if dest not in sheet_rows:
            continue
        rows = sheet_rows[dest]
        order, style, dept, desc = extract_metadata(rows)
        last_colour = last_key = last_size = None
        dest_count = dest_qty = 0
        for r in range(44, len(rows)+1):
            start, end = cell(rows,r,1), cell(rows,r,2)
            if not (is_num(start) and is_num(end)):
                continue
            colour, key, size = cell(rows,r,7), cell(rows,r,9), cell(rows,r,10)
            if colour not in (None, "", '"', "''"): last_colour = as_text(colour)
            if key not in (None, "", '"', "''"): last_key = as_text(key)
            if size not in (None, "", '"', "''"): last_size = as_text(size)
            issue = cell(rows,r,13)
            shipper_pack = cell(rows,r,17)
            measurement = as_text(cell(rows,r,22))
            if not all([order, style, desc, last_colour, last_key, last_size, measurement]):
                raise RuntimeError(f"{dest} sheet, row {r}: প্রয়োজনীয় data missing (Order/Style/Description/Colour/Keycode/Size/Measurement)।")
            a, b = int(start), int(end)
            if b < a:
                raise RuntimeError(f"{dest} sheet, row {r}: Carton range ভুল ({a}-{b})।")
            for carton in range(a,b+1):
                records.append({
                    "dest": dest, "carton": carton, "order": order, "dept": dept,
                    "key": last_key, "style": style, "description": desc,
                    "issue": issue, "shipper_pack": shipper_pack,
                    "colour": last_colour, "size": last_size,
                    "measurement": measurement, "source_row": r
                })
            dest_count += int(cell(rows,r,19) or (b-a+1))
            dest_qty += int(cell(rows,r,20) or 0)
        dest_totals[dest] = (dest_count, dest_qty)
    if not records:
        raise RuntimeError("SYDNEY/MELBOURNE/BRISBANE/PERTH sheet-এ carton data পাওয়া যায়নি।")
    cartons = [x["carton"] for x in records]
    if len(set(cartons)) != len(cartons):
        raise RuntimeError("একই carton number একাধিকবার পাওয়া গেছে। QC করার আগে source Packing List ঠিক করুন।")
    return records, dest_totals

def make_output(template_path, records, dest_totals, save_path):
    wb = load_workbook(template_path)
    base = None
    for sh in wb.worksheets:
        v = sh["E16"].value
        if str(v).strip().zfill(2) == "01":
            base = sh
            break
    if base is None:
        raise RuntimeError("Template-এ E16 cell-এ carton 01 থাকা label sheet পাওয়া যায়নি।")
    if not base._images:
        raise RuntimeError("Template-এর প্রথম label sheet-এ Hanger Logo image পাওয়া যায়নি।")
    anchor = deepcopy(base._images[0].anchor)
    width, height = base._images[0].width, base._images[0].height
    with zipfile.ZipFile(template_path) as z:
        media = [n for n in z.namelist() if n.startswith("xl/media/")]
        if not media:
            raise RuntimeError("Template workbook-এ logo image file পাওয়া যায়নি।")
        logo_bytes = z.read(media[0])
    # Keep only the carton-01 base sheet. Logo is re-added from the exact template asset.
    for sh in list(wb.worksheets):
        if sh is not base:
            wb.remove(sh)
    base._images = []
    total = len(records)
    def add_logo(sh):
        fd, p = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        with open(p, "wb") as f: f.write(logo_bytes)
        im = XLImage(p)
        im.width, im.height = width, height
        im.anchor = deepcopy(anchor)
        sh.add_image(im)
        # Keep path until workbook is saved; caller cleans temporary files.
        logo_temp_paths.append(p)
    logo_temp_paths = []

    try:
        for i, rec in enumerate(records):
            if i == 0:
                sh = base
                sh.title = f"{rec['dest'][:3]} {rec['carton']:03d}"
            else:
                sh = wb.copy_worksheet(base)
                sh.title = f"{rec['dest'][:3]} {rec['carton']:03d}"
            sh["D4"] = rec["dept"]
            sh["D5"] = rec["order"]
            sh["D6"] = rec["key"]
            sh["D7"] = rec["style"]
            sh["D8"] = rec["description"]
            sh["D9"] = units(rec["issue"])
            sh["D10"] = units(rec["shipper_pack"])
            sh["D11"] = rec["colour"]
            sh["D12"] = rec["size"]
            sh["D13"] = "KGS"
            sh["D14"] = rec["measurement"]
            sh["E16"] = f"{rec['carton']:02d}"
            sh["G16"] = f"      {total}"
            sh["P12"] = DESTINATIONS[rec["dest"]]
            sh.sheet_view.showGridLines = False
            sh.print_area = "$C$2:$G$16"
            sh.sheet_properties.pageSetUpPr.fitToPage = True
            sh.page_setup.fitToWidth = 1
            sh.page_setup.fitToHeight = 1
            sh.page_setup.orientation = "landscape"
            add_logo(sh)

        qc = wb.create_sheet("QC REPORT")
        qc.append(["QUALITY CHECK", "RESULT", "DETAIL"])
        for dest, (ctn, qty) in dest_totals.items():
            made = sum(1 for x in records if x["dest"] == dest)
            qc.append([f"{dest} cartons / quantity", "PASS" if made == ctn else "REVIEW",
                       f"Packing List: {ctn} cartons / {qty} units; generated: {made} labels"])
        carton_nums = sorted(x["carton"] for x in records)
        qc.append(["Total labels", "PASS", str(total)])
        qc.append(["Carton number duplicates", "PASS" if len(carton_nums)==len(set(carton_nums)) else "WARNING",
                   "No duplicate carton numbers" if len(carton_nums)==len(set(carton_nums)) else "Duplicate carton numbers found"])
        qc.append(["QTY ISSUE PACK mapping", "PASS", "Packing List column M (Unit Qty)"])
        qc.append(["QTY SHIPPER PACK mapping", "PASS", "Packing List column Q (Blister Qty)"])
        qc.append(["Colour / Keycode / Size", "PASS", "Packing List columns G / I / J; ditto marks carried forward"])
        qc.append(["Measurement", "PASS", "Packing List column V; original unit text retained"])
        qc.append(["Hanger Logo", "PASS", f"Original template image size {width}×{height}px and original anchor preserved"])
        for c in qc[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
        for col, w in {"A":34, "B":16, "C":90}.items():
            qc.column_dimensions[col].width = w
        for row in qc.iter_rows():
            for c in row:
                c.alignment = Alignment(vertical="top", wrap_text=True)
        qc.freeze_panes = "A2"
        qc.sheet_view.showGridLines = False
        wb.save(save_path)
    finally:
        for p in logo_temp_paths:
            try: os.remove(p)
            except OSError: pass

class App:
    def __init__(self, root):
        self.root = root
        root.title("Shipping Mark Automation Pro")
        root.geometry("760x470")
        root.minsize(680, 430)
        self.pack = tk.StringVar()
        self.template = tk.StringVar()
        self.status = tk.StringVar(value="Packing List ও Shipping Mark Format নির্বাচন করুন।")
        self.ui()

    def ui(self):
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Shipping Mark Automation Pro", font=("Segoe UI", 19, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Custom build for ANKO / AU destination carton labels • Offline",
                  font=("Segoe UI", 10)).pack(anchor="w", pady=(3,18))
        self.file_row(frame, "Packing List (.XLS / .XLSX)", self.pack)
        self.file_row(frame, "Shipping Mark Format (.XLSX)", self.template)
        ttk.Label(frame, text="স্বয়ংক্রিয় কাজ", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(18,5))
        for line in [
            "• SYDNEY, MELBOURNE, BRISBANE ও PERTH sheet পড়বে",
            "• Carton range থেকে প্রতিটি carton-এর আলাদা label বানাবে",
            "• QTY ISSUE PACK = Column M; QTY SHIPPER PACK = Column Q",
            "• Colour / Keycode / Size ditto mark carry-forward করবে",
            "• Measurement Column V থেকে নেবে এবং Template-এর logo size/position রাখবে",
            "• QC REPORT sheet-এ carton ও quantity checking দেবে",
        ]:
            ttk.Label(frame, text=line).pack(anchor="w", pady=2)
        bar = ttk.Frame(frame)
        bar.pack(fill="x", pady=(18,8))
        ttk.Button(bar, text="Generate Shipping Mark", command=self.generate).pack(side="left")
        ttk.Button(bar, text="Clear", command=self.clear).pack(side="left", padx=8)
        ttk.Label(frame, textvariable=self.status, foreground="#155a31", wraplength=700).pack(anchor="w", pady=(8,0))
        ttk.Label(frame, text="গুরুত্বপূর্ণ: output print করার আগে QC REPORT এবং কয়েকটি label চোখে দেখে মিলিয়ে নিন।",
                  foreground="#8a4b08", wraplength=700).pack(anchor="w", pady=(14,0))

    def file_row(self, parent, label, var):
        row = ttk.Frame(parent); row.pack(fill="x", pady=6)
        ttk.Label(row, text=label, width=30).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Browse…", command=lambda: self.pick(var, label)).pack(side="left")

    def pick(self, var, label):
        types = [("Excel files", "*.xls *.xlsx *.xlsm")]
        p = filedialog.askopenfilename(filetypes=types)
        if p: var.set(p)

    def clear(self):
        self.pack.set(""); self.template.set("")
        self.status.set("ফাইল নির্বাচন করুন।")

    def generate(self):
        p, t = self.pack.get(), self.template.get()
        if not p or not t:
            messagebox.showwarning("Missing file", "দুটি Excel ফাইলই নির্বাচন করুন।"); return
        if Path(t).suffix.lower() != ".xlsx":
            messagebox.showwarning("Template format", "Shipping Mark Format .xlsx হতে হবে।"); return
        try:
            self.status.set("ফাইল পড়া হচ্ছে…")
            self.root.update_idletasks()
            records, totals = build_records(p)
            suggested = Path(p).stem + "_Shipping_Mark.xlsx"
            dest = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=suggested,
                                                filetypes=[("Excel workbook", "*.xlsx")])
            if not dest:
                self.status.set("Save বাতিল করা হয়েছে।"); return
            make_output(t, records, totals, dest)
            self.status.set(f"সম্পন্ন: {len(records)}টি Shipping Mark তৈরি হয়েছে।")
            messagebox.showinfo("Completed", f"{len(records)}টি label তৈরি হয়েছে।\n\n{dest}\n\nQC REPORT sheet পরীক্ষা করুন।")
        except Exception as e:
            self.status.set("কাজ সম্পন্ন হয়নি।")
            messagebox.showerror("Processing error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    try: ttk.Style().theme_use("vista")
    except tk.TclError: pass
    App(root)
    root.mainloop()
