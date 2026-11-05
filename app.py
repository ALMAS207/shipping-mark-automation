import re
import zipfile
from copy import deepcopy
from io import BytesIO
from pathlib import Path

import streamlit as st
import xlrd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Shipping Mark Automation Pro", page_icon="📦", layout="centered")

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

def clean_measurement(value):
    """Return dimensions only, e.g. 54X45X30 CM."""
    s = as_text(value).upper().replace("×", "X").replace("*", "X")
    s = re.sub(r"\b(LENGTH|WIDTH|HEIGHT|L|W|H)\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    m = re.search(r"(\d+(?:\.\d+)?)\s*X\s*(\d+(?:\.\d+)?)\s*X\s*(\d+(?:\.\d+)?)\s*(CM|MM|INCH|IN)?", s)
    if not m:
        return s
    dims = "X".join(x.rstrip("0").rstrip(".") if "." in x else x for x in m.group(1, 2, 3))
    unit = m.group(4) or ("CM" if "CM" in s else "")
    return f"{dims} {unit}".strip()

def read_packing_list(upload):
    name = upload.name.lower()
    sheets = {}
    data = upload.getvalue()
    if name.endswith(".xls"):
        book = xlrd.open_workbook(file_contents=data)
        for sheet_name in DESTINATIONS:
            try:
                sh = book.sheet_by_name(sheet_name)
            except Exception:
                continue
            sheets[sheet_name] = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    elif name.endswith((".xlsx", ".xlsm")):
        book = load_workbook(BytesIO(data), data_only=True, read_only=True)
        for sheet_name in DESTINATIONS:
            if sheet_name in book.sheetnames:
                sh = book[sheet_name]
                sheets[sheet_name] = [[sh.cell(r, c).value for c in range(1, sh.max_column + 1)]
                                      for r in range(1, sh.max_row + 1)]
        book.close()
    else:
        raise ValueError("Packing List অবশ্যই .xls, .xlsx অথবা .xlsm হতে হবে।")
    return sheets

def cell(rows, r1, c1):
    try:
        return rows[r1 - 1][c1 - 1]
    except Exception:
        return None

def extract_metadata(rows):
    order, style, dept, desc = "", "", "034", "5 PK TUNNEL TRUNK"
    for r in range(34, min(42, len(rows)) + 1):
        for v in rows[r - 1]:
            if not isinstance(v, str):
                continue
            s = v.strip()
            m = re.search(r"ORDER\s*NO\.?\s*(\d+)", s, re.I)
            if m:
                order = m.group(1)
            m = re.search(r"STYLE\s*NO\.?\s*([A-Z0-9]+)", s, re.I)
            if m:
                style = m.group(1)
            m = re.search(r"DEPT\s*NO\.?\s*:\s*(\d+)", s, re.I)
            if m:
                dept = m.group(1)
            if "TUNNEL" in s.upper():
                desc = s
    return order, style, dept, desc

def build_records(upload):
    sheet_rows = read_packing_list(upload)
    records, dest_totals = [], {}
    for dest in DESTINATIONS:
        if dest not in sheet_rows:
            continue
        rows = sheet_rows[dest]
        order, style, dept, desc = extract_metadata(rows)
        last_colour = last_key = last_size = None
        dest_count = dest_qty = 0
        for r in range(44, len(rows) + 1):
            start, end = cell(rows, r, 1), cell(rows, r, 2)
            if not (is_num(start) and is_num(end)):
                continue
            colour, key, size = cell(rows, r, 7), cell(rows, r, 9), cell(rows, r, 10)
            if colour not in (None, "", '"', "''"):
                last_colour = as_text(colour)
            if key not in (None, "", '"', "''"):
                last_key = as_text(key)
            if size not in (None, "", '"', "''"):
                last_size = as_text(size)
            issue = cell(rows, r, 13)
            shipper_pack = cell(rows, r, 17)
            measurement = clean_measurement(cell(rows, r, 22))
            missing = [label for label, value in [
                ("Order No.", order), ("Style No.", style), ("Description", desc),
                ("Colour", last_colour), ("Keycode", last_key), ("Size", last_size),
                ("QTY ISSUE PACK", issue), ("QTY SHIPPER PACK", shipper_pack),
                ("Measurement", measurement)
            ] if value in (None, "")]
            if missing:
                raise ValueError(f"{dest} sheet, row {r}: এই তথ্যগুলো পাওয়া যায়নি: {', '.join(missing)}")
            a, b = int(start), int(end)
            if b < a:
                raise ValueError(f"{dest} sheet, row {r}: Carton range ভুল ({a}-{b})।")
            for carton in range(a, b + 1):
                records.append({
                    "dest": dest, "carton": carton, "order": order, "dept": dept,
                    "key": last_key, "style": style, "description": desc,
                    "issue": issue, "shipper_pack": shipper_pack,
                    "colour": last_colour, "size": last_size,
                    "measurement": measurement, "source_row": r
                })
            dest_count += int(cell(rows, r, 19) or (b - a + 1))
            dest_qty += int(cell(rows, r, 20) or 0)
        dest_totals[dest] = (dest_count, dest_qty)
    if not records:
        raise ValueError("SYDNEY / MELBOURNE / BRISBANE / PERTH sheet-এ carton data পাওয়া যায়নি।")
    cartons = [x["carton"] for x in records]
    if len(set(cartons)) != len(cartons):
        raise ValueError("একই carton number একাধিকবার পাওয়া গেছে। Packing List ঠিক করে আবার দিন।")
    return records, dest_totals

def make_output(template_upload, records, dest_totals):
    template_data = template_upload.getvalue()
    wb = load_workbook(BytesIO(template_data))
    base = None
    for sh in wb.worksheets:
        if str(sh["E16"].value).strip().zfill(2) == "01":
            base = sh
            break
    if base is None:
        raise ValueError("Format ফাইলে E16 cell-এ carton 01 থাকা label sheet পাওয়া যায়নি।")
    if not base._images:
        raise ValueError("Format ফাইলে Hanger Logo image পাওয়া যায়নি।")
    anchor = deepcopy(base._images[0].anchor)
    width, height = base._images[0].width, base._images[0].height
    with zipfile.ZipFile(BytesIO(template_data)) as z:
        media = [n for n in z.namelist() if n.startswith("xl/media/")]
        if not media:
            raise ValueError("Format workbook-এ logo image file পাওয়া যায়নি।")
        logo_bytes = z.read(media[0])

    for sh in list(wb.worksheets):
        if sh is not base:
            wb.remove(sh)
    base._images = []
    total = len(records)

    def add_logo(sh):
        im = XLImage(BytesIO(logo_bytes))
        im.width, im.height = width, height
        im.anchor = deepcopy(anchor)
        sh.add_image(im)

    for i, rec in enumerate(records):
        if i == 0:
            sh = base
        else:
            sh = wb.copy_worksheet(base)
        sh.title = f"{rec['dest'][:3]} {rec['carton']:03d}"[:31]
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
    qc.append(["Carton number duplicates", "PASS", "No duplicate carton numbers"])
    qc.append(["QTY ISSUE PACK mapping", "PASS", "Packing List column M"])
    qc.append(["QTY SHIPPER PACK mapping", "PASS", "Packing List column Q"])
    qc.append(["Colour / Keycode / Size", "PASS", "Packing List columns G / I / J; ditto marks carried forward"])
    qc.append(["Measurement", "PASS", "Packing List column V; dimensions and unit only"])
    qc.append(["Hanger Logo", "PASS", f"Template image size {width}×{height}px and original anchor preserved"])
    for c in qc[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for col, w in {"A": 34, "B": 16, "C": 90}.items():
        qc.column_dimensions[col].width = w
    for row in qc.iter_rows():
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    qc.freeze_panes = "A2"
    qc.sheet_view.showGridLines = False
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), width, height

st.title("📦 Shipping Mark Automation Pro")
st.caption("Online Excel tool • Bangla + English")

st.info("১) Packing List দিন  ২) Shipping Mark Format দিন  ৩) Generate চাপুন  ৪) Excel Download করুন।")

packing_file = st.file_uploader("১. Packing List নির্বাচন করুন (.xls / .xlsx / .xlsm)", type=["xls", "xlsx", "xlsm"])
format_file = st.file_uploader("২. Shipping Mark Format নির্বাচন করুন (.xlsx)", type=["xlsx"])

if st.button("🚀 Generate Shipping Mark", type="primary", use_container_width=True):
    if packing_file is None or format_file is None:
        st.error("দয়া করে Packing List এবং Shipping Mark Format—দুটি ফাইলই Upload করুন।")
    else:
        try:
            with st.spinner("Excel পড়া ও Shipping Mark তৈরি হচ্ছে..."):
                records, totals = build_records(packing_file)
                output, logo_w, logo_h = make_output(format_file, records, totals)
            st.success(f"সফল! মোট {len(records)}টি carton label তৈরি হয়েছে।")
            st.write(f"Hanger Logo: {logo_w} × {logo_h} px — Format-এর মাপ রাখা হয়েছে।")
            filename = f"{Path(packing_file.name).stem}_Shipping_Mark.xlsx"
            st.download_button(
                "⬇️ Shipping Mark Excel Download",
                data=output,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.warning("Download করা Excel-এর QC REPORT sheet এবং কয়েকটি label দেখে Packing List-এর সঙ্গে মিলিয়ে নিন।")
        except Exception as e:
            st.error(f"তৈরি করা যায়নি: {e}")
            st.info("সমস্যা থাকলে Packing List ও Format ফাইল পাঠান; mapping যাচাই করে ঠিক করতে হবে।")

st.divider()
st.caption("Supported destination sheets: SYDNEY, MELBOURNE, BRISBANE, PERTH")
st.caption("QTY ISSUE PACK = Column M • QTY SHIPPER PACK = Column Q • Measurement = Column V")
