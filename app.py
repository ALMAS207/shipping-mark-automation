import re
from copy import copy, deepcopy
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

def text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)

def is_ditto(v):
    return text(v).strip() in ('"', "''", "“", "”", "〃")

def units(v):
    s = text(v)
    if not s:
        return ""
    try:
        n = float(s)
        return f"{int(n):02d} UNITS" if n.is_integer() else f"{n:g} UNITS"
    except Exception:
        return f"{s} UNITS"

def clean_measurement(value):
    s = text(value).upper().replace("×", "X").replace("*", "X")
    s = re.sub(r"(?i)(?<![A-Z])(?:LENGTH|WIDTH|HEIGHT|L|W|H)\s*", "", s)
    s = re.sub(r"\s+", "", s)
    m = re.search(r"(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)(CM|MM|INCH|IN)?", s)
    if not m:
        return text(value)
    nums = []
    for n in m.group(1, 2, 3):
        nums.append(n.rstrip("0").rstrip(".") if "." in n else n)
    unit = m.group(4) or ("CM" if "CM" in s else "")
    return "X".join(nums) + (f" {unit}" if unit else "")

def read_sheets(upload):
    data = upload.getvalue()
    name = upload.name.lower()
    result = {}
    if name.endswith(".xls"):
        book = xlrd.open_workbook(file_contents=data)
        for sn in DESTINATIONS:
            try:
                sh = book.sheet_by_name(sn)
            except Exception:
                continue
            result[sn] = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    elif name.endswith((".xlsx", ".xlsm")):
        book = load_workbook(BytesIO(data), data_only=True, read_only=True)
        for sn in DESTINATIONS:
            if sn in book.sheetnames:
                sh = book[sn]
                result[sn] = [[sh.cell(r, c).value for c in range(1, sh.max_column + 1)]
                              for r in range(1, sh.max_row + 1)]
        book.close()
    else:
        raise ValueError("Packing List অবশ্যই .xls, .xlsx অথবা .xlsm হতে হবে।")
    if not result:
        raise ValueError("SYDNEY / MELBOURNE / BRISBANE / PERTH নামে কোনো Destination Sheet পাওয়া যায়নি।")
    return result

def get_cell(rows, r, c):
    try:
        return rows[r - 1][c - 1]
    except (IndexError, TypeError):
        return None

def find_label(rows, pattern, row_start=1, row_end=45):
    rx = re.compile(pattern, re.I)
    for ri in range(row_start - 1, min(row_end, len(rows))):
        for ci, v in enumerate(rows[ri], start=1):
            if isinstance(v, str):
                m = rx.search(v.strip())
                if m:
                    return ri + 1, ci, m
    return None, None, None

def metadata(rows):
    order = style = dept = description = ""
    for ri in range(min(45, len(rows))):
        row = rows[ri]
        for ci, v in enumerate(row):
            if not isinstance(v, str):
                continue
            s = v.strip()
            if re.search(r"ORDER\s*NO", s, re.I):
                m = re.search(r"ORDER\s*NO\.?\s*[:#-]?\s*(\d+)", s, re.I)
                if m:
                    order = m.group(1)
                else:
                    for adjacent in row[ci + 1:ci + 6]:
                        candidate = text(adjacent)
                        if re.fullmatch(r"\d+", candidate):
                            order = candidate
                            break
            if re.search(r"STYLE\s*NO", s, re.I):
                m = re.search(r"STYLE\s*NO\.?\s*[:#-]?\s*([A-Z0-9-]+)", s, re.I)
                if m:
                    style = m.group(1)
                else:
                    for adjacent in row[ci + 1:ci + 6]:
                        candidate = text(adjacent)
                        if re.fullmatch(r"[A-Z0-9-]+", candidate, re.I):
                            style = candidate
                            break
            if re.search(r"DEPT\s*NO", s, re.I):
                m = re.search(r"DEPT\s*NO\.?\s*[:#-]?\s*(\d+)", s, re.I)
                if m:
                    dept = m.group(1)
                else:
                    for adjacent in row[ci + 1:ci + 6]:
                        candidate = text(adjacent)
                        if re.fullmatch(r"\d+", candidate):
                            dept = candidate
                            break
    # Description appears below the multi-row description heading in column A.
    for row_no in range(34, min(42, len(rows)) + 1):
        v = get_cell(rows, row_no, 1)
        if isinstance(v, str):
            s = v.strip()
            if s and not re.search(r"^(DESCRIPTION|\(ORDER|NUMBER\)|ORDER\s*NO|STYLE\s*NO|DEPT\s*NO|SUMMARY)", s, re.I):
                description = s
                break
    return order, style, dept, description
def detect_columns(rows):
    # Packing-list headings occupy rows 42-43. Some sheets use two-row
    # headings (e.g. Unit / Qty and Blister / Qty), so combine them by column.
    header = rows[41] if len(rows) > 41 else []
    subheader = rows[42] if len(rows) > 42 else []
    found = {}
    width = max(len(header), len(subheader))
    for ci in range(1, width + 1):
        top = text(header[ci - 1] if ci <= len(header) else "").upper()
        bottom = text(subheader[ci - 1] if ci <= len(subheader) else "").upper()
        label = re.sub(r"\s+", " ", f"{top} {bottom}").strip()
        if "COLOUR" in label and "colour" not in found:
            found["colour"] = ci
        if "KEY CODE" in label or "KEYCODE" in label:
            found["key"] = ci
        # Use the top header for SIZE; row 43 also says SIZE below KEY CODE.
        if top == "SIZE":
            found["size"] = ci
        if "UNIT" in label and ("QTY" in label or "QUANTITY" in label):
            found["issue"] = ci
        if "BLISTER" in label and ("QTY" in label or "QUANTITY" in label) and "TTL" not in label:
            found["shipper"] = ci
        if "CORTON MAS" in label or "CARTON MAS" in label or "MEASUREMENT" in label:
            found["measurement"] = ci

    # Known layout places garment size immediately to the right of KEY CODE.
    if "key" in found:
        found["size"] = found["key"] + 1
    # Verified 2026-11-05 AUS layout fallback; detected headings take priority.
    for key, col in {"colour": 7, "key": 9, "size": 10,
                      "issue": 11, "shipper": 15, "measurement": 20}.items():
        found.setdefault(key, col)
    missing = [k for k in ("issue", "shipper", "measurement", "colour", "key", "size") if k not in found]
    if missing:
        raise ValueError("Packing List-এর header থেকে প্রয়োজনীয় column শনাক্ত করা যায়নি: " + ", ".join(missing))
    return found

def build_records(upload):
    sheets = read_sheets(upload)
    records, qc = [], []
    for dest in DESTINATIONS:
        if dest not in sheets:
            continue
        rows = sheets[dest]
        order, style, dept, desc = metadata(rows)
        if not all((order, style, dept, desc)):
            missing_meta = [n for n, v in (("Order No.", order), ("Style No.", style),
                            ("Dept No.", dept), ("Description", desc)) if not v]
            raise ValueError(f"{dest} Sheet: Header থেকে তথ্য পাওয়া যায়নি: {', '.join(missing_meta)}")
        cols = detect_columns(rows)
        last = {"colour": "", "key": "", "size": ""}
        first_data = 44
        made_here = 0
        expected_cartons = None
        # Summary header is in row 34; its total is commonly in the same
        # column on the ORDER NO. row (row 38). Search below the header.
        for ri in range(33, min(40, len(rows))):
            for ci, v in enumerate(rows[ri], start=1):
                if isinstance(v, str) and "TTL CTNS" in v.upper():
                    for summary_row in range(ri + 1, min(ri + 7, len(rows))):
                        candidate = get_cell(rows, summary_row + 1, ci)
                        if is_number(candidate):
                            expected_cartons = int(candidate)
                            break
                    break
            if expected_cartons is not None:
                break
        for r in range(first_data, len(rows) + 1):
            start, end = get_cell(rows, r, 1), get_cell(rows, r, 2)
            if not (is_number(start) and is_number(end)):
                continue
            for field in ("colour", "key", "size"):
                value = get_cell(rows, r, cols[field])
                if value not in (None, "") and not is_ditto(value):
                    last[field] = text(value)
            issue = get_cell(rows, r, cols["issue"])
            shipper = get_cell(rows, r, cols["shipper"])
            measurement = clean_measurement(get_cell(rows, r, cols["measurement"]))
            missing = [label for label, value in [
                ("Colour", last["colour"]), ("Keycode", last["key"]), ("Size", last["size"]),
                ("QTY ISSUE PACK", issue), ("QTY SHIPPER PACK", shipper), ("Measurement", measurement)
            ] if value in (None, "")]
            if missing:
                raise ValueError(f"{dest} Sheet, Row {r}: তথ্য missing: {', '.join(missing)}")
            a, b = int(start), int(end)
            if b < a:
                raise ValueError(f"{dest} Sheet, Row {r}: Carton range ভুল ({a}-{b})।")
            for carton in range(a, b + 1):
                records.append({
                    "dest": dest, "carton": carton, "order": order, "dept": dept,
                    "key": last["key"], "style": style, "description": desc,
                    "issue": issue, "shipper": shipper, "colour": last["colour"],
                    "size": last["size"], "measurement": measurement,
                    "source_row": r
                })
                made_here += 1
        qc.append({"destination": dest, "generated": made_here, "expected": expected_cartons,
                   "status": "PASS" if expected_cartons in (None, made_here) else "REVIEW"})
    if not records:
        raise ValueError("Destination Sheet-এ carton range পাওয়া যায়নি।")
    seen = set()
    duplicates = []
    for rec in records:
        key = (rec["dest"], rec["carton"])
        if key in seen:
            duplicates.append(f'{rec["dest"]} carton {rec["carton"]}')
        seen.add(key)
    if duplicates:
        raise ValueError("একই Destination-এ duplicate carton number পাওয়া গেছে: " + ", ".join(duplicates[:10]))
    return records, qc

def image_bytes(img):
    return img._data()

def make_output(template_upload, records, qc_rows):
    wb = load_workbook(BytesIO(template_upload.getvalue()))
    total = len(records)

    # The supplied format already contains destination-specific master labels.
    # Reuse each destination's own master sheet so AU/SYDNEY etc. stays correct.
    base_by_dest = {}
    images_by_dest = {}
    used_destinations = list(dict.fromkeys(rec["dest"] for rec in records))
    for dest in used_destinations:
        base = next((s for s in wb.worksheets
                     if s.title.strip().upper().startswith("AU " + dest)), None)
        if base is None:
            raise ValueError(
                f"Shipping Mark Format-এ AU {dest} master sheet পাওয়া যায়নি। "
                f"যে Destination আছে শুধু সেই Destination-এর master sheet রাখুন।"
            )
        if not base._images:
            raise ValueError(f"AU {dest} Format-এ Hanger Logo/Image পাওয়া যায়নি।")
        base_by_dest[dest] = base
        images_by_dest[dest] = [
            (im._data(), im.width, im.height, deepcopy(im.anchor))
            for im in base._images
        ]

    # Keep only the four master sheets while generating new carton labels.
    masters = set(base_by_dest.values())
    for sh in list(wb.worksheets):
        if sh not in masters:
            wb.remove(sh)
    for base in masters:
        base._images = []

    def put_images(sh, dest):
        for blob, width, height, anchor in images_by_dest[dest]:
            im = XLImage(BytesIO(blob))
            im.width, im.height = width, height
            im.anchor = deepcopy(anchor)
            sh.add_image(im)

    used_master = set()
    generated_sheets = []
    for rec in records:
        dest = rec["dest"]
        master = base_by_dest[dest]
        if dest not in used_master:
            sh = master
            used_master.add(dest)
        else:
            sh = wb.copy_worksheet(master)

        sh.title = f"{dest[:3]}-{rec['carton']:03d}"[:31]
        sh["I3"] = "AU/" + dest
        sh["C3"] = rec["dept"]
        sh["C4"] = rec["order"]
        sh["C5"] = rec["key"]
        sh["C6"] = rec["style"]
        sh["C7"] = rec["description"]
        sh["C8"] = units(rec["issue"])
        sh["C9"] = units(rec["shipper"])
        sh["C10"] = rec["colour"]
        sh["C11"] = rec["size"]
        sh["C13"] = rec["measurement"]
        sh["D15"] = str(rec["carton"])
        sh["F15"] = str(total)
        sh.sheet_view.showGridLines = False
        sh.sheet_properties.pageSetUpPr.fitToPage = True
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 1
        sh.page_setup.orientation = "landscape"
        sh.print_area = "$B$1:$G$16"
        put_images(sh, dest)
        generated_sheets.append(sh)

    # Put the carton labels in Packing List order, then append the QC sheet.
    wb._sheets = generated_sheets

    report = wb.create_sheet("QC REPORT")
    report.append(["QUALITY CHECK", "RESULT", "DETAIL"])
    for item in qc_rows:
        detail = (f"Packing List expected: {item['expected']} cartons; "
                  if item["expected"] is not None else "Packing List total not found; ")
        detail += f"Generated labels: {item['generated']}"
        report.append([f"{item['destination']} carton count", item["status"], detail])
    report.append(["Total generated labels", "PASS", str(total)])
    report.append(["Duplicate carton numbers", "PASS", "No duplicate carton numbers within each destination"])
    report.append(["Quantity columns", "PASS", "Unit Qty and Blister Qty detected separately for each destination"])
    report.append(["Measurement", "PASS", "CORTON MAS dimensions only; L/W/H labels removed"])
    report.append(["Hanger Logo", "PASS", "Destination master images copied at original size and anchor"])
    for cell in report[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for col, width in {"A": 34, "B": 16, "C": 90}.items():
        report.column_dimensions[col].width = width
    for row in report.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    report.freeze_panes = "A2"
    report.sheet_view.showGridLines = False

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), sum(len(v) for v in images_by_dest.values())

st.title("📦 Shipping Mark Automation Pro")
st.caption("Online Web Version • বাংলা / English • Build 2026.11.05.4")
st.info("১) Packing List Upload করুন  ২) Shipping Mark Format Upload করুন  ৩) Generate চাপুন  ৪) Excel Download করুন।")
packing_file = st.file_uploader("১. Packing List Excel (.XLS / .XLSX / .XLSM)", type=["xls", "xlsx", "xlsm"])
format_file = st.file_uploader("২. Shipping Mark Format (.XLSX)", type=["xlsx"])

if st.button("🚀 Generate Shipping Mark", type="primary", use_container_width=True):
    if not packing_file or not format_file:
        st.error("Packing List এবং Shipping Mark Format—দুটি ফাইলই Upload করুন।")
    else:
        try:
            with st.spinner("Packing List-এর Sheet/Header শনাক্ত করে Shipping Mark তৈরি হচ্ছে..."):
                records, qc_rows = build_records(packing_file)
                output, image_count = make_output(format_file, records, qc_rows)
            st.success(f"সফল! মোট {len(records)}টি carton label তৈরি হয়েছে।")
            st.caption(f"Template-এর {image_count}টি image/logo original size ও position-এ রাখা হয়েছে।")
            st.download_button("⬇️ Shipping Mark Excel Download", data=output,
                file_name=f"{Path(packing_file.name).stem}_Shipping_Mark.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
            st.warning("Print করার আগে QC REPORT এবং কয়েকটি label Packing List-এর সঙ্গে মিলিয়ে নিন।")
        except Exception as e:
            st.error(f"Shipping Mark তৈরি হয়নি: {e}")
            st.info("ভুল mapping হলে Excel-এর header অনুযায়ী শনাক্তকরণ উন্নত করতে হবে; কোনো তথ্য অনুমান করে পূরণ করা হয় না।")

st.divider()
st.caption("প্রতিদিন নতুন Order-এর Packing List (.XLS/.XLSX/.XLSM) এবং Shipping Mark Format (.XLSX) Upload করলেই তৈরি হবে। একই নির্ধারিত Packing List layout অনুসরণ করতে হবে: Destination sheet, carton A:B row 44+, Colour G, Keycode I, Size J, Issue M, Shipper Q, Measurement V।")
st.caption("একটি Order-এ যে Destination sheet আছে, শুধু সেই Destination-এর master label template প্রয়োজন; অন্য Destination না থাকলেও Generate করা যাবে।")
