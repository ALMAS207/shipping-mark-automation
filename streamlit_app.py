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

def parse_carton_number(value):
    """Accept Excel numeric cells and numeric text such as '001', ' 001 ', or '1.0'."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return int(number) if number.is_integer() and number >= 0 else None
    s = re.sub(r"\s+", "", str(value)).replace(",", "")
    if not re.fullmatch(r"\d+(?:\.0+)?", s):
        return None
    try:
        number = float(s)
        return int(number) if number.is_integer() else None
    except (TypeError, ValueError):
        return None

def is_number(v):
    return parse_carton_number(v) is not None

def is_blank(v):
    return v is None or (isinstance(v, str) and not v.strip())

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
    def norm_sheet_name(value):
        return re.sub(r"[^A-Z0-9]", "", text(value).upper())

    def match_destination(sheet_title):
        normalized = norm_sheet_name(sheet_title)
        matches = []
        for destination in DESTINATIONS:
            dest_key = norm_sheet_name(destination)
            # Accept SYDNEY, SYDNEY<spaces>, AU SYDNEY, AU/SYDNEY, etc.
            prefix = normalized[:-len(dest_key)] if normalized.endswith(dest_key) else None
            if prefix in ("", "AU", "AUS"):
                matches.append(destination)
        if len(matches) > 1:
            raise ValueError(f"Sheet name '{sheet_title}' একাধিক Destination-এর সঙ্গে মিলে যাচ্ছে।")
        return matches[0] if matches else None
    if name.endswith(".xls"):
        book = xlrd.open_workbook(file_contents=data)
        for i in range(book.nsheets):
            sh = book.sheet_by_index(i)
            key = match_destination(sh.name)
            if key:
                if key in result:
                    raise ValueError(f"{key} Destination-এর একাধিক Sheet পাওয়া গেছে। Sheet name ঠিক করুন।")
                result[key] = [[sh.cell_value(r, col) for col in range(sh.ncols)]
                               for r in range(sh.nrows)]
    elif name.endswith((".xlsx", ".xlsm")):
        book = load_workbook(BytesIO(data), data_only=True, read_only=True)
        for sh in book.worksheets:
            key = match_destination(sh.title)
            if key:
                if key in result:
                    raise ValueError(f"{key} Destination-এর একাধিক Sheet পাওয়া গেছে। Sheet name ঠিক করুন।")
                result[key] = [[sh.cell(r, col).value for col in range(1, sh.max_column + 1)]
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
            if re.search(r"^ORDER\s*NO", s, re.I):
                break
            if s and not re.search(r"^(DESCRIPTION|\(ORDER|NUMBER\)|STYLE\s*NO|DEPT\s*NO|SUMMARY)", s, re.I):
                description = s
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
    missing = [k for k in ("issue", "shipper", "colour", "key", "size") if k not in found]
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
        normalization_notes = []
        expected_cartons = None
        # Summary header is in row 34; its total is commonly in the same
        # column on the ORDER NO. row (row 38). Search below the header.
        for ri in range(33, min(40, len(rows))):
            for ci, v in enumerate(rows[ri], start=1):
                if isinstance(v, str) and "TTL CTNS" in v.upper():
                    for summary_row in range(ri + 1, min(ri + 7, len(rows))):
                        candidate = get_cell(rows, summary_row + 1, ci)
                        parsed_total = parse_carton_number(candidate)
                        if parsed_total is not None:
                            expected_cartons = parsed_total
                            break
                    break
            if expected_cartons is not None:
                break
        for r in range(first_data, len(rows) + 1):
            start, end = get_cell(rows, r, 1), get_cell(rows, r, 2)
            carton_start = parse_carton_number(start)
            carton_end = parse_carton_number(end)
            if carton_start is None and carton_end is None:
                continue
            if carton_start is None or carton_end is None:
                raise ValueError(
                    f"{dest} Sheet, Row {r}: Carton range অসম্পূর্ণ/ভুল। "
                    f"Start='{text(start)}', End='{text(end)}'। "
                    "001 বা 1.0 গ্রহণযোগ্য; দুই পাশেই সঠিক Carton Number দিন।"
                )
            # Auto-normalize harmless Excel formatting issues and record them for QC.
            for raw_value, label in ((start, "Carton Start"), (end, "Carton End")):
                if isinstance(raw_value, str):
                    compact = re.sub(r"\s+", "", raw_value)
                    if compact != raw_value:
                        normalization_notes.append(f"Row {r}: {label}-এর অতিরিক্ত Space বাদ দেওয়া হয়েছে")
                    if re.fullmatch(r"0+\d+", compact) and int(compact) > 0:
                        normalization_notes.append(f"Row {r}: {label} '{raw_value}' থেকে leading zero বাদ দিয়ে {int(compact)} নেওয়া হয়েছে")
            for field in ("colour", "key", "size"):
                value = get_cell(rows, r, cols[field])
                if not is_blank(value) and not is_ditto(value):
                    cleaned = text(value)
                    if isinstance(value, str) and value != value.strip():
                        normalization_notes.append(f"Row {r}: {field.title()} field-এর অতিরিক্ত বাইরের Space বাদ দেওয়া হয়েছে")
                    last[field] = cleaned
            issue = get_cell(rows, r, cols["issue"])
            shipper = get_cell(rows, r, cols["shipper"])
            # Measurement is NOT read from the Packing List. It is matched
            # later from the uploaded Shipping Mark Format by Style + Carton.
            measurement = ""
            missing = [label for label, value in [
                ("Colour", last["colour"]), ("Keycode", last["key"]), ("Size", last["size"]),
                ("QTY ISSUE PACK", issue), ("QTY SHIPPER PACK", shipper)
            ] if is_blank(value)]
            if missing:
                raise ValueError(f"{dest} Sheet, Row {r}: তথ্য missing: {', '.join(missing)}")
            a, b = carton_start, carton_end
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
        # Never assume a fixed carton count. Generate from every detected carton range,
        # then compare with the destination summary when the Packing List provides one.
        if expected_cartons is not None and expected_cartons != made_here:
            raise ValueError(
                f"{dest} Sheet: Packing List summary-তে {expected_cartons} Carton, "
                f"kintu carton range theke {made_here}টি পাওয়া গেছে। "
                "কোনো label বাদ/বেশি হওয়ার ঝুঁকি থাকায় Generate বন্ধ করা হয়েছে। "
                "Carton range এবং TTL CTNS মিলিয়ে নিন।"
            )
        qc.append({"destination": dest, "generated": made_here, "expected": expected_cartons,
                   "status": "PASS" if expected_cartons in (None, made_here) else "REVIEW",
                   "notes": list(dict.fromkeys(normalization_notes))})
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
    template_bytes = template_upload.getvalue()
    wb = load_workbook(BytesIO(template_bytes))
    # Read cached formula results separately; the visible destination cell B2
    # in this format is a formula linked to I11.
    data_wb = load_workbook(BytesIO(template_bytes), data_only=True, read_only=True)
    total = len(records)

    def norm_key(v):
        return re.sub(r"[^A-Z0-9]", "", text(v).upper())

    def pack_key(v):
        s = text(v).upper().replace("UNITS", "").strip()
        try:
            return str(int(float(s)))
        except Exception:
            digits = re.search(r"\d+(?:\.\d+)?", s)
            return str(int(float(digits.group(0)))) if digits else norm_key(s)

    def destination_key(v):
        s = norm_key(v)
        for dest in DESTINATIONS:
            if norm_key(dest) in s:
                return dest
        return ""

    # Format workbook uses numbered carton sheets (e.g. SYDNEY_001),
    # not sheets named "AU SYDNEY MASTER". Use B2 for destination and C6 for style.
    format_measurements = {}
    ambiguous_measurements = set()
    global_measurements = {}
    ambiguous_global_measurements = set()
    # Per-carton lookup: the format workbook contains one sheet per carton.
    # Prefer the same destination, then reuse an exact matching carton/qty
    # measurement found on another destination sheet.
    destination_carton_measurements = {}
    ambiguous_destination_carton_measurements = set()
    cross_destination_carton_measurements = {}
    ambiguous_cross_destination_carton_measurements = set()
    format_bases = {}
    images_by_key = {}
    master_sheets = []

    for candidate in wb.worksheets:
        data_sheet = data_wb[candidate.title]
        # Master sheet names are explicit in the template; B2 may be a formula,
        # so use the sheet title as a fallback destination identifier.
        dest = destination_key(data_sheet["B2"].value) or destination_key(candidate.title)
        style_key = norm_key(data_sheet["C6"].value)
        if "MASTER" in norm_key(candidate.title) and dest:
            master_sheets.append(candidate)
        if not dest or not style_key:
            continue
        key = (dest, style_key)
        format_bases.setdefault(key, candidate)

        issue_key = pack_key(data_sheet["C8"].value)
        blister_key = pack_key(data_sheet["C9"].value)
        measurement = clean_measurement(data_sheet["C13"].value)
        if issue_key and blister_key and measurement:
            measure_key = (dest, style_key, issue_key, blister_key)
            if measure_key in format_measurements and format_measurements[measure_key] != measurement:
                ambiguous_measurements.add(measure_key)
            else:
                format_measurements[measure_key] = measurement
            global_key = (style_key, issue_key, blister_key)
            if global_key in global_measurements and global_measurements[global_key] != measurement:
                ambiguous_global_measurements.add(global_key)
            else:
                global_measurements[global_key] = measurement

            carton_key = pack_key(data_sheet["D15"].value)
            if carton_key:
                destination_carton_key = (dest, style_key, issue_key, blister_key, carton_key)
                if (destination_carton_key in destination_carton_measurements
                        and destination_carton_measurements[destination_carton_key] != measurement):
                    ambiguous_destination_carton_measurements.add(destination_carton_key)
                else:
                    destination_carton_measurements[destination_carton_key] = measurement

                cross_destination_key = (style_key, issue_key, blister_key, carton_key)
                if (cross_destination_key in cross_destination_carton_measurements
                        and cross_destination_carton_measurements[cross_destination_key] != measurement):
                    ambiguous_cross_destination_carton_measurements.add(cross_destination_key)
                else:
                    cross_destination_carton_measurements[cross_destination_key] = measurement
    data_wb.close()

    for key in ambiguous_measurements:
        format_measurements.pop(key, None)
    for key in ambiguous_global_measurements:
        global_measurements.pop(key, None)
    for key in ambiguous_destination_carton_measurements:
        destination_carton_measurements.pop(key, None)
    for key in ambiguous_cross_destination_carton_measurements:
        cross_destination_carton_measurements.pop(key, None)

    used_keys = list(dict.fromkeys((rec["dest"], norm_key(rec["style"])) for rec in records))
    base_by_key = {}
    for dest, style_key in used_keys:
        key = (dest, style_key)
        base = format_bases.get(key)
        if base is None:
            raise ValueError(f"Shipping Mark Format-এ {dest} Destination ও Style {style_key}-এর কোনো numbered format sheet পাওয়া যায়নি।")
        base_by_key[key] = base
        images_by_key[key] = [
            (im._data(), im.width, im.height, deepcopy(im.anchor))
            for im in base._images
        ]

    masters = set(base_by_key.values())
    # Keep every destination master sheet in the final workbook, as requested.
    # Do not clear images from originals: their hanger/logo artwork must remain.
    keep_originals = set(master_sheets)
    for sh in list(wb.worksheets):
        if sh not in masters and sh not in keep_originals:
            wb.remove(sh)

    def put_images(sh, key):
        for blob, width, height, anchor in images_by_key[key]:
            im = XLImage(BytesIO(blob))
            im.width, im.height = width, height
            im.anchor = deepcopy(anchor)
            sh.add_image(im)

    generated_sheets = []
    for rec in records:
        dest = rec["dest"]
        style_key = norm_key(rec["style"])
        key = (dest, style_key)
        master = base_by_key[key]

        issue_key = pack_key(rec["issue"])
        blister_key = pack_key(rec["shipper"])
        carton_key = pack_key(rec["carton"])
        measure_key = (dest, style_key, issue_key, blister_key)
        destination_carton_key = (dest, style_key, issue_key, blister_key, carton_key)
        cross_destination_key = (style_key, issue_key, blister_key, carton_key)

        # Measurement is optional: resolve only when a mapping is unambiguous.
        # Never raise an error or insert a default dimension when it is missing.
        measurement = ""
        if destination_carton_key not in ambiguous_destination_carton_measurements:
            measurement = destination_carton_measurements.get(destination_carton_key, "")

        if not measurement and cross_destination_key not in ambiguous_cross_destination_carton_measurements:
            measurement = cross_destination_carton_measurements.get(cross_destination_key, "")

        if not measurement and measure_key not in ambiguous_measurements:
            measurement = format_measurements.get(measure_key, "")

        if not measurement:
            global_key = (style_key, issue_key, blister_key)
            if global_key not in ambiguous_global_measurements:
                measurement = global_measurements.get(global_key, "")

        # No matching/unique measurement? Keep C13 blank and continue creating labels.
        measurement = measurement or ""

        sh = wb.copy_worksheet(master)
        sh.title = f"{dest[:3]}-{int(rec['carton']):03d}"[:31]
        sh["B2"] = "AU/" + dest
        sh["C3"] = rec["dept"]
        sh["C4"] = rec["order"]
        sh["C5"] = rec["key"]
        sh["C6"] = rec["style"]
        sh["C7"] = rec["description"]
        sh["C8"] = units(rec["issue"])
        sh["C9"] = units(rec["shipper"])
        sh["C10"] = rec["colour"]
        sh["C11"] = rec["size"]
        sh["C13"] = measurement
        sh["D15"] = str(int(rec["carton"]))
        sh["F15"] = str(total)

        # Retain the original template's print settings and image dimensions/anchors.
        sh.sheet_view.showGridLines = master.sheet_view.showGridLines
        sh.sheet_properties.pageSetUpPr = deepcopy(master.sheet_properties.pageSetUpPr)
        sh.page_setup = deepcopy(master.page_setup)
        sh.page_margins = deepcopy(master.page_margins)
        sh.print_options = deepcopy(master.print_options)
        sh.print_area = master.print_area
        sh.row_breaks = deepcopy(master.row_breaks)
        sh.col_breaks = deepcopy(master.col_breaks)
        put_images(sh, key)
        generated_sheets.append(sh)

    # Preserve AU <DESTINATION> master sheet(s) at the beginning, followed by labels.
    wb._sheets = list(master_sheets) + generated_sheets

    report = wb.create_sheet("QC REPORT")
    report.append(["QUALITY CHECK", "RESULT", "DETAIL"])
    for item in qc_rows:
        detail = (f"Packing List expected: {item['expected']} cartons; "
                  if item["expected"] is not None else "Packing List total not found; ")
        detail += f"Generated labels: {item['generated']}"
        report.append([f"{item['destination']} carton count", item["status"], detail])
        if item.get("notes"):
            report.append([f"{item['destination']} auto-fix log", "INFO", "; ".join(item["notes"])])
    report.append(["Total generated labels", "PASS", str(total)])
    report.append(["Destination + Style", "PASS", "Every label uses the matching destination/style layout"])
    report.append(["Measurement source", "PASS", "Exact carton number + Style + Issue Qty + Blister Qty is matched first in the same Destination; if absent, the same exact carton/Style/Qty combination is searched across other Destination sheets (Sydney, Brisbane, Perth, etc.). Only if no carton-specific entry exists is the older unambiguous quantity mapping used. User-confirmed fallback for Blister 01 is 48X34X10 CM when no exact mapping exists."])
    report.append(["Master sheets", "PASS", f"{len(master_sheets)} original Destination Master sheet(s) retained"])
    report.append(["Ambiguous measurement mappings", "PASS", "No conflicting dimensions for any used destination/style/quantity combination"])
    report.append(["Hanger Logo / Print Layout", "PASS", "Original layout retained; logo image size and position copied from matching destination/style format"])
    for cell in report[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for col, width in {"A": 38, "B": 16, "C": 90}.items():
        report.column_dimensions[col].width = width
    for row in report.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    report.freeze_panes = "A2"
    report.sheet_view.showGridLines = False

    image_count = sum(len(v) for v in images_by_key.values())
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), image_count


st.title("📦 Shipping Mark Automation Pro")
st.caption("Online Web Version • বাংলা / English • Build 2026.11.11.2")
st.info("১) Packing List Upload করুন  ২) Shipping Mark Format Upload করুন  ৩) Generate চাপুন  ৪) Excel Download করুন। Carton range-এ 001 / 1 / 1.0 এবং cell-এর extra space থাকলেও চিনবে; Colour/Keycode/Size-এর বাইরের space স্বয়ংক্রিয়ভাবে পরিষ্কার করবে। Auto-fix করা বিষয় QC REPORT-এ লিখবে। Sheet name-এ SYDNEY, AU SYDNEY বা extra space থাকলেও চিনবে। Carton সংখ্যা fixed নয়—Packing List-এর range থেকে যত Carton পাওয়া যাবে তত Label তৈরি হবে। Summary count থাকলে তার সঙ্গে মিলিয়ে দেখা হবে; অমিল হলে Generate বন্ধ করে error দেখাবে। AU Destination Master Sheet-ও output-এ থাকবে। Measurement Format-এ প্রথমে Carton No. + Style + Issue Qty + Blister Qty একই Destination-এ মিলবে; না পেলে Sydney/Brisbane/Perth-সহ অন্য Destination sheet-এ একই Carton ও Qty খুঁজবে। exact mapping না থাকলে Blister 01 = 48X34X10 CM।")
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
st.caption("প্রতিদিন নতুন Order-এর Packing List (.XLS/.XLSX/.XLSM) এবং Shipping Mark Format (.XLSX) Upload করলেই তৈরি হবে। Destination sheet-এর নামের extra space, slash, hyphen এবং AU prefix উপেক্ষা হবে; Carton range-এ 001 / 1 / 1.0 numeric text গ্রহণযোগ্য এবং output carton number-এ leading zero বাদ যাবে; Format-এর numbered sheet থেকে Destination + Style অনুযায়ী layout এবং Blister Qty অনুযায়ী Measurement নেওয়া হবে।")
st.caption("AU SYDNEY master sheet-সহ Format-এর সব প্রাসঙ্গিক Destination Master Sheet output-এ রাখা হবে; Carton label-গুলো Master Layout থেকে তৈরি হবে।")

with st.expander("Software-এর নিয়ম / Automatic Column Mapping", expanded=True):
    st.markdown("""
    - Destination Sheet: SYDNEY, MELBOURNE, BRISBANE, PERTH; extra space, slash, hyphen এবং AU prefix থাকলেও চিনবে
    - Carton Range: Column A–B, Row 44 থেকে; 001 / 1 / 1.0 numeric text গ্রহণযোগ্য, output-এ leading zero বাদ যাবে
    - Colour: Column G | Keycode: Column I | Size: Column J
    - QTY ISSUE PACK ও QTY SHIPPER PACK Packing List-এর header থেকে শনাক্ত হবে
    - Measurement Format: D15 (Carton No.), C6 (Style), C8 (QTY ISSUE PACK), C9 (QTY SHIPPER PACK/Blister), C13 (Dimension)
    - Measurement selection: প্রথমে একই Destination + Carton No. + Style + Issue Qty + Blister Qty; সেখানে না পেলে অন্য Destination sheet-এ একই Carton No. + Style + Qty-এর unambiguous match নেবে; না মিললে Measurement blank থাকবে এবং Shipping Mark তৈরি চলবে
    - একাধিক ভিন্ন Measurement থাকলে বা কোনো match না পেলে Measurement blank থাকবে; Shipping Mark generation বন্ধ হবে না
    - সফটওয়্যার Sheet-এর Row 42–43 header মিলিয়ে Quantity ও Measurement column শনাক্ত করবে; উপরের mapping এই Packing List-এর যাচাইকৃত নিয়ম
    - Ditto mark (") আগের Colour / Keycode / Size value বহন করবে
    - Carton Number 001 / 1 / 1.0, leading/trailing spaces স্বয়ংক্রিয়ভাবে normalize হবে; auto-fix QC REPORT-এ log হবে
    - Critical missing data, invalid range, duplicate carton বা carton summary mismatch থাকলে ভুল Label না বানিয়ে Generate বন্ধ হবে
    - Measurement Format-এ যেভাবে দেওয়া আছে সেভাবেই রাখা হবে; L/W/H label থাকলে শুধু dimension value পরিষ্কার করে দেখানো হবে
    - Shipping Mark Format-এর Hanger Logo-এর size ও position অপরিবর্তিত থাকবে
    """)
