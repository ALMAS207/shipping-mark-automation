import re
from copy import deepcopy
from io import BytesIO
from pathlib import Path
import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment

DESTINATIONS = {"SYDNEY":"AU/SYDNEY","MELBOURNE":"AU/MELBOURNE","BRISBANE":"AU/BRISBANE","PERTH":"AU/PERTH"}

def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)

def as_text(v):
    if v is None: return ""
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v).strip()

def units(v):
    if v in (None, ""): return ""
    try: return f"{int(float(v)):02d} UNITS"
    except Exception: return f"{as_text(v)} UNITS"

def clean_measurement(v):
    s = as_text(v).upper().replace(" ", "")
    s = re.sub(r"([0-9])X[A-Z]([0-9])", r"\1X\2", s)
    m = re.search(r"(\d+(?:\.\d+)?X\d+(?:\.\d+)?X\d+(?:\.\d+)?)(CM|MM|IN)?", s)
    if m: return f"{m.group(1)} {m.group(2) or 'CM'}"
    return as_text(v)

def read_packing_list(uploaded):
    ext = Path(uploaded.name).suffix.lower()
    data, sheets = uploaded.getvalue(), {}
    if ext == ".xls":
        import xlrd
        book = xlrd.open_workbook(file_contents=data)
        for name in DESTINATIONS:
            try: sh = book.sheet_by_name(name)
            except Exception: continue
            sheets[name] = [[sh.cell_value(r,c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    elif ext in (".xlsx", ".xlsm"):
        book = load_workbook(BytesIO(data), data_only=True, read_only=True)
        for name in DESTINATIONS:
            if name in book.sheetnames:
                sh = book[name]
                sheets[name] = [[sh.cell(r,c).value for c in range(1,sh.max_column+1)] for r in range(1,sh.max_row+1)]
        book.close()
    else: raise ValueError("Packing List .xls, .xlsx অথবা .xlsm হতে হবে।")
    return sheets

def cell(rows, r, c):
    try: return rows[r-1][c-1]
    except Exception: return None

def extract_metadata(rows):
    order, style, dept, desc = "", "", "034", "5 PK TUNNEL TRUNK"
    for r in range(34,42):
        for v in (rows[r-1] if len(rows)>=r else []):
            if not isinstance(v,str): continue
            s=v.strip()
            m=re.search(r"ORDER\s*NO\.?\s*(\d+)",s,re.I)
            if m: order=m.group(1)
            m=re.search(r"STYLE\s*NO\.?\s*([A-Z0-9]+)",s,re.I)
            if m: style=m.group(1)
            m=re.search(r"DEPT\s*NO\.?\s*:\s*(\d+)",s,re.I)
            if m: dept=m.group(1)
            if "TUNNEL" in s.upper(): desc=s
    return order,style,dept,desc

def build_records(uploaded):
    sheet_rows=read_packing_list(uploaded)
    records,totals,warnings=[],{},[]
    for dest in DESTINATIONS:
        if dest not in sheet_rows: continue
        rows=sheet_rows[dest]
        order,style,dept,desc=extract_metadata(rows)
        last_colour=last_key=last_size=None
        dest_count=dest_qty=0
        for r in range(44,len(rows)+1):
            start,end=cell(rows,r,1),cell(rows,r,2)
            if not (is_num(start) and is_num(end)): continue
            colour,key,size=cell(rows,r,7),cell(rows,r,9),cell(rows,r,10)
            if colour not in (None,"",'"',"''"): last_colour=as_text(colour)
            if key not in (None,"",'"',"''"): last_key=as_text(key)
            if size not in (None,"",'"',"''"): last_size=as_text(size)
            issue,shipper=cell(rows,r,13),cell(rows,r,17)
            measurement=clean_measurement(cell(rows,r,22))
            if not all([order,style,desc,last_colour,last_key,last_size,measurement]):
                raise ValueError(f"{dest} sheet, row {r}: Order/Style/Description/Colour/Keycode/Size/Measurement data missing.")
            a,b=int(start),int(end)
            if b<a: raise ValueError(f"{dest} sheet, row {r}: Carton range ভুল ({a}-{b})।")
            if issue in (None,"") or shipper in (None,""): warnings.append(f"{dest} row {r}: Issue/shipper pack quantity খালি।")
            for carton in range(a,b+1):
                records.append({"dest":dest,"carton":carton,"order":order,"dept":dept,"key":last_key,"style":style,
                    "description":desc,"issue":issue,"shipper_pack":shipper,"colour":last_colour,"size":last_size,
                    "measurement":measurement})
            dest_count+=int(cell(rows,r,19) or (b-a+1))
            dest_qty+=int(cell(rows,r,20) or 0)
        if dest_count or dest_qty: totals[dest]=(dest_count,dest_qty)
    if not records: raise ValueError("SYDNEY/MELBOURNE/BRISBANE/PERTH sheet-এ carton data পাওয়া যায়নি।")
    cartons=[x["carton"] for x in records]
    if len(set(cartons))!=len(cartons): raise ValueError("একই carton number একাধিকবার পাওয়া গেছে। Packing List পরীক্ষা করুন।")
    return records,totals,warnings

def make_output(template_upload,records,totals):
    wb=load_workbook(BytesIO(template_upload.getvalue()))
    base=None
    for sh in wb.worksheets:
        if str(sh["E16"].value).strip().zfill(2)=="01": base=sh; break
    if base is None: raise ValueError("Template-এ E16 cell-এ carton 01 থাকা label sheet পাওয়া যায়নি।")
    if not base._images: raise ValueError("Template-এর label sheet-এ Hanger Logo পাওয়া যায়নি।")
    source_img=base._images[0]
    anchor,width,height=deepcopy(source_img.anchor),source_img.width,source_img.height
    logo_bytes=source_img._data()
    for sh in list(wb.worksheets):
        if sh is not base: wb.remove(sh)
    base._images=[]
    total=len(records)
    logo_streams=[]
    for i,rec in enumerate(records):
        if i==0:
            sh=base
            sh.title=f"{rec['dest'][:3]} {rec['carton']:03d}"
        else:
            sh=wb.copy_worksheet(base)
            sh.title=f"{rec['dest'][:3]} {rec['carton']:03d}"
        sh["D4"]=rec["dept"]; sh["D5"]=rec["order"]; sh["D6"]=rec["key"]; sh["D7"]=rec["style"]
        sh["D8"]=rec["description"]; sh["D9"]=units(rec["issue"]); sh["D10"]=units(rec["shipper_pack"])
        sh["D11"]=rec["colour"]; sh["D12"]=rec["size"]; sh["D13"]="KGS"; sh["D14"]=rec["measurement"]
        sh["E16"]=str(rec["carton"]).lstrip("0") or "0"; sh["G16"]=f"      {total}"; sh["P12"]=DESTINATIONS[rec["dest"]]
        sh.sheet_view.showGridLines=False; sh.print_area="$C$2:$G$16"
        sh.sheet_properties.pageSetUpPr.fitToPage=True; sh.page_setup.fitToWidth=1
        sh.page_setup.fitToHeight=1; sh.page_setup.orientation="landscape"
        stream=BytesIO(logo_bytes); logo_streams.append(stream)
        img=XLImage(stream); img.width,img.height=width,height; img.anchor=deepcopy(anchor); sh.add_image(img)
    qc=wb.create_sheet("QC REPORT")
    qc.append(["QUALITY CHECK","RESULT","DETAIL"])
    for dest,(cartons,qty) in totals.items():
        made=sum(1 for x in records if x["dest"]==dest)
        qc.append([f"{dest} cartons / quantity","PASS" if made==cartons else "REVIEW",
                   f"Packing List: {cartons} cartons / {qty} units; generated labels: {made}"])
    qc.append(["Total labels","PASS",str(total)])
    qc.append(["Duplicate carton numbers","PASS","No duplicate carton numbers"])
    qc.append(["QTY ISSUE PACK","CHECK","Packing List column M"])
    qc.append(["QTY SHIPPER PACK","CHECK","Packing List column Q"])
    qc.append(["Colour / Keycode / Size","CHECK","Packing List columns G / I / J"])
    qc.append(["Measurement","CHECK","Packing List column V; dimensions only"])
    qc.append(["Hanger Logo","CHECK",f"Template image size {width}x{height}px; original anchor retained"])
    for c in qc[1]:
        c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="1F4E78")
    for col,w in {"A":34,"B":16,"C":90}.items(): qc.column_dimensions[col].width=w
    for row in qc.iter_rows():
        for c in row: c.alignment=Alignment(vertical="top",wrap_text=True)
    qc.freeze_panes="A2"; qc.sheet_view.showGridLines=False
    output=BytesIO(); wb.save(output)
    return output.getvalue(),total

st.set_page_config(page_title="Shipping Mark Automation Pro",page_icon="📦",layout="centered")
st.title("📦 Shipping Mark Automation Pro")
st.caption("Online Web Version • বাংলা / English")
st.info("Chrome Browser থেকেই Packing List এবং Shipping Mark Format Upload করে Shipping Mark তৈরি করুন।")
with st.form("generate_form"):
    packing=st.file_uploader("১. Packing List Excel (.XLS / .XLSX / .XLSM)",type=["xls","xlsx","xlsm"])
    template=st.file_uploader("২. Shipping Mark Format (.XLSX)",type=["xlsx"])
    submitted=st.form_submit_button("Generate Shipping Mark",type="primary",use_container_width=True)
if submitted:
    if not packing or not template: st.error("দুটি Excel ফাইলই নির্বাচন করুন।")
    else:
        try:
            with st.spinner("Excel পরীক্ষা করে Shipping Mark তৈরি হচ্ছে..."):
                records,totals,warnings=build_records(packing)
                output,count=make_output(template,records,totals)
            st.success(f"সম্পন্ন! {count}টি Shipping Mark তৈরি হয়েছে।")
            st.write("QC Summary")
            st.write(f"- Total labels: {count}")
            st.write(f"- Destination: {', '.join(totals.keys())}")
            st.write("- Measurement: dimensions-only format with unit")
            if warnings: st.warning("কিছু quantity cell খালি পাওয়া গেছে:\n\n"+"\n".join(warnings))
            st.download_button("⬇️ Download Shipping Mark Excel",data=output,
                file_name=Path(packing.name).stem+"_Shipping_Mark.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
            st.warning("প্রিন্টের আগে QC REPORT এবং কয়েকটি label মূল Packing List-এর সঙ্গে মিলিয়ে নিন।")
        except Exception as exc: st.error(f"তৈরি করা যায়নি: {exc}")
with st.expander("Software-এর নিয়ম"):
    st.markdown("""
- Sheet: SYDNEY, MELBOURNE, BRISBANE, PERTH
- Carton range: Column A–B, row 44 থেকে
- Colour / Keycode / Size: Column G / I / J
- QTY ISSUE PACK: Column M
- QTY SHIPPER PACK: Column Q
- Measurement: Column V
- Hanger Logo: Template-এর image size ও অবস্থান বজায় রাখা হবে
- Measurement উদাহরণ: 54X45X30 CM; L X W X H লেখা হবে না
""")
