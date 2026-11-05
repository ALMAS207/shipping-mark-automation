SHIPPING MARK AUTOMATION PRO — CUSTOM BUILD

এই সংস্করণটি আপনার দেওয়া sample:
- AUS 22163703-34NSUTB3NNFD.XLS
- formet(8).xlsx
দিয়ে নির্ধারিত layout ও column mapping অনুযায়ী তৈরি।

INSTALLATION (Windows 10/11)
1. Python 3.10+ ইনস্টল করুন এবং installation-এর সময় "Add Python to PATH" নির্বাচন করুন।
2. ZIP Extract করুন।
3. install_requirements.bat চালান।
4. run_app.bat চালান।
5. Packing List এবং Shipping Mark Format নির্বাচন করে Generate Shipping Mark চাপুন।

SUPPORTED INPUT
- Packing List: .XLS, .XLSX, .XLSM
- Shipping Mark Template: .XLSX
- Destination sheets: SYDNEY, MELBOURNE, BRISBANE, PERTH

CUSTOM MAPPING
- Order / Style / Department / Description: Packing List-এর header section
- Carton range: Column A to B, rows 44 onward
- Colour: Column G
- Keycode: Column I
- Size: Column J
- QTY ISSUE PACK: Column M (Unit Qty)
- QTY SHIPPER PACK: Column Q (Blister Qty)
- Carton Measurement: Column V
- Ditto marks in Colour/Keycode/Size: previous valid value carry-forward
- One label per carton in the carton range
- Carton number total: all destination sheets combined
- Hanger logo: template image size and anchor copied to each generated label

OUTPUT
- One worksheet per carton
- QC REPORT worksheet

IMPORTANT
This is a custom build for the sample structure supplied. If future Packing Lists change sheet names, columns, or row positions, the mapping will need updating. Always inspect QC REPORT and visually check sample labels before printing. The software does not run in the background; launch it when you need to process files.
