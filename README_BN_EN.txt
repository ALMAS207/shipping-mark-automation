SHIPPING MARK AUTOMATION PRO — ONLINE BROWSER VERSION

এখন app.py হলো Streamlit Browser App। Chrome থেকে চালাতে Streamlit Community Cloud-এ একবার Deploy করতে হবে।

ONLINE DEPLOY (একবারই করতে হবে)
1. https://share.streamlit.io/ খুলুন।
2. Sign in with GitHub চাপুন এবং ALMAS207/shipping-mark-automation repository access দিন।
3. Create app / Deploy an app নির্বাচন করুন।
4. Repository: ALMAS207/shipping-mark-automation
5. Branch: main
6. Main file path: app.py
7. Deploy চাপুন এবং build শেষ হওয়া পর্যন্ত অপেক্ষা করুন।
8. Deploy শেষ হলে Streamlit একটি web link দেবে। সেই link Chrome-এ Bookmark করে রাখুন।

প্রতিদিন ব্যবহার
1. Streamlit-এর পাওয়া web link খুলুন।
2. Packing List Upload করুন (.xls / .xlsx / .xlsm)।
3. Shipping Mark Format Upload করুন (.xlsx)।
4. Generate Shipping Mark চাপুন।
5. Shipping Mark Excel Download করুন।
6. Download করা Excel-এর QC REPORT এবং কয়েকটি label মিলিয়ে নিন।

SUPPORTED INPUT
- Packing List: .XLS, .XLSX, .XLSM
- Shipping Mark Template: .XLSX
- Destination sheets: SYDNEY, MELBOURNE, BRISBANE, PERTH

CURRENT DATA MAPPING
- Order / Style / Department / Description: Packing List header section
- Carton range: Column A to B, rows 44 onward
- Colour: Column G
- Keycode: Column I
- Size: Column J
- QTY ISSUE PACK: Column M
- QTY SHIPPER PACK: Column Q
- Carton Measurement: Column V; dimensions and unit only (example: 54X45X30 CM)
- Colour / Keycode / Size ditto marks: previous valid value carried forward
- One label per carton
- Hanger Logo image size and anchor copied from the selected format
- Output includes QC REPORT sheet

IMPORTANT
এই app বর্তমান sample Packing List-এর নির্দিষ্ট sheet name, column এবং row layout অনুসরণ করে। অন্য order-এর Packing List-এ একই structure না থাকলে mapping পরিবর্তন করতে হবে। প্রথমবার deploy করার পর test file দিয়ে output যাচাই করুন। QC REPORT ও label না দেখে সরাসরি print করবেন না।

Windows-এর .bat ফাইলগুলো Desktop version-এর জন্য; Online Browser version চালাতে Streamlit deploy করতে হবে।
