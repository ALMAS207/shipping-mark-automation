SHIPPING MARK AUTOMATION PRO — ONLINE BROWSER VERSION

এখন app.py হলো Streamlit Browser App। Chrome থেকে চালাতে Streamlit Community Cloud-এ একবার Deploy করতে হবে।

ONLINE DEPLOY (একবারই করতে হবে)
1. https://share.streamlit.io/ খুলুন।
2. Sign in with GitHub চাপুন এবং ALMAS207/shipping-mark-automation repository access দিন।
3. Create app / Deploy an app নির্বাচন করুন।
4. Repository: ALMAS207/shipping-mark-automation
5. Branch: main
6. Main file path: streamlit_app.py
7. Deploy চাপুন এবং build শেষ হওয়া পর্যন্ত অপেক্ষা করুন।
8. Deploy শেষ হলে Streamlit একটি web link দেবে। সেই link Chrome-এ Bookmark করে রাখুন।

প্রতিদিন ব্যবহার
1. Streamlit-এর পাওয়া web link খুলুন।
2. প্রতিটি Style-এর Packing List আলাদা Excel ফাইল হিসেবে একসঙ্গে Upload করুন (১০–২০টি পর্যন্ত)।
3. প্রতিটি Style-এর Shipping Mark Format-ও আলাদা Excel ফাইল হিসেবে একসঙ্গে Upload করুন।
4. Software Style No. (Packing List header ও Format-এর C6) মিলিয়ে দুই ধরনের ফাইল pair করবে।
5. Generate Shipping Mark চাপুন।
6. সব Style-এর Shipping Mark ZIP ফাইল হিসেবে Download করুন। ZIP-এর ভিতরে প্রতিটি Style-এর আলাদা Excel এবং QC REPORT থাকবে।
6. Download করা Excel-এর QC REPORT এবং কয়েকটি label মিলিয়ে নিন।

SUPPORTED INPUT
- একসঙ্গে একাধিক Packing List: .XLS, .XLSX, .XLSM (প্রতি ফাইলে একটি Style)
- একসঙ্গে একাধিক Shipping Mark Format: .XLSX (প্রতি ফাইলে একটি Style)
- Pairing: Style No. দিয়ে স্বয়ংক্রিয়ভাবে মিলানো হয়; missing বা duplicate Style pair হলে Generate বন্ধ করে QC error দেখায়
- Destination sheets: SYDNEY, MELBOURNE, BRISBANE, PERTH

CURRENT DATA MAPPING
- Order / Style / Department / Description: Packing List header section
- Carton range: Column A to B, rows 44 onward
- Colour: Column G
- Keycode: Column I
- Size: Column J
- QTY ISSUE PACK: Column M
- QTY SHIPPER PACK: Column Q
- Carton Measurement: Shipping Mark Format-এর matching sheet থেকে; dimensions and unit only (example: 54X45X30 CM). No unique match হলে blank থাকবে; কোনো default dimension বসবে না।
- Colour / Keycode / Size ditto marks: previous valid value carried forward
- One label per carton
- Hanger Logo image size and anchor copied from the selected format
- Output includes QC REPORT sheet

IMPORTANT
এই app বর্তমান sample Packing List-এর নির্দিষ্ট sheet name, column এবং row layout অনুসরণ করে। অন্য order-এর Packing List-এ একই structure না থাকলে mapping পরিবর্তন করতে হবে। প্রথমবার deploy করার পর test file দিয়ে output যাচাই করুন। QC REPORT ও label না দেখে সরাসরি print করবেন না।

Windows-এর .bat ফাইলগুলো Desktop version-এর জন্য; Online Browser version চালাতে Streamlit deploy করতে হবে।
