#📘 README – Stock Register Excel Generator
📌 Overview

Ye program automatically multiple CSV files ko read karta hai aur:

Specific SKU ka data filter karta hai
Daily sales calculate karta hai
Missing dates detect karta hai (OFF days)
Excel file banata hai with:
Proper stock data
OFF days highlighted 🔴


🧾 Requirements

Is program ko chalane ke liye aapko ye cheezen chahiye:

1️⃣ Install Python

👉 Download Python:
https://www.python.org/downloads/

📌 Important:

Install karte waqt "Add Python to PATH" tick zaroor karein
2️⃣ Install VS Code (Recommended)

👉 Download VS Code:
https://code.visualstudio.com/

📌 VS Code ek simple software hai jahan aap code run kar sakte hain

3️⃣ Install Required Python Libraries

VS Code open karein → Terminal open karein → ye command run karein:

pip install pandas openpyxl xlsxwriter
📁 Folder Setup

Ek folder banayein (example: Stock_Project)
Uske andar ye files rakhein:

Stock_Project/
│
├── 2026-03-01.csv
├── 2026-03-02.csv
├── 2026-03-03.csv
├── ...
├── script.py

📌 Important:

CSV files ka naam date format mein hona chahiye
Example: 2026-03-01.csv
⚙️ Configuration (Script Settings)

Script ke andar ye values change kar sakte hain:

TARGET_SKU = "Capstan by Pall Mall 20HL"
START_DATE = "2026-03-01"
END_DATE   = "2026-03-31"

📌 Explanation:

TARGET_SKU → jis product ka data chahiye
START_DATE → start date
END_DATE → end date
▶️ Program Run Karna
Step 1:

VS Code open karein

Step 2:

Folder open karein (File → Open Folder)

Step 3:

Script open karein (script.py)

Step 4:

Run button dabayein ya terminal mein likhein:

python script.py
📊 Output

Program run hone ke baad ek Excel file generate hogi:

Capstan by Pall Mall 20HL_Stock_Register_2026-03-01_to_2026-03-31.xlsx
🎯 Excel Features

✔ Daily stock data
✔ Total sales calculated
✔ Missing dates auto add
✔ OFF days highlighted (Red color) 🔴

❗ Important Notes
CSV files ka format same hona chahiye
Column names change na karein
Date format file name se automatically pick hota hai
🧠 Simple Samajh

Ye program:

👉 CSV files uthata hai
👉 Ek SKU filter karta hai
👉 Har din ka sales nikalta hai
👉 Missing din add karta hai
👉 Excel bana deta hai

🛠 Troubleshooting
❌ Error: "No module found"
👉 Run:

pip install pandas openpyxl xlsxwriter
❌ Excel file nahi bani

👉 Check karein:

CSV files folder mein hain?
File names correct hain? (YYYY-MM-DD.csv)
💡 Future Upgrade Ideas

👨‍💻 Developer Note

Ye tool manual reporting ka time bachane ke liye banaya gaya hai — especially daily stock tracking ke liye.

BY WALID ISLAM
👉 Run:

