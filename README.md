# 🧾 First Inspection Report Generator

A desktop app to **automatically generate First Inspection Reports** with embedded photos and insurance data using an Excel/CSV input.

---

## 🖥️ GUI Instructions (No Coding Needed)

### ▶️ How to Use

1. **Run the App**  
   - From source:  
     ```bash
     python ReportGenerator.py
     ```  
   - Or double-click the `ReportGenerator.exe` (if using the EXE version)

2. **Main Interface Overview**
   [ First Inspection Report Generator ]
   Select Input CSV/Excel File [Browse]
   
   Select Images Folder [Browse]
   
   Select Output Folder [Browse]
   
   [ Generate Reports ] (Button)
   
   Status: Ready


3. **What to Do**
- 📄 Select your Excel/CSV file that has insured party details.
- 🖼️ Select a folder that contains:
  - Room folders (e.g., `bedroom1/`, `kitchen/`)
  - Optional `header.jpg` and `footer.jpg` images in the root
- 📂 Choose output folder for saving generated Word reports

4. **Click “Generate Reports”**  
Watch progress in the status bar. Files will be saved as:
   Output/
└── FIRST INSPECTION REPORT - CLAIM# PR12345 - SMITH - 123_MAIN_ST.docx


---

## 🛠️ Create an EXE (No Python Needed for Users)

### 🔧 Method 1: PyInstaller (Recommended)

1. Install PyInstaller:
```bash
pip install pyinstaller

Build the EXE:

bash
Copy
Edit
pyinstaller --onefile --windowed ReportGenerator.py
Advanced build:

bash
Copy
Edit
pyinstaller --onefile --windowed --icon=app.ico --add-data "assets;assets" ReportGenerator.py
