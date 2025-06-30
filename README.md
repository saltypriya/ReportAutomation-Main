# Report Automation Tool

This Python application generates inspection reports from CSV/Excel data and associated photos, with automatic room detection from folder structure.

## Features

- Creates professional Word documents from template
- Automatically organizes photos by room folders
- Generates placeholder images when photos are missing
- Includes header/footer images support
- Simple GUI interface

## Requirements

- Python 3.8+
- Windows (or macOS/Linux with minor adjustments)

## Installation

1. **Install Python** (if not already installed):
   - Download from [python.org](https://www.python.org/downloads/)
   - Check "Add Python to PATH" during installation

2. **Install dependencies**:
   ```bash
   pip install pandas python-docx pillow openpyxl tk

How to Use
Basic Usage
Prepare your files:

Input data: CSV or Excel file with claim information

Photos: Organize in folders by room name (e.g., kitchen/, bedroom1/)

Optional: Add header.jpg and footer.jpg in main photos folder

Run the application:

bash
python ReportGenerator.py
Using the GUI:

Click "Select Input CSV/Excel File" and choose your data file

Click "Select Images Folder" and choose your photos directory

Click "Select Output Folder" for where to save reports

Click "Generate Reports"


Creating an Executable (Optional)
To create a standalone .exe file:

Install PyInstaller:

bash
pip install pyinstaller
Build the executable:

bash
pyinstaller --onefile --windowed ReportGenerator.py
The executable will be in the dist/ folder
