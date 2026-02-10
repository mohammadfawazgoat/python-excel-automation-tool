# Excel Data Processor & Chart Generator

A lightweight Python automation tool built with openpyxl that processes Excel spreadsheets, applies mathematical transformations to data, and generates visual reports.

---

## 🚀 Features
-Batch Processing: Multiplies values in a specific column by a user-defined factor.
-Data Visualization: Automatically generates and embeds a Bar Chart into the spreadsheet.
-Safety Checks: Includes validation to ensure input files exist and prevents accidental overwriting of existing files.
-Non-Destructive: Saves processed data into a new workbook, leaving your original source file untouched.

---

## 🛠️ Prerequisites
Before running the script, ensure you have Python installed along with the openpyxl library.

```Bash
   cd project
   pip install openpyxl
   python app.py
