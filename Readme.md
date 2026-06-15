# Duplicate Data Checker

A Python utility that reads **Excel (.xlsx)** and **Word (.docx)** files from a folder, checks for duplicate records based on one or more user-selected columns, and generates separate files for duplicate and non-duplicate records.

---

# Features

* Read multiple `.xlsx` files
* Read `.docx` files containing tabular data
* User selects one or more columns for duplicate detection
* Detect duplicates across all loaded files
* Generate:

  * `duplicates.xlsx`
  * `non_duplicates.xlsx`
* Detailed execution logging
* Error handling and validation
* Production-ready folder structure

---

# Project Structure

```text
duplicate_checker/
│
├── duplicate_checker.py
│
├── input_files/
│   ├── customers.xlsx
│   ├── orders.xlsx
│   └── inventory.docx
│
├── output/
│   ├── duplicates.xlsx
│   ├── non_duplicates.xlsx
│   └── duplicate_check.log
│
└── README.md
```

---

# Prerequisites

* Python 3.9 or higher
* pip package manager

Verify installation:

```bash
python --version
```

Example:

```text
Python 3.11.8
```

---

# Create Virtual Environment

## Windows

```bash
python -m venv venv
```

## Linux / Mac

```bash
python3 -m venv venv
```

---

# Activate Virtual Environment

## Windows CMD

```bash
venv\Scripts\activate
```

## Windows PowerShell

```powershell
.\venv\Scripts\Activate.ps1
```

## Linux / Mac

```bash
source venv/bin/activate
```

After activation you should see:

```text
(venv)
```

at the beginning of your terminal prompt.

---

# Install Dependencies

```bash
pip install pandas openpyxl python-docx xlsxwriter xlrd python-dotenv
```

---

# Create requirements.txt

```bash
pip freeze > requirements.txt
```

Install later using:

```bash
pip install -r requirements.txt
```

---

# Input Files

Place all Excel and Word files inside a folder.

Example:

```text
input_files/
├── file1.xlsx
├── file2.xlsx
├── file3.xlsx
└── file4.docx
```
---

---
# Run the Application

```bash
python duplicate_checker.py
```
---

# Example Execution

```text
================================
 DUPLICATE DATA CHECKER
================================

Enter folder path containing files:
D:\input_files

Available Columns:

customer_id
customer_name
city
source_file

Enter columns for duplicate check:
customer_id
```

Or multiple columns:

```text
customer_id,customer_name
```

---

# Install Dependencies on Another Machine

```bash
python -m venv venv

venv\Scripts\activate

pip install -r requirements.txt
```

---

# Deactivate Environment

```bash
deactivate
```

---

# Supported Formats

| Format      | Supported |
| ----------- | --------- |
| XLSX        | Yes       |
| DOCX Tables | Yes       |
| XLS         | Yes       |
| CSV         | Yes        |
| PDF         | No        |

---

# Author

Duplicate Data Checker Utility

Python + Pandas + OpenPyXL + python-docx

# .ENV
INPUT_FOLDER=INPUT_FOLDER_PATH
OUTPUT_FOLDER=OUTPUT_FOLDER_PATH
