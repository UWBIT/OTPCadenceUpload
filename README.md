# Cadence Data Cleaning Script

## Overview

This Python script processes and cleans student contact data exported from an Excel file. It is designed to prepare the data for upload into **Cadence**, a texting platform used for communicating with incoming first-year and transfer students. The script ensures data consistency, validates phone numbers, and handles opt-out preferences.

---

## Features

- **Name Reconciliation**: Compares three name fields (Admissions, OTP, Legal) and prompts the user to choose the correct one if they differ. Automatically fills in missing names when possible.
- **Phone Number Validation**: Flags invalid phone numbers (non-numeric or not 10 digits) with an orange highlight.
- **Opt-Out Handling**: Moves students who opted out of texting (`sms_allowed == 0`) to a separate sheet titled **NO SMS**.
- **Quarter and Application Type Mapping**: Converts numeric codes to readable text (e.g., `1` → `Winter`, `2` → `Transfer`).
- **Column Cleanup**: Removes unnecessary columns and inserts new ones with default values (e.g., `Application_Campus` set to "Bothell", `Import_Year` set to "2025").
- **Header Renaming**: Updates column headers to match Cadence's required format.
- **Output**: Saves the cleaned data to a new Excel file with `_Upload` appended to the original filename.

---

## Requirements

- Python 3.x
- `openpyxl` library
- GUI environment (uses `tkinter` for file selection)

Install dependencies with:
pip install openpyxl
import PatternFill
from tkinter import Tk
from tkinter.filedialog import askopenfilename


