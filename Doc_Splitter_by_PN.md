# Documentation for PDF Splitter by Person Number

## Overview

This script splits large PDF files (like salary review letters) into individual PDFs per employee, based on detecting:

- The employee’s name (using "Dear ...")
- Their person number (from explicit "Person Number:" field or directly below the name)

It uses PyPDF2 for PDF processing and Tkinter for a simple GUI with progress reporting.

## Features

- Select multiple PDFs to process at once
- Extract employee first/last name and person number
- Save each employee’s PDF separately with structured filenames:

```py
	YYYY-MM_Salary Review_LASTNAME_Firstname_PersonNumber.pdf
```

- Handles continuation pages (multiple pages per employee
- GUI shows progress and automatically closes when finished

## Requirements
Install dependencies:

```bash
pip install PyPDF2
```
(Tkinter is included with standard Python on Windows/Mac; on Linux you may need to install python3-tk.)

## Usage

1. Run the script (python pdf_splitter.py)
2. Use the file dialog to select input PDF(s)
3. Choose an output folder
4. Wait while it processes — progress is shown in the GUI
5. Individual PDFs are saved into a timestamped subfolder