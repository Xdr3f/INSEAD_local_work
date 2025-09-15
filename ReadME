# PDF Utilities Suite

This project contains three Python scripts for working with PDF files:

1. Duplicate Scanned Files Finder – Detects exact and near-duplicate PDFs in a folder and generates a detailed PDF report with optional thumbnails.

2. Matching Name Verification – Checks whether the name in a PDF filename appears in the PDF content, generating a report for files with no match or partial match.

3. PDF Splitter by Person Number – Splits PDFs into individual files based on a person number in the filename.

## Requirements

Python 3.13+ is recommended.

Install the required packages using ``pip``:

``` bash
pip install -r requirements.txt
```

**requirements.txt** example:

```shell
pdfplumber>=0.7.6
rapidfuzz>=2.15.0
reportlab>=4.0
Pillow>=10.0
PyMuPDF>=1.22
```

The versions indicate minimum compatibility. Newer versions are allowed.

## Script 1: Duplicate Scanned Files Finder

File: ``Duplicate_scanned_files_finder.py``

### Description:
Scans a folder of PDFs to identify duplicates. It distinguishes between exact duplicates and near-duplicates (using perceptual hashing) and generates a PDF report with optional thumbnails.

### Usage:

```bash
python Duplicate_scanned_files_finder.py
```

1. Select the folder containing PDFs.
2. The script scans for duplicates.
3. Select a save location for the PDF report.
4. The report contains Exact Duplicates and Near-Duplicates tables.

## Script 2: Matching Name Verification

File: ``Matching_Name_Verification.py``

### Description:
Verifies if the names in PDF filenames appear in the PDF content. Identifies perfect matches, partial matches, and files with no match, generating a detailed report.

### Usage:
```bash
python Matching_Name_Verification.py
```

1. Select the folder containing PDFs.
2. The script scans each PDF for name matches.
3. Select a save location for the PDF report.
4. The report contains tables for:
	- No Match – Names not found in content.
	- Partial Match – Close matches.
	- Errors – Files that could not be processed.

## Script 3: PDF Splitter by Person Number

File: ``PDF_Splitter_by_Person_Number.py``

### Description:
Splits PDFs based on a person number in the filename. Useful for separating multi-person documents into individual PDFs.

### Usage:

```bash
python PDF_Splitter_by_Person_Number.py
```

Select the folder containing PDFs.
The script splits each PDF into separate files for each person number found in the filename.
Output files are saved in a structured folder system to avoid overwriting.

### Notes
- All scripts use Tkinter for folder selection dialogs and progress updates.
- Reports are generated in PDF format using ReportLab.
- Ensure your PDFs are readable and not password-protected.
- Scripts allow customization:
	- Duplicate Finder: threshold for near-duplicates.
	- Name Verification: number of threads, match thresholds.
	- PDF Splitter: output folder structure.

### Example Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run duplicate PDF finder
python Duplicate_scanned_files_finder.py

# Run filename match verifier
python Matching_Name_Verification.py

# Run PDF splitter by person number
python PDF_Splitter_by_Person_Number.py
```