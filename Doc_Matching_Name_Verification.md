# Mathing_Name_Verification.py
## Purpose

This script verifies if the person's name embedded in a PDF filename appears in the PDF content. It categorizes files into:

- **Perfect match** – Name found exactly in text
- **Partial match** – Name partially matches using fuzzy similarity
- **No match** – Name not found or below threshold
- **Errors** – Any processing issues

It generates a PDF report summarizing results for human review.

## Dependencies

- Python 3.10+
- pdfplumber
	– for extracting text from PDFs
- rapidfuzz
	– for fuzzy string matching
- reportlab
	– for PDF report generation
- tkinter – for folder selection & GUI

## Core Components
### 1. ``Config`` Class

Holds configurable parameters:
- ``NO_MATCH_THRESHOLD = 30``
- ``PERFECT_MATCH_THRESHOLD = 100``
- ``MAX_WORKERS = 4`` – Threads for parallel processing
- ``CHUNK_SIZE = 2000`` – Characters processed at a time

### 2. ``extract_text_from_pdf(pdf_path)``

- Extracts all text from a PDF using pdfplumber.
- Returns ``(text_content, error_message)``
- Handles errors per page and logs warnings.

### 3. ``normalize_text(text, preserve_accents=False)``

- Converts text to lowercase
- Removes punctuation and optionally accents
- Normalizes whitespace

### 4. ``extract_name_tokens(filename)``

- Converts filename into candidate name tokens
- Supports ``_``, ``-``, ``.``, separators
- Returns list of (``normalized_no_accents``, ``normalized_with_accents``)

### 5. ``check_name_in_pdf(pdf_path)``

- Core function for matching PDF text against filename tokens
- Handles single-token and multi-token names
- Performs exact and fuzzy matching using RapidFuzz
- Returns a dictionary:
```py
{
  "best_pair": tuple of matched name tokens,
  "best_score": similarity score (0-100),
  "matched_text": best matching substring,
  "error": None or error string
}
```
### 6. ``PDFScannerGUI`` Class

- Tkinter-based GUI
- Features:
	- Folder selection
	- Progress bar
	- Status & detail messages
- Uses ``ThreadPoolExecutor`` to scan PDFs in parallel
- Generates enhanced PDF report using ``generate_enhanced_pdf_report()``
- Color-codes results: green (perfect), orange (partial), red (no match)

### 7. ``generate_enhanced_pdf_report(results, save_path)``

- Builds PDF report including:
	- Summary of all files
	- Tables for ``Errors``, ``No Match``, ``Partial Match``
	- Highlights scores with colors
- Supports large datasets and handles text truncation for readability

## Example Usage
```bash
python Mathing_Name_Verification.py
```

1. Select folder containing PDFs.
2. Wait while GUI updates progress.
3. Select location to save report.
4. Report opens with categorized results.

### Key Implementation Notes

- Handles large PDFs efficiently using chunking
- Supports accented and non-accented names
- Parallel processing improves speed on many PDFs
- Includes robust error handling and logging
- PDF report highlights critical issues for quick review

### Required Packages (``requirements.txt``)
```ini
PyMuPDF==1.22.5
Pillow==10.1.0
imagehash==4.3.1
reportlab==4.0.0
pdfplumber==0.9.0
rapidfuzz==2.16.0
```

**Note:**
	tkinter is included with standard Python installations.