# Duplicate_scanned_files_finder.py

## Purpose
This script scans a folder of PDF files to identify:

1. **Exact duplicates** – PDFs that are visually identical.
2. **Near duplicates** – PDFs that are visually similar but may differ slightly (e.g., scanned multiple times, minor modifications).

It then generates a comprehensive PDF report with:

- Exact duplicates
- Near duplicates
- Optional thumbnail previews
- Summary totals

## Dependencies

Python 3.10+

- PyMuPDF (fitz)
 – for reading PDF pages

- Pillow
 – for image handling

- imagehash
 – for perceptual hashing

- reportlab
 – for PDF report generation

- tkinter (built-in in standard Python) – for GUI dialogs

## Functions & Classes

### 1. ``hash_pdf_file(filepath: str) -> str``

**Purpose:** Generate a SHA-256 hash of the PDF file’s binary content.

**Parameters:**
- ``filepath`` – Path to the PDF.

**Returns:** SHA-256 hash string.

**Use case:** Detects exact duplicates by comparing file content.

### 2. ``perceptual_hash_pdf(filepath: str) -> Optional[imagehash.ImageHash]``

**Purpose:** Compute a perceptual hash of the first page of a PDF.

**Parameters:**
``filepath`` – Path to the PDF.

**Returns:** ``imagehash.ImageHash`` object for visual similarity comparison. Returns None if page cannot be read.

**Notes:**
- Converts the first page to grayscale and resizes to 256x256 for hashing.

- Used to detect near duplicates visually.

### 3. ``find_duplicate_pdfs(folder_path: str, threshold: int = 5) -> tuple``

**Purpose:** Find exact and near-duplicate PDFs in a folder.

**Parameters:**
- ``folder_path`` – Path containing PDFs.
- ``threshold`` – Maximum Hamming distance for near-duplicate detection.

**Returns:** Tuple (``exact_duplicates``, ``near_duplicates``) where each is a list of tuples (``duplicate_file``, ``original_file``, ``distance``)

**Notes:**
- ``distance = 0`` → exact match
- ``distance ≤ threshold`` → near duplicate

### 4. ``generate_pdf_report(folder_path, exact_duplicates, near_duplicates, save_path, thumbnail=True)``

**Purpose:** Generate a visual PDF report of duplicates.

**Parameters:**
- ``folder_path`` – Folder containing PDFs

- ``exact_duplicates`` – List from ``find_duplicate_pdfs()``
- ``near_duplicates`` – List from ``find_duplicate_pdfs()``
- ``save_path`` – File path to save the report
- ``thumbnail`` – If True, include small image previews

**Behavior:**
- Creates a title page with totals
- Adds tables for exact & near duplicates
- Optionally embeds thumbnail images of the first page
- Handles errors safely and cleans up temporary files

### 5. Main Execution

**Behavior:**
- Opens a folder selection dialog via ``tkinter``.
- Scans PDFs for duplicates using ``find_duplicate_pdfs()``.
- Opens a "Save As" dialog for the report.
- Generates PDF report using ``generate_pdf_report()``.
- Prints completion message.

### Example Usage
```bash
python Duplicate_scanned_files_finder.py
```

1. Select folder containing PDFs.

2. Select save location for the report.

3. Wait for the process to complete.

### Key Implementation Notes
- Uses perceptual hashing to handle scanned PDFs and minor variations.
- Thumbnail generation allows visual confirmation of duplicates.
- Threshold parameter lets you tune sensitivity for near duplicates.
- Temporary thumbnail images are cleaned up automatically.
