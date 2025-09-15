"""
PDF Name Matching & Verification Script
---------------------------------------

This script scans PDF files in a selected folder and checks whether the
name embedded in the PDF filename appears in the PDF content.

It generates a PDF report listing files with:
- No match
- Partial match (fuzzy similarity)
- Errors
- Perfect matches are excluded from the report but counted in the summary

Dependencies:
- pdfplumber: Extract text from PDFs
- rapidfuzz: Perform fuzzy string matching
- reportlab: Generate PDF reports
- tkinter: GUI for folder selection and progress display
"""

import os
import re
import unicodedata
import logging
import pdfplumber
from rapidfuzz import fuzz
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Tuple, Optional
import datetime
import threading
import sys
import traceback

# ------------------------------
# Logging Setup
# ------------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ------------------------------
# Configuration
# ------------------------------
class Config:
    """Configuration class with adjustable parameters"""
    NO_MATCH_THRESHOLD = 30        # Score below which we consider as no match
    PERFECT_MATCH_THRESHOLD = 100  # Score at which a match is considered perfect
    MAX_WORKERS = 4                # Max number of threads for parallel PDF processing
    CHUNK_SIZE = 2000              # Characters processed at a time for large PDFs

# ------------------------------
# PDF Text Extraction
# ------------------------------
def extract_text_from_pdf(pdf_path: str) -> Tuple[str, Optional[str]]:
    """
    Extracts text from a PDF file.

    Args:
        pdf_path (str): Path to PDF file

    Returns:
        Tuple[str, Optional[str]]: Extracted text and any error message
    """
    text_content = []
    error_message = None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                try:
                    page_text = page.extract_text()
                    if page_text and not page_text.isspace():
                        text_content.append(page_text)
                except Exception as e:
                    # Log page-specific extraction error but continue
                    error_message = f"Error on page {page_num}: {str(e)}"
                    logging.warning(error_message)

    except Exception as e:
        error_message = f"Error opening PDF: {str(e)}"
        logging.error(error_message)

    return " ".join(text_content), error_message

# ------------------------------
# Text Normalization
# ------------------------------
def normalize_text(text: str, preserve_accents: bool = False) -> str:
    """
    Normalize text for comparison:
    - Lowercase
    - Remove punctuation
    - Optional: preserve accents
    - Normalize whitespace

    Args:
        text (str): Input text
        preserve_accents (bool): Whether to preserve accents

    Returns:
        str: Normalized text
    """
    text = (text or "").lower()
    if preserve_accents:
        text = unicodedata.normalize("NFKC", text)
    else:
        text = unicodedata.normalize("NFKD", text).encode('ascii', 'ignore').decode('ascii')
    # Keep only letters, digits, spaces, and hyphens
    text = re.sub(r'[^a-z0-9\s-]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

# ------------------------------
# Filename Tokenization
# ------------------------------
def extract_name_tokens(filename: str) -> List[str]:
    """
    Extract potential name tokens from filename
    Converts separators (_ - . space) to underscore and splits

    Args:
        filename (str): Filename to parse

    Returns:
        List[str]: List of tuples (no_accents, with_accents)
    """
    base = os.path.splitext(os.path.basename(filename))[0]

    for sep in ['_', '-', '.', ' ']:
        base = base.replace(sep, '_')

    parts = base.split('_')
    tokens = []

    for part in parts:
        norm_no_accents = normalize_text(part, preserve_accents=False)
        norm_with_accents = normalize_text(part, preserve_accents=True)

        if (norm_no_accents and not norm_no_accents.isspace()) or \
           (norm_with_accents and not norm_with_accents.isspace()):
            tokens.append((norm_no_accents, norm_with_accents))

    return tokens

# ------------------------------
# Matching Functions
# ------------------------------
def is_bidirectional_match(a: str, b: str) -> bool:
    """
    Simple exact match check after normalization
    """
    return normalize_text(a) == normalize_text(b)

def find_exact_word_match(word: str, text: str) -> Tuple[bool, str, List[str]]:
    """
    Checks if the exact phrase appears in the text (normalized whitespace)

    Returns:
        Tuple[bool, str, List[str]]: (found, full_match_text, context_matches)
    """
    norm_word = normalize_text(word)
    norm_text = normalize_text(text)
    if f" {norm_word} " in f" {norm_text} ":
        return True, word, []
    return False, "", []

def get_best_ngram_match(name: str, chunk: str, ngram_range=(2, 5)):
    """
    For a chunk of text, find the n-gram substring that best matches a name
    using fuzzy matching.
    """
    words = chunk.split()
    candidates = set()
    for n in range(ngram_range[0], ngram_range[1]+1):
        for i in range(len(words) - n + 1):
            candidates.add(" ".join(words[i:i+n]))
    candidates.add(chunk)  # fallback
    from rapidfuzz import process
    match, score, _ = process.extractOne(name, candidates)
    return match if match else chunk

# ------------------------------
# Main PDF Name Checker
# ------------------------------
def check_name_in_pdf(pdf_path: str) -> Dict:
    """
    Check PDF text against filename name tokens.

    Returns a dictionary with:
    - best_pair: matched name tokens
    - best_score: similarity score 0-100
    - matched_text: substring from PDF that best matches
    - error: any error encountered
    """
    pdf_text, error = extract_text_from_pdf(pdf_path)
    if error:
        return {"best_pair": None, "best_score": 0, "error": error}

    # Normalize text for matching
    pdf_no_accents = normalize_text(pdf_text, preserve_accents=False)
    pdf_with_accents = normalize_text(pdf_text, preserve_accents=True)
    pdf_no_accents_norm = re.sub(r'\s+', ' ', pdf_no_accents)
    pdf_with_accents_norm = re.sub(r'\s+', ' ', pdf_with_accents)

    tokens = extract_name_tokens(pdf_path)
    alpha_tokens = [(no_acc, with_acc) for no_acc, with_acc in tokens if no_acc.isalpha()]

    if not alpha_tokens:
        return {"best_pair": None, "best_score": 0, "error": "No valid alphabetic name tokens found"}

    best_pair = None
    best_score = 0
    best_matched_text = ""
    perfect_match_found = False

    # Case 1: single-token names
    if len(alpha_tokens) == 1:
        no_acc, with_acc = alpha_tokens[0]
        for text in [pdf_no_accents_norm, pdf_with_accents_norm]:
            for name in [no_acc, with_acc]:
                found, full_match, context_matches = find_exact_word_match(name, text)
                if found and is_bidirectional_match(full_match, name):
                    perfect_match_found = True
                    best_score = 100
                    best_pair = (name, "")
                    best_matched_text = full_match
                    break
                # Fuzzy matching for partial match
                for chunk_start in range(0, len(text), Config.CHUNK_SIZE):
                    chunk = text[chunk_start:chunk_start + Config.CHUNK_SIZE]
                    ratio = fuzz.ratio(name, chunk)
                    partial = fuzz.partial_ratio(name, chunk)
                    score = max(ratio, partial)
                    if not perfect_match_found and score > best_score:
                        best_score = min(score, Config.PERFECT_MATCH_THRESHOLD - 1)
                        best_pair = (name, "")
                        best_matched_text = get_best_ngram_match(name, chunk)
            if perfect_match_found:
                break

    # Case 2: multi-token names
    if not perfect_match_found:
        for i in range(len(alpha_tokens)):
            for j in range(i + 1, len(alpha_tokens)):
                first_no_acc, first_with_acc = alpha_tokens[i]
                last_no_acc, last_with_acc = alpha_tokens[j]

                combinations = [
                    (f"{first_no_acc} {last_no_acc}", f"{first_with_acc} {last_with_acc}"),
                    (f"{last_no_acc} {first_no_acc}", f"{last_with_acc} {first_with_acc}"),
                    (f"{first_no_acc}, {last_no_acc}", f"{first_with_acc}, {last_with_acc}")
                ]

                for no_acc_combined, with_acc_combined in combinations:
                    for text in [pdf_no_accents_norm, pdf_with_accents_norm]:
                        for name in [no_acc_combined, with_acc_combined]:
                            found, full_match, _ = find_exact_word_match(name, text)
                            if found and is_bidirectional_match(full_match, name):
                                perfect_match_found = True
                                best_score = 100
                                best_pair = (first_with_acc, last_with_acc)
                                best_matched_text = full_match
                                break
                            for chunk_start in range(0, len(text), Config.CHUNK_SIZE):
                                chunk = text[chunk_start:chunk_start + Config.CHUNK_SIZE]
                                ratio = fuzz.ratio(name, chunk)
                                partial = fuzz.partial_ratio(name, chunk)
                                score = max(ratio, partial)
                                if not perfect_match_found and score > best_score:
                                    best_score = min(score, Config.PERFECT_MATCH_THRESHOLD - 1)
                                    best_pair = (first_with_acc, last_with_acc)
                                    best_matched_text = get_best_ngram_match(name, chunk)
                        if perfect_match_found:
                            break
                    if perfect_match_found:
                        break
                if perfect_match_found:
                    break
            if perfect_match_found:
                break

    return {
        "best_pair": best_pair,
        "best_score": best_score,
        "matched_text": best_matched_text,
        "error": None
    }

# ------------------------------
# GUI Class
# ------------------------------
class PDFScannerGUI:
    """
    Tkinter GUI for PDF folder selection, progress display,
    and generating a PDF report.
    """
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF Name Scanner")
        self.setup_gui()
        self.processing_lock = threading.Lock()

    def setup_gui(self):
        """Initialize GUI layout with buttons, progress bar, and labels"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Button to select folder and start processing
        ttk.Button(main_frame, text="Select Folder", command=self.process_folder).grid(row=0, column=0, pady=5)

        # Progress bar section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="5")
        progress_frame.grid(row=1, column=0, pady=5, sticky=(tk.W, tk.E))
        self.progress = ttk.Progressbar(progress_frame, length=300, mode='determinate')
        self.progress.grid(row=0, column=0, pady=5)
        self.status_var = tk.StringVar()
        self.detail_var = tk.StringVar()
        ttk.Label(progress_frame, textvariable=self.status_var).grid(row=1, column=0, pady=2)
        ttk.Label(progress_frame, textvariable=self.detail_var).grid(row=2, column=0, pady=2)

    def safe_update_gui(self, status=None, detail=None, progress=None):
        """Thread-safe GUI updates"""
        def update():
            if status is not None:
                self.status_var.set(status)
            if detail is not None:
                self.detail_var.set(detail)
            if progress is not None:
                self.progress["value"] = progress
            self.root.update()
        self.root.after(0, update)

    # ------------------------------
    # Folder Processing
    # ------------------------------
    def process_folder(self):
        """Main processing function to scan PDFs and generate report"""
        folder_path = filedialog.askdirectory(title="Select the folder containing PDFs")
        if not folder_path:
            return

        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
        total_files = len(pdf_files)
        if total_files == 0:
            messagebox.showerror("No PDFs Found", "No PDF files found in the selected folder.")
            self.safe_update_gui(status="No PDF files found in selected folder")
            return

        # Initialize results storage
        results = {
            "no_match": [], 
            "partial_match": [], 
            "perfect_match": [],
            "errors": []
        }

        self.progress["maximum"] = total_files
        self.progress["value"] = 0

        # Parallel PDF processing
        with ThreadPoolExecutor(max_workers=Config.MAX_WORKERS) as executor:
            futures = {executor.submit(check_name_in_pdf, os.path.join(folder_path, f)): f for f in pdf_files}
            completed = 0
            for future in as_completed(futures):
                file = futures[future]
                try:
                    res = future.result()
                    score = res["best_score"]
                    pair = res["best_pair"]
                    error = res.get("error")
                    matched_text = res.get("matched_text", "")

                    if error:
                        results["errors"].append((file, error))
                    elif score < Config.NO_MATCH_THRESHOLD:
                        results["no_match"].append((file, pair, score, matched_text))
                    elif score < Config.PERFECT_MATCH_THRESHOLD:
                        results["partial_match"].append((file, pair, score, matched_text))
                    else:
                        results["perfect_match"].append((file, pair, score, matched_text))
                except Exception as e:
                    results["errors"].append((file, f"Processing error: {str(e)}"))

                completed += 1
                self.safe_update_gui(
                    status=f"Processing: {completed}/{total_files}",
                    detail=f"Current file: {file}",
                    progress=completed
                )

        # Save PDF report
        default_name = f"PDF_Name_Duplicate_Report_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.pdf"
        save_path = filedialog.asksaveasfilename(
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save PDF report as..."
        )
        if save_path:
            self.generate_enhanced_pdf_report(results, save_path)
            self.safe_update_gui(
                status="Processing complete",
                detail=f"Report generated: {save_path}"
            )
            # Auto-close after 1.5s
            self.root.after(1500, self.root.destroy)

    # ------------------------------
    # PDF Report Generation
    # ------------------------------
    def generate_enhanced_pdf_report(self, results: Dict, save_path: str):
        """
        Generates a PDF report with all results and color-coded matches
        """
        doc = SimpleDocTemplate(save_path, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        elements = []

        styles = getSampleStyleSheet()
        style_title = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=20, leading=24)
        style_h2 = ParagraphStyle('Heading2', parent=styles['Heading2'], fontSize=16, leading=20)
        style_normal = ParagraphStyle('Normal', parent=styles['Normal'], fontSize=12, leading=16)

        elements.append(Paragraph("PDF Filename Match Report", style_title))
        elements.append(Spacer(1, 12))

        # Summary
        total_processed = sum(len(x) for x in results.values())
        total_issues = len(results['no_match']) + len(results['partial_match']) + len(results['errors'])
        summary_text = (
            f"Total files processed: {total_processed}<br/>"
            f"Perfect matches: {len(results['perfect_match'])}<br/>"
            f"Total files with issues: {total_issues}<br/>"
            f"&nbsp;&nbsp;- No matches: {len(results['no_match'])}<br/>"
            f"&nbsp;&nbsp;- Partial matches: {len(results['partial_match'])}<br/>"
            f"&nbsp;&nbsp;- Errors: {len(results['errors'])}"
        )
        elements.append(Paragraph(summary_text, style_normal))
        elements.append(Spacer(1, 12))

        # Internal helper to build tables
        def create_section(title, items):
            elements.append(PageBreak())
            elements.append(Paragraph(f"{title} ({len(items)})", style_h2))
            elements.append(Spacer(1, 6))

            if not items:
                elements.append(Paragraph("None", style_normal))
                elements.append(Spacer(1, 12))
                return

            data = []
            if title == "Errors":
                data.append([Paragraph("Filename", style_normal), Paragraph("Error", style_normal)])
                for file, error in items:
                    data.append([Paragraph(file, style_normal), Paragraph(error, style_normal)])
            else:
                data.append([
                    Paragraph("Filename", style_normal),
                    Paragraph("Name Detected", style_normal),
                    Paragraph("Score (%)", style_normal),
                    Paragraph("Best Filename", style_normal),
                    Paragraph("Closest Text", style_normal)
                ])
                for item in items:
                    file, pair, score, matched_text = item
                    pair_text = " ".join(p for p in pair if p) if pair else "N/A"
                    matched_text_display = matched_text if len(matched_text) <= 500 else matched_text[:497] + "..."
                    data.append([
                        Paragraph(file, style_normal),
                        Paragraph(pair_text, style_normal),
                        Paragraph(f"{score:.0f}", style_normal),
                        Paragraph(pair_text, style_normal),
                        Paragraph(matched_text_display, style_normal)
                    ])

            col_widths = [120, 120, 60, 120, 120]
            t = Table(data, colWidths=col_widths)
            t.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey)
            ]))
            elements.append(t)
            elements.append(Spacer(1, 12))

        create_section("Errors", results['errors'])
        create_section("No Match", results['no_match'])
        create_section("Partial Match", results['partial_match'])
        doc.build(elements)

# ------------------------------
# Run GUI
# ------------------------------
if __name__ == "__main__":
    app = PDFScannerGUI()
    app.root.mainloop()
