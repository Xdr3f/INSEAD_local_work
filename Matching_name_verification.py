"""This code searches if the name of the person written in the PDF name is present in the PDF content.
It generates a PDF report listing files with no match or partial match for later review by the personnel.
If a perfect match is found, the file is ignored in the report."""

import os
import re
import unicodedata
import logging
import pdfplumber
from rapidfuzz import fuzz
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Tuple, Optional
from functools import partial
import threading
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
import datetime

import sys
import traceback

# Set up logging
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(levelname)s - %(message)s')

class Config:
    """Configuration class for easily adjustable parameters"""
    NO_MATCH_THRESHOLD = 30
    PERFECT_MATCH_THRESHOLD = 100
    MAX_WORKERS = 4  # For parallel processing
    CHUNK_SIZE = 2000  # Number of characters to process at a time for large PDFs

def extract_text_from_pdf(pdf_path: str) -> Tuple[str, Optional[str]]:
    """Extract text from PDF document"""
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
                    error_message = f"Error on page {page_num}: {str(e)}"
                    logging.warning(error_message)
                    
    except Exception as e:
        error_message = f"Error opening PDF: {str(e)}"
        logging.error(error_message)
        
    return " ".join(text_content), error_message

def normalize_text(text: str, preserve_accents: bool = False) -> str:
    """Normalize text for comparison with configurable accent handling and normalized whitespace"""
    text = (text or "").lower()
    
    if preserve_accents:
        text = unicodedata.normalize("NFKC", text)
    else:
        text = unicodedata.normalize("NFKD", text).encode('ascii', 'ignore').decode('ascii')
    
    text = re.sub(r'[^a-z0-9\s-]', ' ', text)    # Keep letters, digits, spaces, hyphens
    text = re.sub(r'\s+', ' ', text).strip()     # Normalize all whitespace to single space
    return text

def extract_name_tokens(filename: str) -> List[str]:
    """Extract potential name tokens from filename with improved handling"""
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


def is_bidirectional_match(a: str, b: str) -> bool:
    return normalize_text(a) == normalize_text(b)


def find_exact_word_match(word: str, text: str) -> Tuple[bool, str, List[str]]:
    """
    Returns True if the exact combination of words in `word` appears in `text`.
    Whitespace normalized for strict matching.
    """
    norm_word = normalize_text(word)
    norm_text = normalize_text(text)
    # Add spaces around to ensure phrase boundaries
    if f" {norm_word} " in f" {norm_text} ":
        return True, word, []
    return False, "", []

def get_best_ngram_match(name, chunk, ngram_range=(2, 5)):
    words = chunk.split()
    candidates = set()
    for n in range(ngram_range[0], ngram_range[1]+1):
        for i in range(len(words) - n + 1):
            candidates.add(" ".join(words[i:i+n]))
    # Always include the whole chunk as fallback
    candidates.add(chunk)
    from rapidfuzz import process
    match, score, _ = process.extractOne(name, candidates)
    return match if match else chunk


def check_name_in_pdf(pdf_path: str) -> Dict:
    pdf_text, error = extract_text_from_pdf(pdf_path)
    if error:
        return {"best_pair": None, "best_score": 0, "error": error}

    # ...removed debug print...

    # Normalize text
    pdf_no_accents = normalize_text(pdf_text, preserve_accents=False)
    pdf_with_accents = normalize_text(pdf_text, preserve_accents=True)

    # Normalize whitespace for strict matching
    pdf_no_accents_norm = re.sub(r'\s+', ' ', pdf_no_accents)
    pdf_with_accents_norm = re.sub(r'\s+', ' ', pdf_with_accents)

    tokens = extract_name_tokens(pdf_path)
    # ...removed debug print...

    # Keep only alphabetic tokens for names
    alpha_tokens = [(no_acc, with_acc) for no_acc, with_acc in tokens if no_acc.isalpha()]
    # ...removed debug print...

    if not alpha_tokens:
        return {"best_pair": None, "best_score": 0, "error": "No valid alphabetic name tokens found"}

    best_pair = None
    best_score = 0
    best_matched_text = ""
    perfect_match_found = False

    # Case 1: Single-token names
    if len(alpha_tokens) == 1:
        no_acc, with_acc = alpha_tokens[0]
    # ...removed debug print...
        for text in [pdf_no_accents_norm, pdf_with_accents_norm]:
            for name in [no_acc, with_acc]:
                found, full_match, context_matches = find_exact_word_match(name, text)
                # ...removed debug print...
                if found and is_bidirectional_match(full_match, name):
                    perfect_match_found = True
                    best_score = 100
                    best_pair = (name, "")
                    best_matched_text = full_match
                    break

                # Fuzzy similarity → partial match only
                for chunk_start in range(0, len(text), Config.CHUNK_SIZE):
                    chunk = text[chunk_start:chunk_start + Config.CHUNK_SIZE]
                    ratio = fuzz.ratio(name, chunk)
                    partial = fuzz.partial_ratio(name, chunk)
                    score = max(ratio, partial)
                    if not perfect_match_found and score > best_score:
                        best_score = min(score, Config.PERFECT_MATCH_THRESHOLD - 1)
                        best_pair = (name, "")
                        # Find the actual best-matching substring in the chunk
                        best_matched_text = get_best_ngram_match(name, chunk)

            if perfect_match_found:
                break

    # Case 2: Two+ tokens → try combinations
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
                            # ...removed debug print...
                            found, full_match, _ = find_exact_word_match(name, text)
                            # ...removed debug print...
                            if found and is_bidirectional_match(full_match, name):
                                perfect_match_found = True
                                best_score = 100
                                best_pair = (first_with_acc, last_with_acc)
                                best_matched_text = full_match
                                # ...removed debug print...
                                break

                            # Fuzzy similarity → partial match only
                            for chunk_start in range(0, len(text), Config.CHUNK_SIZE):
                                chunk = text[chunk_start:chunk_start + Config.CHUNK_SIZE]
                                ratio = fuzz.ratio(name, chunk)
                                partial = fuzz.partial_ratio(name, chunk)
                                score = max(ratio, partial)
                                if not perfect_match_found and score > best_score:
                                    best_score = min(score, Config.PERFECT_MATCH_THRESHOLD - 1)
                                    best_pair = (first_with_acc, last_with_acc)
                                    # Use RapidFuzz to find the closest substring in the chunk
                                    best_matched_text = get_best_ngram_match(name, chunk)
                                    # ...removed debug print...

                        if perfect_match_found:
                            break
                    if perfect_match_found:
                        break
                if perfect_match_found:
                    break
            if perfect_match_found:
                break

    # ...removed debug print...
    return {
        "best_pair": best_pair,
        "best_score": best_score,
        "matched_text": best_matched_text,
        "error": None
    }


class PDFScannerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF Name Scanner")
        self.setup_gui()
        self.processing_lock = threading.Lock()

    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Button(main_frame, text="Select Folder", command=self.process_folder).grid(row=0, column=0, pady=5)
        
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="5")
        progress_frame.grid(row=1, column=0, pady=5, sticky=(tk.W, tk.E))
        
        self.progress = ttk.Progressbar(progress_frame, length=300, mode='determinate')
        self.progress.grid(row=0, column=0, pady=5)
        
        self.status_var = tk.StringVar()
        self.detail_var = tk.StringVar()
        ttk.Label(progress_frame, textvariable=self.status_var).grid(row=1, column=0, pady=2)
        ttk.Label(progress_frame, textvariable=self.detail_var).grid(row=2, column=0, pady=2)

    def safe_update_gui(self, status=None, detail=None, progress=None):
        def update():
            if status is not None:
                self.status_var.set(status)
            if detail is not None:
                self.detail_var.set(detail)
            if progress is not None:
                self.progress["value"] = progress
            self.root.update()
        
        self.root.after(0, update)

    def process_folder(self):
        folder_path = filedialog.askdirectory(title="Select the folder containing PDFs")
        if not folder_path:
            return

        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
        total_files = len(pdf_files)
        
        if total_files == 0:
            messagebox.showerror("No PDFs Found", "No PDF files found in the selected folder.")
            self.safe_update_gui(status="No PDF files found in selected folder")
            return

        results = {
            "no_match": [], 
            "partial_match": [], 
            "perfect_match": [],
            "errors": []
        }
        
        self.progress["maximum"] = total_files
        self.progress["value"] = 0
        
        with ThreadPoolExecutor(max_workers=Config.MAX_WORKERS) as executor:
            futures = {}
            for file in pdf_files:
                full_path = os.path.join(folder_path, file)
                future = executor.submit(check_name_in_pdf, full_path)
                futures[future] = file
            
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

        # Default report filename with date
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
            # --- Auto-close window after 1.5 seconds ---
            self.root.after(1500, self.root.destroy)

    def generate_enhanced_pdf_report(self, results: Dict, save_path: str):
        doc = SimpleDocTemplate(save_path, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        elements = []
        
        # --- Styles ---
        styles = getSampleStyleSheet()
        style_title = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=20, leading=24)
        style_h2 = ParagraphStyle('Heading2', parent=styles['Heading2'], fontSize=16, leading=20)
        style_normal = ParagraphStyle('Normal', parent=styles['Normal'], fontSize=12, leading=16)

        # --- Title ---
        elements.append(Paragraph("PDF Filename Match Report", style_title))
        elements.append(Spacer(1, 12))

        # --- Summary ---
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

        # --- Function to create table sections ---
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

            # --- Flexible column widths ---
            colWidths = [120, 100, 40, 120, 180] if title != "Errors" else [200, 320]

            # --- Create Table with repeating header ---
            table = Table(data, colWidths=colWidths, repeatRows=1)
            tbl_style = TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 14),    # header font size
                ('FONTSIZE', (0,1), (-1,-1), 12),   # body font size
                ('BACKGROUND', (0,0), (-1,0), colors.grey),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ])

            # Color-code rows for non-error tables
            if title != "Errors":
                for i, row in enumerate(data[1:], start=1):
                    try:
                        score = int(row[2].text)  # extract text from Paragraph
                        if score >= 100:
                            tbl_style.add('TEXTCOLOR', (0,i), (-1,i), colors.green)
                        elif score >= 30:
                            tbl_style.add('TEXTCOLOR', (0,i), (-1,i), colors.orange)
                        else:
                            tbl_style.add('TEXTCOLOR', (0,i), (-1,i), colors.red)
                    except:
                        pass

            table.setStyle(tbl_style)
            elements.append(table)
            elements.append(Spacer(1, 12))

        # --- Build all sections ---
        create_section("Errors", results["errors"])
        create_section("No Match", results["no_match"])
        create_section("Partial Match", results["partial_match"])

        # --- Build PDF ---
        doc.build(elements)

if __name__ == "__main__":
    try:
        app = PDFScannerGUI()
        app.root.mainloop()
    except Exception as e:
        # Gather full traceback
        tb_str = "".join(traceback.format_exception(type(e), e, e.__traceback__))
        logging.error(f"Unhandled exception:\n{tb_str}")
        
        # Show error to user
        try:
            root = tk.Tk()
            root.withdraw()  # Hide main window
            messagebox.showerror(
                "Critical Error",
                f"An unexpected error occurred:\n\n{str(e)}\n\n"
                "The application will now close."
            )
            root.destroy()
        except:
            # If Tkinter itself fails, fallback to console output
            print(f"Critical Error: {e}", file=sys.stderr)
        
        sys.exit(1)  # Exit with error code