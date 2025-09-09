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
from tkinter import filedialog, ttk
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Tuple, Optional
from functools import partial
import threading

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
    """Normalize text for comparison with configurable accent handling"""
    text = (text or "").lower()
    
    if preserve_accents:
        text = unicodedata.normalize("NFKC", text)
    else:
        text = unicodedata.normalize("NFKD", text).encode('ascii', 'ignore').decode('ascii')
    
    text = re.sub(r'[^a-z0-9\s-]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
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

def find_exact_word_match(word: str, text: str) -> Tuple[bool, str, List[str]]:
    """Find exact word match ensuring the entire word/phrase matches exactly.
    Also finds all possible name variations in the surrounding context.
    
    Args:
        word: The word/phrase to search for
        text: The text to search in
    
    Returns:
        Tuple of (found, matched_text, all_possible_matches)
        Only returns True if the entire word/phrase matches exactly AND
        there are no longer variations of the name in context
    """
    # Normalize both word and text
    norm_word = normalize_text(word)
    norm_text = normalize_text(text)
    
    # Create pattern that matches the exact word/phrase with word boundaries
    pattern = r'\b' + re.escape(norm_word) + r'\b'
    matches = list(re.finditer(pattern, norm_text, re.IGNORECASE))
    
    if matches:
        all_context_matches = set()
        
        for match in matches:
            # Get context around the match (50 chars before and after)
            start = max(0, match.start() - 50)
            end = min(len(norm_text), match.end() + 50)
            context = norm_text[start:end]
            
            # Find all possible word combinations in context that could be names
            # This helps detect if our match is part of a longer name
            context_pattern = r'\b[A-Za-z]+(?:\s+[A-Za-z]+){0,3}\b'
            
            for m in re.finditer(context_pattern, context, re.IGNORECASE):
                context_text = m.group().strip()
                if context_text:  # Ignore empty matches
                    # Get original text for this match from the un-normalized text
                    orig_start = start + m.start()
                    orig_end = start + m.end()
                    original_text = text[orig_start:orig_end].strip()
                    all_context_matches.add(original_text)
            
            # Get original text for the exact match
            match_start = match.start()
            match_end = match.end()
            matched_original = text[match_start:match_end].strip()
            
            # Only return true if:
            # 1. The normalized matched text exactly equals our search word
            # 2. It's not part of a longer name
            is_part_of_longer_name = any(
                (normalize_text(matched_original) != normalize_text(other_text) and 
                 normalize_text(matched_original) in normalize_text(other_text))
                for other_text in all_context_matches
            )
            
            if normalize_text(matched_original) == normalize_text(word) and not is_part_of_longer_name:
                return True, matched_original, list(all_context_matches)
    
    return False, "", []

def check_name_in_pdf(pdf_path: str) -> Dict:
    pdf_text, error = extract_text_from_pdf(pdf_path)
    if error:
        return {"best_pair": None, "best_score": 0, "error": error}

    pdf_no_accents = normalize_text(pdf_text, preserve_accents=False)
    pdf_with_accents = normalize_text(pdf_text, preserve_accents=True)
    
    tokens = extract_name_tokens(pdf_path)
    
    if not tokens:
        return {"best_pair": None, "best_score": 0, "error": "No valid name tokens found"}

    best_pair = None
    best_score = 0
    best_matched_text = ""
    perfect_match_found = False

    if len(tokens) == 1:
        no_acc, with_acc = tokens[0]
        for text in [pdf_no_accents, pdf_with_accents]:
            for name in [no_acc, with_acc]:
                found, full_match, context_matches = find_exact_word_match(name, text)
                if found:
                    # For perfect match, require:
                    # 1. Exact equality after normalization
                    # 2. No other variations of the name exist in context
                    norm_name = normalize_text(name)
                    
                    # Check all possible matches in context to ensure no variations exist
                    has_variations = any(
                        normalize_text(other_match) != norm_name 
                        and norm_name in normalize_text(other_match)
                        for other_match in context_matches
                    )
                    
                    if not has_variations and normalize_text(full_match) == norm_name:
                        perfect_match_found = True
                        best_score = 100
                        best_pair = (name, "")
                        best_matched_text = full_match
                        break
                
                for chunk_start in range(0, len(text), Config.CHUNK_SIZE):
                    chunk = text[chunk_start:chunk_start + Config.CHUNK_SIZE]
                    ratio = fuzz.ratio(name, chunk)
                    partial = fuzz.partial_ratio(name, chunk)
                    score = max(ratio, partial)
                    if score > best_score:
                        best_score = score
                        best_pair = (name, "")
                        best_matched_text = name
            if perfect_match_found:
                break

    if not perfect_match_found:
        for i in range(len(tokens)):
            for j in range(i + 1, len(tokens)):
                first_no_acc, first_with_acc = tokens[i]
                last_no_acc, last_with_acc = tokens[j]
                
                combinations = [
                    (f"{first_no_acc} {last_no_acc}", f"{first_with_acc} {last_with_acc}"),
                    (f"{last_no_acc} {first_no_acc}", f"{last_with_acc} {first_with_acc}"),
                    (f"{first_no_acc}, {last_no_acc}", f"{first_with_acc}, {last_with_acc}")
                ]
                
                for no_acc_combined, with_acc_combined in combinations:
                    for text in [pdf_no_accents, pdf_with_accents]:
                        for name in [no_acc_combined, with_acc_combined]:
                            found, full_match = find_exact_word_match(name, text)
                            if found:
                                # For perfect match, require exact equality
                                if normalize_text(full_match) == normalize_text(name):
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
                                if score > best_score:
                                    best_score = score
                                    best_pair = (first_with_acc, last_with_acc)
                                    best_matched_text = name
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

        save_path = filedialog.asksaveasfilename(
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
        c = canvas.Canvas(save_path, pagesize=A4)
        width, height = A4
        margin_x = 50
        y = height - 50

        c.setFont("Helvetica-Bold", 16)
        c.drawString(margin_x, y, "PDF Filename Match Report")
        y -= 30

        c.setFont("Helvetica-Bold", 12)
        total_processed = sum(len(x) for x in results.values())
        total_issues = len(results['no_match']) + len(results['partial_match']) + len(results['errors'])
        
        summary_text = (
            f"Total files processed: {total_processed}\n"
            f"Perfect matches: {len(results['perfect_match'])}\n"
            f"Total files with issues: {total_issues}\n"
            f"  - No matches: {len(results['no_match'])}\n"
            f"  - Partial matches: {len(results['partial_match'])}\n"
            f"  - Errors: {len(results['errors'])}"
        )
        for line in summary_text.split('\n'):
            c.drawString(margin_x, y, line)
            y -= 20

        sections = [
            ("Errors", results["errors"]),
            ("No Match", results["no_match"]),
            ("Partial Match", results["partial_match"])
        ]

        for title, items in sections:
            if y < 100:
                c.showPage()
                y = height - 50

            y -= 20
            c.setFont("Helvetica-Bold", 12)
            c.drawString(margin_x, y, f"{title} ({len(items)})")
            y -= 15

            if not items:
                c.setFont("Helvetica", 10)
                c.drawString(margin_x + 10, y, "None")
                y -= 15
                continue

            c.setFont("Helvetica", 10)
            for item in items:
                if isinstance(item, tuple) and len(item) == 2:
                    file, error = item
                    text = f"- {file}: {error}"
                else:
                    file, pair, score, matched_text = item
                    pair_text = " ".join(p for p in pair if p) if pair else "N/A"
                    text = f"- {file} — {pair_text} — {score:.0f}% — Best match: '{matched_text}'"

                words = text.split()
                line = ""
                for word in words:
                    test_line = line + " " + word if line else word
                    if c.stringWidth(test_line, "Helvetica", 10) < width - 2*margin_x:
                        line = test_line
                    else:
                        c.drawString(margin_x + 10, y, line)
                        y -= 12
                        line = word
                    if y < 60:
                        c.showPage()
                        y = height - 50
                if line:
                    c.drawString(margin_x + 10, y, line)
                    y -= 12

        c.save()

if __name__ == "__main__":
    app = PDFScannerGUI()
    app.root.mainloop()
