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
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, List, Tuple, Optional

# Set up logging
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(levelname)s - %(message)s')

class Config:
    """Configuration class for easily adjustable parameters"""
    NO_MATCH_THRESHOLD = 30
    PERFECT_MATCH_THRESHOLD = 100
    MAX_WORKERS = 4  # For parallel processing
    CHUNK_SIZE = 1000  # Number of characters to process at a time for large PDFs

def extract_text_from_pdf(pdf_path: str) -> Tuple[str, Optional[str]]:
    """Extract text from PDF with error handling"""
    text_content = ""
    error_message = None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                try:
                    page_text = page.extract_text()
                    if page_text:
                        text_content += " " + page_text
                except Exception as e:
                    error_message = f"Error on page {page.page_number}: {str(e)}"
                    logging.warning(error_message)
    except Exception as e:
        error_message = f"Error opening PDF: {str(e)}"
        logging.error(error_message)
    return text_content, error_message

def normalize_text(text: str) -> str:
    """Normalize text for comparison"""
    text = (text or "").lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if ch.isalpha() or ch.isspace())
    text = re.sub(r"\\s+", " ", text).strip()
    return text

def extract_name_tokens(filename: str) -> List[str]:
    """Extract potential name tokens from filename with multiple separators"""
    base = os.path.splitext(os.path.basename(filename))[0]
    # Handle multiple possible separators
    for sep in ['_', '-', ' ']:
        base = base.replace(sep, '_')
    parts = base.split('_')
    return [normalize_text(p) for p in parts if normalize_text(p) and normalize_text(p).isalpha()]

def check_name_in_pdf(pdf_path: str) -> Dict:
    """Check for name matches in PDF content"""
    pdf_text, error = extract_text_from_pdf(pdf_path)
    if error:
        return {"best_pair": None, "best_score": 0, "error": error}

    normalized_pdf = normalize_text(pdf_text)
    tokens = extract_name_tokens(pdf_path)
    
    if len(tokens) < 2:
        return {"best_pair": None, "best_score": 0, "error": "Not enough name tokens found"}

    best_pair = None
    best_score = 0

    # Check all possible pairs, not just consecutive
    for i in range(len(tokens)):
        for j in range(i + 1, len(tokens)):
            first, last = tokens[i], tokens[j]
            # Try different combinations
            combinations = [
                f"{first} {last}",
                f"{last} {first}",
                f"{first}, {last}"
            ]
            for combined in combinations:
                score = fuzz.partial_ratio(combined, normalized_pdf)
                if score > best_score:
                    best_score = score
                    best_pair = (first, last)

    return {"best_pair": best_pair, "best_score": best_score, "error": None}

class PDFScannerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF Name Scanner")
        self.setup_gui()

    def setup_gui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Buttons
        ttk.Button(main_frame, text="Select Folder", command=self.process_folder).grid(row=0, column=0, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, length=300, mode='determinate')
        self.progress.grid(row=1, column=0, pady=5)
        
        # Status label
        self.status_var = tk.StringVar()
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=2, column=0, pady=5)

    def process_folder(self):
        folder_path = filedialog.askdirectory(title="Select the folder containing PDFs")
        if not folder_path:
            return

        # Get PDF files
        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
        total_files = len(pdf_files)
        
        if total_files == 0:
            self.status_var.set("No PDF files found in selected folder")
            return

        results = {"no_match": [], "partial_match": [], "errors": []}
        
        # Setup progress bar
        self.progress["maximum"] = total_files
        self.progress["value"] = 0
        
        # Process files with ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=Config.MAX_WORKERS) as executor:
            futures = []
            for file in pdf_files:
                full_path = os.path.join(folder_path, file)
                future = executor.submit(check_name_in_pdf, full_path)
                futures.append((file, future))
            
            # Process results as they complete
            for file, future in futures:
                res = future.result()
                score = res["best_score"]
                pair = res["best_pair"]
                error = res.get("error")

                if error:
                    results["errors"].append((file, error))
                elif score < Config.NO_MATCH_THRESHOLD:
                    results["no_match"].append((file, pair, score))
                elif score < Config.PERFECT_MATCH_THRESHOLD:
                    results["partial_match"].append((file, pair, score))
                
                self.progress["value"] += 1
                self.status_var.set(f"Processing: {self.progress['value']}/{total_files}")
                self.root.update()

        # Generate report
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save PDF report as..."
        )
        
        if save_path:
            self.generate_enhanced_pdf_report(results, save_path)
            self.status_var.set(f"Report generated: {save_path}")

    def generate_enhanced_pdf_report(self, results: Dict, save_path: str):
        """Generate an enhanced PDF report with more detailed information"""
        c = canvas.Canvas(save_path, pagesize=A4)
        width, height = A4
        margin_x = 50
        y = height - 50

        # Title
        c.setFont("Helvetica-Bold", 16)
        c.drawString(margin_x, y, "PDF Filename Match Report")
        y -= 30

        # Summary
        c.setFont("Helvetica-Bold", 12)
        summary_text = (
            f"Total files processed: {sum(len(x) for x in results.values())}\n"
            f"No matches: {len(results['no_match'])}\n"
            f"Partial matches: {len(results['partial_match'])}\n"
            f"Errors: {len(results['errors'])}"
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
            if y < 100:  # New page if needed
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
                if isinstance(item, tuple) and len(item) == 2:  # Error entry
                    file, error = item
                    text = f"- {file}: {error}"
                else:  # Match entry
                    file, pair, score = item
                    pair_text = " ".join(pair) if pair else "N/A"
                    text = f"- {file} — {pair_text} — {score:.0f}%"

                # Handle long lines
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
