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
import threading

# --------- Logging ---------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --------- Configuration ---------
class Config:
    NO_MATCH_THRESHOLD = 30
    PERFECT_MATCH_THRESHOLD = 100
    MAX_WORKERS = 4
    CHUNK_SIZE = 2000  # Process PDFs in chunks to save memory

# --------- Utilities ---------
def extract_text_from_pdf(pdf_path: str) -> Tuple[str, Optional[str]]:
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
    text = (text or "").lower()
    if preserve_accents:
        text = unicodedata.normalize("NFKC", text)
    else:
        text = unicodedata.normalize("NFKD", text).encode('ascii', 'ignore').decode('ascii')
    text = re.sub(r'[^a-z0-9\s-]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def extract_name_tokens(filename: str) -> List[Tuple[str, str]]:
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

def check_name_in_pdf(pdf_path: str) -> Dict:
    pdf_text, error = extract_text_from_pdf(pdf_path)
    if error:
        return {"best_pair": None, "best_score": 0, "matched_text": "", "error": error}

    pdf_no_accents = normalize_text(pdf_text, preserve_accents=False)
    pdf_with_accents = normalize_text(pdf_text, preserve_accents=True)
    tokens = extract_name_tokens(pdf_path)

    if not tokens:
        return {"best_pair": None, "best_score": 0, "matched_text": "", "error": "No valid name tokens found"}

    best_pair = None
    best_score = 0
    best_matched_text = ""

    # Single token names
    if len(tokens) == 1:
        no_acc, with_acc = tokens[0]
        for text in [pdf_no_accents, pdf_with_accents]:
            for name in [no_acc, with_acc]:
                for chunk_start in range(0, len(text), Config.CHUNK_SIZE):
                    chunk = text[chunk_start:chunk_start + Config.CHUNK_SIZE]
                    score = fuzz.partial_ratio(name, chunk)
                    if score > best_score:
                        best_score = score
                        best_pair = (name, "")
                        best_matched_text = chunk[max(0, chunk.find(name)-20):chunk.find(name)+len(name)+20]

    # All token pairs
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
                        for chunk_start in range(0, len(text), Config.CHUNK_SIZE):
                            chunk = text[chunk_start:chunk_start + Config.CHUNK_SIZE]
                            score = fuzz.partial_ratio(name, chunk)
                            if score > best_score:
                                best_score = score
                                best_pair = (first_with_acc, last_with_acc)
                                start = max(0, chunk.find(name) - 20)
                                end = min(len(chunk), chunk.find(name) + len(name) + 20)
                                best_matched_text = chunk[start:end]

    return {
        "best_pair": best_pair,
        "best_score": best_score,
        "matched_text": best_matched_text,
        "error": None
    }

# --------- GUI Class ---------
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

        results = {"no_match": [], "partial_match": [], "errors": []}
        self.progress["maximum"] = total_files
        self.progress["value"] = 0

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
                except Exception as e:
                    results["errors"].append((file, f"Processing error: {str(e)}"))
                completed += 1
                self.safe_update_gui(
                    status=f"Processing: {completed}/{total_files}",
                    detail=f"Current file: {file}",
                    progress=completed
                )

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                 filetypes=[("PDF files", "*.pdf")],
                                                 title="Save PDF report as...")
        if save_path:
            self.generate_enhanced_pdf_report(results, save_path)
            self.safe_update_gui(status="Processing complete", detail=f"Report generated: {save_path}")
        # Close the GUI automatically after 3 seconds
        self.root.after(3000, self.root.destroy)

    def generate_enhanced_pdf_report(self, results: Dict, save_path: str):
        c = canvas.Canvas(save_path, pagesize=A4)
        width, height = A4
        margin_x = 50
        y = height - 50

        c.setFont("Helvetica-Bold", 16)
        c.drawString(margin_x, y, "PDF Filename Match Report")
        y -= 30

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
                    text = f"- {file} — {pair_text} — {score:.0f}% — Snippet: '{matched_text}'"

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
