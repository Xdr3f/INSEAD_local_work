import os
import string
import re
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
import traceback
import threading
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilenames, askdirectory

# -------------------------
# Utilities
# -------------------------
def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars)

def page_verification(text):
    required_phrases = ["dear", "chief people officer"]
    return any(phrase in text.lower() for phrase in required_phrases)

def extract_name(text, max_lines_above=8):
    first_name = None
    last_name = None
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    dear_idx = None
    for idx, line in enumerate(lines):
        dear_match = re.search(
            r"Dear[,]?\s+([A-Za-zÀ-ÖØ-öø-ÿ\-']+(?:\s+[A-Za-zÀ-ÖØ-öø-ÿ\-']+)?)",
            line, re.IGNORECASE
        )
        if dear_match:
            first_name = dear_match.group(1).strip()
            dear_idx = idx
            break
    if not first_name:
        return None, None
    start_idx = max(0, dear_idx - max_lines_above)
    candidate_lines = lines[start_idx:dear_idx][::-1]
    candidate_lines.append(lines[dear_idx])
    for line in candidate_lines:
        pattern = re.compile(
            rf"{re.escape(first_name)}\s+([A-Za-zÀ-ÖØ-öø-ÿ\-']+(?:\s+[A-Za-zÀ-ÖØ-öø-ÿ\-']+){{0,2}})"
        )
        match = pattern.search(line)
        if match:
            last_name = match.group(1).strip()
            break
    return first_name, last_name

# -------------------------
# GUI
# -------------------------
class PDFSplitterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF Splitter by Person Number")
        self.setup_gui()

    def setup_gui(self):
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.select_button = ttk.Button(self.main_frame, text="Select PDFs", command=self.start_processing)
        self.select_button.grid(row=0, column=0, pady=5)

        self.progress_frame = ttk.LabelFrame(self.main_frame, text="Progress", padding=5)
        self.progress_frame.grid(row=1, column=0, pady=5, sticky=(tk.W, tk.E))

        self.progress = ttk.Progressbar(self.progress_frame, length=400, mode='determinate')
        self.progress.grid(row=0, column=0, pady=5)

        self.status_var = tk.StringVar()
        self.detail_var = tk.StringVar()
        ttk.Label(self.progress_frame, textvariable=self.status_var).grid(row=1, column=0, pady=2)
        ttk.Label(self.progress_frame, textvariable=self.detail_var).grid(row=2, column=0, pady=2)

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

    def start_processing(self):
        # Hide the button to simplify UI
        self.select_button.grid_forget()

        # Run processing in a thread to avoid freezing GUI
        thread = threading.Thread(target=self.process_pdfs)
        thread.start()

    def process_pdfs(self):
        try:
            # Ask for PDFs
            input_pdfs = askopenfilenames(title="Select PDF files", filetypes=[("PDF Files", "*.pdf")])
            if not input_pdfs:
                self.safe_update_gui(status="No files selected. Exiting.")
                return

            # Ask for output folder
            save_folder = askdirectory(title="Select folder to save PDFs")
            if not save_folder:
                self.safe_update_gui(status="No folder selected. Exiting.")
                return

            output_subfolder_name = f"Splitted_PDFs_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
            output_subfolder = os.path.join(save_folder, output_subfolder_name)
            os.makedirs(output_subfolder, exist_ok=True)

            total_files = len(input_pdfs)
            all_saved_files = []
            current_prefix = datetime.now().strftime("%Y-%m")

            for idx, input_pdf in enumerate(input_pdfs, 1):
                reader = PdfReader(input_pdf)
                saved_files = []
                writer = None
                current_filename = None

                for i, page in enumerate(reader.pages):
                    text = page.extract_text() or ""
                    pn_match = re.search(r"Person Number[:\s]+(\d+)", text)
                    verified = page_verification(text)

                    if pn_match or verified:
                        if writer and current_filename:
                            with open(current_filename, "wb") as f_out:
                                writer.write(f_out)
                            saved_files.append(current_filename)

                        first_name, last_name = extract_name(text)
                        if not first_name:
                            first_name = "UNKNOWN"
                        if not last_name:
                            last_name = "UNKNOWN"

                        if pn_match:
                            person_number = sanitize_filename(pn_match.group(1))
                        else:
                            person_number = f"PN_NotFound_page_{i+1:04d}"

                        base_name = f"{current_prefix}_Salary Review_{last_name.upper()}_{first_name}_{person_number}.pdf"
                        current_filename = os.path.join(output_subfolder, sanitize_filename(base_name))
                        writer = PdfWriter()
                        writer.add_page(page)

                    else:
                        if writer:
                            writer.add_page(page)
                        else:
                            fallback_name = f"{current_prefix}_Salary Review_UNKNOWN_UNKNOWN_PN_NotFound_page_{i+1:04d}.pdf"
                            current_filename = os.path.join(output_subfolder, sanitize_filename(fallback_name))
                            writer = PdfWriter()
                            writer.add_page(page)

                if writer and current_filename:
                    with open(current_filename, "wb") as f_out:
                        writer.write(f_out)
                    saved_files.append(current_filename)

                all_saved_files.extend(saved_files)
                self.safe_update_gui(
                    status=f"Processing {idx}/{total_files}",
                    detail=f"Current file: {os.path.basename(input_pdf)}",
                    progress=(idx/total_files)*100
                )

            self.safe_update_gui(
                status=f"Done! {len(all_saved_files)} PDFs saved.",
                detail=f"Output folder: {output_subfolder}",
                progress=100
            )

        except Exception as e:
            tb_str = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            self.safe_update_gui(status="An unexpected error occurred. See console.")
            print("An unexpected error occurred:\n", tb_str)

if __name__ == "__main__":
    app = PDFSplitterGUI()
    app.root.mainloop()
