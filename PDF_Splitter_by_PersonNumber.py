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
    """
    Remove/replace invalid characters so filenames are safe across OS.
    Keeps letters, digits, spaces, dashes, underscores, dots, and parentheses.
    """
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars)


def page_verification(text):
    """
    Lightweight heuristic to confirm if a page starts a new employee letter.
    Looks for key phrases like "Dear" or "Chief People Officer".
    """
    required_phrases = ["dear", "chief people officer"]
    return any(phrase in text.lower() for phrase in required_phrases)


def find_number_below_name(text, first_name, last_name, max_lines_below=2):
    """
    Try to find a numeric identifier (person number) in lines immediately below the name.
    Example:
        Dear John,
        John Doe
        104
        Human Resources
    Returns: the number as string, or None if not found.
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not first_name or not last_name:
        return None

    # Regex: look for "Firstname ... Lastname" on the same line
    name_pattern = re.compile(rf"{re.escape(first_name)}.*{re.escape(last_name)}", re.IGNORECASE)

    for idx, line in enumerate(lines):
        if name_pattern.search(line):
            # Check the next few lines after the name for a number
            for offset in range(1, max_lines_below + 1):
                if idx + offset < len(lines):
                    candidate = lines[idx + offset]
                    # Match whole line if it’s only digits
                    number_match = re.match(r"^\d{1,}$", candidate)
                    if number_match:
                        return number_match.group(0)
    return None


def extract_name(text, max_lines_above=8):
    """
    Extract first name and last name from page text.
    Logic:
    - Find the line starting with "Dear ..."
    - Capture the first name from that greeting
    - Look up a few lines above 'Dear' to find a last name that matches

    Returns: (first_name, last_name) or (None, None) if not found
    """
    first_name = None
    last_name = None
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    dear_idx = None

    # Find "Dear <firstname>"
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

    # Search upwards (within N lines above "Dear") for last name
    start_idx = max(0, dear_idx - max_lines_above)
    candidate_lines = lines[start_idx:dear_idx][::-1]  # reverse order
    candidate_lines.append(lines[dear_idx])  # also check the "Dear ..." line

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
# GUI Application
# -------------------------

class PDFSplitterGUI:
    """
    Tkinter-based GUI for splitting PDF files into employee-specific PDFs
    using extracted names and person numbers.
    """

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF Splitter by Person Number")
        self.setup_gui()

    def setup_gui(self):
        """
        Create and layout GUI elements:
        - File select button
        - Progress bar
        - Status labels
        """
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Button to start selecting PDFs
        self.select_button = ttk.Button(self.main_frame, text="Select PDFs", command=self.start_processing)
        self.select_button.grid(row=0, column=0, pady=5)

        # Progress section
        self.progress_frame = ttk.LabelFrame(self.main_frame, text="Progress", padding=5)
        self.progress_frame.grid(row=1, column=0, pady=5, sticky=(tk.W, tk.E))

        self.progress = ttk.Progressbar(self.progress_frame, length=400, mode='determinate')
        self.progress.grid(row=0, column=0, pady=5)

        # Status + details
        self.status_var = tk.StringVar()
        self.detail_var = tk.StringVar()
        ttk.Label(self.progress_frame, textvariable=self.status_var).grid(row=1, column=0, pady=2)
        ttk.Label(self.progress_frame, textvariable=self.detail_var).grid(row=2, column=0, pady=2)

    def safe_update_gui(self, status=None, detail=None, progress=None):
        """
        Thread-safe way to update the GUI (status, detail, progress bar).
        Uses `after(0, ...)` to schedule update in Tkinter’s event loop.
        """
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
        """
        Triggered when user clicks 'Select PDFs'.
        - Hide button
        - Launch processing in a separate thread
        """
        self.select_button.grid_forget()
        thread = threading.Thread(target=self.process_pdfs)
        thread.start()

    def process_pdfs(self):
        """
        Core logic:
        - Ask user for PDFs and output folder
        - Process each PDF page by page
        - Detect when a new employee starts
        - Extract names + person number
        - Save individual employee PDFs into a timestamped folder
        - Update GUI with progress
        """
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

            # Create timestamped subfolder
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
                        # Save previous employee PDF if open
                        if writer and current_filename:
                            with open(current_filename, "wb") as f_out:
                                writer.write(f_out)
                            saved_files.append(current_filename)

                        # Extract name
                        first_name, last_name = extract_name(text)
                        if not first_name:
                            first_name = "UNKNOWN"
                        if not last_name:
                            last_name = "UNKNOWN"

                        # Get Person Number
                        if pn_match:
                            person_number = sanitize_filename(pn_match.group(1))
                        else:
                            candidate_pn = find_number_below_name(text, first_name, last_name)
                            if candidate_pn:
                                person_number = sanitize_filename(candidate_pn)
                            else:
                                person_number = f"PN_NotFound_page_{i+1:04d}"

                        # Build filename
                        base_name = f"{current_prefix}_Salary Review_{last_name.upper()}_{first_name}_{person_number}.pdf"
                        current_filename = os.path.join(output_subfolder, sanitize_filename(base_name))

                        # Start a new PDF
                        writer = PdfWriter()
                        writer.add_page(page)

                    else:
                        # Continuation page of current employee
                        if writer:
                            writer.add_page(page)
                        else:
                            # Fallback in case no employee detected yet
                            fallback_name = f"{current_prefix}_Salary Review_UNKNOWN_UNKNOWN_PN_NotFound_page_{i+1:04d}.pdf"
                            current_filename = os.path.join(output_subfolder, sanitize_filename(fallback_name))
                            writer = PdfWriter()
                            writer.add_page(page)

                # Save the last file for this input PDF
                if writer and current_filename:
                    with open(current_filename, "wb") as f_out:
                        writer.write(f_out)
                    saved_files.append(current_filename)

                all_saved_files.extend(saved_files)

                # Update GUI progress
                self.safe_update_gui(
                    status=f"Processing {idx}/{total_files}",
                    detail=f"Current file: {os.path.basename(input_pdf)}",
                    progress=(idx / total_files) * 100
                )

            # Final status
            self.safe_update_gui(
                status=f"Done! {len(all_saved_files)} PDFs saved.",
                detail=f"Output folder: {output_subfolder}",
                progress=100
            )

            # Auto-close after 2 seconds
            self.root.after(2000, self.root.destroy)

        except Exception as e:
            # Print traceback for debugging
            tb_str = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            self.safe_update_gui(status="An unexpected error occurred. See console.")
            print("An unexpected error occurred:\n", tb_str)


# -------------------------
# Run Application
# -------------------------
if __name__ == "__main__":
    app = PDFSplitterGUI()
    app.root.mainloop()
