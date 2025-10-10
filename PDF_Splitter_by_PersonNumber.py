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
    Looks for key phrases like "Dear", "Cher" or "Chère".
    """
    required_phrases = ["dear", "cher", "chère"]
    return any(phrase in text.lower() for phrase in required_phrases)


def find_number_below_name(text, first_names, last_name, max_lines_below=3):
    """
    Try to find a numeric identifier (person number) in lines immediately below the name.
    Example:
        Full Name
        PN : 104
        Dear John
    Returns: the number as string, or None if not found.
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not first_names or not last_name:
        return None

    # Build a regex that matches the entire full name (first names + last name)
    name_pattern = re.compile(
        r"\b" + r"\s+".join(map(re.escape, first_names)) + r"\s+" + re.escape(last_name) + r"\b",
        re.IGNORECASE
    )

    for idx, line in enumerate(lines):
        if name_pattern.search(line):
            # Check the next few lines after the name for a number
            for offset in range(1, max_lines_below + 1):
                if idx + offset < len(lines):
                    candidate = lines[idx + offset]
                    # Match whole line if it’s only digits
                    if re.fullmatch(r"\d+", candidate):
                        return candidate
    return None


def extract_name(text):
    """
    Extract first name(s) and last name from page text.
    
    Process:
    1. Identify the greeting line starting with "Dear", "Cher", or "Chère".
    2. Extract all first names from the greeting, stripping punctuation.
    3. Search all lines above the greeting to find a line containing all first names.
       - Everything after the last first name in that line is assumed to be the last name.
    4. If no line contains the full name, the last name defaults to 'UNKNOWN'.
    
    Returns:
        first_names_str (str): All extracted first names joined by space.
        last_name (str): Extracted last name, or 'UNKNOWN' if not found.
    """
    first_names = []
    last_name = "UNKNOWN"
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    dear_idx = None

    # Find greeting line
    for idx, line in enumerate(lines):
        match = re.match(r"^(Dear|Cher|Chère)[,]?\s+(.*)", line, flags=re.I)
        if match:
            dear_idx = idx
            # Extract all first names from greeting, stripping punctuation
            first_names = [fn.strip(string.punctuation) for fn in re.split(r"\s+", match.group(2).strip())]
            break

    if dear_idx is None or not first_names:
        return None, None

    # Search all previous lines for full name containing all first names
    for line in reversed(lines[:dear_idx]):
        found_all = all(re.search(rf"\b{re.escape(fn)}\b", line) for fn in first_names)
        if found_all:
            temp = line
            for fn in first_names:
                temp = re.sub(rf"\b{re.escape(fn)}\b", "", temp, count=1).strip()
            if temp:
                last_name = temp
            break

    return " ".join(first_names), last_name




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
            timestamp_str = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            output_subfolder_name = f"Splitted_PDFs_{timestamp_str}"
            output_subfolder = os.path.join(save_folder, output_subfolder_name)
            os.makedirs(output_subfolder, exist_ok=True)

            total_files = len(input_pdfs)
            all_saved_files = []
            current_prefix = datetime.now().strftime("%Y-%m")

            # Collect information for summary GUI
            missing_cpo_list = []
            unknown_name_list = []

            for idx, input_pdf in enumerate(input_pdfs, 1):
                reader = PdfReader(input_pdf)
                saved_files = []
                writer = None
                current_filename = None
                cpo_found = False
                person_number_current = "UNKNOWN"
                first_name_current = "UNKNOWN"
                last_name_current = "UNKNOWN"

                for i, page in enumerate(reader.pages):
                    text = page.extract_text() or ""
                    has_dear = bool(re.search(r"^(Dear|Cher|Chère)[,]?\s+", text, flags=re.I))
                    pn_match = re.search(r"PN[:\s]+(\d+)", text)
                    has_cpo = "chief people officer" in text.lower()

                    # Decide if we need a new PDF
                    if pn_match or has_dear:
                        # Save previous PDF
                        if writer and current_filename:
                            with open(current_filename, "wb") as f_out:
                                writer.write(f_out)
                            saved_files.append(current_filename)

                            # Track missing CPO
                            if not cpo_found:
                                missing_cpo_list.append({
                                    "Filename": os.path.basename(current_filename),
                                    "First Name": first_name_current,
                                    "Last Name": last_name_current
                                })

                            # Track UNKNOWN names
                            if "UNKNOWN" in first_name_current or "UNKNOWN" in last_name_current:
                                unknown_name_list.append({
                                    "Filename": os.path.basename(current_filename),
                                    "First Name": first_name_current,
                                    "Last Name": last_name_current
                                })

                        # Extract name
                        first_name_str, last_name = extract_name(text)
                        first_name_current = first_name_str if first_name_str else "UNKNOWN"
                        last_name_current = last_name if last_name else "UNKNOWN"

                        # Get Person Number
                        if pn_match:
                            person_number_current = sanitize_filename(pn_match.group(1))
                        else:
                            candidate_pn = find_number_below_name(text, first_name_current.split(), last_name_current)
                            if candidate_pn:
                                person_number_current = sanitize_filename(candidate_pn)
                            else:
                                person_number_current = f"PN_NotFound_page_{i+1:04d}"

                        # Build filename
                        base_name = f"{current_prefix}_Salary Review_{last_name_current.upper()}_{first_name_current}_{person_number_current}.pdf"
                        current_filename = os.path.join(output_subfolder, sanitize_filename(base_name))

                        # Start new PDF
                        writer = PdfWriter()
                        writer.add_page(page)
                        cpo_found = has_cpo

                    else:
                        # Continuation page
                        if writer:
                            writer.add_page(page)
                            if has_cpo:
                                cpo_found = True
                        else:
                            # Fallback if first page doesn't have PN/Dear
                            fallback_name = f"{current_prefix}_Salary Review_UNKNOWN_UNKNOWN_PN_NotFound_page_{i+1:04d}.pdf"
                            current_filename = os.path.join(output_subfolder, sanitize_filename(fallback_name))
                            writer = PdfWriter()
                            writer.add_page(page)
                            if has_cpo:
                                cpo_found = True

                # Save last PDF
                if writer and current_filename:
                    with open(current_filename, "wb") as f_out:
                        writer.write(f_out)
                    saved_files.append(current_filename)

                    # Track missing CPO
                    if not cpo_found:
                        missing_cpo_list.append({
                            "Filename": os.path.basename(current_filename),
                            "First Name": first_name_current,
                            "Last Name": last_name_current
                        })

                    # Track UNKNOWN names
                    if "UNKNOWN" in first_name_current or "UNKNOWN" in last_name_current:
                        unknown_name_list.append({
                            "Filename": os.path.basename(current_filename),
                            "First Name": first_name_current,
                            "Last Name": last_name_current
                        })

                all_saved_files.extend(saved_files)

                # Update GUI progress
                self.safe_update_gui(
                    status=f"Processing {idx}/{total_files}",
                    detail=f"Current file: {os.path.basename(input_pdf)}",
                    progress=(idx / total_files) * 100
                )

            # Display summary GUI
            summary_window = tk.Toplevel(self.root)
            summary_window.title("PDF Summary Report")

            summary_text = tk.Text(summary_window, width=100, height=30)
            summary_text.pack(padx=10, pady=10)

            summary_text.insert(tk.END, f"Total PDFs processed: {len(all_saved_files)}\n\n")
            summary_text.insert(tk.END, f"PDFs missing CPO: {len(missing_cpo_list)}\n")
            for item in missing_cpo_list:
                summary_text.insert(tk.END, f"  - {item['Filename']} ({item['First Name']} {item['Last Name']})\n")

            summary_text.insert(tk.END, f"\nPDFs with UNKNOWN names: {len(unknown_name_list)}\n")
            for item in unknown_name_list:
                summary_text.insert(tk.END, f"  - {item['Filename']} ({item['First Name']} {item['Last Name']})\n")

            summary_text.config(state=tk.DISABLED)

            # Close button
            close_button = ttk.Button(summary_window, text="Close", command=summary_window.destroy)
            close_button.pack(pady=5)

            # Final status
            self.safe_update_gui(
                status=f"Done! {len(all_saved_files)} PDFs saved.",
                detail=f"Output folder: {output_subfolder}",
                progress=100
            )

        except Exception as e:
            tb_str = "".join(traceback.format_exception(type(e), e, e.__traceback__))
            self.safe_update_gui(status="An unexpected error occurred. See console.")
            print("An unexpected error occurred:\n", tb_str)


# -------------------------
# Run Application
# -------------------------
if __name__ == "__main__":
    app = PDFSplitterGUI()
    app.root.mainloop()
