""" 
--This script allows the user to select one or more PDF files and a destination folder. 
--It then processes each PDF, searching for "Person Number" entries using a regular expression. 
--The script is designed to split or organize the PDFs based on these person numbers,
  saving the resulting files with sanitized filenames in the chosen folder.

--Note : If an error occurs with a syntax error, try running the script in a terminal instead of an IDE. If on VSCode,
  you can run it in the "dedicated terminal".
  
--Important : The script assumes that the "Person Number" is formatted as "Person Number: XXXXX" or 
    Person Number XXXXX" in the PDF text.
--The script also attempts to extract first and last names from the text, looking for lines containing "Dear"
  and lines above it. If names cannot be found, it defaults to "UNKNOWN".
--Pages without "dear", "chief people officer" and "person number" are grouped with the preceding page.
"""
import re
from PyPDF2 import PdfReader, PdfWriter
from tkinter import Tk
from tkinter.filedialog import askopenfilenames, askdirectory
import os
import string
from datetime import datetime

# -------------------------
# Utilities
# -------------------------
def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars)

def page_verification(text):
    """
    Returns True if page contains required phrases to start a new PDF.
    """
    required_phrases = ["dear", "chief people officer"]
    return any(phrase in text.lower() for phrase in required_phrases)

def extract_name(text, max_lines_above=8):
    """
    Extracts first and last name from text.
    - First name from 'Dear ...'
    - Last name from lines above (up to max_lines_above) or the Dear line
    - Returns (first_name, last_name) or (None, None)
    """
    first_name = None
    last_name = None

    lines = [line.strip() for line in text.splitlines() if line.strip()]

    # Step 1: find line containing "Dear"
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

    # Step 2: look for last name in lines above (up to max_lines_above)
    start_idx = max(0, dear_idx - max_lines_above)
    candidate_lines = lines[start_idx:dear_idx][::-1]  # nearest line first
    candidate_lines.append(lines[dear_idx])  # fallback to Dear line

    # Match first name followed by 1-3 words (last name)
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
# Main processing
# -------------------------
def main():
    Tk().withdraw()

    # Select PDFs
    input_pdfs = askopenfilenames(title="Select PDF files", filetypes=[("PDF Files", "*.pdf")])
    if not input_pdfs:
        print("No files selected. Exiting.")
        return

    # Select output folder
    save_folder = askdirectory(title="Select folder to save PDFs")
    if not save_folder:
        print("No folder selected. Exiting.")
        return
    os.makedirs(save_folder, exist_ok=True)

    # Regex for Person Number
    pattern = re.compile(r"Person Number[:\s]+(\d+)")
    all_saved_files = []

    current_prefix = datetime.now().strftime("%Y-%m")

    for input_pdf in input_pdfs:
        reader = PdfReader(input_pdf)
        saved_files = []

        writer = None
        current_filename = None

        # Process pages
        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            pn_match = pattern.search(text)
            verified = page_verification(text)

            if pn_match or verified:  # New PDF start
                # Save previous
                if writer and current_filename:
                    with open(current_filename, "wb") as f_out:
                        writer.write(f_out)
                    saved_files.append(current_filename)

                # Extract info
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
                current_filename = os.path.join(save_folder, sanitize_filename(base_name))

                writer = PdfWriter()
                writer.add_page(page)

            else:  # Continuation page
                if writer:
                    writer.add_page(page)
                else:
                    # Unknown standalone
                    fallback_name = f"{current_prefix}_Salary Review_UNKNOWN_UNKNOWN_PN_NotFound_page_{i+1:04d}.pdf"
                    current_filename = os.path.join(save_folder, sanitize_filename(fallback_name))
                    writer = PdfWriter()
                    writer.add_page(page)

        # Save last PDF
        if writer and current_filename:
            with open(current_filename, "wb") as f_out:
                writer.write(f_out)
            saved_files.append(current_filename)

        all_saved_files.extend(saved_files)
        print(f"Processed {os.path.basename(input_pdf)}: {len(saved_files)} files saved.")

    print(f"\nDone! {len(all_saved_files)} PDFs saved to {save_folder}")

if __name__ == "__main__":
    main()