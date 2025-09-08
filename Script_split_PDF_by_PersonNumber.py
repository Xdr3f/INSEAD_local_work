"""
This script allows the user to select one or more PDF files and a destination folder. 
It then processes each PDF, searching for "Person Number" entries using a regular expression. 
The script is designed to split or organize the PDFs based on these person numbers, 
saving the resulting files with sanitized filenames in the chosen folder.
"""

import re
from PyPDF2 import PdfReader, PdfWriter
from tkinter import Tk
from tkinter.filedialog import askopenfilenames, askdirectory
import os
import string

def sanitize_filename(name):
    valid_chars = f"-_.() {string.ascii_letters}{string.digits}"
    return "".join(c for c in name if c in valid_chars)

def main():
    Tk().withdraw()
    
    # Select multiple PDF files
    input_pdfs = askopenfilenames(title="Select PDF files", filetypes=[("PDF Files", "*.pdf")])
    if not input_pdfs:
        print("No files selected. Exiting.")
        return

    # Select folder
    save_folder = askdirectory(title="Select folder to save PDFs")
    if not save_folder:
        print("No folder selected. Exiting.")
        return
    os.makedirs(save_folder, exist_ok=True)

    # Regex pattern
    pattern = re.compile(r"Person Number[:\s]+(\d+)")
    all_saved_files = []

    # Loop over each selected PDF
    for input_pdf in input_pdfs:
        reader = PdfReader(input_pdf)
        saved_files = []

        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            match = pattern.search(text)
            person_number = match.group(1) if match else f"PN_NotFound_page_{i+1:04d}"
            if not match:
                print(f"Warning: No Person Number found on page {i+1} of {os.path.basename(input_pdf)}")

            person_number = sanitize_filename(person_number)
            writer = PdfWriter()
            writer.add_page(page)

            output_filename = os.path.join(save_folder, f"{person_number}.pdf")
            with open(output_filename, "wb") as f_out:
                writer.write(f_out)

            saved_files.append(output_filename)

        all_saved_files.extend(saved_files)
        print(f"Processed {os.path.basename(input_pdf)}: {len(saved_files)} pages saved.")

    print(f"\nDone! {len(all_saved_files)} PDFs saved to {save_folder}")

if __name__ == "__main__":
    main()
