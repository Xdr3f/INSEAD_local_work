import pdfplumber
import os
import re
import unicodedata
import tkinter as tk
from tkinter import filedialog
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# ===== PDF Text Extraction =====
def extract_text_from_pdf(pdf_path):
    text_content = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                text_content += " " + text
    return text_content

# ===== Normalize Text =====
def normalize_text(text):
    """
    Normalize text for comparison:
    - Lowercase
    - Normalize accents (é -> e, ñ -> n, etc.)
    - Keep only letters and spaces
    - Collapse multiple spaces
    """
    text = text.lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if c.isalpha() or c.isspace())
    text = re.sub(r"\s+", " ", text).strip()
    return text

# ===== Extract First & Last Name from Filename =====
def extract_name_from_filename(filename):
    """
    Expected format: {date}_{role}_{firstName}_{lastName}_{ID}
    Returns (firstName, lastName) or (None, None) if not matched
    """
    base_name = os.path.splitext(os.path.basename(filename))[0]
    parts = base_name.split("_")
    if len(parts) >= 4:
        first_name = parts[2]
        last_name = parts[3]
        return first_name, last_name
    return None, None

# ===== Check Match =====
def check_name_in_pdf(pdf_path):
    pdf_text = extract_text_from_pdf(pdf_path)
    normalized_pdf = normalize_text(pdf_text)

    first_name, last_name = extract_name_from_filename(pdf_path)
    if not first_name or not last_name:
        return False  # Can't extract name properly

    first_name = normalize_text(first_name)
    last_name = normalize_text(last_name)

    # Check if both names are present somewhere in the PDF
    return (first_name in normalized_pdf and last_name in normalized_pdf)

# ===== Scan Folder =====
def scan_folder(folder_path):
    results = {"no_match": []}
    for file in os.listdir(folder_path):
        if file.lower().endswith(".pdf"):
            full_path = os.path.join(folder_path, file)
            if not check_name_in_pdf(full_path):
                results["no_match"].append(file)
    return results

# ===== Generate PDF Report =====
def generate_pdf_report(files_list, save_path):
    c = canvas.Canvas(save_path, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 50, "PDF Filename Match Report")
    c.setFont("Helvetica", 12)

    y = height - 80
    if not files_list:
        c.drawString(50, y, "All PDF filenames match the content. No mismatches found.")
    else:
        c.drawString(50, y, "Files that did NOT match the content:")
        y -= 20
        for file in files_list:
            c.drawString(60, y, "- " + file)
            y -= 20
            if y < 50:
                c.showPage()
                c.setFont("Helvetica", 12)
                y = height - 50

    c.save()

# ===== Main Program =====
root = tk.Tk()
root.withdraw()

# Select folder containing PDFs
folder_path = filedialog.askdirectory(title="Select the folder containing PDFs")
if not folder_path:
    print("No folder selected. Exiting...")
    exit()

# Scan PDFs
results = scan_folder(folder_path)
no_match_files = results["no_match"]

# Ask where to save the report PDF
save_path = filedialog.asksaveasfilename(
    defaultextension=".pdf",
    filetypes=[("PDF files", "*.pdf")],
    title="Save PDF report as..."
)
if not save_path:
    print("No save location selected. Exiting...")
    exit()

# Generate PDF report
generate_pdf_report(no_match_files, save_path)
print(f"PDF report generated at: {save_path}")
