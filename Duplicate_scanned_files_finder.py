import os
import hashlib
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

import fitz  # PyMuPDF
from PIL import Image
import imagehash


def hash_pdf_file(filepath):
    """Compute a hash of the PDF file's binary content."""
    hasher = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while True:
            chunk = f.read(8192)
            if not chunk:
                break
            hasher.update(chunk)
    return hasher.hexdigest()


def perceptual_hash_pdf(filepath):
    """Compute a perceptual hash based on the first page of a PDF."""
    try:
        doc = fitz.open(filepath)
        if len(doc) == 0:
            return None

        page = doc[0]
        pix = page.get_pixmap(dpi=100)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        img = img.convert("L").resize((256, 256), Image.LANCZOS)
        return imagehash.phash(img)
    except Exception as e:
        print(f"Error hashing {filepath}: {e}")
        return None


def find_duplicate_pdfs(folder_path, threshold=5):
    """Find visually duplicate PDF files (perceptual hash)"""
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    hashes = {}
    duplicates = []

    for file in pdf_files:
        full_path = os.path.join(folder_path, file)
        file_hash = perceptual_hash_pdf(full_path)
        if not file_hash:
            continue

        found = False
        for h, orig in hashes.items():
            distance = file_hash - h  # Hamming distance
            if distance <= threshold:
                duplicates.append((file, orig))
                found = True
                break

        if not found:
            hashes[file_hash] = file

    return duplicates


def generate_pdf_report(duplicates, save_path):
    c = canvas.Canvas(save_path, pagesize=A4)
    width, height = A4
    margin_x = 50
    y = height - 50

    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin_x, y, "Duplicate PDF Files Report")
    y -= 30

    c.setFont("Helvetica-Bold", 12)
    summary = f"Total duplicates found: {len(duplicates)}"
    c.drawString(margin_x, y, summary)
    y -= 20

    c.setFont("Helvetica", 10)
    if not duplicates:
        c.drawString(margin_x, y, "No duplicate files found.")
    else:
        for dup, orig in duplicates:
            line = f"Duplicate: {dup}  |  Matches: {orig}"
            c.drawString(margin_x, y, line)
            y -= 12
            if y < 60:
                c.showPage()
                y = height - 50
    c.save()


if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()  # hide the main window

    # Select folder with PDFs
    folder_path = filedialog.askdirectory(title="Select folder containing PDFs")
    if not folder_path:
        print("No folder selected. Exiting.")
        exit()

    print("Scanning PDFs for duplicates...")

    duplicates = find_duplicate_pdfs(folder_path)

    # Select where to save the report
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Save PDF report as..."
    )
    if not save_path:
        print("No save path selected. Exiting.")
        exit()

    generate_pdf_report(duplicates, save_path)
    print(f"Report generated: {save_path}")
