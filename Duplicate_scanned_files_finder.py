import os
import hashlib
import fitz  # PyMuPDF for reading PDFs
from PIL import Image
import imagehash
import tempfile
import shutil
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image as RLImage, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ===== PDF Hashing =====
def hash_pdf_file(filepath):
    """
    Computes a SHA-256 hash of the PDF file's binary content.
    Useful for detecting exact binary duplicates.
    """
    hasher = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while True:
            chunk = f.read(8192)
            if not chunk:
                break
            hasher.update(chunk)
    return hasher.hexdigest()


def perceptual_hash_pdf(filepath):
    """
    Generates a perceptual hash based on the first page of the PDF.
    Converts the page to grayscale, resizes to 256x256, and calculates pHash.
    Useful for detecting visually similar (near-duplicate) PDFs.
    """
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

# ===== Find Duplicates =====
def find_duplicate_pdfs(folder_path, threshold=5):
    """
    Scan the folder and find exact and near-duplicate PDFs.
    
    Args:
        folder_path (str): Path to the folder containing PDFs
        threshold (int): Maximum hamming distance for near-duplicates
    
    Returns:
        exact_duplicates (list): [(duplicate_filename, original_filename, distance)]
        near_duplicates (list): [(duplicate_filename, original_filename, distance)]
    """
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    hashes = {}
    exact_duplicates = []
    near_duplicates = []

    for file in pdf_files:
        full_path = os.path.join(folder_path, file)
        file_hash = perceptual_hash_pdf(full_path)
        if not file_hash:
            continue

        found_exact = False
        for h, orig in hashes.items():
            distance = file_hash - h
            if distance == 0:
                exact_duplicates.append((file, orig, distance))
                found_exact = True
                break
            elif distance <= threshold:
                near_duplicates.append((file, orig, distance))
                found_exact = True
                break
        if not found_exact:
            hashes[file_hash] = file
    return exact_duplicates, near_duplicates

# ===== PDF Report Generation =====
def generate_pdf_report(folder_path, exact_duplicates, near_duplicates, save_path, thumbnail=True):
    """
    Generates a PDF report summarizing duplicates found in the folder.
    
    Args:
        folder_path (str): Scanned folder path
        exact_duplicates (list): List of exact duplicates
        near_duplicates (list): List of near duplicates
        save_path (str): Output PDF path
        thumbnail (bool): Whether to include first-page thumbnails
    """
    # Temporary folder for thumbnails
    thumb_tempdir = tempfile.mkdtemp(prefix="pdf_thumbs_")
    doc = SimpleDocTemplate(save_path, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    total_style = ParagraphStyle(
        'total_style',
        parent=styles['Heading2'],
        fontSize=14,
        leading=16,
        textColor=colors.darkblue,
        spaceAfter=12,
    )

    elements.append(Paragraph("Duplicate PDF Files Report", styles['Title']))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Folder scanned: {folder_path}", styles['Normal']))
    elements.append(Paragraph(f"Total exact duplicates: {len(exact_duplicates)}", total_style))
    elements.append(Paragraph(f"Total near-duplicates (threshold applied): {len(near_duplicates)}", total_style))
    elements.append(Spacer(1, 24))

    def add_table_section(title, data_list):
        """Add a table for a section (exact or near duplicates)"""
        if not data_list:
            elements.append(Paragraph(f"No {title.lower()} found.", styles['Normal']))
            return

        elements.append(Paragraph(title, styles['Heading2']))
        elements.append(Spacer(1, 12))

        cell_style = ParagraphStyle('cell_style', fontSize=9, leading=11)
        table_data = [["Duplicate", "Original", "Distance", "Thumbnail"]]

        for dup, orig, dist in data_list:
            row = [
                Paragraph(dup, cell_style),
                Paragraph(orig, cell_style),
                Paragraph(str(dist), cell_style)
            ]

            if thumbnail:
                try:
                    doc_pdf = fitz.open(os.path.join(folder_path, dup))
                    page = doc_pdf[0]
                    pix = page.get_pixmap(dpi=50)
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    thumb_path = os.path.join(thumb_tempdir, f"thumb_{dup}.png")
                    img.thumbnail((80, 100))
                    img.save(thumb_path)
                    row.append(RLImage(thumb_path))
                except Exception:
                    row.append(Paragraph("Error", cell_style))
            else:
                row.append("")

            table_data.append(row)

        table = Table(table_data, colWidths=[150, 150, 60, 100])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.grey),
            ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 8),
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND',(0,1),(-1,-1),colors.beige),
        ]))
        elements.append(table)
        elements.append(PageBreak())

    add_table_section("Exact Duplicates", exact_duplicates)
    add_table_section("Near-Duplicates", near_duplicates)

    doc.build(elements)
    shutil.rmtree(thumb_tempdir, ignore_errors=True)

# ===== Main Execution =====
if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog
    import datetime

    root = tk.Tk()
    root.withdraw()

    folder_path = filedialog.askdirectory(title="Select folder containing PDFs")
    if not folder_path:
        print("No folder selected. Exiting.")
        exit()

    print("Scanning PDFs for duplicates...")
    exact_duplicates, near_duplicates = find_duplicate_pdfs(folder_path, threshold=5)

    default_name = f"PDF_Duplicate_Report_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')}.pdf"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        initialfile=default_name,
        filetypes=[("PDF files", "*.pdf")],
        title="Save PDF report as..."
    )
    if not save_path:
        print("No save path selected. Exiting.")
        exit()

    generate_pdf_report(folder_path, exact_duplicates, near_duplicates, save_path)
    print(f"Report generated: {save_path}")
