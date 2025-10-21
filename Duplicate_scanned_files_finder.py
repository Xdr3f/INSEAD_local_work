import os
import hashlib
import fitz  # PyMuPDF
from PIL import Image
import imagehash
import tempfile
import shutil
import pdfplumber
import re
import unicodedata
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image as RLImage, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import tkinter as tk
from tkinter import filedialog

# ------------------------------
# PDF Type Detection
# ------------------------------
def is_text_pdf(pdf_path: str) -> bool:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text and text.strip():
                    return True
    except Exception:
        return False
    return False

def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip().lower()

def text_hash_pdf(pdf_path: str) -> str:
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            full_text += page.extract_text() or ""
    return hashlib.sha256(normalize_text(full_text).encode("utf-8")).hexdigest()

# ------------------------------
# Perceptual Hash for Images
# ------------------------------
def perceptual_hash_pdf(pdf_path: str) -> imagehash.ImageHash:
    try:
        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            return None
        page = doc[0]
        pix = page.get_pixmap(dpi=100)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        img = img.convert("L").resize((256, 256), Image.LANCZOS)
        return imagehash.phash(img)
    except Exception as e:
        print(f"Error hashing {pdf_path}: {e}")
        return None

# ------------------------------
# Duplicate Detection
# ------------------------------
def find_duplicate_pdfs(folder_path, phash_threshold=5):
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    
    text_hashes = {}
    image_hashes = {}
    exact_text_duplicates = []
    exact_visual_duplicates = []
    near_visual_duplicates = []

    for file in pdf_files:
        full_path = os.path.join(folder_path, file)
        if is_text_pdf(full_path):
            h = text_hash_pdf(full_path)
            if h in text_hashes:
                exact_text_duplicates.append((file, text_hashes[h], 0))
            else:
                text_hashes[h] = file
        else:
            h = perceptual_hash_pdf(full_path)
            if not h:
                continue
            found = False
            for ih, orig_file in image_hashes.items():
                dist = h - ih
                if dist == 0:
                    exact_visual_duplicates.append((file, orig_file, dist))
                    found = True
                    break
                elif dist <= phash_threshold:
                    near_visual_duplicates.append((file, orig_file, dist))
                    found = True
                    break
            if not found:
                image_hashes[h] = file

    return exact_text_duplicates, exact_visual_duplicates, near_visual_duplicates

# ------------------------------
# PDF Report Generation
# ------------------------------
def generate_pdf_report(folder_path, text_dups, exact_visual, near_visual, save_path, thumbnail=True):
    tempdir = tempfile.mkdtemp(prefix="pdf_thumbs_")
    doc = SimpleDocTemplate(save_path, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    title_style = styles['Title']
    heading_style = styles['Heading2']
    normal_style = styles['Normal']
    cell_style = ParagraphStyle('cell', fontSize=9, leading=11)

    elements.append(Paragraph("PDF Duplicates Report", title_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Folder scanned: {folder_path}", normal_style))
    elements.append(Paragraph(f"Total exact text duplicates: {len(text_dups)}", heading_style))
    elements.append(Paragraph(f"Total exact visual duplicates: {len(exact_visual)}", heading_style))
    elements.append(Paragraph(f"Total near-duplicate visual PDFs: {len(near_visual)}", heading_style))
    elements.append(Spacer(1, 24))

    def add_table(title, duplicates, include_thumbnails):
        elements.append(PageBreak())
        elements.append(Paragraph(f"{title} ({len(duplicates)})", heading_style))
        elements.append(Spacer(1, 12))

        if not duplicates:
            elements.append(Paragraph("None", normal_style))
            elements.append(Spacer(1, 12))
            return

        table_data = [["Duplicate", "Original", "Distance", "Thumbnail"]]

        for dup, orig, dist in duplicates:
            row = [Paragraph(dup, cell_style), Paragraph(orig, cell_style), Paragraph(str(dist), cell_style)]
            if include_thumbnails and dist >= 0:
                try:
                    doc_pdf = fitz.open(os.path.join(folder_path, dup))
                    page = doc_pdf[0]
                    pix = page.get_pixmap(dpi=50)
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    thumb_path = os.path.join(tempdir, f"thumb_{dup}.png")
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
        elements.append(Spacer(1, 12))

    add_table("Exact Text Duplicates", text_dups, include_thumbnails=False)
    add_table("Exact Visual Duplicates", exact_visual, include_thumbnails=thumbnail)
    add_table("Near-Duplicate Visual PDFs", near_visual, include_thumbnails=thumbnail)

    doc.build(elements)
    shutil.rmtree(tempdir, ignore_errors=True)

# ------------------------------
# Main Execution
# ------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    folder_path = filedialog.askdirectory(title="Select folder containing PDFs")
    if not folder_path:
        print("No folder selected. Exiting.")
        exit()

    print("Scanning PDFs for duplicates...")
    text_dups, exact_visual, near_visual = find_duplicate_pdfs(folder_path, phash_threshold=5)

    default_name = f"PDF_Duplicates_Report_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M')}.pdf"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        initialfile=default_name,
        filetypes=[("PDF files", "*.pdf")],
        title="Save PDF report as..."
    )
    if not save_path:
        print("No save path selected. Exiting.")
        exit()

    generate_pdf_report(folder_path, text_dups, exact_visual, near_visual, save_path)
    print(f"Report generated: {save_path}")