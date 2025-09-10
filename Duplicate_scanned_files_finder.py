import os
import hashlib
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
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
    """Compute a perceptual hash based on rendered first page of PDF (PyMuPDF)."""
    try:
        doc = fitz.open(filepath)
        if len(doc) == 0:
            return None

        page = doc[0]  # first page only
        pix = page.get_pixmap(dpi=100)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        img = img.convert("L").resize((256, 256), Image.LANCZOS)
        return imagehash.phash(img)  # returns ImageHash object
    except Exception as e:
        print(f"Error hashing {filepath}: {e}")
        return None


def find_duplicate_pdfs(folder_path, use_visual=False, threshold=5):
    """
    Find duplicate PDF files (binary or perceptual).
    For visual comparison, considers files duplicates if hash distance <= threshold.
    """
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    hashes = {}
    duplicates = []

    for file in pdf_files:
        full_path = os.path.join(folder_path, file)

        if use_visual:
            file_hash = perceptual_hash_pdf(full_path)
        else:
            file_hash = hash_pdf_file(full_path)

        if not file_hash:
            continue

        found = False
        for h, orig in hashes.items():
            if use_visual:
                distance = file_hash - h  # Hamming distance
                if distance <= threshold:
                    duplicates.append((file, orig))
                    found = True
                    break
            else:
                if file_hash == h:
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


class DuplicatePDFScannerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Duplicate PDF Scanner")
        self.use_visual = tk.BooleanVar()
        self.setup_gui()

    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Button(main_frame, text="Select Folder", command=self.process_folder).grid(row=0, column=0, pady=5)
        ttk.Checkbutton(main_frame, text="Use visual similarity (slower)", variable=self.use_visual).grid(row=1, column=0, pady=5)

        self.status_var = tk.StringVar()
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=2, column=0, pady=5)

    def process_folder(self):
        folder_path = filedialog.askdirectory(title="Select the folder containing PDFs")
        if not folder_path:
            return

        try:
            duplicates = find_duplicate_pdfs(folder_path, use_visual=self.use_visual.get())
        except Exception as e:
            messagebox.showerror("Error", f"Failed to scan folder: {e}")
            return

        self.status_var.set(f"Found {len(duplicates)} duplicate(s). Generating report...")

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save PDF report as..."
        )
        if save_path:
            generate_pdf_report(duplicates, save_path)
            self.status_var.set(f"Report generated: {save_path}")
            self.root.after(1500, self.root.destroy)


if __name__ == "__main__":
    app = DuplicatePDFScannerGUI()
    app.root.mainloop()
