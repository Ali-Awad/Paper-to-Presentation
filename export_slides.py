#!/usr/bin/env python3
"""Export each page of presentation.pdf as a PNG image."""
import os
import fitz  # PyMuPDF

pdf_path = "presentation.pdf"
out_dir = "slides_export"
# Very high resolution: 8x scale (~576 DPI equivalent) for large displays and print
zoom = 8.0

if not os.path.isfile(pdf_path):
    print(f"Error: {pdf_path} not found. Compile the .tex file first.")
    exit(1)

os.makedirs(out_dir, exist_ok=True)
doc = fitz.open(pdf_path)
n = len(doc)
mat = fitz.Matrix(zoom, zoom)
for i, page in enumerate(doc):
    pix = page.get_pixmap(matrix=mat, alpha=False)
    out_path = os.path.join(out_dir, f"slide_{i+1:02d}.png")
    pix.save(out_path)
    print(f"Saved {out_path}")
doc.close()
print(f"Exported {n} slides to {out_dir}/")
