#!/usr/bin/env python3
"""Put slide images from slides_export/ into a PowerPoint file."""
import os
import glob
from pptx import Presentation
from pptx.util import Emu

slide_dir = "slides_export"
out_pptx = "presentation.pptx"

# Discover all slide_XX.png (sorted)
paths = sorted(glob.glob(os.path.join(slide_dir, "slide_*.png")))
if not paths:
    print(f"Error: No slide_*.png found in {slide_dir}. Run export_slides.py first.")
    exit(1)

prs = Presentation()
prs.slide_width = Emu(12192000)   # 16:9
prs.slide_height = Emu(6858000)
blank_slide_layout = prs.slide_layouts[6]

for path in paths:
    slide = prs.slides.add_slide(blank_slide_layout)
    slide.shapes.add_picture(
        path,
        left=0,
        top=0,
        width=prs.slide_width,
        height=prs.slide_height,
    )
    print(f"Added {os.path.basename(path)}")

prs.save(out_pptx)
print(f"Saved {out_pptx} ({len(paths)} slides)")
