# Paper to Presentation

![Paper to Presentation](Assets/banner.png)

Turn a research paper into an editable PowerPoint presentation in one command. This repo bundles a clean **LaTeX Beamer** template (16:9) with a Python converter (`beamer_to_pptx.py`) that rebuilds each slide as native PowerPoint objects - preserving speaker notes, the footer logo, and tightly-cropped figures pulled straight from your paper. The visual style is inspired by the [Michigan Technological University](https://www.mtu.edu/) PowerPoint template.

The workflow in one line: drop your paper in `Paper_files/`, export its figures into `Extracted_figures/`, write the narrative in `presentation.tex`, run the converter, and hand the `.pptx` off to co-authors who don't use LaTeX.

## Features

- **Paper-centric layout**: dedicated folders for the source paper (`Paper_files/`), extracted figures (`Extracted_figures/`), and reusable assets (`Assets/`)
- **One-step conversion** to an editable PPTX with `beamer_to_pptx.py` - no image-flattening, each slide is a real PowerPoint slide with real bullets, titles, and textboxes
- **Smart table / TikZ cropping** driven by the PDF's own vector primitives, so figures extracted from Beamer pages don't include surrounding captions or bullet text
- **Speaker notes ON by default** via `\note{...}` after each frame, carried into the PPTX Notes pane
- **Footer logo ON by default**, auto-discovered from `Assets/`
- 16:9 widescreen layout with generous margins and gold accent bullets
- Custom footer: gray page count `(1/6)`, centered date, right-aligned logo
- Title page skips page numbering automatically
- Side-by-side column layout for text + figures
- Section separators for organized source

## Folder layout

```
.
|- presentation.tex          # Main Beamer template
|- presentation.pdf          # Compiled PDF (output of latexmk)
|- beamer_to_pptx.py         # LaTeX -> editable PPTX converter
|- requirements.txt          # Python deps (PyMuPDF, python-pptx, Pillow)
|- Assets/                   # Logo, banner, reusable graphics
|   |- logo.jpg
|   `- banner.png
|- Extracted_figures/        # Figures you pull from your paper to put on slides
|- Paper_files/              # Reference paper (PDF/tex/supplementary) you are presenting
`- .cursor/rules/            # Cursor rules for this template
```

`\graphicspath{{Extracted_figures/}{Assets/}}` is set in the template, so `\includegraphics{figure1}` works without a path prefix.

## Installation

### 1. Install LaTeX

You need a TeX distribution with `pdflatex` and `latexmk`.

**Ubuntu / Debian / WSL:**

```bash
sudo apt install texlive-full latexmk
```

**macOS (Homebrew):**

```bash
brew install --cask mactex
```

**Windows:** install [MiKTeX](https://miktex.org/) or [TeX Live](https://tug.org/texlive/).

### 2. Editor setup (VS Code / Cursor)

Install the [LaTeX Workshop](https://marketplace.visualstudio.com/items?itemName=James-Yu.latex-workshop) extension. Once installed, saving any `.tex` file will automatically compile the PDF using `latexmk`.

### 3. Python packages for PPTX export

```bash
pip install -r requirements.txt
```

## Quick start

```bash
git clone https://github.com/Ali-Awad/paper-to-presentation.git
cd paper-to-presentation
pip install -r requirements.txt
```

Then:

1. Drop the paper you are presenting into `Paper_files/` (PDF, source tex, or supplementary material).
2. Export the figures you want on slides from that paper into `Extracted_figures/` (PNG / PDF / JPG).
3. Put your institution logo in `Assets/` as `logo.jpg` (or `logo.png`) - a logo is already shipped for you to replace.
4. Open `presentation.tex`, edit title / author / date, reference your extracted figures (`\includegraphics{figure1}`), and save. LaTeX Workshop compiles the PDF automatically.
5. Build the PPTX:

```bash
python3 beamer_to_pptx.py presentation.tex
```

That single command produces `presentation_editable.pptx` alongside the `.tex`. The compiled PDF is auto-detected from the same basename when present, which enables high-fidelity cropping for TikZ and tables.

## How to modify the template

### Title, author, date

Edit these lines near the top of the document section in `presentation.tex`:

```tex
\title{Your Presentation Title}
\author{Your Name}
\institute{Your Institution}
\date{02/25/2026}
```

### Logo

The logo is on by default and is read from `Assets/logo.jpg`.

- To change the logo, replace `Assets/logo.jpg` with your image (use a short, horizontal file so it fits the footer bar).
- To use a different path or filename, edit the `\renewcommand{\mylogo}{...}` line in the preamble.
- To remove the logo entirely, comment out the `\renewcommand{\mylogo}` line.

The converter will auto-discover any `\includegraphics` path containing `logo`, and falls back to `Assets/logo.jpg` / `Assets/logo.png` if nothing is found in the `.tex`.

### Figures from the paper

Save any figure you extract from the paper into `Extracted_figures/`. Because `\graphicspath` already lists that folder, you can reference them by filename only:

```tex
\includegraphics[width=\linewidth,height=0.58\textheight,keepaspectratio]{figure1}
```

### Adding slides

Each slide is a `\begin{frame}...\end{frame}` block followed by a `\note{...}` for speaker notes:

```tex
%----------------------------------------------------------------------
\section{Your Section Name}
%----------------------------------------------------------------------
\begin{frame}{Slide Title}
  \begin{itemize}
    \item First point.
    \item Second point.
  \end{itemize}
\end{frame}
\note{Explain the first point in more detail here. These notes appear in the PPTX Notes pane.}
```

Beamer silently ignores `\note{...}` in the normal PDF output, so your audience-facing PDF stays clean while the PPTX export retains your speaker notes.

### Accent color

The default bullet color is gold (`#FFC000`). Change it in the preamble:

```tex
\definecolor{accent}{HTML}{FFC000}
```

## Compiling from the command line

If you prefer not to use LaTeX Workshop:

```bash
latexmk -pdf -interaction=nonstopmode presentation.tex
```

## How the PPTX export works

`beamer_to_pptx.py` parses the `.tex` into slide structures (title, bullets, columns, images, notes) and rebuilds each slide using editable PowerPoint objects - not as a flat image.

For frames that use `\begin{tikzpicture}` or `\begin{tabular}` (which have no direct PPTX equivalent), the script clip-rasterizes just the relevant region of the compiled PDF:

- It reads the PDF page's **vector drawings** (including booktabs rule lines) and **embedded image rects** to locate the element's own bounding box.
- It merges in only text blocks that lie inside that bbox's vertical range and horizontally overlap - so captions, side text, and bullet lists below the table are **not** pulled into the extracted image.
- When bullets coexist with a table on the same slide, the bottom of the crop is snapped to the table's last rule line.

This replaces the previous full-page whitespace-band heuristic, which tended to leak captions and stray text into the extracted figure.

Usage:

```bash
python3 beamer_to_pptx.py presentation.tex                     # auto-detects presentation.pdf
python3 beamer_to_pptx.py presentation.tex out.pptx            # custom output name
python3 beamer_to_pptx.py presentation.tex my.pdf out.pptx     # explicit PDF + output
```

The PPTX output works in PowerPoint, Google Slides, or Keynote.

## Files

| Path | Description |
|------|-------------|
| `presentation.tex` | Main Beamer template (with `\note{...}` per frame) |
| `presentation.pdf` | Compiled PDF, also fed to the converter for fallback rasterization |
| `beamer_to_pptx.py` | Single-script LaTeX -> editable PPTX converter |
| `requirements.txt` | Python dependencies (PyMuPDF, python-pptx, Pillow) |
| `Paper_files/` | The paper you are presenting (PDF / tex / supplementary) |
| `Extracted_figures/` | Figures pulled from the paper, usable as `\includegraphics{...}` |
| `Assets/` | Logo, banner, and reusable graphics |
| `.cursor/rules/` | Cursor rules that enforce this repo's conventions |
