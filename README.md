# Paper to Presentation

![Paper to Presentation](Assets/banner.png)

Turn a research paper into an editable PowerPoint presentation in one command. This repo bundles a clean **LaTeX Beamer** template (16:9) with a Python converter (`beamer_to_pptx.py`) that rebuilds each slide as native PowerPoint objects - preserving speaker notes, the footer logo, and tightly-cropped figures pulled straight from your paper. The visual style is inspired by the [Michigan Technological University](https://www.mtu.edu/) PowerPoint template.

The workflow in one line: drop your paper in `Paper_files/`, export its figures into `Extracted_figures/`, write the narrative in `presentation.tex`, run the converter, and hand the `.pptx` off to co-authors who don't use LaTeX.

## Features

- **Paper-centric layout**: dedicated folders for the source paper (`Paper_files/`), extracted figures (`Extracted_figures/`), and reusable assets (`Assets/`)
- **Automatic figure collection**: drop your figures anywhere on the `\graphicspath` (e.g. alongside the paper in `Paper_files/`) and the converter copies every `\includegraphics` it encounters into `Extracted_figures/` for you - no manual file shuffling
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
|- Paper_files/              # YOU drop your paper here (starts empty)
|- Extracted_figures/        # CONVERTER writes here: copies of every figure used + TikZ/table crops (starts empty)
|- Assets/                   # Logo, banner, and reusable graphics
|   |- logo.jpg
|   `- banner.png
`- .cursor/rules/            # Cursor rules that enforce this repo's conventions
```

Two of these folders ship **empty on purpose**:

- `Paper_files/` - **you** put your paper here before you start. Drop the source PDF (and any `.tex` / `.bib` / supplementary you have) into it, along with any figure files you've pulled out of the paper (PNG / PDF / JPG). Keeping paper + figures together means you can reference them immediately from `presentation.tex` without moving anything.
- `Extracted_figures/` - **populated by the converter, not by you**. Every time you run `beamer_to_pptx.py`, the script:
  - walks every `\includegraphics{...}` in `presentation.tex`,
  - resolves each one via `\graphicspath` (so figures in `Paper_files/` are found), and
  - **copies the source file into `Extracted_figures/`** if it isn't already there.

  It also saves the auto-cropped TikZ/table PNGs here as `slide_content_<N>.png`. So after conversion, `Extracted_figures/` is a canonical snapshot of every figure actually used in the deck - easy to share, archive, or review. Auto-cropped filenames are git-ignored; your own figures are tracked.

`\graphicspath{{Extracted_figures/}{Paper_files/}{Assets/}}` is set in the template: the first run picks figures up from `Paper_files/` and copies them to `Extracted_figures/`; subsequent runs prefer the copies. Either way, you write `\includegraphics{figure1}` with no path prefix.

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

Then, in order:

1. **Put the paper and its figures into `Paper_files/`** - this folder starts empty. Drop the source PDF, any `.tex` / `.bib` / supplementary files, **and any figure files you pull out of the paper** (PNG / PDF / JPG) here. Keeping them together means the next step can reference figures by filename alone.
2. **Put your institution logo in `Assets/`** as `logo.jpg` (or `logo.png`) - a placeholder logo is already shipped; replace it with your own.
3. **Edit `presentation.tex`**: update title / author / date and reference your figures with `\includegraphics{figure1}` (no path needed - `\graphicspath` searches `Paper_files/` for you). Save the file; LaTeX Workshop compiles the PDF automatically.
4. **Build the PPTX**:

```bash
python3 beamer_to_pptx.py presentation.tex
```

That single command produces `presentation_editable.pptx` alongside the `.tex`. It also **copies every referenced figure into `Extracted_figures/` automatically** (so you never have to move files by hand) and writes any auto-cropped TikZ/table images there as `slide_content_<N>.png` (git-ignored). The compiled PDF is auto-detected from the same basename when present, which enables high-fidelity cropping for TikZ and tables.

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

Save any figure you pull out of the paper into `Paper_files/` (alongside the paper itself). Because `\graphicspath` lists `Paper_files/`, you reference the figure by filename only:

```tex
\includegraphics[width=\linewidth,height=0.58\textheight,keepaspectratio]{figure1}
```

When you run `beamer_to_pptx.py`, the converter **copies every referenced figure into `Extracted_figures/` for you** - you don't copy them yourself. On subsequent runs LaTeX will find the copy in `Extracted_figures/` first (it's earlier in `\graphicspath`), so you can safely delete the original from `Paper_files/` once you're happy with the deck if you want a tighter archive.

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

For every `\includegraphics{figure}` it encounters, it:

- resolves the path against `\graphicspath` (searching `Extracted_figures/`, then `Paper_files/`, then `Assets/`, then the repo root);
- if the resolved file lives outside `Extracted_figures/`, **copies it into `Extracted_figures/`** (no overwrite if an identically-sized copy already exists);
- embeds the copy in the generated slide.

The footer logo is treated as chrome and is **not** copied.

For frames that use `\begin{tikzpicture}` or `\begin{tabular}` (which have no direct PPTX equivalent), the script clip-rasterizes just the relevant region of the compiled PDF:

- It reads the PDF page's **vector drawings** (including booktabs rule lines) and **embedded image rects** to locate the element's own bounding box.
- It merges in only text blocks that lie inside that bbox's vertical range and horizontally overlap - so captions, side text, and bullet lists below the table are **not** pulled into the extracted image.
- When bullets coexist with a table on the same slide, the bottom of the crop is snapped to the table's last rule line.
- The cropped PNG is saved into `Extracted_figures/` as `slide_content_<N>.png` so your extracted-figure folder doubles as the converter's output cache. Auto-generated filenames are in `.gitignore` by default.

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
| `Paper_files/` | Paper + its figures (PDF / tex / PNG / JPG / supplementary). **Starts empty - you fill it.** |
| `Extracted_figures/` | Canonical copies of every figure actually used in the deck + auto-cropped TikZ/table PNGs. **Populated by the converter, not by you.** |
| `Assets/` | Logo, banner, and reusable graphics |
| `.cursor/rules/` | Cursor rules that enforce this repo's conventions |
