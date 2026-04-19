# Paper to Presentation

![Paper to Presentation](Assets/banner.png)

Turn a research paper into an editable PowerPoint presentation in one prompt. This repo bundles a clean **LaTeX Beamer** template (16:9), a Python converter (`beamer_to_pptx.py`), and Cursor rules that let the coding agent do the heavy lifting: read your paper, extract its figures, write the slide narrative, and produce both a PDF and an editable PPTX. The visual style is inspired by the [Michigan Technological University](https://www.mtu.edu/) PowerPoint template.

## The workflow

1. **You** drop the paper into `Paper_files/`. The PDF alone works, but dropping the source `.tex` / `.bib` / supplementary alongside it is encouraged, the richer the input, the smarter the slides (better section structure, cleaner equations, more accurate figure captions).
2. **You** tell the Cursor agent *"generate the slides"*.
3. **The agent** sets up `.venv`, extracts the figures it needs from the PDF into `Extracted_figures/`, edits `presentation.tex`, compiles the PDF, builds the editable PPTX, and self-checks the result against the rules in `.cursor/rules/presentation-slides.mdc`.

You do not hand-extract figures or hand-copy files.

## Features

- **One-prompt generation**: drop the paper, ask for slides, get PDF + editable PPTX
- **Editable PPTX**: each slide is a real PowerPoint slide with real bullets, titles, and textboxes (no image-flattening)
- **Smart TikZ / table cropping** driven by the PDF's own vector primitives, so captions and bullet text don't leak into extracted images
- **Speaker notes ON by default** via `\note{...}` after each frame, carried into the PPTX Notes pane
- **Footer logo ON by default**, auto-discovered from `Assets/`
- **16:9 widescreen** layout with generous margins, gold accent bullets, gray page count, centered date, right-aligned logo

## Folder layout

```
.
|- presentation.tex          # Main Beamer template (agent edits this)
|- presentation.pdf          # Compiled PDF (latexmk output)
|- beamer_to_pptx.py         # LaTeX -> editable PPTX converter
|- requirements.txt          # Python deps (PyMuPDF, python-pptx, Pillow)
|- Paper_files/              # YOU drop the paper here (starts empty)
|- Extracted_figures/        # AGENT + CONVERTER write here (starts empty)
|- Assets/                   # Logo, banner, reusable graphics
|   |- logo.jpg
|   `- banner.png
`- .cursor/rules/            # Cursor rules that pin the workflow
```

- `Paper_files/` - starts empty. Drop the paper here. The PDF alone is enough to get slides, but dropping the source `.tex` / `.bib` / supplementary too is encouraged — the agent uses whatever's available, and more source material yields smarter, more faithful slides.
- `Extracted_figures/` - starts empty. The agent writes figures it pulls from the paper here; the converter also writes auto-cropped TikZ/table PNGs (`slide_content_<N>.png`, git-ignored). Agent-extracted figures are tracked.
- `Assets/` - deck chrome (logo, banner). Not figure storage.

## Prerequisites

**LaTeX** (any distribution with `pdflatex` + `latexmk`):

- Ubuntu / Debian / WSL: `sudo apt install texlive-full latexmk`
- macOS: `brew install --cask mactex`
- Windows: [MiKTeX](https://miktex.org/) or [TeX Live](https://tug.org/texlive/)

**Python 3** with venv support (everything else is installed by the agent into a local `.venv/`).

Optional: the [LaTeX Workshop](https://marketplace.visualstudio.com/items?itemName=James-Yu.latex-workshop) extension for VS Code / Cursor so saving `.tex` auto-compiles the PDF.

## Quick start

```bash
git clone https://github.com/Ali-Awad/Paper-to-Presentation.git
cd Paper-to-Presentation
```

Then:

1. Drop the paper into `Paper_files/`. The PDF alone is enough to get started; dropping the source `.tex` / `.bib` / supplementary alongside it is encouraged for smarter output.
2. (Optional) replace `Assets/logo.jpg` with your institution's logo.
3. Ask the Cursor agent: *"generate the slides from the paper I just dropped."*

The agent will create `.venv/`, install `requirements.txt`, extract figures from the PDF into `Extracted_figures/`, fill `presentation.tex`, compile `presentation.pdf`, build `presentation_editable.pptx`, and run a post-generation self-check against the style/content rules in `.cursor/rules/`.

### Running the toolchain manually

If you'd rather drive it yourself (e.g., after hand-editing `presentation.tex`):

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt

latexmk -pdf -interaction=nonstopmode presentation.tex
python3 beamer_to_pptx.py presentation.tex
```

Usage variants:

```bash
python3 beamer_to_pptx.py presentation.tex                     # auto-detects presentation.pdf
python3 beamer_to_pptx.py presentation.tex out.pptx            # custom output name
python3 beamer_to_pptx.py presentation.tex my.pdf out.pptx     # explicit PDF + output
```

Recompile the PDF before rebuilding the PPTX: the converter uses the PDF for high-fidelity TikZ / table cropping.

## Tweaking the template

**Title / author / date** in `presentation.tex`:

```tex
\title{Your Presentation Title}
\author{Your Name}
\institute{Your Institution}
\date{02/25/2026}
```

**Logo**: replace `Assets/logo.jpg` with your own horizontal image, or comment out `\renewcommand{\mylogo}` to disable.

**Accent color** (default gold):

```tex
\definecolor{accent}{HTML}{FFC000}
```

**Figure resolution**: `beamer_to_pptx.py` rasterizes TikZ / table regions at `PDF_DPI = 450`. Raise it for sharper output (at the cost of PPTX size), lower it for smaller files.

**Adding slides** - each slide is a `\begin{frame}...\end{frame}` block followed by a `\note{...}`:

```tex
\begin{frame}{Slide Title}
  \begin{itemize}
    \item First point.
    \item Second point.
  \end{itemize}
\end{frame}
\note{Speaker notes appear in the PPTX Notes pane, not in the audience PDF.}
```

Keep each body bullet to a short, scannable line; put detail in `\note{...}` or sub-bullets.

## How the PPTX export works

`beamer_to_pptx.py` parses the `.tex` into slide structures (title, bullets, columns, images, notes) and rebuilds each slide using editable PowerPoint objects.

For frames that use `\begin{tikzpicture}` or `\begin{tabular}` (no direct PPTX equivalent), the script clip-rasterizes just the relevant region of the compiled PDF using the element's own vector bounding box (booktabs rule lines, TikZ paths) and embedded image rects, then writes the crop to `Extracted_figures/slide_content_<N>.png`. Captions, side text, and unrelated bullets are not pulled into the image.

The PPTX output works in PowerPoint, Google Slides, and Keynote.

## Files

| Path | Description |
|------|-------------|
| `presentation.tex` | Main Beamer template (with `\note{...}` per frame) |
| `presentation.pdf` | Compiled PDF, also fed to the converter for TikZ / table cropping |
| `beamer_to_pptx.py` | LaTeX -> editable PPTX converter |
| `requirements.txt` | Python dependencies (PyMuPDF, python-pptx, Pillow) |
| `Paper_files/` | **You** drop the paper PDF here. Starts empty. |
| `Extracted_figures/` | **Agent + converter** write figures here. Starts empty. |
| `Assets/` | Logo, banner, reusable deck graphics |
| `.cursor/rules/` | Cursor rules pinning the generation workflow |
