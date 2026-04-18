# Notes for AI assistants working in this repo

## Where figures from the paper live

Slide figures are **not** only "referenced in the text" of `presentation.tex`. The TeX file uses **filename-only** includes (e.g. `\includegraphics{figure1}`) and `\graphicspath` in `presentation.tex` tells LaTeX (and the converter) where to look.

**Always check `Paper_files/` first** for image files the author placed next to the paper: PNG, JPG, JPEG, PDF. List that directory (or `glob` `Paper_files/*`) when you need to know what exists. The author is expected to drop pulled-out figures there alongside the paper PDF.

**Then** `Extracted_figures/`: the converter copies every referenced figure here after a run. If a name appears in `\includegraphics` but the file is missing, it may still be in `Paper_files/` under a different name; ask the user or match by stem.

**`Assets/`** is for logo, banner, and reusable deck chrome, not paper figures.

## Bullets (when you edit `presentation.tex`)

**Must:** keep each frame body `\item` **short** (about one line at 16:9; no paragraph-style bullets). **Must** put detail, nuance, caveats, and asides in `\note{...}` right after `\end{frame}`. If a point is still dense, use sub-bullets; do not stretch a single `\item` into a long block of text.
