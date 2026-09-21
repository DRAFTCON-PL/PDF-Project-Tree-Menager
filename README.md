# PDF Project Tree Manager
<img width="770" height="859" alt="Screenshot_2" src="https://github.com/user-attachments/assets/2f5a8441-6121-4a1e-8ff1-2bab74386c4f" />

A lightweight Windows desktop tool for organizing and merging PDF files into a project tree — built for anyone who repeatedly assembles the same set of output PDFs (e.g. project documentation, technical drawings, reports) from many source files.

![Python](https://img.shields.io/badge/python-3.x-blue) ![Platform](https://img.shields.io/badge/platform-Windows-lightgrey) ![License](https://img.shields.io/badge/license-Freeware-green)

## Overview

PDF Project Tree Manager lets you define several **output PDF sets** ("targets"), each built by merging a list of **source PDF files** in a chosen order. You arrange everything visually in a tree, drag & drop PDFs straight from Windows Explorer, and merge all sets with one click into a shared output directory.

Projects are saved as JSON files, so you can reopen and re-merge them any time — the app also tracks whether each source file has changed since it was added, warning you before you merge stale content.

## Features

- **Project tree view** — one shared output directory containing multiple named output PDF files, each with an ordered list of source PDFs.
- **Drag & drop** — reorder files within the tree, move them between output sets, or drop PDF files directly from Windows Explorer.
- **Live PDF preview** — first-page thumbnail of the selected source file, cached and re-rendered on demand.
- **File status tracking** — color-coded indicators (green / orange / red) show whether a source file is up to date, modified since it was added, or missing on disk.
- **One-click merge** — merges every output set in the background with a progress bar, without freezing the UI.
- **Date suffix automation** — optionally appends a `YYYY.MM.DD` date to output file names, with a built-in date picker.
- **Undo / Redo** — full undo/redo stack for tree edits (add, delete, move, rename, sort).
- **Sorting** — sort selected files alphabetically (A-Z / Z-A) within a set.
- **File details** — inspect page size (with standard paper format detection: A0–A5, Letter, Legal), modification status, and full path.
- **Project save/load** — projects are stored as human-readable `.json` files, including the shared output directory and file modification timestamps.
- **Bilingual interface** — switch between Polish and English at any time via a flag button in the toolbar; the whole UI (menus, dialogs, tooltips, messages) updates instantly.
- **Light/Dark theme** — follows the Windows system appearance automatically.

## Requirements

- Windows
- Python 3.x

Required third-party packages are installed automatically on first run if missing:

- `customtkinter` — modern UI widgets
- `pypdf` — PDF merging
- `tkinterdnd2` — drag & drop from Explorer
- `PyMuPDF` (`pymupdf`) — PDF page rendering for previews
- `Pillow` — image handling
- `tkcalendar` — date picker widget

## Getting started

```bash
python "PDF Project Tree Menager_EN_02.py"
```

On first launch, the app checks for the required packages and installs any that are missing via `pip`.

## Basic workflow

1. Set a **shared output directory** (🎯) where the merged PDFs will be saved.
2. Add one or more **output targets** (🆕) — these are the final PDF files you want to produce.
3. Add **source PDF files** (➕) to each target, or drag them in directly from Windows Explorer.
4. Reorder files by dragging them within the tree; use **Sort (A-Z/Z-A)** from the right-click menu if needed.
5. Click **Merge sets** (🔄) to generate all output PDFs in one go.
6. Save the project (💾) to continue working on it later.

## License

Freeware — free to use.

## Feedback & updates

Website: [www.draftcon.pl](https://www.draftcon.pl)
Email: biuro@draftcon.pl

If you'd like to share feedback or request a newer version of the program, feel free to reach out.
