# Amesmarkdown

`Amesmarkdown` is a local macOS-friendly Python app that converts Microsoft Office files, PDF documents, and Markdown across common document formats. It includes:

- A CLI for converting individual files or whole folders
- A small GUI for choosing files/folders, previewing Markdown, and saving results locally
- A menu bar app for quick conversions from the macOS status bar
- Asset extraction so embedded images are exported into `assets/<document-name>/`

## Supported input

- `.md`
- `.docx`
- `.pdf`
- `.pptx`
- `.xlsx`

## Output behavior

- Converts `.docx`, `.pdf`, `.pptx`, and `.xlsx` into `.md`
- Converts `.md` into `.docx`, `.pptx`, `.xlsx`, or `.pdf`
- Exports embedded images into an `assets/` folder next to the Markdown file
- Preserves common structure:
  - headings
  - lists
  - tables
  - images

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

## Command Names

Preferred commands:

- `amesmarkdown`
- `amesmarkdown-gui`
- `amesmarkdown-menubar`

Legacy `mdconvert`, `mdconvert-gui`, and `mdconvert-menubar` commands are still available as aliases.

## CLI

Single file:

```bash
amesmarkdown input.docx output.md
```

Markdown to Word:

```bash
amesmarkdown notes.md report.docx
```

Markdown to PowerPoint using an output directory:

```bash
amesmarkdown slides.md ./exports --to pptx
```

Single file to output directory:

```bash
amesmarkdown input.docx ./output_dir
```

Folder mode:

```bash
amesmarkdown ./input_folder ./output_folder
```

Markdown folder to PDF folder:

```bash
amesmarkdown ./markdown_folder ./pdf_output --to pdf
```

PDF example:

```bash
amesmarkdown input.pdf output.md
```

## GUI

```bash
amesmarkdown-gui
```

The GUI includes:

- file picker
- folder picker
- output folder selector
- output format selector
- default output folder button
- preview pane
- conversion status log

The main window also shows a short supported-format hint so users can quickly see that Markdown, Word, PowerPoint, Excel, and PDF files are accepted.

Notes:

- This build uses Tkinter to stay simple and local.
- The GUI window is fixed-size so accidental resize gestures do not change the layout.
- On some Homebrew Python builds, Tkinter may not be bundled. If `amesmarkdown-gui` reports missing Tk support, either install a Python distribution with Tk support or use the CLI.
- Drag-and-drop is left out of the initial version to keep setup lightweight. If you want, we can add `tkinterdnd2` next.

## Menu Bar App

```bash
amesmarkdown-menubar
```

The menu bar app adds a lightweight macOS status bar entry with:

- `Convert File...`
- `Convert Folder...`
- `Open Output Folder`
- `Quit`

It reuses the same conversion service as the CLI and GUI, so output behavior stays consistent across all three modes.

## macOS App Launcher

The project now includes a minimal macOS app bundle at `macos/Amesmarkdown.app`.

After the project folder is renamed to `Amesmarkdown` and the app bundle is copied into `/Applications`, you can launch the GUI from Finder, Spotlight, Launchpad, or the Applications folder without opening Terminal.

## Project layout

```text
macos/
  Amesmarkdown.app/
src/mdconvert_app/
  cli.py
  gui.py
  service.py
  converters/
tests/
```
