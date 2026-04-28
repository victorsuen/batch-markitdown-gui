# Batch MarkItDown GUI

A desktop-friendly batch converter that exports supported files to Markdown using [Microsoft MarkItDown](https://github.com/microsoft/markitdown).

This project includes:

- `batch_markitdown.py` - CLI batch converter
- `batch_markitdown_gui.py` - Windows GUI with drag-and-drop
- Custom app icons (`.ico`, `.png`)

## Features

- Batch convert supported files to `.md`
- Drag-and-drop files/folders in GUI
- Progress, logs, and timeout control
- Remember output folder option
- OCR fallback for scanned/image-based PDFs
- Traditional Chinese OCR text cleanup
- Financial table mode for better numeric row formatting

## Supported Input Types

Depends on MarkItDown installed extras, commonly:

- PDF, DOCX, PPTX, XLSX, XLS
- CSV, JSON, XML, TXT, HTML/HTM, EPUB, ZIP, MD

## Requirements

- Python 3.10+
- Windows (GUI tested on Windows)

Install dependencies:

```bash
pip install -r requirements.txt
```

## CLI Usage

```bash
python batch_markitdown.py "D:\input-folder"
```

Optional parameters:

```bash
python batch_markitdown.py "D:\input-folder" -o "D:\output-folder" --no-recursive --use-plugins
```

Default output folder is Desktop.

## GUI Usage

```bash
python batch_markitdown_gui.py
```

In GUI you can:

- Select source/output folders
- Drag files/folders to drop area
- Enable fast mode (`convert_local`)
- Set per-file timeout
- Enable financial table mode

## Build EXE (PyInstaller)

```bash
python -m PyInstaller --noconfirm --clean --windowed --onefile --name BatchMarkItDownGUI --icon batch_markitdown_icon.ico --collect-all tkinterdnd2 --collect-all markitdown --collect-all magika --collect-all rapidocr_onnxruntime batch_markitdown_gui.py
```

Output EXE will be under:

`dist/BatchMarkItDownGUI.exe`

## Download EXE

Use the latest binary from GitHub Releases.

## Release Process

See `RELEASE_CHECKLIST.md` for the step-by-step release workflow.

## Notes

- Scanned PDFs with no text layer require OCR; this app auto-fallbacks to OCR for such PDFs.
- If conversion appears slow, reduce timeout or keep fast mode enabled.
- For huge repositories, avoid committing virtual environments and build artifacts.

## License

MIT

