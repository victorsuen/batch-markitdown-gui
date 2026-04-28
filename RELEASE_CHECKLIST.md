# Release Checklist

Use this checklist for each new release.

## 1) Prepare Version

- [ ] Decide next version (example: `v1.11`)
- [ ] Update `APP_VERSION` in `batch_markitdown_gui.py`
- [ ] Update `CHANGELOG.md` with release notes and date

## 2) Local Validation

- [ ] Install dependencies:
  - `pip install -r requirements.txt`
- [ ] Run GUI locally:
  - `python batch_markitdown_gui.py`
- [ ] Verify key flows:
  - [ ] Drag/drop files and folders
  - [ ] OCR fallback for scanned PDF
  - [ ] Traditional Chinese formatting
  - [ ] Financial table mode output
  - [ ] Timeout handling

## 3) Build EXE

- [ ] Build executable:
  - `python -m PyInstaller --noconfirm --clean --windowed --onefile --name BatchMarkItDownGUI --icon batch_markitdown_icon.ico --collect-all tkinterdnd2 --collect-all markitdown --collect-all magika --collect-all rapidocr_onnxruntime batch_markitdown_gui.py`
- [ ] Confirm output exists:
  - `dist/BatchMarkItDownGUI.exe`
- [ ] Smoke-test EXE on desktop

## 4) Commit & Push

- [ ] Commit code/docs for release
- [ ] Push to `main`

## 5) GitHub Release

- [ ] Create tag/release (example `v1.11`)
- [ ] Upload EXE asset (rename clearly, e.g. `BatchMarkItDownGUI_v111.exe`)
- [ ] Add release notes:
  - Highlights
  - Fixes
  - Known limitations

## 6) Post-Release Verification

- [ ] Open release page and verify asset downloadable
- [ ] Verify README still matches latest usage/build steps
- [ ] Announce release link to users

