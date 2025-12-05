# AI Agent Phase 1

A Windows-focused desktop utility that lets a user pick any open window, take a screenshot, and process it locally with OCR. The GUI is built with PyQt5 and relies on local tools only.

## Features
- **Start/Settings hub**: initial view with Start (window selection) and Settings (Tesseract path) buttons.
- **Window selection**: lists current top-level windows on Windows using `pygetwindow`.
- **Capture & preview**: "Take" captures the selected window and shows it in a preview pane.
- **Local OCR**: "Process" runs OCR on the captured image with `pytesseract` (Tesseract CLI required). Intended to read text, charts, math, and other on-screen content.

## Requirements
- Windows 10/11
- Python 3.10+
- Tesseract OCR installed locally (https://github.com/tesseract-ocr/tesseract). Note the installation path (e.g., `C:\\Program Files\\Tesseract-OCR\\tesseract.exe`).
- Dependencies listed in `requirements.txt`.

Install dependencies:
```bash
python -m pip install -r requirements.txt
```

## Running the app
From the repository root:
```bash
python app/main.py
```

### Workflow
1. Click **Start** and pick a window from the list.
2. Click **Take** to capture that window. The preview updates on the right.
3. Click **Process** to run local OCR. Results appear in a dialog (full text in the Details section).
4. Use **Settings** to set the `tesseract.exe` path if it is not on `PATH`.

## Notes
- The app degrades gracefully if dependencies are missing by prompting the user to install them.
- Screen capture relies on `pyautogui`; ensure the selected window is not minimized.
- OCR accuracy depends on your Tesseract installation and language packs.
