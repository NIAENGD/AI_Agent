# AI Agent Phase 2

A Windows-focused desktop utility that lets a user pick any open window, take a screenshot, and process it locally with OCR. The GUI is built with PyQt5 and relies on local tools only. Phase 2 adds a single Windows `.bat` helper to set up, update, and run the project in one place.

## Features
- **Start/Settings hub**: initial view with Start (window selection) and Settings (Tesseract path) buttons.
- **Window selection**: lists current top-level windows on Windows using `pygetwindow`.
- **Capture & preview**: "Take" captures the selected window and shows it in a preview pane.
- **Local OCR**: "Process" runs OCR on the captured image with `pytesseract` (Tesseract CLI required). Intended to read text, charts, math, and other on-screen content.

## Requirements
- Windows 10/11
- Python 3.10+
- Tesseract OCR installed locally (https://github.com/tesseract-ocr/tesseract). Note the installation path (e.g., `C:\Program Files\Tesseract-OCR\tesseract.exe`).
- Dependencies listed in `requirements.txt`.

### One-file Windows setup (.bat)
Use the included `windows_setup.bat` file to handle setup, updates, and running the app:
1. Place `windows_setup.bat` in the repository root (next to `app/` and `requirements.txt`). If you only downloaded the script, drop it into the cloned/extracted project folder.
2. Double-click the file or run it from `cmd` with `windows_setup.bat`.
3. Choose from the menu:
   - **Create or refresh virtual environment**: sets up `.venv` with the correct Python interpreter.
   - **Install/Update Python dependencies**: installs from `requirements.txt` inside `.venv`.
   - **Update project from Git (pull)**: grabs the latest code if Git is available.
   - **Run AI Agent**: activates `.venv` (installing dependencies if missing) and launches `app\main.py`.
   - **Full setup (venv + deps + run)**: performs install steps and starts the app in one go.
   - **Clean __pycache__ folders**: removes Python bytecode caches.

The script verifies Python 3.10+ is on `PATH`. Git is only required for the "Update project" option. After the app launches, use **Settings** to point to `tesseract.exe` if it is not on `PATH`.

### Manual install (without the .bat helper)
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
- On startup the app now performs a dependency check (PyQt5, pygetwindow, pyautogui, Pillow, pytesseract) and shows a single
  actionable message if anything is missing—run `pip install -r requirements.txt` to resolve them.
- Screen capture relies on `pyautogui`; ensure the selected window is not minimized.
- OCR accuracy depends on your Tesseract installation and language packs.
