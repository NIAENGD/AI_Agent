# AI Agent Phase 2

A Windows-focused desktop utility that lets a user pick any open window, take a screenshot, and process it locally with OCR. The GUI is built with wxPython and relies on local tools only. This phase ships with a single Windows PowerShell launcher, `agent.ps1`, that installs everything in-place, keeps itself updated, and always starts the latest version of the app.

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

### One-file Windows setup (`agent.ps1`)
Download **only** `agent.ps1` to the folder where you want the app to live, then double-click it (or run it from PowerShell). Everything happens automatically in subfolders next to the script—no prompts or extra tools required:

- Private Python runtime at `.\\.python`.
- Virtual environment and dependencies in `.venv`.
- Local Tesseract OCR install at `.\\.tesseract` (a system Tesseract install is reused if already present).
- Project source checked out in `ai_agent\\source`.

Behavior:
1. **First run**: downloads the latest code (via Git if available, otherwise a zip), installs Python, creates the virtual environment, installs dependencies, provisions Tesseract, and launches the app.
2. **Subsequent runs**: checks for updates first. If new code is found, it updates the source, refreshes the launcher if needed, restarts itself, revalidates dependencies, and then opens the app so you always use the newest version.

Everything stays self-contained in the folder beside `agent.ps1`, making the launcher the only entry point you need.

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
- On startup the app now performs a dependency check (wxPython, pygetwindow, pyautogui, Pillow, pytesseract, pywin32) and shows a single actionable message if anything is missing—run `pip install -r requirements.txt` to resolve them.
- Screen capture first uses Win32's `PrintWindow` via `pywin32` for compatibility with hardware-accelerated windows. If that fails, it falls back to `pyautogui` and requires the window to be visible and not minimized.
- OCR accuracy depends on your Tesseract installation and language packs.
