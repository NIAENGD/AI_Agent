"""Web-based UI for window capture and OCR.

This version replaces the original wxPython desktop app with a Flask-powered
web interface. It exposes the same capabilities—selecting a window, capturing
its contents, optionally cropping the capture, and running OCR—without any
fullscreen mode. The server listens on a browser-safe port (default 8000) and binds to 0.0.0.0 so it can
be reached over the network.
"""
from __future__ import annotations

import io
import importlib
import importlib.util
import os
import socket
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple


def _install_requirements() -> bool:
    requirements_path = Path(__file__).resolve().parent.parent / "requirements.txt"
    if not requirements_path.exists():
        return False

    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "-r", str(requirements_path)],
            check=True,
            capture_output=True,
            text=True,
        )
    except subprocess.CalledProcessError:
        return False

    return True


def _ensure_flask_installed() -> None:
    if importlib.util.find_spec("flask") is not None:
        return

    if _install_requirements() and importlib.util.find_spec("flask") is not None:
        return

    raise RuntimeError("Flask is not installed. Please run: pip install -r requirements.txt")


_ensure_flask_installed()

from flask import (
    Flask,
    jsonify,
    render_template_string,
    request,
    send_file,
)

try:  # type: ignore
    import pygetwindow as gw
except Exception:  # pragma: no cover - optional dependency
    gw = None

try:  # type: ignore
    import pyautogui
except Exception:  # pragma: no cover - optional dependency
    pyautogui = None

try:  # type: ignore
    import pytesseract
except Exception:  # pragma: no cover - optional dependency
    pytesseract = None

try:  # type: ignore
    from PIL import Image
except Exception:  # pragma: no cover - optional dependency
    Image = None


@dataclass
class SelectedWindow:
    title: str
    left: int
    top: int
    width: int
    height: int
    hwnd: Optional[int] = None

    @property
    def region(self) -> tuple[int, int, int, int]:
        return self.left, self.top, self.width, self.height


@dataclass
class AppState:
    selected_window: Optional[SelectedWindow] = None
    captured_image: Optional["Image.Image"] = None
    crop_box: Optional[Tuple[int, int, int, int]] = None
    tesseract_path: Optional[str] = None


state = AppState()
app = Flask(__name__)


# ---------- Dependency handling ----------

def _detect_dependency_state() -> Dict[str, bool]:
    return {
        "pygetwindow": gw is not None,
        "pyautogui": pyautogui is not None,
        "pillow": Image is not None,
        "pytesseract": pytesseract is not None,
        "pywin32": _has_pywin32(),
    }


def _has_pywin32() -> bool:
    required_modules = ("win32gui", "win32ui", "win32con")
    return all(importlib.util.find_spec(module) is not None for module in required_modules)


def _detect_local_tesseract() -> Optional[str]:
    repo_root = Path(__file__).resolve().parent.parent
    candidates = [repo_root / ".tesseract" / "tesseract.exe"]

    program_files = os.environ.get("PROGRAMFILES")
    if program_files:
        candidates.append(Path(program_files) / "Tesseract-OCR" / "tesseract.exe")

    program_files_x86 = os.environ.get("PROGRAMFILES(X86)")
    if program_files_x86:
        candidates.append(Path(program_files_x86) / "Tesseract-OCR" / "tesseract.exe")

    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return None


def _apply_tesseract_path() -> None:
    if pytesseract and state.tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = state.tesseract_path


def _ensure_dependency(key: str, friendly_name: str) -> Optional[str]:
    """Return an error message if a dependency is missing; otherwise None."""

    missing = _detect_dependency_state()
    if missing.get(key, False):
        return None

    if _attempt_install_requirements():
        missing = _detect_dependency_state()
        if missing.get(key, False):
            return None

    return (
        f"{friendly_name} is not installed. Please run: pip install -r requirements.txt"
    )


def _attempt_install_requirements() -> bool:
    if not _install_requirements():
        return False

    _refresh_optional_dependencies()
    return True


def _refresh_optional_dependencies() -> None:
    global gw, pyautogui, pytesseract, Image

    if gw is None:
        try:
            gw = importlib.import_module("pygetwindow")
        except Exception:  # pragma: no cover
            gw = None
    if pyautogui is None:
        try:
            pyautogui = importlib.import_module("pyautogui")
        except Exception:  # pragma: no cover
            pyautogui = None
    if pytesseract is None:
        try:
            pytesseract = importlib.import_module("pytesseract")
        except Exception:  # pragma: no cover
            pytesseract = None
    if Image is None:
        try:
            from PIL import Image as PilImage

            Image = PilImage
        except Exception:  # pragma: no cover
            Image = None


# ---------- Window helpers ----------

def _list_windows() -> List[SelectedWindow]:
    if gw is None:
        return []
    titles: List[str] = [title for title in gw.getAllTitles() if title.strip()]
    titles.sort()
    windows: List[SelectedWindow] = []
    for title in titles:
        window = gw.getWindowsWithTitle(title)[0]
        windows.append(
            SelectedWindow(
                title=title,
                left=window.left,
                top=window.top,
                width=window.width,
                height=window.height,
                hwnd=getattr(window, "_hWnd", None),
            )
        )
    return windows


def _capture_selected_window(selection: SelectedWindow) -> Optional["Image.Image"]:
    if not selection:
        return None
    if Image is None:
        raise RuntimeError("Install Pillow to capture windows.")

    if selection.hwnd:
        try:
            win32_image = _capture_hwnd(selection.hwnd, selection.width, selection.height)
            if win32_image:
                return win32_image
        except Exception:
            # Fall through to pyautogui
            pass

    if pyautogui is None:
        raise RuntimeError("Install pyautogui to capture windows without Win32 support.")

    left, top, width, height = selection.region
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    return screenshot


def _capture_hwnd(hwnd: int, width: int, height: int) -> Optional["Image.Image"]:
    if not _has_pywin32():
        return None

    import win32con  # type: ignore
    import win32gui  # type: ignore
    import win32ui  # type: ignore

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    if not hwnd_dc:
        return None

    try:
        window_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        mem_dc = window_dc.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(window_dc, width, height)
        mem_dc.SelectObject(bitmap)

        result = win32gui.PrintWindow(hwnd, mem_dc.GetSafeHdc(), 0)
        if result != 1:
            return None

        bmpinfo = bitmap.GetInfo()
        bmpstr = bitmap.GetBitmapBits(True)
        image = Image.frombuffer(
            "RGB",
            (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
            bmpstr,
            "raw",
            "BGRX",
            0,
            1,
        )
        return image.crop((0, 0, width, height))
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        mem_dc.DeleteDC()
        window_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)


# ---------- OCR ----------

def _run_ocr(image: "Image.Image") -> str:
    if pytesseract is None:
        raise RuntimeError("pytesseract is not installed.")

    return pytesseract.image_to_string(image).strip()


# ---------- API routes ----------


@app.route("/api/windows")
def api_windows() -> tuple[str, int]:
    if message := _ensure_dependency("pygetwindow", "pygetwindow"):
        return jsonify({"error": message}), 400

    windows = _list_windows()
    return jsonify({"windows": [w.title for w in windows]})


@app.route("/api/settings", methods=["POST"])
def api_settings() -> tuple[str, int]:
    data = request.get_json(silent=True) or {}
    path = data.get("tesseract_path")
    state.tesseract_path = path or _detect_local_tesseract()
    _apply_tesseract_path()
    return jsonify({"message": "Settings saved.", "tesseract_path": state.tesseract_path})


@app.route("/api/capture", methods=["POST"])
def api_capture() -> tuple[str, int]:
    if message := _ensure_dependency("pillow", "Pillow"):
        return jsonify({"error": message}), 400

    data = request.get_json(silent=True) or {}
    title = data.get("title")
    if not title:
        return jsonify({"error": "No window title provided."}), 400

    windows = _list_windows()
    match = next((w for w in windows if w.title == title), None)
    if match is None:
        return jsonify({"error": "Window not found. Refresh the list and try again."}), 404

    try:
        screenshot = _capture_selected_window(match)
    except Exception as exc:  # pragma: no cover - user environment specific
        return jsonify({"error": str(exc)}), 500

    if screenshot is None:
        return jsonify({"error": "Capture failed."}), 500

    state.selected_window = match
    state.captured_image = screenshot
    state.crop_box = None
    return jsonify({"message": "Capture ready."})


@app.route("/api/ocr", methods=["POST"])
def api_ocr() -> tuple[str, int]:
    if state.captured_image is None:
        return jsonify({"error": "Take a capture first."}), 400

    if message := _ensure_dependency("pytesseract", "pytesseract"):
        return jsonify({"error": message}), 400

    data = request.get_json(silent=True) or {}
    crop = data.get("crop")

    image_for_ocr = state.captured_image
    if crop:
        try:
            l, t, r, b = [int(value) for value in (crop.get("left"), crop.get("top"), crop.get("right"), crop.get("bottom"))]
            image_for_ocr = image_for_ocr.crop((l, t, r, b))
            state.crop_box = (l, t, r, b)
        except Exception as exc:
            return jsonify({"error": f"Invalid crop: {exc}"}), 400
    else:
        state.crop_box = None

    try:
        text = _run_ocr(image_for_ocr)
    except Exception as exc:  # pragma: no cover - user environment specific
        return jsonify({"error": str(exc)}), 500

    return jsonify({"text": text})


@app.route("/image")
def image() -> "Response":
    if state.captured_image is None:
        return jsonify({"error": "No capture available."}), 404

    buffer = io.BytesIO()
    state.captured_image.save(buffer, format="PNG")
    buffer.seek(0)
    return send_file(buffer, mimetype="image/png")


@app.route("/")
def index() -> str:
    template = _PAGE_TEMPLATE
    current_path = state.tesseract_path or ""
    crop_box = state.crop_box
    return render_template_string(
        template,
        tesseract_path=current_path,
        crop_box=crop_box,
        has_capture=state.captured_image is not None,
    )


# ---------- Page template ----------


_PAGE_TEMPLATE = """
<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>AI Agent - Web Capture & OCR</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 0; padding: 0; background: #f5f5f5; color: #222; }
    header { background: #0a84ff; color: white; padding: 16px 24px; }
    main { padding: 20px; max-width: 1200px; margin: 0 auto; }
    section { background: white; border-radius: 10px; padding: 16px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.12); }
    h1 { margin: 0 0 6px; }
    h2 { margin-top: 0; }
    button { background: #0a84ff; color: white; border: none; padding: 10px 16px; border-radius: 8px; cursor: pointer; font-size: 14px; }
    button.secondary { background: #555; }
    button:disabled { background: #b0c9ff; cursor: not-allowed; }
    label { display: block; margin-bottom: 6px; font-weight: bold; }
    input, select, textarea { width: 100%; padding: 8px; border-radius: 6px; border: 1px solid #ccc; box-sizing: border-box; font-size: 14px; }
    .row { display: flex; gap: 16px; flex-wrap: wrap; }
    .col { flex: 1 1 300px; }
    #status { margin-top: 8px; color: #0a84ff; font-weight: bold; }
    #preview-container { position: relative; display: inline-block; }
    #crop-overlay { position: absolute; border: 2px dashed #ff5c00; background: rgba(255, 92, 0, 0.2); display: none; pointer-events: none; }
    #captureImage { max-width: 100%; border-radius: 8px; border: 1px solid #ddd; }
    pre { background: #f0f0f0; padding: 12px; border-radius: 8px; white-space: pre-wrap; }
  </style>
</head>
<body>
  <header>
    <h1>AI Agent - Web Capture & OCR</h1>
    <div>Workflow: Select a window → Capture → (Optional) Crop → OCR.</div>
  </header>
  <main>
    <section>
      <div class=\"row\">
        <div class=\"col\">
          <h2>Window selection</h2>
          <label for=\"windowSelect\">Open windows</label>
          <select id=\"windowSelect\"></select>
          <div style=\"margin-top:8px; display:flex; gap:8px; flex-wrap: wrap;\">
            <button id=\"refreshBtn\" type=\"button\">Refresh</button>
            <button id=\"captureBtn\" type=\"button\">Capture</button>
          </div>
        </div>
        <div class=\"col\">
          <h2>OCR settings</h2>
          <label for=\"tesseractPath\">Tesseract executable</label>
          <input id=\"tesseractPath\" type=\"text\" placeholder=\"Path to tesseract.exe\" value=\"{{ tesseract_path }}\" />
          <div style=\"margin-top:8px;\">
            <button id=\"saveSettingsBtn\" type=\"button\">Save settings</button>
          </div>
        </div>
      </div>
      <div id=\"status\"></div>
    </section>

    <section>
      <h2>Preview & crop</h2>
      {% if has_capture %}
      <div id=\"preview-container\">
        <img id=\"captureImage\" src=\"/image\" alt=\"Capture preview\" />
        <div id=\"crop-overlay\"></div>
      </div>
      {% else %}
      <div>No capture yet. Select a window and click Capture.</div>
      <div id=\"preview-container\" style=\"display:none;\">
        <img id=\"captureImage\" src=\"\" alt=\"Capture preview\" />
        <div id=\"crop-overlay\"></div>
      </div>
      {% endif %}
      <div style=\"margin-top: 12px; display:flex; gap:8px; flex-wrap:wrap;\">
        <button id=\"clearCropBtn\" type=\"button\" class=\"secondary\">Clear crop</button>
        <button id=\"ocrBtn\" type=\"button\">Run OCR</button>
      </div>
      <div id=\"cropInfo\" style=\"margin-top:6px; color:#444;\"></div>
    </section>

    <section>
      <h2>OCR output</h2>
      <pre id=\"ocrOutput\"></pre>
    </section>
  </main>

  <script>
    let cropBox = null;
    let naturalWidth = 0;
    let naturalHeight = 0;

    async function refreshWindows() {
      const res = await fetch('/api/windows');
      const status = document.getElementById('status');
      if (!res.ok) {
        const data = await res.json();
        status.textContent = data.error || 'Unable to load windows.';
        return;
      }
      const data = await res.json();
      const select = document.getElementById('windowSelect');
      select.innerHTML = '';
      data.windows.forEach(title => {
        const option = document.createElement('option');
        option.value = title;
        option.textContent = title;
        select.appendChild(option);
      });
      status.textContent = `Found ${data.windows.length} window(s).`;
    }

    async function saveSettings() {
      const path = document.getElementById('tesseractPath').value;
      const res = await fetch('/api/settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tesseract_path: path })
      });
      const data = await res.json();
      const status = document.getElementById('status');
      status.textContent = data.message || data.error || 'Settings updated';
    }

    async function capture() {
      const select = document.getElementById('windowSelect');
      const title = select.value;
      const status = document.getElementById('status');
      if (!title) {
        status.textContent = 'Select a window first.';
        return;
      }

      const res = await fetch('/api/capture', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ title })
      });
      const data = await res.json();
      if (!res.ok) {
        status.textContent = data.error || 'Capture failed.';
        return;
      }

      status.textContent = data.message || 'Capture ready.';
      const img = document.getElementById('captureImage');
      img.src = '/image?cache=' + Date.now();
      img.onload = () => {
        naturalWidth = img.naturalWidth;
        naturalHeight = img.naturalHeight;
      };
      document.getElementById('preview-container').style.display = 'inline-block';
      clearCrop();
    }

    async function runOcr() {
      const status = document.getElementById('status');
      const payload = cropBox ? { crop: cropBox } : {};
      const res = await fetch('/api/ocr', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (!res.ok) {
        status.textContent = data.error || 'OCR failed.';
        return;
      }
      status.textContent = 'OCR complete.';
      document.getElementById('ocrOutput').textContent = data.text || '';
    }

    function clearCrop() {
      cropBox = null;
      document.getElementById('cropInfo').textContent = 'No crop set; OCR will use the full image.';
      const overlay = document.getElementById('crop-overlay');
      overlay.style.display = 'none';
    }

    function setupCropping() {
      const img = document.getElementById('captureImage');
      const overlay = document.getElementById('crop-overlay');
      let start = null;

      img.addEventListener('load', () => {
        naturalWidth = img.naturalWidth;
        naturalHeight = img.naturalHeight;
      });

      img.addEventListener('mousedown', (ev) => {
        if (!naturalWidth || !naturalHeight) return;
        const rect = img.getBoundingClientRect();
        start = { x: ev.clientX - rect.left, y: ev.clientY - rect.top, rect };
        overlay.style.display = 'block';
        overlay.style.left = `${start.x}px`;
        overlay.style.top = `${start.y}px`;
        overlay.style.width = '1px';
        overlay.style.height = '1px';
      });

      img.addEventListener('mousemove', (ev) => {
        if (!start) return;
        const rect = start.rect;
        const x = ev.clientX - rect.left;
        const y = ev.clientY - rect.top;
        const left = Math.min(start.x, x);
        const top = Math.min(start.y, y);
        const width = Math.abs(x - start.x);
        const height = Math.abs(y - start.y);
        overlay.style.left = `${left}px`;
        overlay.style.top = `${top}px`;
        overlay.style.width = `${width}px`;
        overlay.style.height = `${height}px`;
      });

      img.addEventListener('mouseup', (ev) => {
        if (!start) return;
        const rect = start.rect;
        const x = ev.clientX - rect.left;
        const y = ev.clientY - rect.top;
        const left = Math.min(start.x, x);
        const top = Math.min(start.y, y);
        const width = Math.abs(x - start.x);
        const height = Math.abs(y - start.y);
        start = null;

        if (width < 10 || height < 10) {
          clearCrop();
          return;
        }

        const scaleX = naturalWidth / img.clientWidth;
        const scaleY = naturalHeight / img.clientHeight;
        cropBox = {
          left: Math.round(left * scaleX),
          top: Math.round(top * scaleY),
          right: Math.round((left + width) * scaleX),
          bottom: Math.round((top + height) * scaleY)
        };
        document.getElementById('cropInfo').textContent = `Crop: (${cropBox.left}, ${cropBox.top}) → (${cropBox.right}, ${cropBox.bottom})`;
      });
    }

    document.getElementById('refreshBtn').addEventListener('click', refreshWindows);
    document.getElementById('captureBtn').addEventListener('click', capture);
    document.getElementById('ocrBtn').addEventListener('click', runOcr);
    document.getElementById('saveSettingsBtn').addEventListener('click', saveSettings);
    document.getElementById('clearCropBtn').addEventListener('click', clearCrop);

    setupCropping();
    refreshWindows();
    clearCrop();
  </script>
</body>
</html>
"""


# ---------- App entry ----------


def _is_unsafe_browser_port(port: int) -> bool:
    """Return True if modern Chromium-based browsers commonly block this port.

    Chrome/Edge intentionally block several ports (including 6000, used by X11),
    which surfaces in the browser as ERR_UNSAFE_PORT even if Flask is running.
    """
    if port <= 0 or port > 65535:
        return True
    # X11 ports (6000-6063) are blocked by Chromium.
    if 6000 <= port <= 6063:
        return True
    # Commonly blocked legacy/IRC ports.
    if port in {6665, 6666, 6667, 6668, 6669}:
        return True
    return False


def _is_port_available(host: str, port: int) -> bool:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind((host, port))
        return True
    except OSError:
        return False


def _choose_bind(host: str, requested_port: int) -> tuple[str, int, Optional[str]]:
    """Return (host, port, warning_message)."""
    warning: Optional[str] = None

    port = requested_port
    if _is_unsafe_browser_port(port):
        warning = (
            f"Requested port {port} is blocked by Chrome/Edge (ERR_UNSAFE_PORT). "
            "Switching to 8000."
        )
        port = 8000

    # If the chosen port is already taken, scan forward a bit.
    if port != 0 and not _is_port_available(host, port):
        for candidate in range(port + 1, port + 51):
            if _is_port_available(host, candidate) and not _is_unsafe_browser_port(candidate):
                if warning:
                    warning += f" (Port {port} was in use; using {candidate} instead.)"
                else:
                    warning = f"Port {port} was in use; using {candidate} instead."
                port = candidate
                break
        else:
            # Let the OS pick a free port.
            if warning:
                warning += " (Falling back to an OS-assigned free port.)"
            else:
                warning = "Falling back to an OS-assigned free port."
            port = 0

    return host, port, warning


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument(
        "--host",
        default=os.environ.get("AI_AGENT_HOST", "0.0.0.0"),
        help="Bind host (default: 0.0.0.0)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.environ.get("AI_AGENT_PORT", os.environ.get("PORT", "8000"))),
        help="Bind port (default: 8000)",
    )
    parser.add_argument(
        "--public",
        action="store_true",
        help="Alias for --host 0.0.0.0 (useful for LAN access)",
    )

    args, _unknown = parser.parse_known_args()

    host = "0.0.0.0" if args.public else args.host
    host, port, warning = _choose_bind(host, args.port)

    if warning:
        print(warning, file=sys.stderr)

    state.tesseract_path = _detect_local_tesseract()
    _apply_tesseract_path()
    app.run(host=host, port=port)
    return 0



if __name__ == "__main__":
    sys.exit(main())
