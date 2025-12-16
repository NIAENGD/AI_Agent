"""Web-based UI for window capture and OCR.

This version replaces the original wxPython desktop app with a Flask-powered
web interface. It exposes the same capabilities—selecting a window, capturing
its contents, optionally cropping the capture, and running OCR—without any
fullscreen mode. The server listens on a browser-safe port (default 8000) and binds to 0.0.0.0 so it can
be reached over the network.
"""
from __future__ import annotations

import base64
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

CONFIG_DIR = Path(__file__).resolve().parent / "configs"
CONFIG_FILE = CONFIG_DIR / "prompts.txt"
API_KEY_FILE = CONFIG_DIR / "api_key.txt"
DEFAULT_PROMPT = {"title": "Example", "prompt": "This is an example prompt"}


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


# ---------- Prompt configuration ----------


def _ensure_prompt_store() -> None:
    _ensure_config_dir()
    if not CONFIG_FILE.exists():
        _write_prompt_entries([DEFAULT_PROMPT])


def _write_prompt_entries(entries: List[Dict[str, str]]) -> None:
    lines = [f"{entry['title']};{entry['prompt']};" for entry in entries]
    CONFIG_FILE.write_text("\n".join(lines), encoding="utf-8")


def _ensure_config_dir() -> None:
    CONFIG_DIR.mkdir(exist_ok=True)


def _load_api_key() -> Optional[str]:
    env_key = os.environ.get("OPENAI_API_KEY")
    if env_key and env_key.strip():
        return env_key.strip()

    if API_KEY_FILE.exists():
        key = API_KEY_FILE.read_text(encoding="utf-8").strip()
        if key:
            return key
    return None


def _save_api_key(raw_key: str) -> bool:
    _ensure_config_dir()
    key = (raw_key or "").strip()
    if not key:
        if API_KEY_FILE.exists():
            API_KEY_FILE.unlink()
        return False

    API_KEY_FILE.write_text(key, encoding="utf-8")
    return True


def _parse_prompt_lines(raw: str) -> List[Dict[str, str]]:
    prompts: List[Dict[str, str]] = []
    for line in raw.splitlines():
        parts = [segment.strip() for segment in line.split(";") if segment.strip()]
        if len(parts) >= 2:
            prompts.append({"title": parts[0], "prompt": parts[1]})
    return prompts


def _load_prompt_entries() -> List[Dict[str, str]]:
    _ensure_prompt_store()
    try:
        raw = CONFIG_FILE.read_text(encoding="utf-8")
    except FileNotFoundError:
        _write_prompt_entries([DEFAULT_PROMPT])
        return [DEFAULT_PROMPT]

    prompts = _parse_prompt_lines(raw)
    return prompts


def _upsert_prompt_entry(title: str, prompt: str) -> List[Dict[str, str]]:
    prompts = _load_prompt_entries()
    replacement = {"title": title, "prompt": prompt}
    for idx, entry in enumerate(prompts):
        if entry["title"] == title:
            prompts[idx] = replacement
            break
    else:
        prompts.append(replacement)
    _write_prompt_entries(prompts)
    return prompts


def _delete_prompt_entry(title: str) -> Tuple[List[Dict[str, str]], bool]:
    prompts = _load_prompt_entries()
    new_prompts = [entry for entry in prompts if entry["title"] != title]
    if len(new_prompts) == len(prompts):
        return prompts, False

    _write_prompt_entries(new_prompts)
    return new_prompts, True


def _image_to_data_url(image: "Image.Image") -> str:
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    buffer.seek(0)
    encoded = base64.b64encode(buffer.read()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def _get_openai_client():  # type: ignore[return-type]
    if importlib.util.find_spec("openai") is None:
        return None, "OpenAI SDK is not installed. Please add it to requirements.txt."

    from openai import OpenAI
    try:
        import httpx  # type: ignore
    except Exception:
        httpx = None  # type: ignore

    api_key = _load_api_key()
    if not api_key:
        return None, (
            "Set OPENAI_API_KEY or save a key in configs/api_key.txt via Settings."
        )

    if httpx is not None:
        version_parts = []
        for part in str(getattr(httpx, "__version__", "0")).split("."):
            if part.isdigit():
                version_parts.append(int(part))
        if version_parts and version_parts[0] >= 1:
            return (
                None,
                "Installed httpx>=1.0 is incompatible with the OpenAI SDK here. "
                "Please install httpx<1.0 (pip install 'httpx<1.0').",
            )

    try:
        return OpenAI(api_key=api_key), None
    except TypeError as exc:
        if "proxies" in str(exc):
            return (
                None,
                "Unable to initialize OpenAI client because the httpx version "
                "is incompatible. Please install httpx<1.0.",
            )
        return None, f"Unable to initialize OpenAI client: {exc}"


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

    _refresh_window_bounds(selection)

    if selection.hwnd:
        try:
            win32_image = _capture_hwnd(selection.hwnd)
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


def _refresh_window_bounds(selection: SelectedWindow) -> None:
    if not selection.hwnd or not _has_pywin32():
        return

    import win32gui  # type: ignore

    try:
        left, top, right, bottom = win32gui.GetWindowRect(selection.hwnd)
    except Exception:
        return

    width = max(0, right - left)
    height = max(0, bottom - top)
    if width and height:
        selection.left = left
        selection.top = top
        selection.width = width
        selection.height = height


def _capture_hwnd(hwnd: int) -> Optional["Image.Image"]:
    if not _has_pywin32():
        return None

    import win32con  # type: ignore
    import win32gui  # type: ignore
    import win32ui  # type: ignore

    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    except Exception:
        return None

    width = max(0, right - left)
    height = max(0, bottom - top)
    if not width or not height:
        return None

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    if not hwnd_dc:
        return None

    bitmap = None
    try:
        window_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        mem_dc = window_dc.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(window_dc, width, height)
        mem_dc.SelectObject(bitmap)

        result = win32gui.PrintWindow(hwnd, mem_dc.GetSafeHdc(), win32con.PW_RENDERFULLCONTENT)
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
        if bitmap:
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


@app.route("/api/settings", methods=["GET", "POST"])
def api_settings() -> tuple[str, int]:
    if request.method == "GET":
        if state.tesseract_path is None:
            state.tesseract_path = _detect_local_tesseract()
            _apply_tesseract_path()

        return jsonify(
            {
                "tesseract_path": state.tesseract_path,
                "api_key_present": _load_api_key() is not None,
                "api_key_path": str(API_KEY_FILE),
            }
        )

    data = request.get_json(silent=True) or {}
    path = data.get("tesseract_path")
    state.tesseract_path = path or _detect_local_tesseract()
    _apply_tesseract_path()

    api_key_present = None
    if "api_key" in data:
        api_key_present = _save_api_key(str(data.get("api_key", "")))

    return jsonify(
        {
            "message": "Settings saved.",
            "tesseract_path": state.tesseract_path,
            "api_key_present": api_key_present
            if api_key_present is not None
            else _load_api_key() is not None,
        }
    )


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
        data_url = _image_to_data_url(image_for_ocr)
    except Exception as exc:  # pragma: no cover - user environment specific
        return jsonify({"error": str(exc)}), 500

    return jsonify({"text": text, "image_data_url": data_url})


@app.route("/api/configs", methods=["GET", "POST", "PUT", "DELETE"])
def api_configs() -> tuple[str, int]:
    if request.method == "GET":
        return jsonify({"prompts": _load_prompt_entries()})

    data = request.get_json(silent=True) or {}
    title = (data.get("title") or "").strip()
    prompt = (data.get("prompt") or "").strip()

    if request.method == "DELETE":
        if not title:
            return jsonify({"error": "Title is required to delete a prompt."}), 400
        prompts, removed = _delete_prompt_entry(title)
        if not removed:
            return jsonify({"error": "Prompt not found."}), 404
        return jsonify({"prompts": prompts, "message": f"Prompt '{title}' deleted."})

    if request.method == "PUT":
        original_title = (data.get("original_title") or title).strip()
        if not original_title or not title or not prompt:
            return jsonify({"error": "Original title, new title, and prompt are required."}), 400

        prompts = _load_prompt_entries()
        if not any(entry["title"] == original_title for entry in prompts):
            return jsonify({"error": "Prompt not found."}), 404

        updated = []
        for entry in prompts:
            if entry["title"] == original_title:
                updated.append({"title": title, "prompt": prompt})
            else:
                updated.append(entry)

        _write_prompt_entries(updated)
        return jsonify({"prompts": updated, "message": "Prompt updated."})

    if not title or not prompt:
        return jsonify({"error": "Title and prompt are required."}), 400

    prompts = _upsert_prompt_entry(title, prompt)
    return jsonify({"prompts": prompts, "message": "Prompt saved."})


@app.route("/api/configs/upload", methods=["POST"])
def api_configs_upload() -> tuple[str, int]:
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    uploaded = request.files["file"]
    try:
        content = uploaded.read().decode("utf-8")
    except Exception:
        return jsonify({"error": "Unable to read the uploaded file as text."}), 400

    new_prompts = _parse_prompt_lines(content)
    if not new_prompts:
        return jsonify({"error": "No valid prompts found in the file."}), 400

    prompts = _load_prompt_entries() + new_prompts
    _write_prompt_entries(prompts)
    return jsonify({"prompts": prompts, "message": "Prompts uploaded."})


@app.route("/api/ai_response", methods=["POST"])
def api_ai_response() -> tuple[str, int]:
    data = request.get_json(silent=True) or {}
    queue = data.get("queue") or []
    include_images = bool(data.get("include_images", True))
    prompt_text = (data.get("prompt") or "").strip()

    if not isinstance(queue, list):
        return jsonify({"error": "Queue must be a list."}), 400
    if len(queue) > 10:
        return jsonify({"error": "Queue limit is 10 items per request."}), 400

    texts: List[str] = []
    for item in queue:
        text = str(item.get("text", "")).strip()
        if text:
            texts.append(text)

    combined_text = "\n\n".join(texts)
    if prompt_text and combined_text:
        combined_prompt = f"{prompt_text}\n\n{combined_text}"
    else:
        combined_prompt = prompt_text or combined_text

    if not combined_prompt:
        combined_prompt = "Please review the provided OCR text and images."

    content: List[Dict[str, str]] = [{"type": "input_text", "text": combined_prompt}]

    if include_images:
        for item in queue:
            image_data = item.get("image")
            if image_data:
                content.append({"type": "input_image", "image_url": image_data})

    client, error = _get_openai_client()
    if error:
        return jsonify({"error": error}), 400

    model = os.environ.get("AI_AGENT_MODEL", "gpt-5.2")

    try:
        response = client.responses.create(
            model=model,
            input=[{"role": "user", "content": content}],
        )
    except Exception as exc:  # pragma: no cover - network/env specific
        return jsonify({"error": str(exc)}), 500

    return jsonify({"response": getattr(response, "output_text", "") or ""})


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
    :root {
      --bg: #f2f2f2;
      --panel: #ffffff;
      --border: #cfcfcf;
      --text: #1f1f1f;
      --muted: #6b6b6b;
      --accent: #3d3d3d;
      --accent-strong: #2b2b2b;
      --highlight: #ececec;
    }
    body { font-family: Arial, sans-serif; margin: 0; padding: 0; background: var(--bg); color: var(--text); }
    header { background: var(--accent-strong); color: #f5f5f5; padding: 16px 24px; }
    main { padding: 20px; max-width: 1200px; margin: 0 auto; }
    section { background: var(--panel); border: 1px solid var(--border); border-radius: 10px; padding: 16px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }
    h1 { margin: 0 0 6px; }
    h2 { margin-top: 0; color: var(--accent-strong); }
    button { background: var(--accent); color: #f5f5f5; border: none; padding: 10px 16px; border-radius: 8px; cursor: pointer; font-size: 14px; }
    button.secondary { background: var(--muted); color: #f5f5f5; }
    button:disabled { background: #b5b5b5; cursor: not-allowed; }
    label { display: block; margin-bottom: 6px; font-weight: bold; color: var(--accent-strong); }
    input, select, textarea { width: 100%; padding: 8px; border-radius: 6px; border: 1px solid var(--border); box-sizing: border-box; font-size: 14px; background: #fdfdfd; color: var(--text); }
    textarea { min-height: 90px; }
    .row { display: flex; gap: 16px; flex-wrap: wrap; }
    .col { flex: 1 1 300px; }
    #status { margin-top: 8px; color: var(--accent-strong); font-weight: bold; }
    #preview-container { position: relative; display: inline-block; }
    #crop-overlay { position: absolute; border: 2px dashed #707070; background: rgba(112, 112, 112, 0.2); display: none; pointer-events: none; }
    #captureImage { max-width: 100%; border-radius: 8px; border: 1px solid var(--border); }
    pre { background: var(--highlight); padding: 12px; border-radius: 8px; white-space: pre-wrap; border: 1px solid var(--border); }
    #orientationNotice { display: none; margin-top: 10px; background: #5a5a5a; color: #f5f5f5; padding: 8px 12px; border-radius: 8px; }
    body.portrait-warning #orientationNotice { display: block; }
    .queue-list { list-style: none; padding: 0; margin: 0; display: grid; gap: 8px; }
    .queue-item { border: 1px solid var(--border); padding: 8px; border-radius: 8px; background: var(--highlight); }
    .queue-item-title { font-weight: bold; color: var(--accent-strong); margin-bottom: 4px; }
    .stack { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
  </style>
</head>
<body>
  <header>
    <div style="display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap;">
      <div>
        <h1 style="margin-bottom:4px;">AI Agent - Web Capture & OCR</h1>
        <div>Workflow: Select a window → Capture → (Optional) Crop → OCR.</div>
      </div>
      <div style="display:flex; gap:8px;">
        <button id="fullscreenBtn" type="button" class="secondary">Fullscreen</button>
      </div>
    </div>
    <div id="orientationNotice">Best viewed in landscape on mobile. Rotate your device if needed.</div>
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
          <h2>App settings</h2>
          <label for=\"tesseractPath\">Tesseract executable</label>
          <input id=\"tesseractPath\" type=\"text\" placeholder=\"Path to tesseract.exe\" value=\"{{ tesseract_path }}\" />
          <label for=\"apiKey\" style=\"margin-top:8px;\">OpenAI API key</label>
          <input id=\"apiKey\" type=\"password\" placeholder=\"sk-...\" autocomplete=\"off\" />
          <div id=\"apiKeyStatus\" style=\"margin-top:6px; color: var(--muted); font-size: 13px;\">No API key saved.</div>
          <div style=\"margin-top:4px; color: var(--muted); font-size: 13px;\">Stored in configs/api_key.txt. Leave blank to clear.</div>
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
          <button id=\"startCropBtn\" type=\"button\">Select crop area</button>
          <button id=\"clearCropBtn\" type=\"button\" class=\"secondary\">Clear crop</button>
          <button id=\"ocrBtn\" type=\"button\">Run OCR</button>
        </div>
      <div id=\"cropInfo\" style=\"margin-top:6px; color:#444;\"></div>
    </section>

    <section>
      <h2>OCR output</h2>
      <pre id=\"ocrOutput\"></pre>
      <div class=\"stack\" style=\"margin-top:8px;\">
        <button id=\"queueBtn\" type=\"button\">Save to queue</button>
        <button id=\"clearQueueBtn\" type=\"button\" class=\"secondary\">Clear queue</button>
      </div>
      <div id=\"queueStatus\" style=\"color: var(--muted); margin-top:6px;\"></div>
    </section>

    <section>
      <h2>Prompts &amp; AI response</h2>
      <div class=\"row\">
        <div class=\"col\">
          <label for=\"promptSelect\">Saved prompts</label>
          <select id=\"promptSelect\"></select>
          <div class=\"stack\" style=\"margin-top:8px;\">
            <input id=\"newPromptTitle\" type=\"text\" placeholder=\"New prompt title\" />
            <div style=\"display:flex; gap:8px; flex-wrap:wrap;\">
              <button id=\"savePromptBtn\" type=\"button\">Save prompt</button>
              <button id=\"updatePromptBtn\" type=\"button\" class=\"secondary\">Update selected</button>
              <button id=\"deletePromptBtn\" type=\"button\" class=\"secondary\">Delete selected</button>
            </div>
          </div>
          <label for=\"promptText\" style=\"margin-top:8px;\">Prompt text</label>
          <textarea id=\"promptText\" placeholder=\"Select a saved prompt or type your own\"></textarea>
          <div class=\"stack\" style=\"margin-top:8px;\">
            <input id=\"configFileInput\" type=\"file\" accept=\"text/plain\" />
            <button id=\"uploadConfigBtn\" type=\"button\" class=\"secondary\">Upload .txt prompts</button>
          </div>
        </div>
        <div class=\"col\">
          <label>AI request</label>
          <div class=\"stack\" style=\"margin-bottom:8px;\">
            <label style=\"display:flex; align-items:center; gap:6px; font-weight: normal; color: var(--text);\"><input type=\"checkbox\" id=\"includeImages\" checked />Include images</label>
            <span id=\"queueCount\" style=\"color: var(--muted);\"></span>
          </div>
          <ul class=\"queue-list\" id=\"queueList\"></ul>
          <div class=\"stack\" style=\"margin-top:8px;\">
            <button id=\"sendAiBtn\" type=\"button\">Generate response</button>
          </div>
        </div>
      </div>
    </section>
  </main>

  <script>
      let cropBox = null;
      let naturalWidth = 0;
      let naturalHeight = 0;
      let isSelectingCrop = false;
      let firstCropPoint = null;
      let lastOcrText = '';
      let lastOcrImage = null;
      let queueItems = [];
      const maxQueueItems = 10;
      let promptEntries = [];

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

    function updateApiKeyStatus(isPresent) {
      const apiKeyStatus = document.getElementById('apiKeyStatus');
      apiKeyStatus.textContent = isPresent ? 'API key saved.' : 'No API key saved.';
    }

    async function loadSettings() {
      const status = document.getElementById('status');
      const res = await fetch('/api/settings');
      const data = await res.json();
      if (!res.ok) {
        status.textContent = data.error || 'Unable to load settings.';
        return;
      }

      document.getElementById('tesseractPath').value = data.tesseract_path || '';
      updateApiKeyStatus(!!data.api_key_present);
    }

    async function saveSettings() {
      const path = document.getElementById('tesseractPath').value;
      const apiKey = document.getElementById('apiKey').value;
      const res = await fetch('/api/settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tesseract_path: path, api_key: apiKey })
      });
      const data = await res.json();
      const status = document.getElementById('status');
      status.textContent = data.message || data.error || 'Settings updated';
      if ('api_key_present' in data) {
        updateApiKeyStatus(!!data.api_key_present);
      }
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
      lastOcrText = data.text || '';
      lastOcrImage = data.image_data_url || null;
      document.getElementById('ocrOutput').textContent = lastOcrText;
      document.getElementById('queueStatus').textContent = '';
    }

      function clearCrop() {
        cropBox = null;
        firstCropPoint = null;
        isSelectingCrop = false;
        document.getElementById('cropInfo').textContent = 'No crop set; OCR will use the full image.';
        const overlay = document.getElementById('crop-overlay');
        overlay.style.display = 'none';
      }

      function startCropSelection() {
        if (!naturalWidth || !naturalHeight) {
          document.getElementById('status').textContent = 'Capture an image before cropping.';
          return;
        }
        isSelectingCrop = true;
        firstCropPoint = null;
        document.getElementById('cropInfo').textContent = 'Click the top-left corner, then the bottom-right corner of your crop.';
        const overlay = document.getElementById('crop-overlay');
        overlay.style.display = 'none';
      }

      function setupCropping() {
        const img = document.getElementById('captureImage');
        const overlay = document.getElementById('crop-overlay');

        img.addEventListener('load', () => {
          naturalWidth = img.naturalWidth;
          naturalHeight = img.naturalHeight;
        });

        img.addEventListener('click', (ev) => {
          if (!isSelectingCrop) return;
          const rect = img.getBoundingClientRect();
          const point = { x: ev.clientX - rect.left, y: ev.clientY - rect.top };

          if (!firstCropPoint) {
            firstCropPoint = point;
            document.getElementById('cropInfo').textContent = 'First corner set. Click the bottom-right corner to finish.';
            overlay.style.display = 'block';
            overlay.style.left = `${point.x}px`;
            overlay.style.top = `${point.y}px`;
            overlay.style.width = '2px';
            overlay.style.height = '2px';
            return;
          }

          const left = Math.min(firstCropPoint.x, point.x);
          const top = Math.min(firstCropPoint.y, point.y);
          const width = Math.abs(point.x - firstCropPoint.x);
          const height = Math.abs(point.y - firstCropPoint.y);
          firstCropPoint = null;
          isSelectingCrop = false;

          if (width < 10 || height < 10) {
            clearCrop();
            return;
          }

          overlay.style.display = 'block';
          overlay.style.left = `${left}px`;
          overlay.style.top = `${top}px`;
          overlay.style.width = `${width}px`;
          overlay.style.height = `${height}px`;

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

      async function goFullscreen() {
        const root = document.documentElement;
        if (!document.fullscreenElement) {
          await root.requestFullscreen().catch(() => {});
        } else {
          await document.exitFullscreen().catch(() => {});
        }
      }

      function enforceLandscape() {
        if (window.matchMedia('(max-width: 900px)').matches && 'orientation' in screen && screen.orientation.lock) {
          screen.orientation.lock('landscape').catch(() => {});
        }

        const portrait = window.matchMedia('(orientation: portrait)').matches;
        document.body.classList.toggle('portrait-warning', portrait);
      }

      function updateQueueUI() {
        const list = document.getElementById('queueList');
        list.innerHTML = '';
        if (!queueItems.length) {
          const empty = document.createElement('li');
          empty.className = 'queue-item';
          empty.textContent = 'Queue is empty. Save OCR results to build a request.';
          list.appendChild(empty);
        } else {
          queueItems.forEach((item, idx) => {
            const li = document.createElement('li');
            li.className = 'queue-item';
            const title = document.createElement('div');
            title.className = 'queue-item-title';
            title.textContent = `Item ${idx + 1}`;
            const text = document.createElement('div');
            text.textContent = item.text.slice(0, 140) + (item.text.length > 140 ? '…' : '');
            li.appendChild(title);
            li.appendChild(text);
            list.appendChild(li);
          });
        }
        document.getElementById('queueCount').textContent = `${queueItems.length}/${maxQueueItems} queued`;
      }

      function addToQueue() {
        if (!lastOcrText) {
          document.getElementById('queueStatus').textContent = 'Run OCR before adding to the queue.';
          return;
        }
        if (queueItems.length >= maxQueueItems) {
          document.getElementById('queueStatus').textContent = 'Queue is full (10 items max).';
          return;
        }
        queueItems.push({ text: lastOcrText, image: lastOcrImage });
        document.getElementById('queueStatus').textContent = 'Saved to queue. Capture and OCR the next image if needed.';
        updateQueueUI();
      }

      function clearQueue() {
        queueItems = [];
        document.getElementById('queueStatus').textContent = 'Queue cleared.';
        updateQueueUI();
      }

      async function loadPrompts() {
        const select = document.getElementById('promptSelect');
        const status = document.getElementById('status');
        try {
          const res = await fetch('/api/configs');
          const data = await res.json();
          if (!res.ok) {
            status.textContent = data.error || 'Unable to load prompts.';
            return;
          }
          promptEntries = data.prompts || [];
          populatePromptSelect();
        } catch (err) {
          status.textContent = 'Unable to load prompts.';
          select.innerHTML = '';
        }
      }

      function populatePromptSelect(preferredTitle) {
        const select = document.getElementById('promptSelect');
        const previousSelection = preferredTitle || select.value;
        select.innerHTML = '';
        const placeholder = document.createElement('option');
        placeholder.value = '';
        placeholder.textContent = 'Choose a prompt';
        select.appendChild(placeholder);

        let target = null;
        promptEntries.forEach(entry => {
          const option = document.createElement('option');
          option.value = entry.title;
          option.textContent = entry.title;
          select.appendChild(option);
          if (!target && (entry.title === previousSelection || previousSelection === '')) {
            target = entry;
          }
        });

        if (target) {
          select.value = target.title;
          document.getElementById('promptText').value = target.prompt;
          document.getElementById('newPromptTitle').value = target.title;
        } else {
          select.value = '';
          document.getElementById('promptText').value = '';
          document.getElementById('newPromptTitle').value = '';
        }
      }

      function onPromptChange() {
        const select = document.getElementById('promptSelect');
        const target = promptEntries.find(p => p.title === select.value);
        document.getElementById('promptText').value = target ? target.prompt : '';
        document.getElementById('newPromptTitle').value = target ? target.title : '';
      }

      async function savePrompt() {
        const title = document.getElementById('newPromptTitle').value.trim();
        const prompt = document.getElementById('promptText').value.trim();
        const status = document.getElementById('status');
        if (!title || !prompt) {
          status.textContent = 'Provide both a title and prompt text.';
          return;
        }
        const res = await fetch('/api/configs', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ title, prompt })
        });
        const data = await res.json();
        if (!res.ok) {
          status.textContent = data.error || 'Unable to save prompt.';
          return;
        }
        status.textContent = data.message || 'Prompt saved.';
        promptEntries = data.prompts || [];
        populatePromptSelect(title);
      }

      async function updatePrompt() {
        const select = document.getElementById('promptSelect');
        const originalTitle = select.value;
        const title = document.getElementById('newPromptTitle').value.trim() || originalTitle;
        const prompt = document.getElementById('promptText').value.trim();
        const status = document.getElementById('status');

        if (!originalTitle) {
          status.textContent = 'Select a prompt to update.';
          return;
        }
        if (!prompt || !title) {
          status.textContent = 'Provide both a title and prompt text.';
          return;
        }

        const res = await fetch('/api/configs', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ original_title: originalTitle, title, prompt })
        });
        const data = await res.json();
        if (!res.ok) {
          status.textContent = data.error || 'Unable to update prompt.';
          return;
        }
        status.textContent = data.message || 'Prompt updated.';
        promptEntries = data.prompts || [];
        populatePromptSelect(title);
      }

      async function deletePrompt() {
        const select = document.getElementById('promptSelect');
        const title = select.value;
        const status = document.getElementById('status');
        if (!title) {
          status.textContent = 'Select a prompt to delete.';
          return;
        }

        const res = await fetch('/api/configs', {
          method: 'DELETE',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ title })
        });
        const data = await res.json();
        if (!res.ok) {
          status.textContent = data.error || 'Unable to delete prompt.';
          return;
        }
        status.textContent = data.message || 'Prompt deleted.';
        promptEntries = data.prompts || [];
        populatePromptSelect();
      }

      async function uploadPromptFile() {
        const input = document.getElementById('configFileInput');
        const status = document.getElementById('status');
        if (!input.files || !input.files.length) {
          status.textContent = 'Choose a .txt file to upload prompts.';
          return;
        }
        const form = new FormData();
        form.append('file', input.files[0]);
        const res = await fetch('/api/configs/upload', { method: 'POST', body: form });
        const data = await res.json();
        if (!res.ok) {
          status.textContent = data.error || 'Unable to upload prompts.';
          return;
        }
        status.textContent = data.message || 'Prompts uploaded.';
        promptEntries = data.prompts || [];
        populatePromptSelect();
        input.value = '';
      }

      async function sendAiRequest() {
        const status = document.getElementById('status');
        const prompt = document.getElementById('promptText').value.trim();
        const includeImages = document.getElementById('includeImages').checked;
        status.textContent = 'Sending to AI…';
        const res = await fetch('/api/ai_response', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ prompt, include_images: includeImages, queue: queueItems })
        });
        const data = await res.json();
        if (!res.ok) {
          status.textContent = data.error || 'AI request failed.';
          return;
        }
        status.textContent = 'AI response ready.';
        document.getElementById('ocrOutput').textContent = data.response || '';
      }

    document.getElementById('refreshBtn').addEventListener('click', refreshWindows);
      document.getElementById('captureBtn').addEventListener('click', capture);
      document.getElementById('ocrBtn').addEventListener('click', runOcr);
      document.getElementById('saveSettingsBtn').addEventListener('click', saveSettings);
      document.getElementById('clearCropBtn').addEventListener('click', clearCrop);
      document.getElementById('startCropBtn').addEventListener('click', startCropSelection);
      document.getElementById('fullscreenBtn').addEventListener('click', goFullscreen);
      document.getElementById('queueBtn').addEventListener('click', addToQueue);
      document.getElementById('clearQueueBtn').addEventListener('click', clearQueue);
      document.getElementById('promptSelect').addEventListener('change', onPromptChange);
      document.getElementById('savePromptBtn').addEventListener('click', savePrompt);
      document.getElementById('updatePromptBtn').addEventListener('click', updatePrompt);
      document.getElementById('deletePromptBtn').addEventListener('click', deletePrompt);
      document.getElementById('uploadConfigBtn').addEventListener('click', uploadPromptFile);
      document.getElementById('sendAiBtn').addEventListener('click', sendAiRequest);
      window.addEventListener('orientationchange', enforceLandscape);
      window.addEventListener('resize', enforceLandscape);

      setupCropping();
      loadSettings();
      refreshWindows();
      clearCrop();
      loadPrompts();
      updateQueueUI();
      enforceLandscape();
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
