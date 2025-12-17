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
GOOGLE_API_KEY_FILE = CONFIG_DIR / "google_api_key.txt"
PROVIDER_FILE = CONFIG_DIR / "ai_provider.txt"
DEFAULT_PROVIDER = "openai"
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


def _load_google_api_key() -> Optional[str]:
    env_key = os.environ.get("GOOGLE_API_KEY")
    if env_key and env_key.strip():
        return env_key.strip()

    if GOOGLE_API_KEY_FILE.exists():
        key = GOOGLE_API_KEY_FILE.read_text(encoding="utf-8").strip()
        if key:
            return key
    return None


def _save_google_api_key(raw_key: str) -> bool:
    _ensure_config_dir()
    key = (raw_key or "").strip()
    if not key:
        if GOOGLE_API_KEY_FILE.exists():
            GOOGLE_API_KEY_FILE.unlink()
        return False

    GOOGLE_API_KEY_FILE.write_text(key, encoding="utf-8")
    return True


def _load_ai_provider() -> str:
    env_value = os.environ.get("AI_AGENT_PROVIDER")
    if env_value and env_value.lower() in {"openai", "google"}:
        return env_value.lower()

    if PROVIDER_FILE.exists():
        stored = PROVIDER_FILE.read_text(encoding="utf-8").strip().lower()
        if stored in {"openai", "google"}:
            return stored

    return DEFAULT_PROVIDER


def _save_ai_provider(provider: str) -> str:
    normalized = (provider or "").strip().lower()
    if normalized not in {"openai", "google"}:
        return _load_ai_provider()

    _ensure_config_dir()
    PROVIDER_FILE.write_text(normalized, encoding="utf-8")
    return normalized


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


def _data_url_to_inline_data(data_url: str) -> Optional[Dict[str, str]]:
    if not data_url.startswith("data:"):
        return None

    try:
        header, encoded = data_url.split(",", 1)
    except ValueError:
        return None

    mime_type = "image/png"
    if ";base64" in header:
        mime_type = header.split("data:", 1)[1].split(";", 1)[0] or mime_type

    try:
        decoded = base64.b64decode(encoded)
    except Exception:
        return None

    return {"mime_type": mime_type, "data": base64.b64encode(decoded).decode("ascii")}


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


def _get_google_client():  # type: ignore[return-type]
    if importlib.util.find_spec("google.genai") is None:
        return None, (
            "Google Generative AI SDK is not installed. Please add google-genai "
            "to requirements.txt."
        )

    try:
        from google import genai
    except Exception as exc:  # pragma: no cover - import failure
        return None, f"Unable to import Google Generative AI client: {exc}"

    api_key = _load_google_api_key()
    if not api_key:
        return None, "Set GOOGLE_API_KEY or save a key for Google AI in settings."

    try:
        return genai.Client(api_key=api_key), None
    except Exception as exc:  # pragma: no cover - network/env specific
        return None, f"Unable to initialize Google Generative AI client: {exc}"


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

    _apply_tesseract_path()

    if image.mode == "RGBA":
        # Tesseract expects 3-channel data; flatten any alpha channel first.
        image = image.convert("RGB")
    elif image.mode not in ("RGB", "L"):
        image = image.convert("RGB")

    return pytesseract.image_to_string(image).strip()


def _parse_crop_request(data: Dict[str, object], image: "Image.Image") -> Tuple[Optional[Tuple[int, int, int, int]], Optional[str]]:
    """Parse and normalize crop payloads from the browser.

    The UI historically sent `crop_box` as {x, y, width, height}. Some clients
    may send `crop` as {left, top, right, bottom}. We accept both.
    """
    crop = None
    if isinstance(data, dict):
        crop = data.get("crop") or data.get("crop_box") or data.get("cropBox")

    if not crop:
        return None, None

    if not isinstance(crop, dict):
        return None, "Crop must be an object."

    try:
        if all(key in crop for key in ("left", "top", "right", "bottom")):
            l = int(crop.get("left"))  # type: ignore[arg-type]
            t = int(crop.get("top"))   # type: ignore[arg-type]
            r = int(crop.get("right")) # type: ignore[arg-type]
            b = int(crop.get("bottom"))# type: ignore[arg-type]
        elif all(key in crop for key in ("x", "y", "width", "height")):
            l = int(crop.get("x"))      # type: ignore[arg-type]
            t = int(crop.get("y"))      # type: ignore[arg-type]
            r = l + int(crop.get("width"))   # type: ignore[arg-type]
            b = t + int(crop.get("height"))  # type: ignore[arg-type]
        else:
            return None, "Crop object must contain either left/top/right/bottom or x/y/width/height."
    except Exception as exc:
        return None, f"Invalid crop values: {exc}"

    # Ensure left<right and top<bottom
    if r < l:
        l, r = r, l
    if b < t:
        t, b = b, t

    width, height = image.size
    l = max(0, min(width, l))
    r = max(0, min(width, r))
    t = max(0, min(height, t))
    b = max(0, min(height, b))

    if r - l <= 0 or b - t <= 0:
        return None, "Crop area is empty after normalization."

    return (l, t, r, b), None


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
                "google_api_key_present": _load_google_api_key() is not None,
                "api_key_path": str(API_KEY_FILE),
                "provider": _load_ai_provider(),
            }
        )

    data = request.get_json(silent=True) or {}
    path = data.get("tesseract_path")
    state.tesseract_path = path or _detect_local_tesseract()
    _apply_tesseract_path()

    provider = data.get("provider") or _load_ai_provider()
    provider = _save_ai_provider(provider)

    api_key_present = None
    google_api_key_present = None
    if "api_key" in data:
        if provider == "google":
            google_api_key_present = _save_google_api_key(str(data.get("api_key", "")))
        else:
            api_key_present = _save_api_key(str(data.get("api_key", "")))

    return jsonify(
        {
            "message": "Settings saved.",
            "tesseract_path": state.tesseract_path,
            "api_key_present": api_key_present
            if api_key_present is not None
            else _load_api_key() is not None,
            "google_api_key_present": google_api_key_present
            if google_api_key_present is not None
            else _load_google_api_key() is not None,
            "provider": provider,
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


@app.route("/api/clear_capture", methods=["POST"])
def api_clear_capture() -> tuple[str, int]:
    state.captured_image = None
    state.crop_box = None
    return jsonify({"message": "Capture cleared."})


@app.route("/api/ocr", methods=["POST"])
def api_ocr() -> tuple[str, int]:
    if state.captured_image is None:
        return jsonify({"error": "Take a capture first."}), 400

    if message := _ensure_dependency("pytesseract", "pytesseract"):
        return jsonify({"error": message}), 400

    data = request.get_json(silent=True) or {}

    image_for_ocr = state.captured_image
    crop_box, crop_error = _parse_crop_request(data, image_for_ocr)
    if crop_error:
        return jsonify({"error": crop_error}), 400

    if crop_box:
        image_for_ocr = image_for_ocr.crop(crop_box)
        state.crop_box = crop_box
    else:
        state.crop_box = None

    try:
        text = _run_ocr(image_for_ocr)
        data_url = _image_to_data_url(image_for_ocr)
    except Exception as exc:  # pragma: no cover - user environment specific
        return jsonify({"error": str(exc)}), 500

    response_payload: Dict[str, object] = {
        "text": text,
        # Keep both keys for backward compatibility with older clients.
        "image_data_url": data_url,
        "image_data": data_url,
    }
    if crop_box:
        l, t, r, b = crop_box
        response_payload["crop_box"] = {"left": l, "top": t, "right": r, "bottom": b}

    return jsonify(response_payload)


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

    provider = _load_ai_provider()
    if provider == "google":
        client, error = _get_google_client()
        if error:
            return jsonify({"error": error}), 400

        model = os.environ.get("AI_AGENT_GOOGLE_MODEL", "gemini-3-pro-preview")
        parts: List[Dict[str, str]] = []
        for entry in content:
            if entry.get("type") == "input_text" and entry.get("text"):
                parts.append({"text": entry["text"]})
            if include_images and entry.get("type") == "input_image":
                inline = _data_url_to_inline_data(entry.get("image_url", ""))
                if inline:
                    parts.append({"inline_data": inline})

        if not parts:
            parts.append({"text": "Please review the provided OCR text and images."})

        try:
            response = client.models.generate_content(
                model=model,
                contents=[{"role": "user", "parts": parts}],
            )
        except Exception as exc:  # pragma: no cover - network/env specific
            return jsonify({"error": str(exc)}), 500

        text = getattr(response, "output_text", None) or getattr(response, "text", "")
        return jsonify({"response": text or ""})

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


_PAGE_TEMPLATE = r"""
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>AI Agent - Web Capture & OCR</title>
  <script>
    window.MathJax = {
      tex: { inlineMath: [['$', '$'], ['\\(', '\\)']] },
      svg: { fontCache: 'global' }
    };
  </script>
  <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/dompurify@3.0.6/dist/purify.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js"></script>
  <style>
    :root {
      --bg: #e8e8e8;
      --panel: #ffffff;
      --border: #cfcfcf;
      --text: #1f1f1f;
      --muted: #6b6b6b;
      --accent: #3d3d3d;
      --accent-strong: #2b2b2b;
      --highlight: #ececec;
      --success: #0a7d4d;
    }
    * { box-sizing: border-box; }
    body { font-family: Arial, sans-serif; margin: 0; padding: 16px; background: var(--bg); color: var(--text); min-height: 100vh; display: flex; align-items: flex-start; justify-content: center; overflow: auto; }
    h2 { margin: 0; color: var(--accent-strong); }
    h3 { margin: 0 0 6px; color: var(--accent-strong); }
    button { background: var(--accent); color: #f5f5f5; border: none; padding: 10px 16px; border-radius: 8px; cursor: pointer; font-size: 14px; }
    button.secondary { background: var(--muted); color: #f5f5f5; }
    button.ghost { background: transparent; color: var(--accent-strong); border: 1px solid var(--border); }
    button:disabled { background: #b5b5b5; cursor: not-allowed; }
    label { display: block; margin-bottom: 6px; font-weight: bold; color: var(--accent-strong); }
    input, select, textarea { width: 100%; padding: 8px; border-radius: 6px; border: 1px solid var(--border); box-sizing: border-box; font-size: 14px; background: #fdfdfd; color: var(--text); }
    textarea { min-height: 90px; }
    .frame { width: min(1200px, 100%); background: linear-gradient(135deg, #f7f7f7 0%, #ededed 100%); border: 1px solid var(--border); border-radius: 16px; box-shadow: 0 8px 24px rgba(0,0,0,0.08); padding: 12px; display: flex; align-items: stretch; justify-content: center; }
    .viewport { position: relative; width: 100%; min-height: calc(100vh - 32px); }
    .page { display: none; position: relative; padding: 16px; background: var(--panel); border: 1px solid var(--border); border-radius: 12px; box-shadow: inset 0 1px 2px rgba(0,0,0,0.04); overflow: auto; min-height: inherit; }
    .page.active { display: flex; flex-direction: column; gap: 12px; }
    .page-heading { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; flex-wrap: wrap; }
    .page-heading .title { font-size: 18px; font-weight: bold; color: var(--accent-strong); }
    .grid { display: grid; gap: 12px; }
    .grid.two { grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); }
    .stack { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
    .panel { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }
    #status { position: absolute; left: 16px; bottom: 16px; padding: 8px 12px; background: rgba(0,0,0,0.7); color: #f5f5f5; border-radius: 10px; max-width: calc(100% - 32px); font-size: 14px; }
    #orientationNotice { position: absolute; top: 16px; right: 16px; padding: 8px 12px; background: #5a5a5a; color: #f5f5f5; border-radius: 8px; display: none; }
    body.portrait-warning #orientationNotice { display: block; }
    #preview-container { position: relative; display: none; max-width: 80%; }
    #crop-overlay { position: absolute; border: 2px dashed #707070; background: rgba(112, 112, 112, 0.2); display: none; pointer-events: none; }
    #captureImage { width: 100%; max-width: 100%; border-radius: 8px; border: 1px solid var(--border); height: auto; }
    #capturePreview { display: none; width: 80%; max-width: 640px; border-radius: 8px; border: 1px solid var(--border); }
    pre { background: var(--highlight); padding: 12px; border-radius: 8px; white-space: pre-wrap; border: 1px solid var(--border); }
    .queue-list { list-style: none; padding: 0; margin: 0; display: grid; gap: 8px; }
    .queue-item { border: 1px solid var(--border); padding: 8px; border-radius: 8px; background: var(--highlight); }
    .queue-item-title { font-weight: bold; color: var(--accent-strong); margin-bottom: 4px; }
    .label-muted { color: var(--muted); font-size: 13px; }
    .result-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 12px; }
    .rendered-box { background: #fff; border: 1px solid var(--border); border-radius: 10px; padding: 12px; min-height: 240px; overflow: auto; }
    .math-block { margin: 12px 0; }
    .tag { display: inline-block; background: var(--highlight); border-radius: 6px; padding: 4px 8px; border: 1px solid var(--border); font-size: 12px; }
    .page-nav { margin-top: auto; display: flex; justify-content: space-between; align-items: center; gap: 8px; flex-wrap: wrap; padding-top: 8px; border-top: 1px solid var(--border); }
    .nav-buttons { display: flex; gap: 8px; flex-wrap: wrap; }
  </style>
</head>
<body>
  <div class="frame">
    <div class="viewport">
      <div id="orientationNotice">Best viewed in landscape on mobile. Rotate your device if needed.</div>
      <div id="status">Ready.</div>

      <section class="page active" data-page="1">
        <div class="page-heading">
          <div class="title">Page 1 · Setup and configuration</div>
          <div class="label-muted">Choose a target window and confirm settings.</div>
        </div>
        <div class="grid two">
          <div class="panel">
            <h2>Window selection</h2>
            <label for="windowSelect">Open windows</label>
            <select id="windowSelect"></select>
            <div class="stack" style="margin-top:8px;">
              <button id="refreshBtn" type="button">Refresh</button>
            </div>
          </div>
          <div class="panel">
            <h2>App settings</h2>
            <label for="aiProvider">AI provider</label>
            <select id="aiProvider">
              <option value="openai">OpenAI</option>
              <option value="google">Google Gemini</option>
            </select>
            <div class="label-muted" style="margin-bottom:6px;">Switch between OpenAI and Google AI for responses.</div>
            <label for="tesseractPath">Tesseract executable</label>
            <input id="tesseractPath" type="text" placeholder="Path to tesseract.exe" value="{{ tesseract_path }}" />
            <label for="apiKey" style="margin-top:8px;">API key</label>
            <input id="apiKey" type="password" placeholder="sk-..." autocomplete="off" />
            <div class="stack" style="margin: 6px 0 0 0; align-items: flex-start;">
              <input type="checkbox" id="clearApiKey" />
              <label for="clearApiKey" class="label-muted" style="font-weight: normal; margin: 0;">Clear stored API key for the selected provider</label>
            </div>
            <div id="apiKeyStatus" class="label-muted" style="margin-top:6px;">No API key saved.</div>
            <div class="label-muted" style="margin-top:4px;">Keys are stored per provider in the configs folder. Leave the field blank to keep the saved key.</div>
            <div style="margin-top:8px;">
              <button id="saveSettingsBtn" type="button">Save settings</button>
            </div>
            <div class="label-muted" style="margin-top:10px;">Detected crop box: {{ crop_box if crop_box else 'None' }}</div>
          </div>
        </div>
        <div class="page-nav">
          <div class="label-muted">Need fullscreen? <button id="fullscreenBtn" type="button" class="ghost">Toggle</button></div>
          <div class="nav-buttons">
            <button type="button" data-page-target="2">Go to capture</button>
          </div>
        </div>
      </section>

      <section class="page" data-page="2">
        <div class="page-heading">
          <div class="title">Page 2 · Capture</div>
          <div class="label-muted">Run a capture, preview it, and head to cropping.</div>
        </div>
        <div class="grid two">
          <div class="panel">
            <h2>Capture controls</h2>
            <p class="label-muted">Selected window: <span id="selectedWindowLabel">None</span></p>
            <div class="stack" style="margin-top:8px;">
              <button id="captureBtn" type="button">Capture</button>
              <button class="ghost" type="button" data-page-target="3">Go to crop</button>
            </div>
            <div style="margin-top:8px;">
              <img id="capturePreview" alt="Capture preview" />
            </div>
          </div>
          <div class="panel">
            <h2>Notes</h2>
            <p class="label-muted">Capture once, then continue to the dedicated crop and OCR pages.</p>
            <p class="label-muted">You can return here anytime to retake the shot.</p>
          </div>
        </div>
        <div class="page-nav">
          <div class="nav-buttons">
            <button class="ghost" type="button" data-page-target="1">Back to setup</button>
            <button type="button" data-page-target="3">Next: crop & OCR</button>
          </div>
        </div>
      </section>

      <section class="page" data-page="3">
        <div class="page-heading">
          <div class="title">Page 3 · Crop & OCR</div>
          <div class="label-muted">Crop your capture and review OCR before saving.</div>
        </div>
        <div class="grid two">
          <div class="panel">
            <h2>Preview & crop</h2>
            {% if has_capture %}
            <div id="preview-container" style="display:inline-block;">
              <img id="captureImage" src="/image" alt="Capture preview" />
              <div id="crop-overlay"></div>
            </div>
            {% else %}
            <p class="label-muted">No capture yet. Use Capture to grab a window first.</p>
            <div id="preview-container">
              <img id="captureImage" src="" alt="Capture preview" />
              <div id="crop-overlay"></div>
            </div>
            {% endif %}
            <div class="stack" style="margin-top: 12px;">
              <button id="startCropBtn" type="button">Select crop area</button>
              <button id="clearCropBtn" type="button" class="secondary">Clear crop</button>
              <button id="ocrBtn" type="button">Run OCR</button>
            </div>
            <div id="cropInfo" class="label-muted" style="margin-top:6px;"></div>
          </div>
          <div class="panel">
            <h2>OCR output</h2>
            <pre id="ocrOutput"></pre>
            <div class="stack" style="margin-top:8px;">
              <button id="queueBtn" type="button">Save to queue</button>
              <button id="clearQueueBtn" type="button" class="secondary">Clear queue</button>
            </div>
            <div id="queueStatus" class="label-muted" style="margin-top:6px;"></div>
          </div>
        </div>
        <div class="page-nav">
          <div class="nav-buttons">
            <button class="ghost" type="button" data-page-target="2">Back to capture</button>
            <button type="button" data-page-target="4">Next: submit</button>
          </div>
        </div>
      </section>

      <section class="page" data-page="4">
        <div class="page-heading">
          <div class="title">Page 4 · Submit</div>
          <div class="label-muted">Queue items on the left, prompt and options on the right.</div>
        </div>
        <div class="grid two">
          <div class="panel">
            <h2>Queued captures</h2>
            <div class="stack" style="margin-bottom:8px;">
              <span class="tag" id="queueCount">0/10 queued</span>
              <span class="label-muted">Add more items from the Crop page if needed.</span>
            </div>
            <ul class="queue-list" id="queueList"></ul>
          </div>
          <div class="panel">
            <h2>Prompt & options</h2>
            <label for="promptSelect">Saved prompts</label>
            <select id="promptSelect"></select>
            <div class="stack" style="margin-top:8px;">
              <input id="newPromptTitle" type="text" placeholder="New prompt title" />
              <div class="stack">
                <button id="savePromptBtn" type="button">Save prompt</button>
                <button id="updatePromptBtn" type="button" class="secondary">Update selected</button>
                <button id="deletePromptBtn" type="button" class="secondary">Delete selected</button>
              </div>
            </div>
            <label for="promptText" style="margin-top:8px;">Prompt text</label>
            <textarea id="promptText" placeholder="Select a saved prompt or type your own"></textarea>
            <div class="stack" style="margin-top:8px;">
              <input id="configFileInput" type="file" accept="text/plain" />
              <button id="uploadConfigBtn" type="button" class="secondary">Upload .txt prompts</button>
            </div>
            <div class="stack" style="margin-top:12px;">
              <label style="display:flex; align-items:center; gap:6px; font-weight: normal; color: var(--text);"><input type="checkbox" id="includeImages" checked />Include images</label>
              <button id="sendAiBtn" type="button">Generate response</button>
            </div>
          </div>
        </div>
        <div class="page-nav">
          <div class="nav-buttons">
            <button class="ghost" type="button" data-page-target="3">Back to crop</button>
            <button type="button" data-page-target="5" id="jumpToResultBtn" class="secondary">View last result</button>
          </div>
        </div>
      </section>

      <section class="page" data-page="5">
        <div class="page-heading">
          <div class="title">Page 5 · Result</div>
          <div class="label-muted">Rendered response with Markdown and LaTeX.</div>
        </div>
        <div class="result-grid">
          <div class="panel">
            <h3>Response</h3>
            <div class="rendered-box" id="aiRenderedResponse">Run a request to see results.</div>
          </div>
        </div>
        <div class="page-nav">
          <div class="nav-buttons">
            <button class="ghost" type="button" data-page-target="2">Capture again</button>
            <button type="button" data-page-target="1">Back to setup</button>
          </div>
        </div>
      </section>
    </div>
  </div>

  <script>
      let cropBox = null;
      let naturalWidth = 0;
      let naturalHeight = 0;
      let isSelectingCrop = false;
      let firstCropPoint = null;
      let activePointerId = null;
      let pointerMoved = false;
      let lastOcrText = '';
      let lastOcrImage = null;
      let queueItems = [];
      const maxQueueItems = 10;
      let promptEntries = [];
      let currentPage = 1;
      let openaiKeyPresent = false;
      let googleKeyPresent = false;

      marked.setOptions({ gfm: true, breaks: true, mangle: false, headerIds: false });

      marked.use({
        extensions: [
          {
            name: 'math',
            level: 'inline',
            start(src) {
              const match = src.match(/\\\(|\\\[|\$\$/);
              return match ? match.index : undefined;
            },
            tokenizer(src) {
              const inline = src.match(/^\\\((.+?)\\\)/s);
              if (inline) {
                return { type: 'math_inline', raw: inline[0], text: inline[1] };
              }

              const displayBracket = src.match(/^\\\[(.+?)\\\]/s);
              if (displayBracket) {
                return { type: 'math_block', raw: displayBracket[0], text: displayBracket[1] };
              }

              const displayDollar = src.match(/^\$\$([\s\S]+?)\$\$/);
              if (displayDollar) {
                return { type: 'math_block', raw: displayDollar[0], text: displayDollar[1] };
              }

              return undefined;
            },
            renderer(token) {
              if (token.type === 'math_block') {
                return `<div class="math-block">\\[${token.text}\\]</div>`;
              }

              return `<span class="math-inline">\\(${token.text}\\)</span>`;
            }
          }
        ]
      });

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
      updateSelectedWindowLabel();
    }

    function updateSelectedWindowLabel() {
      const select = document.getElementById('windowSelect');
      const label = document.getElementById('selectedWindowLabel');
      if (!select || !label) return;
      label.textContent = select.value || 'None';
    }

    function updateApiKeyStatus() {
      const apiKeyStatus = document.getElementById('apiKeyStatus');
      const providerSelect = document.getElementById('aiProvider');
      const apiKeyLabel = document.querySelector('label[for="apiKey"]');
      const provider = providerSelect ? providerSelect.value : 'openai';
      const hasKey = provider === 'google' ? googleKeyPresent : openaiKeyPresent;
      const providerName = provider === 'google' ? 'Google AI' : 'OpenAI';
      apiKeyStatus.textContent = hasKey ? `${providerName} API key saved.` : `No ${providerName} API key saved.`;
      if (apiKeyLabel) {
        apiKeyLabel.textContent = `${providerName} API key`;
      }
      const apiKeyInput = document.getElementById('apiKey');
      if (apiKeyInput) {
        apiKeyInput.placeholder = provider === 'google' ? 'Google API key' : 'sk-...';
      }
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
      openaiKeyPresent = !!data.api_key_present;
      googleKeyPresent = !!data.google_api_key_present;
      document.getElementById('aiProvider').value = data.provider || 'openai';
      updateApiKeyStatus();
      document.getElementById('apiKey').value = '';
      document.getElementById('clearApiKey').checked = false;
    }

    async function saveSettings() {
      const status = document.getElementById('status');
      status.textContent = 'Saving settings...';
      const tesseractPath = document.getElementById('tesseractPath').value;
      const apiKey = document.getElementById('apiKey').value.trim();
      const provider = document.getElementById('aiProvider').value;
      const clearKey = document.getElementById('clearApiKey').checked;
      const payload = { tesseract_path: tesseractPath, provider };
      if (clearKey) {
        payload.api_key = '';
      } else if (apiKey) {
        payload.api_key = apiKey;
      }
      const res = await fetch('/api/settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (!res.ok) {
        status.textContent = data.error || 'Unable to save settings.';
        return;
      }
      openaiKeyPresent = !!data.api_key_present;
      googleKeyPresent = !!data.google_api_key_present;
      document.getElementById('aiProvider').value = data.provider || provider;
      document.getElementById('clearApiKey').checked = false;
      document.getElementById('apiKey').value = '';
      updateApiKeyStatus();
      status.textContent = data.message || 'Settings saved.';
    }

    async function capture() {
      const status = document.getElementById('status');
      const select = document.getElementById('windowSelect');
      const title = select.value;
      if (!title) {
        status.textContent = 'Select a window before capturing.';
        return;
      }
      status.textContent = 'Capturing window...';
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
      status.textContent = 'Capture ready. Continue to crop & OCR when you are ready.';
      const img = document.getElementById('captureImage');
      const newSrc = '/image?' + Date.now();
      img.src = newSrc;
      document.getElementById('preview-container').style.display = 'inline-block';
      const preview = document.getElementById('capturePreview');
      if (preview) {
        preview.src = newSrc;
        preview.style.display = 'block';
      }
      clearCrop(true);
      setPage(2);
    }

    function updateCropOverlay(box) {
      const overlay = document.getElementById('crop-overlay');
      if (!box) {
        overlay.style.display = 'none';
        return;
      }
      overlay.style.display = 'block';
      overlay.style.left = `${box.x}px`;
      overlay.style.top = `${box.y}px`;
      overlay.style.width = `${box.width}px`;
      overlay.style.height = `${box.height}px`;
    }

    function setupCropping() {
      const img = document.getElementById('captureImage');
      const overlay = document.getElementById('crop-overlay');

      function pointFromEvent(event) {
        const rect = img.getBoundingClientRect();
        return {
          x: Math.max(0, Math.min(rect.width, event.clientX - rect.left)),
          y: Math.max(0, Math.min(rect.height, event.clientY - rect.top)),
          rect,
        };
      }

      function commitCrop(endPoint) {
        const { rect, x, y } = endPoint;
        if (!firstCropPoint) return;
        const box = {
          x: Math.min(firstCropPoint.x, x),
          y: Math.min(firstCropPoint.y, y),
          width: Math.abs(x - firstCropPoint.x),
          height: Math.abs(y - firstCropPoint.y),
        };
        cropBox = {
          x: Math.round(box.x * (naturalWidth / rect.width)),
          y: Math.round(box.y * (naturalHeight / rect.height)),
          width: Math.round(box.width * (naturalWidth / rect.width)),
          height: Math.round(box.height * (naturalHeight / rect.height)),
          display: box,
        };
        updateCropOverlay(box);
        isSelectingCrop = false;
        firstCropPoint = null;
        activePointerId = null;
        pointerMoved = false;
        document.getElementById('cropInfo').textContent = `Crop set: (${cropBox.x}, ${cropBox.y}, ${cropBox.width}, ${cropBox.height})`;
      document.getElementById('status').textContent = 'Crop saved. Run OCR to continue.';
    }
      img.addEventListener('load', () => {
        const rect = img.getBoundingClientRect();
        naturalWidth = img.naturalWidth;
        naturalHeight = img.naturalHeight;
        overlay.style.width = `${rect.width}px`;
        overlay.style.height = `${rect.height}px`;
      });

      img.addEventListener('pointerdown', (event) => {
        if (!isSelectingCrop) return;
        const point = pointFromEvent(event);
        pointerMoved = false;
        if (!firstCropPoint) {
          firstCropPoint = { x: point.x, y: point.y };
          activePointerId = event.pointerId;
          img.setPointerCapture(event.pointerId);
          updateCropOverlay({ x: point.x, y: point.y, width: 2, height: 2 });
          document.getElementById('status').textContent = 'Tap the bottom-right corner (or drag) to finish the crop.';
          return;
        }
        commitCrop(point);
      });

      img.addEventListener('pointermove', (event) => {
        if (!isSelectingCrop || !firstCropPoint || event.pointerId !== activePointerId) return;
        const point = pointFromEvent(event);
        pointerMoved = pointerMoved || Math.abs(point.x - firstCropPoint.x) > 2 || Math.abs(point.y - firstCropPoint.y) > 2;
        const box = {
          x: Math.min(firstCropPoint.x, point.x),
          y: Math.min(firstCropPoint.y, point.y),
          width: Math.abs(point.x - firstCropPoint.x),
          height: Math.abs(point.y - firstCropPoint.y),
        };
        updateCropOverlay(box);
      });

      img.addEventListener('pointerup', (event) => {
        if (!isSelectingCrop || !firstCropPoint || event.pointerId !== activePointerId) return;
        const point = pointFromEvent(event);
        if (pointerMoved) {
          commitCrop(point);
        } else {
          // The first tap has finished without movement; allow a second tap to set the opposite corner.
          activePointerId = null;
        }
        img.releasePointerCapture(event.pointerId);
      });
    }

    function hydratePreviewOnLoad() {
      const img = document.getElementById('captureImage');
      const container = document.getElementById('preview-container');
      const preview = document.getElementById('capturePreview');
      if (img && img.getAttribute('src')) {
        container.style.display = 'inline-block';
        if (preview) {
          preview.src = img.getAttribute('src');
          preview.style.display = 'block';
        }
      }
    }

    function startCropSelection() {
      if (!document.getElementById('captureImage').src) {
        document.getElementById('status').textContent = 'Capture first before cropping.';
        return;
      }
      isSelectingCrop = true;
      document.getElementById('status').textContent = 'Tap once for the top-left corner, then tap bottom-right (drag also works).';
    }

    function clearCrop(skipMessage = false) {
      cropBox = null;
      firstCropPoint = null;
      isSelectingCrop = false;
      activePointerId = null;
      pointerMoved = false;
      updateCropOverlay(null);
      document.getElementById('cropInfo').textContent = '';
      if (!skipMessage) {
        document.getElementById('status').textContent = 'Crop cleared.';
      }
    }
async function runOcr() {
  const status = document.getElementById('status');
  status.textContent = 'Running OCR...';

  const payload = {};
  if (cropBox) {
    // Convert UI crop box ({x,y,width,height}) into an explicit crop rectangle.
    payload.crop = {
      left: cropBox.x,
      top: cropBox.y,
      right: cropBox.x + cropBox.width,
      bottom: cropBox.y + cropBox.height
    };
  }

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
  lastOcrText = data.text || '';
  lastOcrImage = data.image_data_url || data.image_data || null;
  document.getElementById('ocrOutput').textContent = lastOcrText || '[No text detected]';
  status.textContent = 'OCR complete. Add to queue or capture again.';
}

    function goFullscreen() {
      if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen().catch(() => {});
      } else {
        document.exitFullscreen().catch(() => {});
      }
    }

    function enforceLandscape() {
      try {
        if (window.matchMedia('(max-width: 900px)').matches && 'orientation' in screen && screen.orientation.lock) {
          screen.orientation.lock('landscape').catch(() => {});
        }
      } catch (err) {
        // Ignore unsupported orientation lock
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
      setPage(3);
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

      function renderMarkdown(content) {
        const target = document.getElementById('aiRenderedResponse');
        if (!content) {
          target.textContent = 'No response yet.';
          return;
        }
        let html = '';
        try {
          html = marked.parse(content);
        } catch (err) {
          target.textContent = content;
          return;
        }

        target.innerHTML = DOMPurify.sanitize(html, {
          ALLOWED_ATTR: ['href', 'name', 'target', 'class', 'rel', 'id'],
        });

        if (window.MathJax && window.MathJax.typesetPromise) {
          window.MathJax.typesetPromise([target]).catch(() => {});
        }
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
        renderMarkdown(data.response || '');
        await clearAfterResult();
        setPage(5);
      }

      async function clearAfterResult() {
        queueItems = [];
        updateQueueUI();
        lastOcrText = '';
        lastOcrImage = null;
        clearCrop(true);
        document.getElementById('ocrOutput').textContent = '';
        document.getElementById('queueStatus').textContent = '';
        const preview = document.getElementById('capturePreview');
        if (preview) {
          preview.src = '';
          preview.style.display = 'none';
        }
        const img = document.getElementById('captureImage');
        if (img) {
          img.src = '';
        }
        document.getElementById('preview-container').style.display = 'none';
        try {
          await fetch('/api/clear_capture', { method: 'POST' });
        } catch (err) {
          // ignore
        }
      }

      function setPage(pageNumber) {
        currentPage = pageNumber;
        document.querySelectorAll('.page').forEach(page => {
          page.classList.toggle('active', Number(page.dataset.page) === pageNumber);
        });
      }

    function bindNavigation() {
      document.querySelectorAll('[data-page-target]').forEach(el => {
        el.addEventListener('click', () => {
            const target = Number(el.dataset.pageTarget);
            if (!Number.isNaN(target)) {
              setPage(target);
            }
          });
        });
      }

      document.getElementById('refreshBtn').addEventListener('click', refreshWindows);
      document.getElementById('windowSelect').addEventListener('change', updateSelectedWindowLabel);
      document.getElementById('captureBtn').addEventListener('click', capture);
      document.getElementById('ocrBtn').addEventListener('click', runOcr);
      document.getElementById('saveSettingsBtn').addEventListener('click', saveSettings);
      document.getElementById('aiProvider').addEventListener('change', () => {
        document.getElementById('apiKey').value = '';
        document.getElementById('clearApiKey').checked = false;
        updateApiKeyStatus();
      });
      document.getElementById('clearCropBtn').addEventListener('click', () => clearCrop());
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
      document.getElementById('jumpToResultBtn').addEventListener('click', () => setPage(5));
      window.addEventListener('orientationchange', enforceLandscape);
      window.addEventListener('resize', enforceLandscape);

      bindNavigation();
      setupCropping();
      loadSettings();
      refreshWindows();
      clearCrop(true);
      loadPrompts();
      updateQueueUI();
      enforceLandscape();
      hydratePreviewOnLoad();
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
