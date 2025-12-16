"""Phase 1 desktop GUI for window capture and OCR.

This application targets Windows and now uses wxPython for the GUI. It allows
users to select an open window, capture its contents, and run OCR on the
capture using a local Tesseract installation.
"""

from __future__ import annotations

import importlib
import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import wx

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

    @property
    def region(self) -> tuple[int, int, int, int]:
        return self.left, self.top, self.width, self.height


def pil_to_bitmap(image: Image.Image) -> wx.Bitmap:
    """Convert a PIL image to a wx.Bitmap for preview rendering."""

    rgba = image.convert("RGBA")
    width, height = rgba.size
    return wx.Bitmap.FromBufferRGBA(width, height, rgba.tobytes())


class SettingsDialog(wx.Dialog):
    """Simple settings dialog allowing tesseract executable configuration."""

    def __init__(self, parent: wx.Window, tesseract_path: Optional[str]):
        super().__init__(parent, title="Settings", size=(520, 180))

        instruction = wx.StaticText(self, label="Tesseract executable:")
        self._path_ctrl = wx.TextCtrl(self, value=tesseract_path or "")
        browse_btn = wx.Button(self, label="Browse…")
        browse_btn.Bind(wx.EVT_BUTTON, self._browse)

        form_sizer = wx.FlexGridSizer(2, 2, 10, 10)
        form_sizer.AddGrowableCol(1, 1)
        form_sizer.Add(instruction, 0, wx.ALIGN_CENTER_VERTICAL)
        form_sizer.Add(self._path_ctrl, 1, wx.EXPAND)

        browse_row = wx.BoxSizer(wx.HORIZONTAL)
        browse_row.AddStretchSpacer(1)
        browse_row.Add(browse_btn, 0)

        btn_sizer = self.CreateSeparatedButtonSizer(wx.OK | wx.CANCEL)

        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(form_sizer, 0, wx.ALL | wx.EXPAND, 12)
        main_sizer.Add(browse_row, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 12)
        if btn_sizer:
            main_sizer.Add(btn_sizer, 0, wx.ALL | wx.EXPAND, 12)

        self.SetSizer(main_sizer)

    def _browse(self, event: wx.CommandEvent) -> None:  # pragma: no cover - UI interaction
        with wx.FileDialog(
            self,
            message="Select tesseract.exe",
            wildcard="Executable (*.exe)|*.exe|All files (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
        ) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                self._path_ctrl.SetValue(dialog.GetPath())

    @property
    def tesseract_path(self) -> Optional[str]:
        text = self._path_ctrl.GetValue().strip()
        return text or None


class WindowSelectionDialog(wx.Dialog):
    """Dialog that lists current top-level windows for selection."""

    def __init__(self, parent: wx.Window):
        super().__init__(parent, title="Select a window to capture", size=(420, 360))

        self._list_box = wx.ListBox(self)
        refresh_btn = wx.Button(self, label="Refresh")
        refresh_btn.Bind(wx.EVT_BUTTON, self._populate_windows)

        btn_sizer = self.CreateSeparatedButtonSizer(wx.OK | wx.CANCEL)

        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(self._list_box, 1, wx.ALL | wx.EXPAND, 12)
        main_sizer.Add(refresh_btn, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.ALIGN_RIGHT, 12)
        if btn_sizer:
            main_sizer.Add(btn_sizer, 0, wx.ALL | wx.EXPAND, 12)

        self.SetSizer(main_sizer)
        self._populate_windows()

    def _populate_windows(self, event: Optional[wx.CommandEvent] = None) -> None:
        self._list_box.Clear()
        if gw is None:
            wx.MessageBox("Install pygetwindow to list windows.", "pygetwindow missing", wx.ICON_WARNING | wx.OK, parent=self)
            return

        titles: List[str] = [title for title in gw.getAllTitles() if title.strip()]
        titles.sort()
        self._list_box.InsertItems(titles, 0)

    def selected_title(self) -> Optional[str]:
        selection = self._list_box.GetSelection()
        if selection == wx.NOT_FOUND:
            return None
        return self._list_box.GetString(selection)

    def get_selection(self) -> Optional[SelectedWindow]:
        if self.ShowModal() != wx.ID_OK:
            return None
        title = self.selected_title()
        if not title or gw is None:
            return None
        window = gw.getWindowsWithTitle(title)[0]
        return SelectedWindow(title, window.left, window.top, window.width, window.height)


class PreviewPanel(wx.Panel):
    """Displays a scaled preview image."""

    def __init__(self, parent: wx.Window):
        super().__init__(parent, size=(660, 380))
        self.SetBackgroundColour(wx.Colour(240, 240, 240))
        self._bitmap: Optional[wx.Bitmap] = None

        self._static_bitmap = wx.StaticBitmap(self)

        border = wx.StaticBoxSizer(wx.StaticBox(self, label="Preview"), wx.VERTICAL)
        border.Add(self._static_bitmap, 1, wx.ALL | wx.EXPAND, 6)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(border, 1, wx.ALL | wx.EXPAND, 8)
        self.SetSizer(sizer)

    def update_image(self, bitmap: wx.Bitmap) -> None:
        self._bitmap = bitmap
        target_size = self.GetClientSize()
        image = bitmap.ConvertToImage()
        scaled = image.Scale(target_size.width, target_size.height, wx.IMAGE_QUALITY_HIGH)
        self._static_bitmap.SetBitmap(wx.Bitmap(scaled))
        self.Layout()


class OCRFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Phase 1 Screen Capture", size=(940, 560))

        self._selected_window: Optional[SelectedWindow] = None
        self._captured_image: Optional[Image.Image] = None
        self._tesseract_path: Optional[str] = self._detect_local_tesseract()
        self._dependency_state: Dict[str, bool] = {}
        self._install_attempted = False

        self._build_ui()
        self._apply_tesseract_path()
        self._check_dependencies()

        self.Bind(wx.EVT_CLOSE, self._on_close)

    def _build_ui(self) -> None:
        header = wx.StaticText(
            self,
            label=(
                "Select a window, take a snapshot, then process it locally.\n"
                "Works best on Windows with Tesseract installed."
            ),
        )
        header.Wrap(840)

        self._preview = PreviewPanel(self)

        self._take_btn = wx.Button(self, label="Take")
        self._take_btn.Disable()
        self._take_btn.Bind(wx.EVT_BUTTON, self._take_capture)

        self._process_btn = wx.Button(self, label="Process")
        self._process_btn.Disable()
        self._process_btn.Bind(wx.EVT_BUTTON, self._process_capture)

        self._start_btn = wx.Button(self, label="Start")
        self._start_btn.Bind(wx.EVT_BUTTON, self._select_window)

        settings_btn = wx.Button(self, label="Settings")
        settings_btn.Bind(wx.EVT_BUTTON, self._open_settings)

        button_col = wx.BoxSizer(wx.VERTICAL)
        button_col.Add(self._start_btn, 0, wx.BOTTOM | wx.EXPAND, 8)
        button_col.Add(settings_btn, 0, wx.BOTTOM | wx.EXPAND, 12)
        button_col.AddStretchSpacer(1)
        button_col.Add(self._take_btn, 0, wx.BOTTOM | wx.EXPAND, 8)
        button_col.Add(self._process_btn, 0, wx.BOTTOM | wx.EXPAND, 8)
        button_col.AddStretchSpacer(1)

        content_layout = wx.BoxSizer(wx.HORIZONTAL)
        content_layout.Add(button_col, 0, wx.ALL | wx.EXPAND, 8)
        content_layout.Add(self._preview, 1, wx.ALL | wx.EXPAND, 4)

        main_layout = wx.BoxSizer(wx.VERTICAL)
        main_layout.Add(header, 0, wx.ALL | wx.EXPAND, 12)
        main_layout.Add(content_layout, 1, wx.ALL | wx.EXPAND, 8)

        self.SetSizer(main_layout)

    def _check_dependencies(self) -> None:
        """Detect missing modules early and give a single actionable message."""

        self._dependency_state = self._detect_dependency_state()

        missing = [name for name, ok in self._dependency_state.items() if not ok]
        if missing:
            if self._attempt_install_requirements():
                self._dependency_state = self._detect_dependency_state()
                missing = [name for name, ok in self._dependency_state.items() if not ok]

        if missing:
            message = (
                "The following Python packages are required but not available: \n"
                + "\n".join(f" • {name}" for name in missing)
                + "\n\nTried installing them automatically but some are still missing."
            )
            wx.MessageBox(message, "Missing dependencies", wx.ICON_ERROR | wx.OK, parent=self)

        self._start_btn.Enable(self._dependency_state.get("pygetwindow", False))
        if not self._dependency_state.get("pygetwindow", False):
            self._status_message("Install pygetwindow to enable window selection.")

    def _detect_dependency_state(self) -> Dict[str, bool]:
        return {
            "pygetwindow": gw is not None,
            "pyautogui": pyautogui is not None,
            "pillow": Image is not None,
            "pytesseract": pytesseract is not None,
        }

    def _detect_local_tesseract(self) -> Optional[str]:
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

    def _apply_tesseract_path(self) -> None:
        if pytesseract and self._tesseract_path:
            pytesseract.pytesseract.tesseract_cmd = self._tesseract_path

    def _ensure_dependency(self, key: str, friendly_name: str) -> bool:
        """Show a clear message if the dependency is missing and stop the action."""

        if self._dependency_state.get(key, False):
            return True

        if self._attempt_install_requirements():
            self._dependency_state = self._detect_dependency_state()
            if self._dependency_state.get(key, False):
                return True

        wx.MessageBox(
            f"{friendly_name} is not installed. Please run: pip install -r requirements.txt",
            "Missing dependency",
            wx.ICON_ERROR | wx.OK,
            parent=self,
        )
        return False

    def _attempt_install_requirements(self) -> bool:
        if self._install_attempted:
            return False

        self._install_attempted = True
        if not self._install_requirements():
            return False

        self._refresh_optional_dependencies()
        wx.MessageBox(
            "Required packages were installed. Please retry your action.",
            "Dependencies installed",
            wx.ICON_INFORMATION | wx.OK,
            parent=self,
        )
        return True

    def _install_requirements(self) -> bool:
        requirements_path = Path(__file__).resolve().parent.parent / "requirements.txt"
        if not requirements_path.exists():
            wx.MessageBox(
                f"Could not find requirements file at {requirements_path}.",
                "Missing requirements.txt",
                wx.ICON_ERROR | wx.OK,
                parent=self,
            )
            return False

        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "-r", str(requirements_path)],
                check=True,
                capture_output=True,
                text=True,
            )
        except subprocess.CalledProcessError as exc:
            details = exc.stderr or exc.stdout or "Unknown error"
            wx.MessageBox(
                f"pip install failed with:\n{details}",
                "Dependency installation failed",
                wx.ICON_ERROR | wx.OK,
                parent=self,
            )
            return False

        return True

    def _refresh_optional_dependencies(self) -> None:
        global gw, pyautogui, pytesseract, Image

        if gw is None:
            try:
                gw = importlib.import_module("pygetwindow")
            except Exception:  # pragma: no cover - optional dependency
                gw = None
        if pyautogui is None:
            try:
                pyautogui = importlib.import_module("pyautogui")
            except Exception:  # pragma: no cover - optional dependency
                pyautogui = None
        if pytesseract is None:
            try:
                pytesseract = importlib.import_module("pytesseract")
            except Exception:  # pragma: no cover - optional dependency
                pytesseract = None
        if Image is None:
            try:
                from PIL import Image as PilImage

                Image = PilImage
            except Exception:  # pragma: no cover - optional dependency
                Image = None

    def _open_settings(self, event: wx.CommandEvent) -> None:  # pragma: no cover - UI interaction
        dialog = SettingsDialog(self, self._tesseract_path)
        if dialog.ShowModal() == wx.ID_OK:
            self._tesseract_path = dialog.tesseract_path
            self._apply_tesseract_path()
        dialog.Destroy()

    def _select_window(self, event: wx.CommandEvent) -> None:  # pragma: no cover - UI interaction
        if not self._ensure_dependency("pygetwindow", "pygetwindow"):
            return
        dialog = WindowSelectionDialog(self)
        selection = dialog.get_selection()
        dialog.Destroy()
        if not selection:
            return
        self._selected_window = selection
        self._take_btn.Enable()
        self._process_btn.Disable()
        self._status_message(f"Selected window: {selection.title}")

    def _take_capture(self, event: wx.CommandEvent) -> None:  # pragma: no cover - UI interaction
        if self._selected_window is None:
            wx.MessageBox("Please select a window first.", "No window", wx.ICON_INFORMATION | wx.OK, parent=self)
            return
        if not self._ensure_dependency("pyautogui", "pyautogui"):
            return
        if not self._ensure_dependency("pillow", "Pillow"):
            return

        try:
            screenshot = pyautogui.screenshot(region=self._selected_window.region)
        except Exception as exc:  # pragma: no cover - user environment specific
            wx.MessageBox(str(exc), "Capture failed", wx.ICON_ERROR | wx.OK, parent=self)
            return

        self._captured_image = screenshot
        bitmap = pil_to_bitmap(screenshot)
        self._preview.update_image(bitmap)
        self._process_btn.Enable()
        self._status_message("Capture ready for processing")

    def _process_capture(self, event: wx.CommandEvent) -> None:  # pragma: no cover - UI interaction
        if self._captured_image is None:
            wx.MessageBox("Take a capture first.", "No capture", wx.ICON_INFORMATION | wx.OK, parent=self)
            return
        if not self._ensure_dependency("pytesseract", "pytesseract"):
            return

        try:
            text = pytesseract.image_to_string(self._captured_image)
        except Exception as exc:  # pragma: no cover - user environment specific
            wx.MessageBox(str(exc), "Processing failed", wx.ICON_ERROR | wx.OK, parent=self)
            return

        with wx.MessageDialog(
            self,
            message="Extracted text:",
            caption="OCR Result",
            style=wx.OK | wx.CENTRE | wx.STAY_ON_TOP,
        ) as dialog:
            dialog.SetExtendedMessage(text)
            dialog.ShowModal()

    def _on_close(self, event: wx.CloseEvent) -> None:  # pragma: no cover - UI interaction
        reason = (
            "AI Agent is about to quit. Any selected window or OCR results will be lost.\n\n"
            "Do you want to exit?"
        )
        if wx.MessageBox(reason, "Exit AI Agent", wx.ICON_QUESTION | wx.YES_NO, parent=self) == wx.YES:
            self.Destroy()
        else:
            event.Veto()
            self._status_message("Close cancelled; continuing session.")

    def _status_message(self, message: str) -> None:
        wx.MessageBox(message, "Status", wx.ICON_INFORMATION | wx.OK, parent=self)


def main() -> int:
    app = wx.App()
    frame = OCRFrame()
    frame.Show()
    app.MainLoop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
