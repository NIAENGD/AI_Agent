"""Desktop GUI for window capture and OCR.

This application targets Windows and uses wxPython for the GUI. It allows users
to select an open window, capture its contents, optionally crop the captured
image, and run OCR on the (optionally cropped) capture using a local Tesseract
installation.

Key UX goals:
- Touchscreen-friendly, DPI-aware, resizable UI.
- No "capture successful" popup after taking a capture; the preview updates in-place.
- Crop selection is drawn on the preview, but OCR runs on the original-resolution image.
"""

from __future__ import annotations

import importlib
import importlib.util
import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

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
    hwnd: Optional[int] = None

    @property
    def region(self) -> tuple[int, int, int, int]:
        return self.left, self.top, self.width, self.height


def pil_to_bitmap(image: "Image.Image") -> wx.Bitmap:
    """Convert a PIL image to a wx.Bitmap for preview rendering."""
    rgba = image.convert("RGBA")
    width, height = rgba.size
    return wx.Bitmap.FromBufferRGBA(width, height, rgba.tobytes())


class SettingsDialog(wx.Dialog):
    """Settings dialog allowing tesseract executable configuration.

    The dialog is DPI-safe, resizable, and scrollable.
    """

    def __init__(
        self,
        parent: wx.Window,
        tesseract_path: Optional[str],
        fullscreen_enabled: bool,
    ):
        super().__init__(parent, title="Settings", style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

        # Scrolled content region for settings fields.
        scroller = wx.ScrolledWindow(self, style=wx.TAB_TRAVERSAL)
        scroller.SetScrollRate(10, 10)

        instruction = wx.StaticText(scroller, label="Tesseract executable:")
        self._path_ctrl = wx.TextCtrl(scroller, value=tesseract_path or "")
        browse_btn = wx.Button(scroller, label="Browse…")
        browse_btn.Bind(wx.EVT_BUTTON, self._browse)

        self._fullscreen_checkbox = wx.CheckBox(scroller, label="Enable full screen (F11)")
        self._fullscreen_checkbox.SetValue(fullscreen_enabled)

        form_sizer = wx.FlexGridSizer(rows=1, cols=2, vgap=10, hgap=10)
        form_sizer.AddGrowableCol(1, 1)
        form_sizer.Add(instruction, 0, wx.ALIGN_CENTER_VERTICAL)
        form_sizer.Add(self._path_ctrl, 1, wx.EXPAND)

        browse_row = wx.BoxSizer(wx.HORIZONTAL)
        browse_row.AddStretchSpacer(1)
        browse_row.Add(browse_btn, 0)

        content_sizer = wx.BoxSizer(wx.VERTICAL)
        content_sizer.Add(form_sizer, 0, wx.EXPAND | wx.ALL, 12)
        content_sizer.Add(self._fullscreen_checkbox, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 12)
        content_sizer.Add(browse_row, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

        scroller.SetSizer(content_sizer)
        content_sizer.Fit(scroller)
        scroller.FitInside()

        # IMPORTANT: buttons must be parented to the dialog, and placed outside the scroller.
        btn_sizer = self.CreateSeparatedButtonSizer(wx.OK | wx.CANCEL)

        outer = wx.BoxSizer(wx.VERTICAL)
        outer.Add(scroller, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 12)
        if btn_sizer:
            outer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 12)

        self.SetSizer(outer)

        self.SetMinSize(self.FromDIP((560, 240)))
        self.SetSize(self.FromDIP((680, 280)))
        self.Layout()
        self.CentreOnParent()

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

    @property
    def fullscreen_enabled(self) -> bool:
        return self._fullscreen_checkbox.GetValue()


class WindowSelectionDialog(wx.Dialog):
    """Dialog that lists current top-level windows for selection."""

    def __init__(self, parent: wx.Window):
        super().__init__(parent, title="Select a window to capture", style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

        self._list_box = wx.ListBox(self)
        refresh_btn = wx.Button(self, label="Refresh")
        refresh_btn.Bind(wx.EVT_BUTTON, self._populate_windows)

        btn_sizer = self.CreateSeparatedButtonSizer(wx.OK | wx.CANCEL)

        # Touch/DPI friendly sizing.
        button_min = self.FromDIP((140, 52))
        refresh_btn.SetMinSize(button_min)

        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(self._list_box, 1, wx.ALL | wx.EXPAND, 12)
        main_sizer.Add(refresh_btn, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.ALIGN_RIGHT, 12)
        if btn_sizer:
            main_sizer.Add(btn_sizer, 0, wx.ALL | wx.EXPAND, 12)

        self.SetSizer(main_sizer)
        self.SetMinSize(self.FromDIP((520, 420)))
        self.SetSize(self.FromDIP((620, 520)))
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
        return SelectedWindow(
            title=title,
            left=window.left,
            top=window.top,
            width=window.width,
            height=window.height,
            hwnd=getattr(window, "_hWnd", None),
        )


class CaptureCanvas(wx.Panel):
    """Preview canvas with optional crop-rectangle selection.

    The displayed bitmap is scaled for preview, but crop coordinates are mapped
    back to the original image resolution so OCR can run on original pixels.
    """

    def __init__(self, parent: wx.Window, on_crop_changed: Optional[Callable[[Optional[Tuple[int, int, int, int]]], None]] = None):
        super().__init__(parent)

        self.SetBackgroundColour(wx.Colour(245, 245, 245))

        self._orig_pil: Optional["Image.Image"] = None
        self._orig_bmp: Optional[wx.Bitmap] = None

        # Display-scaled bitmap and mapping parameters.
        self._display_bmp: Optional[wx.Bitmap] = None
        self._display_rect: wx.Rect = wx.Rect(0, 0, 0, 0)
        self._scale: float = 1.0

        # Crop selection state in original-image coordinates.
        self._crop_box: Optional[Tuple[int, int, int, int]] = None  # (l, t, r, b)
        self._crop_mode: bool = False
        self._dragging: bool = False
        self._drag_start_img: Optional[Tuple[float, float]] = None
        self._drag_curr_img: Optional[Tuple[float, float]] = None

        self._on_crop_changed = on_crop_changed

        self.Bind(wx.EVT_PAINT, self._on_paint)
        self.Bind(wx.EVT_SIZE, self._on_size)
        self.Bind(wx.EVT_LEFT_DOWN, self._on_left_down)
        self.Bind(wx.EVT_MOTION, self._on_motion)
        self.Bind(wx.EVT_LEFT_UP, self._on_left_up)

    def has_image(self) -> bool:
        return self._orig_pil is not None

    def set_image(self, pil_image: "Image.Image") -> None:
        self._orig_pil = pil_image
        self._orig_bmp = pil_to_bitmap(pil_image)
        self.clear_crop()
        self._refresh_scaled_bitmap()
        self.Refresh()

    def clear_image(self) -> None:
        self._orig_pil = None
        self._orig_bmp = None
        self._display_bmp = None
        self._display_rect = wx.Rect(0, 0, 0, 0)
        self._scale = 1.0
        self.clear_crop()
        self._crop_mode = False
        self.Refresh()

    def set_crop_mode(self, enabled: bool) -> None:
        self._crop_mode = enabled
        if not enabled:
            self._dragging = False
            self._drag_start_img = None
            self._drag_curr_img = None
        self.Refresh()

    def crop_mode(self) -> bool:
        return self._crop_mode

    def clear_crop(self) -> None:
        self._crop_box = None
        self._dragging = False
        self._drag_start_img = None
        self._drag_curr_img = None
        if self._on_crop_changed:
            self._on_crop_changed(None)
        self.Refresh()

    def get_crop_box(self) -> Optional[Tuple[int, int, int, int]]:
        return self._crop_box

    def _on_size(self, event: wx.SizeEvent) -> None:
        self._refresh_scaled_bitmap()
        event.Skip()

    def _refresh_scaled_bitmap(self) -> None:
        if self._orig_bmp is None:
            self._display_bmp = None
            self._display_rect = wx.Rect(0, 0, 0, 0)
            self._scale = 1.0
            return

        client = self.GetClientSize()
        if client.width <= 2 or client.height <= 2:
            return

        img = self._orig_bmp.ConvertToImage()
        iw, ih = img.GetSize()
        if iw <= 0 or ih <= 0:
            return

        scale = min(client.width / iw, client.height / ih)

        dw = max(1, int(iw * scale))
        dh = max(1, int(ih * scale))

        x = (client.width - dw) // 2
        y = (client.height - dh) // 2

        scaled = img.Scale(dw, dh, wx.IMAGE_QUALITY_HIGH)
        self._display_bmp = wx.Bitmap(scaled)
        self._display_rect = wx.Rect(x, y, dw, dh)
        self._scale = scale

    def _client_to_image(self, pt: wx.Point) -> Optional[Tuple[float, float]]:
        if self._orig_pil is None or self._display_bmp is None:
            return None
        if not self._display_rect.Contains(pt):
            return None

        ix = (pt.x - self._display_rect.x) / self._scale
        iy = (pt.y - self._display_rect.y) / self._scale

        # Clamp to image bounds.
        w, h = self._orig_pil.size
        ix = min(max(ix, 0.0), float(w))
        iy = min(max(iy, 0.0), float(h))
        return ix, iy

    def _image_to_client(self, ix: float, iy: float) -> wx.Point:
        cx = int(self._display_rect.x + ix * self._scale)
        cy = int(self._display_rect.y + iy * self._scale)
        return wx.Point(cx, cy)

    def _current_drag_box_client(self) -> Optional[wx.Rect]:
        if self._orig_pil is None:
            return None

        if self._drag_start_img and self._drag_curr_img:
            x1, y1 = self._drag_start_img
            x2, y2 = self._drag_curr_img
        elif self._crop_box:
            l, t, r, b = self._crop_box
            x1, y1, x2, y2 = float(l), float(t), float(r), float(b)
        else:
            return None

        p1 = self._image_to_client(x1, y1)
        p2 = self._image_to_client(x2, y2)
        left = min(p1.x, p2.x)
        top = min(p1.y, p2.y)
        right = max(p1.x, p2.x)
        bottom = max(p1.y, p2.y)
        return wx.Rect(left, top, max(1, right - left), max(1, bottom - top))

    def _on_left_down(self, event: wx.MouseEvent) -> None:
        if not self._crop_mode or self._orig_pil is None:
            return
        img_pt = self._client_to_image(event.GetPosition())
        if img_pt is None:
            return
        self._dragging = True
        self._drag_start_img = img_pt
        self._drag_curr_img = img_pt
        self.CaptureMouse()
        self.Refresh()

    def _on_motion(self, event: wx.MouseEvent) -> None:
        if not self._dragging or not self._crop_mode or self._orig_pil is None:
            return
        if not event.Dragging() or not event.LeftIsDown():
            return
        img_pt = self._client_to_image(event.GetPosition())
        if img_pt is None:
            return
        self._drag_curr_img = img_pt
        self.Refresh()

    def _on_left_up(self, event: wx.MouseEvent) -> None:
        if not self._dragging or self._orig_pil is None:
            return

        if self.HasCapture():
            self.ReleaseMouse()

        self._dragging = False
        end_pt = self._client_to_image(event.GetPosition())
        if end_pt is None:
            self._drag_start_img = None
            self._drag_curr_img = None
            self.Refresh()
            return

        if self._drag_start_img is None:
            return

        x1, y1 = self._drag_start_img
        x2, y2 = end_pt

        left = int(min(x1, x2))
        top = int(min(y1, y2))
        right = int(max(x1, x2))
        bottom = int(max(y1, y2))

        # Enforce a minimal crop area.
        if right - left < 10 or bottom - top < 10:
            self._crop_box = None
            if self._on_crop_changed:
                self._on_crop_changed(None)
        else:
            w, h = self._orig_pil.size
            left = max(0, min(left, w - 1))
            top = max(0, min(top, h - 1))
            right = max(left + 1, min(right, w))
            bottom = max(top + 1, min(bottom, h))
            self._crop_box = (left, top, right, bottom)
            if self._on_crop_changed:
                self._on_crop_changed(self._crop_box)

        self._drag_start_img = None
        self._drag_curr_img = None
        # Exit crop mode after a selection to keep the workflow simple.
        self._crop_mode = False
        self.Refresh()

    def _on_paint(self, event: wx.PaintEvent) -> None:
        dc = wx.BufferedPaintDC(self)
        dc.Clear()

        # Frame border.
        client = self.GetClientRect()
        dc.SetPen(wx.Pen(wx.Colour(210, 210, 210), 1))
        dc.SetBrush(wx.Brush(self.GetBackgroundColour()))
        dc.DrawRectangle(client)

        if self._display_bmp is None:
            dc.SetTextForeground(wx.Colour(110, 110, 110))
            msg = "Preview will appear here after you take a capture."
            tw, th = dc.GetTextExtent(msg)
            dc.DrawText(msg, (client.width - tw) // 2, (client.height - th) // 2)
            return

        # Draw image.
        dc.DrawBitmap(self._display_bmp, self._display_rect.x, self._display_rect.y, True)

        # Draw crop rectangle (existing or in-progress).
        crop_rect = self._current_drag_box_client()
        if crop_rect:
            # Thicker pen for touchscreen visibility.
            dc.SetPen(wx.Pen(wx.Colour(0, 120, 215), self.FromDIP(3), style=wx.PENSTYLE_SHORT_DASH))
            dc.SetBrush(wx.Brush(wx.Colour(0, 120, 215, 40)))
            dc.DrawRectangle(crop_rect)


class OCRFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="AI Agent - Window Capture & OCR", size=(1200, 760))

        self._selected_window: Optional[SelectedWindow] = None
        self._captured_image: Optional["Image.Image"] = None
        self._tesseract_path: Optional[str] = self._detect_local_tesseract()
        self._dependency_state: Dict[str, bool] = {}
        self._install_attempted = False

        self._crop_box: Optional[Tuple[int, int, int, int]] = None
        self._is_fullscreen: bool = False
        self._button_base_size = self.FromDIP((170, 64))
        self._button_base_font = None  # Will be captured once buttons are created.
        self._action_buttons: List[wx.Button] = []
        self._header_base_font: Optional[wx.Font] = None
        self._selected_label_base_font: Optional[wx.Font] = None

        self._build_ui()
        self._apply_tesseract_path()
        self._check_dependencies()

        # Keyboard shortcuts for full screen.
        self.Bind(wx.EVT_CHAR_HOOK, self._on_key_press)

        self.Bind(wx.EVT_CLOSE, self._on_close)

    # ---------- UI ----------

    def _build_ui(self) -> None:
        self.CreateStatusBar(number=1)
        self.SetStatusText("Ready.")
        info_panel = wx.Panel(self)

        self._header = wx.StaticText(
            info_panel,
            label=(
                "Workflow: Select a window → Take a capture → (Optional) Crop → OCR.\n"
                "OCR runs locally via Tesseract."
            ),
        )

        self._selected_label = wx.StaticText(info_panel, label="Selected window: (none)")
        self._selected_label.SetForegroundColour(wx.Colour(80, 80, 80))
        self._header_base_font = self._header.GetFont()
        self._selected_label_base_font = self._selected_label.GetFont()

        info_sizer = wx.BoxSizer(wx.VERTICAL)
        info_sizer.Add(self._header, 0, wx.EXPAND | wx.BOTTOM, self.FromDIP(6))
        info_sizer.Add(self._selected_label, 0, wx.EXPAND)

        info_panel.SetSizer(info_sizer)

        # Main split area: preview (left) and OCR output (right).
        splitter = wx.SplitterWindow(self, style=wx.SP_LIVE_UPDATE | wx.SP_3D)
        left_panel = wx.Panel(splitter)
        right_panel = wx.Panel(splitter)

        self._canvas = CaptureCanvas(left_panel, on_crop_changed=self._on_crop_changed)

        left_box = wx.StaticBoxSizer(wx.StaticBox(left_panel, label="Capture Preview"), wx.VERTICAL)
        left_box.Add(self._canvas, 1, wx.EXPAND | wx.ALL, self.FromDIP(8))
        left_panel.SetSizer(left_box)

        self._ocr_output = wx.TextCtrl(
            right_panel,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
        )
        right_box = wx.StaticBoxSizer(wx.StaticBox(right_panel, label="OCR Output"), wx.VERTICAL)
        right_box.Add(self._ocr_output, 1, wx.EXPAND | wx.ALL, self.FromDIP(8))
        right_panel.SetSizer(right_box)

        splitter.SplitVertically(left_panel, right_panel, sashPosition=self.FromDIP(440))
        splitter.SetMinimumPaneSize(self.FromDIP(260))

        # Touch-friendly button row (wraps on narrow windows).
        controls_panel = wx.Panel(self)
        controls_sizer = wx.WrapSizer(wx.HORIZONTAL)

        self._start_btn = self._make_action_button(controls_panel, "Select Window")
        self._start_btn.Bind(wx.EVT_BUTTON, self._select_window)

        self._take_btn = self._make_action_button(controls_panel, "Take")
        self._take_btn.Disable()
        self._take_btn.Bind(wx.EVT_BUTTON, self._take_capture)

        self._crop_btn = self._make_action_button(controls_panel, "Crop")
        self._crop_btn.Disable()
        self._crop_btn.Bind(wx.EVT_BUTTON, self._toggle_crop)

        self._process_btn = self._make_action_button(controls_panel, "OCR")
        self._process_btn.Disable()
        self._process_btn.Bind(wx.EVT_BUTTON, self._process_capture)

        self._settings_btn = self._make_action_button(controls_panel, "Settings")
        self._settings_btn.Bind(wx.EVT_BUTTON, self._open_settings)

        for btn in (self._start_btn, self._take_btn, self._crop_btn, self._process_btn, self._settings_btn):
            self._action_buttons.append(btn)
            if self._button_base_font is None:
                self._button_base_font = btn.GetFont()
            controls_sizer.Add(btn, 0, wx.ALL, self.FromDIP(6))

        controls_panel.SetSizer(controls_sizer)

        main = wx.BoxSizer(wx.VERTICAL)
        main.Add(info_panel, 0, wx.EXPAND | wx.ALL, self.FromDIP(12))
        main.Add(splitter, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, self.FromDIP(12))
        main.Add(controls_panel, 0, wx.EXPAND | wx.ALL, self.FromDIP(12))

        self.SetSizer(main)

        self._apply_touch_scaling(self._is_fullscreen)

        # Dynamic wrap for header on resize.
        self.Bind(wx.EVT_SIZE, self._on_frame_size)

    def _on_frame_size(self, event: wx.SizeEvent) -> None:
        try:
            # Wrap based on available width; keep some margin.
            width = max(200, self.GetClientSize().width - self.FromDIP(40))
            self._header.Wrap(width)
        except Exception:
            pass
        event.Skip()

    def _on_key_press(self, event: wx.KeyEvent) -> None:
        if event.GetKeyCode() == wx.WXK_F11:
            self._toggle_fullscreen()
            return
        event.Skip()

    def _make_action_button(self, parent: wx.Window, label: str) -> wx.Button:
        btn = wx.Button(parent, label=label)
        font = btn.GetFont()
        font.SetPointSize(max(font.GetPointSize(), 11))
        font.SetWeight(wx.FONTWEIGHT_SEMIBOLD)
        btn.SetFont(font)
        btn.SetMinSize(self._scaled_button_size())
        return btn

    def _scaled_button_size(self) -> wx.Size:
        factor = 1.3 if self._is_fullscreen else 1.0
        return wx.Size(int(self._button_base_size.width * factor), int(self._button_base_size.height * factor))

    def _apply_touch_scaling(self, fullscreen: bool) -> None:
        if not self._action_buttons:
            return

        factor = 1.3 if fullscreen else 1.0
        font_factor = 1.2 if fullscreen else 1.0

        for btn in self._action_buttons:
            btn.SetMinSize(wx.Size(int(self._button_base_size.width * factor), int(self._button_base_size.height * factor)))
            font = (self._button_base_font or btn.GetFont()).Bold()
            base_size = self._button_base_font.GetPointSize() if self._button_base_font else font.GetPointSize()
            font.SetPointSize(int(round(base_size * font_factor)))
            btn.SetFont(font)

        # Slightly bump header and label text for readability in full screen.
        for label in (self._header, self._selected_label):
            font = label.GetFont()
            if label is self._header and self._header_base_font:
                base_size = self._header_base_font.GetPointSize()
            elif label is self._selected_label and self._selected_label_base_font:
                base_size = self._selected_label_base_font.GetPointSize()
            else:
                base_size = font.GetPointSize()
            font.SetPointSize(int(round(base_size * font_factor)))
            label.SetFont(font)
        self.Layout()

    def _toggle_fullscreen(self, target_state: Optional[bool] = None) -> None:
        desired = (not self._is_fullscreen) if target_state is None else target_state
        if desired == self._is_fullscreen:
            return

        self._is_fullscreen = desired
        self.ShowFullScreen(desired, style=wx.FULLSCREEN_ALL)
        self._apply_touch_scaling(desired)
        message = "Full screen enabled. Press F11 or use Settings to exit." if desired else "Exited full screen."
        self._set_status(message)

    def _set_status(self, message: str) -> None:
        self.SetStatusText(message)

    # ---------- Dependency / configuration ----------

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
                "The following Python packages are required but not available:\n"
                + "\n".join(f" • {name}" for name in missing)
                + "\n\nTried installing them automatically but some are still missing."
            )
            wx.MessageBox(message, "Missing dependencies", wx.ICON_ERROR | wx.OK, parent=self)

        self._start_btn.Enable(self._dependency_state.get("pygetwindow", False))
        if not self._dependency_state.get("pygetwindow", False):
            self._set_status("Install pygetwindow to enable window selection.")

    def _detect_dependency_state(self) -> Dict[str, bool]:
        return {
            "pygetwindow": gw is not None,
            "pyautogui": pyautogui is not None,
            "pillow": Image is not None,
            "pytesseract": pytesseract is not None,
            "pywin32": self._has_pywin32(),
        }

    def _has_pywin32(self) -> bool:
        required_modules = ("win32gui", "win32ui", "win32con")
        return all(importlib.util.find_spec(module) is not None for module in required_modules)

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

    # ---------- Actions ----------

    def _open_settings(self, event: wx.CommandEvent) -> None:  # pragma: no cover - UI interaction
        dialog = SettingsDialog(self, self._tesseract_path, self._is_fullscreen)
        if dialog.ShowModal() == wx.ID_OK:
            self._tesseract_path = dialog.tesseract_path
            self._apply_tesseract_path()
            status_msg = "Settings saved."
            if dialog.fullscreen_enabled != self._is_fullscreen:
                self._toggle_fullscreen(dialog.fullscreen_enabled)
                status_msg = "Settings saved. Full screen enabled." if dialog.fullscreen_enabled else "Settings saved. Full screen disabled."
            self._set_status(status_msg)
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
        self._selected_label.SetLabel(f"Selected window: {selection.title}")
        self._take_btn.Enable()

        # Reset capture-related state.
        self._captured_image = None
        self._canvas.clear_image()
        self._crop_btn.SetLabel("Crop")
        self._crop_btn.Disable()
        self._process_btn.Disable()
        self._ocr_output.SetValue("")
        self._set_status("Window selected. Tap 'Take' to capture.")

    def _capture_selected_window(self) -> Optional["Image.Image"]:
        if self._selected_window is None:
            return None
        if not self._ensure_dependency("pillow", "Pillow"):
            return None

        # Prefer Win32 capture when possible for better fidelity and no screen overlay issues.
        if sys.platform == "win32" and self._selected_window.hwnd and self._ensure_dependency("pywin32", "pywin32"):
            win32_capture = self._capture_with_win32(self._selected_window)
            if win32_capture is not None:
                return win32_capture

        if not self._ensure_dependency("pyautogui", "pyautogui"):
            return None
        if pyautogui is None:
            return None
        return pyautogui.screenshot(region=self._selected_window.region)

    def _capture_with_win32(self, selection: SelectedWindow) -> Optional["Image.Image"]:
        import win32con
        import win32gui
        import win32ui

        hwnd = selection.hwnd
        if hwnd is None:
            return None

        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width, height = right - left, bottom - top
        if width <= 0 or height <= 0:
            return None

        hwnd_dc = win32gui.GetWindowDC(hwnd)
        if hwnd_dc == 0:
            return None

        window_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        mem_dc = window_dc.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(window_dc, width, height)
        mem_dc.SelectObject(bitmap)

        def _print_window(h: int, hdc: int, flags: int) -> int:
            if hasattr(win32gui, "PrintWindow"):
                return int(win32gui.PrintWindow(h, hdc, flags))
            from ctypes import windll, wintypes

            user32 = windll.user32
            user32.PrintWindow.argtypes = [wintypes.HWND, wintypes.HDC, wintypes.UINT]
            user32.PrintWindow.restype = wintypes.BOOL
            return int(user32.PrintWindow(h, hdc, flags))

        try:
            flags = int(getattr(win32con, "PW_RENDERFULLCONTENT", 0x00000002))
            result = _print_window(hwnd, mem_dc.GetSafeHdc(), flags)
            if result != 1 and flags != 0:
                result = _print_window(hwnd, mem_dc.GetSafeHdc(), 0)
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

    def _take_capture(self, event: wx.CommandEvent) -> None:  # pragma: no cover - UI interaction
        if self._selected_window is None:
            wx.MessageBox("Please select a window first.", "No window", wx.ICON_INFORMATION | wx.OK, parent=self)
            return

        try:
            screenshot = self._capture_selected_window()
        except Exception as exc:  # pragma: no cover - user environment specific
            wx.MessageBox(str(exc), "Capture failed", wx.ICON_ERROR | wx.OK, parent=self)
            return

        if screenshot is None:
            wx.MessageBox(
                "Could not capture the selected window. Make sure it is visible and try again.",
                "Capture failed",
                wx.ICON_ERROR | wx.OK,
                parent=self,
            )
            return

        self._captured_image = screenshot
        self._canvas.set_image(screenshot)

        # Enable crop + OCR now that a capture exists.
        self._crop_btn.Enable()
        self._process_btn.Enable()

        # Reset previous OCR/crop state.
        self._ocr_output.SetValue("")
        self._crop_box = None
        self._crop_btn.SetLabel("Crop")

        # No popup here per requirements; update status only.
        self._set_status("Capture ready. Optional: tap 'Crop' then drag a rectangle; then tap 'OCR'.")

    def _toggle_crop(self, event: wx.CommandEvent) -> None:  # pragma: no cover - UI interaction
        if self._captured_image is None or not self._canvas.has_image():
            return

        # If crop mode is active, pressing again cancels crop mode.
        if self._canvas.crop_mode():
            self._canvas.set_crop_mode(False)
            self._set_status("Crop mode cancelled.")
            return


        # If a crop is already set, this button clears it.
        if self._crop_box is not None:
            self._canvas.clear_crop()
            self._crop_box = None
            self._crop_btn.SetLabel("Crop")
            self._set_status("Crop cleared. OCR will run on the full capture.")
            return

        # Otherwise enable crop mode; user drags on the preview.
        self._canvas.set_crop_mode(True)
        self._set_status("Crop mode enabled: drag a rectangle on the preview to select OCR region.")

    def _on_crop_changed(self, crop_box: Optional[Tuple[int, int, int, int]]) -> None:
        self._crop_box = crop_box
        if crop_box is None:
            self._crop_btn.SetLabel("Crop")
            return
        self._crop_btn.SetLabel("Clear Crop")
        l, t, r, b = crop_box
        self._set_status(f"Crop set: ({l}, {t}) → ({r}, {b}). Tap 'OCR' to process the cropped region.")

    def _process_capture(self, event: wx.CommandEvent) -> None:  # pragma: no cover - UI interaction
        if self._captured_image is None:
            wx.MessageBox("Take a capture first.", "No capture", wx.ICON_INFORMATION | wx.OK, parent=self)
            return
        if not self._ensure_dependency("pytesseract", "pytesseract"):
            return

        # Crop must be applied to the ORIGINAL image (not the preview).
        image_for_ocr = self._captured_image
        if self._crop_box is not None:
            try:
                image_for_ocr = self._captured_image.crop(self._crop_box)
            except Exception as exc:
                wx.MessageBox(str(exc), "Crop failed", wx.ICON_ERROR | wx.OK, parent=self)
                return

        try:
            text = pytesseract.image_to_string(image_for_ocr)
        except Exception as exc:  # pragma: no cover - user environment specific
            wx.MessageBox(str(exc), "Processing failed", wx.ICON_ERROR | wx.OK, parent=self)
            return

        self._ocr_output.SetValue(text.strip())
        self._set_status("OCR complete.")

    # ---------- Lifecycle ----------

    def _on_close(self, event: wx.CloseEvent) -> None:  # pragma: no cover - UI interaction
        reason = (
            "AI Agent is about to quit. Any selected window, capture, crop selection, "
            "or OCR output will be lost.\n\nDo you want to exit?"
        )
        if wx.MessageBox(reason, "Exit", wx.ICON_QUESTION | wx.YES_NO, parent=self) == wx.YES:
            self.Destroy()
        else:
            event.Veto()
            self._set_status("Close cancelled; continuing session.")


def main() -> int:
    app = wx.App()
    frame = OCRFrame()
    frame.Show()
    app.MainLoop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
