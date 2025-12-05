"""Phase 1 desktop GUI for window capture and OCR.

This application targets Windows and uses PyQt5 for the GUI. It allows
users to select an open window, capture its contents, and run OCR on the
capture using a local Tesseract installation.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import List, Optional

from PyQt5 import QtCore, QtGui, QtWidgets

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
    from PIL import Image, ImageQt
except Exception:  # pragma: no cover - optional dependency
    Image = None
    ImageQt = None


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


class SettingsDialog(QtWidgets.QDialog):
    """Simple settings dialog allowing tesseract executable configuration."""

    def __init__(self, tesseract_path: Optional[str], parent: Optional[QtWidgets.QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.resize(500, 150)

        self._path_edit = QtWidgets.QLineEdit(self)
        if tesseract_path:
            self._path_edit.setText(tesseract_path)
        browse_btn = QtWidgets.QPushButton("Browse…", self)
        browse_btn.clicked.connect(self._browse)

        form = QtWidgets.QFormLayout()
        form.addRow("Tesseract executable:", self._path_edit)

        btn_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel, parent=self
        )
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)

        layout = QtWidgets.QVBoxLayout()
        path_layout = QtWidgets.QHBoxLayout()
        path_layout.addWidget(self._path_edit)
        path_layout.addWidget(browse_btn)
        layout.addLayout(form)
        layout.addLayout(path_layout)
        layout.addWidget(btn_box)
        self.setLayout(layout)

    def _browse(self) -> None:
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select tesseract.exe", "", "Executable (*.exe)")
        if file_path:
            self._path_edit.setText(file_path)

    @property
    def tesseract_path(self) -> Optional[str]:
        text = self._path_edit.text().strip()
        return text or None


class WindowSelectionDialog(QtWidgets.QDialog):
    """Dialog that lists current top-level windows for selection."""

    def __init__(self, parent: Optional[QtWidgets.QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Select a window to capture")
        self.resize(400, 300)

        self._list_widget = QtWidgets.QListWidget(self)
        self._list_widget.doubleClicked.connect(self.accept)

        refresh_btn = QtWidgets.QPushButton("Refresh", self)
        refresh_btn.clicked.connect(self.populate_windows)

        btn_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel, parent=self
        )
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)

        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self._list_widget)
        layout.addWidget(refresh_btn)
        layout.addWidget(btn_box)
        self.setLayout(layout)

        self.populate_windows()

    def populate_windows(self) -> None:
        self._list_widget.clear()
        if gw is None:
            QtWidgets.QMessageBox.warning(self, "pygetwindow missing", "Install pygetwindow to list windows.")
            return

        titles: List[str] = [title for title in gw.getAllTitles() if title.strip()]
        titles.sort()
        self._list_widget.addItems(titles)

    def selected_title(self) -> Optional[str]:
        items = self._list_widget.selectedItems()
        if not items:
            return None
        return items[0].text()

    def get_selection(self) -> Optional[SelectedWindow]:
        if self.exec_() != QtWidgets.QDialog.Accepted:
            return None
        title = self.selected_title()
        if not title or gw is None:
            return None
        window = gw.getWindowsWithTitle(title)[0]
        return SelectedWindow(title, window.left, window.top, window.width, window.height)


class PreviewWidget(QtWidgets.QLabel):
    """Displays a scaled preview image."""

    def __init__(self):
        super().__init__()
        self.setFixedSize(640, 360)
        self.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Plain)
        self.setAlignment(QtCore.Qt.AlignCenter)
        self.setText("Preview")

    def update_image(self, pixmap: QtGui.QPixmap) -> None:
        scaled = pixmap.scaled(self.size(), QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.setPixmap(scaled)


class OCRApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Phase 1 Screen Capture")
        self.resize(900, 500)

        self._selected_window: Optional[SelectedWindow] = None
        self._captured_image: Optional[Image.Image] = None
        self._tesseract_path: Optional[str] = None

        self._build_ui()

    def _build_ui(self) -> None:
        header = QtWidgets.QLabel(
            "Select a window, take a snapshot, then process it locally.\n"
            "Works best on Windows with Tesseract installed."
        )
        header.setWordWrap(True)

        self._preview = PreviewWidget()

        self._take_btn = QtWidgets.QPushButton("Take")
        self._take_btn.setEnabled(False)
        self._take_btn.clicked.connect(self.take_capture)

        self._process_btn = QtWidgets.QPushButton("Process")
        self._process_btn.setEnabled(False)
        self._process_btn.clicked.connect(self.process_capture)

        start_btn = QtWidgets.QPushButton("Start")
        start_btn.clicked.connect(self.select_window)

        settings_btn = QtWidgets.QPushButton("Settings")
        settings_btn.clicked.connect(self.open_settings)

        button_col = QtWidgets.QVBoxLayout()
        button_col.addWidget(start_btn)
        button_col.addWidget(settings_btn)
        button_col.addStretch()
        button_col.addWidget(self._take_btn)
        button_col.addWidget(self._process_btn)
        button_col.addStretch()

        main_layout = QtWidgets.QVBoxLayout()
        main_layout.addWidget(header)

        content_layout = QtWidgets.QHBoxLayout()
        content_layout.addLayout(button_col)
        content_layout.addWidget(self._preview, stretch=1)

        main_layout.addLayout(content_layout)
        self.setLayout(main_layout)

    def open_settings(self) -> None:
        dialog = SettingsDialog(self._tesseract_path, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self._tesseract_path = dialog.tesseract_path
            if pytesseract and self._tesseract_path:
                pytesseract.pytesseract.tesseract_cmd = self._tesseract_path

    def select_window(self) -> None:
        if gw is None:
            QtWidgets.QMessageBox.critical(self, "Missing dependency", "Install pygetwindow to list windows.")
            return
        dialog = WindowSelectionDialog(self)
        selection = dialog.get_selection()
        if not selection:
            return
        self._selected_window = selection
        self._take_btn.setEnabled(True)
        self._process_btn.setEnabled(False)
        self.status_message(f"Selected window: {selection.title}")

    def take_capture(self) -> None:
        if self._selected_window is None:
            QtWidgets.QMessageBox.information(self, "No window", "Please select a window first.")
            return
        if pyautogui is None:
            QtWidgets.QMessageBox.critical(self, "Missing dependency", "Install pyautogui to capture screenshots.")
            return
        if ImageQt is None:
            QtWidgets.QMessageBox.critical(self, "Missing dependency", "Install Pillow for image handling.")
            return

        try:
            screenshot = pyautogui.screenshot(region=self._selected_window.region)
        except Exception as exc:  # pragma: no cover - user environment specific
            QtWidgets.QMessageBox.critical(self, "Capture failed", str(exc))
            return

        self._captured_image = screenshot
        qimage = ImageQt.ImageQt(screenshot)
        pixmap = QtGui.QPixmap.fromImage(qimage)
        self._preview.update_image(pixmap)
        self._process_btn.setEnabled(True)
        self.status_message("Capture ready for processing")

    def process_capture(self) -> None:
        if self._captured_image is None:
            QtWidgets.QMessageBox.information(self, "No capture", "Take a capture first.")
            return
        if pytesseract is None:
            QtWidgets.QMessageBox.critical(self, "Missing dependency", "Install pytesseract for OCR.")
            return

        try:
            text = pytesseract.image_to_string(self._captured_image)
        except Exception as exc:  # pragma: no cover - user environment specific
            QtWidgets.QMessageBox.critical(self, "Processing failed", str(exc))
            return

        dialog = QtWidgets.QMessageBox(self)
        dialog.setWindowTitle("OCR Result")
        dialog.setText("Extracted text:")
        dialog.setDetailedText(text)
        dialog.exec_()

    def status_message(self, message: str) -> None:
        QtWidgets.QMessageBox.information(self, "Status", message)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    window = OCRApp()
    window.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
