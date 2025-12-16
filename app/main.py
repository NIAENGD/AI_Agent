"""Phase 1 desktop GUI for window capture and OCR.

This application targets Windows and uses PyQt5 for the GUI. It allows
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
    from PIL import Image
    from PIL.ImageQt import ImageQt as PilImageQt
except Exception:  # pragma: no cover - optional dependency
    Image = None
    PilImageQt = None


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
        self._tesseract_path: Optional[str] = self._detect_local_tesseract()
        self._dependency_state: Dict[str, bool] = {}
        self._install_attempted = False

        self._build_ui()
        self._apply_tesseract_path()
        self._check_dependencies()

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

        self._start_btn = QtWidgets.QPushButton("Start")
        self._start_btn.clicked.connect(self.select_window)

        settings_btn = QtWidgets.QPushButton("Settings")
        settings_btn.clicked.connect(self.open_settings)

        button_col = QtWidgets.QVBoxLayout()
        button_col.addWidget(self._start_btn)
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
            QtWidgets.QMessageBox.critical(self, "Missing dependencies", message)

        self._start_btn.setEnabled(self._dependency_state.get("pygetwindow", False))
        if not self._dependency_state.get("pygetwindow", False):
            self.status_message("Install pygetwindow to enable window selection.")

    def _detect_dependency_state(self) -> Dict[str, bool]:
        return {
            "pygetwindow": gw is not None,
            "pyautogui": pyautogui is not None,
            "pillow": Image is not None and PilImageQt is not None,
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

        QtWidgets.QMessageBox.critical(
            self,
            "Missing dependency",
            f"{friendly_name} is not installed. Please run: pip install -r requirements.txt",
        )
        return False

    def _attempt_install_requirements(self) -> bool:
        if self._install_attempted:
            return False

        self._install_attempted = True
        if not self._install_requirements():
            return False

        self._refresh_optional_dependencies()
        QtWidgets.QMessageBox.information(
            self,
            "Dependencies installed",
            "Required packages were installed. Please retry your action.",
        )
        return True

    def _install_requirements(self) -> bool:
        requirements_path = Path(__file__).resolve().parent.parent / "requirements.txt"
        if not requirements_path.exists():
            QtWidgets.QMessageBox.critical(
                self,
                "Missing requirements.txt",
                f"Could not find requirements file at {requirements_path}.",
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
            QtWidgets.QMessageBox.critical(
                self,
                "Dependency installation failed",
                f"pip install failed with:\n{details}",
            )
            return False

        return True

    def _refresh_optional_dependencies(self) -> None:
        global gw, pyautogui, pytesseract, Image, PilImageQt

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
        if Image is None or PilImageQt is None:
            try:
                from PIL import Image as PilImage
                from PIL.ImageQt import ImageQt as PilImageQtImported
                Image = PilImage
                PilImageQt = PilImageQtImported
            except Exception:  # pragma: no cover - optional dependency
                Image = None
                PilImageQt = None

    def open_settings(self) -> None:
        dialog = SettingsDialog(self._tesseract_path, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self._tesseract_path = dialog.tesseract_path
            self._apply_tesseract_path()

    def select_window(self) -> None:
        if not self._ensure_dependency("pygetwindow", "pygetwindow"):
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
        if not self._ensure_dependency("pyautogui", "pyautogui"):
            return
        if not self._ensure_dependency("pillow", "Pillow"):
            return

        try:
            screenshot = pyautogui.screenshot(region=self._selected_window.region)
        except Exception as exc:  # pragma: no cover - user environment specific
            QtWidgets.QMessageBox.critical(self, "Capture failed", str(exc))
            return

        self._captured_image = screenshot
        qimage = PilImageQt(screenshot)
        pixmap = QtGui.QPixmap.fromImage(qimage)
        self._preview.update_image(pixmap)
        self._process_btn.setEnabled(True)
        self.status_message("Capture ready for processing")

    def process_capture(self) -> None:
        if self._captured_image is None:
            QtWidgets.QMessageBox.information(self, "No capture", "Take a capture first.")
            return
        if not self._ensure_dependency("pytesseract", "pytesseract"):
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

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # type: ignore[override]
        reason = (
            "AI Agent is about to quit. Any selected window or OCR results will be lost.\n\n"
            "Do you want to exit?"
        )
        result = QtWidgets.QMessageBox.question(
            self,
            "Exit AI Agent",
            reason,
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )
        if result == QtWidgets.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
            self.status_message("Close cancelled; continuing session.")

    def status_message(self, message: str) -> None:
        QtWidgets.QMessageBox.information(self, "Status", message)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    window = OCRApp()
    window.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
