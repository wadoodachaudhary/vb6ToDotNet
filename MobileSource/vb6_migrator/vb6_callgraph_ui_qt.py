#!/usr/bin/env python3
"""
Qt shell for VB6 -> Blazor migrator.

Panels (left to right):
1) .NET Blazor project tree
2) VB6 file tree/grid
3) VB6 source
4) .NET Blazor source (+ output/errors/run log)
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path
from typing import Dict, List, Tuple

# Prevent noisy Qt accessibility table warnings on rapid model refreshes.
os.environ.setdefault("QT_ACCESSIBILITY", "0")

try:
    from PySide6.QtCore import Qt, QSize, QTimer, Signal
    from PySide6.QtGui import (
        QAction, QColor, QFont, QIcon, QPainter, QPixmap, QPen,
        QSyntaxHighlighter, QTextCharFormat,
    )
    from PySide6.QtWidgets import (
        QApplication,
        QAbstractItemView,
        QButtonGroup,
        QCheckBox,
        QComboBox,
        QDialog,
        QFileDialog,
        QFormLayout,
        QHBoxLayout,
        QHeaderView,
        QInputDialog,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPlainTextEdit,
        QPushButton,
        QRadioButton,
        QSplitter,
        QStackedWidget,
        QTableWidget,
        QTabWidget,
        QTableWidgetItem,
        QToolBar,
        QTreeWidget,
        QTreeWidgetItem,
        QVBoxLayout,
        QWidget,
    )
    QT_BINDING = "PySide6"
except Exception:
    from PyQt6.QtCore import Qt, QSize, QTimer, pyqtSignal as Signal
    from PyQt6.QtGui import (
        QAction, QColor, QFont, QIcon, QPainter, QPixmap, QPen,
        QSyntaxHighlighter, QTextCharFormat,
    )
    from PyQt6.QtWidgets import (
        QApplication,
        QAbstractItemView,
        QButtonGroup,
        QCheckBox,
        QComboBox,
        QDialog,
        QFileDialog,
        QFormLayout,
        QHBoxLayout,
        QHeaderView,
        QInputDialog,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPlainTextEdit,
        QPushButton,
        QRadioButton,
        QSplitter,
        QStackedWidget,
        QTableWidget,
        QTabWidget,
        QTableWidgetItem,
        QToolBar,
        QTreeWidget,
        QTreeWidgetItem,
        QVBoxLayout,
        QWidget,
    )
    QT_BINDING = "PyQt6"

import vb6_callgraph_ui as backend


ROLE_PATH = int(Qt.ItemDataRole.UserRole)
ROLE_FLAG = ROLE_PATH + 1
ROLE_COMP_ID = ROLE_PATH + 2
ROLE_STACK = ROLE_PATH + 3
ROLE_NODE_KIND = ROLE_PATH + 4
MAX_SEARCH_MATCHES = 200
QT_RUNTIME_LOG = Path(__file__).resolve().parent / ".logs" / "qt-runtime.log"
APP_CONFIG_PATH = Path(__file__).resolve().parent / "migrator_config.json"


def _install_exception_logging() -> None:
    QT_RUNTIME_LOG.parent.mkdir(parents=True, exist_ok=True)

    def _hook(exc_type, exc, tb):
        try:
            with QT_RUNTIME_LOG.open("a", encoding="utf-8") as fh:
                fh.write(f"\n=== Unhandled exception {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")
                traceback.print_exception(exc_type, exc, tb, file=fh)
        except Exception:
            pass
        sys.__excepthook__(exc_type, exc, tb)

    sys.excepthook = _hook


def _default_app_config() -> dict:
    script_dir = Path(__file__).resolve().parent
    projects_root = (script_dir / "projects").resolve()
    return {
        "projects_root": str(projects_root),
        "project_name": "PrescionBuilder",
        "project_dir": str((projects_root / "PrescionBuilder").resolve()),
        "root": "/Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6",
        "prefer_hfest": True,
        "ui": True,
        "output_project": "/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/HomeFront.csproj",
        "output_dir": "/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/Components/Pages/Migrated",
        "context_dir": "/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/",
        "instructions": "Convert teh attached VB6 files to .NET Blazor using the HF Controls. Employ and DBWrapperSqlServer.cs for database queries.",
        "model": "gpt-5.2-codex",
        "temperature": "0.2",
        "vb6_image_dir": "/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/wwwroot/images/vb6",
        "dotnet_image_dir": "/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/wwwroot/images/dotnet",
    }


def _save_app_config(config: dict) -> None:
    APP_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    APP_CONFIG_PATH.write_text(json.dumps(config, indent=2, sort_keys=True), encoding="utf-8")


def _load_app_config() -> dict:
    config = _default_app_config()
    if APP_CONFIG_PATH.exists():
        try:
            raw = json.loads(APP_CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                for key, value in raw.items():
                    if value is not None:
                        config[key] = value
        except Exception:
            pass

    projects_root = Path(str(config.get("projects_root") or _default_app_config()["projects_root"])).expanduser().resolve()
    config["projects_root"] = str(projects_root)

    # Honor user wording "PrecionBuilder" while mapping to existing package "PrescionBuilder".
    requested_name = str(config.get("project_name") or "PrescionBuilder").strip() or "PrescionBuilder"
    normalized = requested_name.replace(" ", "")
    if normalized.lower() in {"precionbuilder", "prescionbuilder"}:
        requested_name = "PrescionBuilder"
    config["project_name"] = requested_name

    preferred_dir = projects_root / requested_name
    configured_dir_raw = str(config.get("project_dir") or "").strip()
    configured_dir = Path(configured_dir_raw).expanduser().resolve() if configured_dir_raw else preferred_dir.resolve()
    if configured_dir.name.lower() in {"precionbuilder", "prescionbuilder"}:
        configured_dir = (projects_root / "PrescionBuilder").resolve()
    config["project_dir"] = str(configured_dir)

    for key in (
        "root",
        "output_project",
        "output_dir",
        "context_dir",
        "vb6_image_dir",
        "dotnet_image_dir",
    ):
        raw = str(config.get(key) or "").strip()
        if raw:
            config[key] = str(Path(raw).expanduser())

    config["prefer_hfest"] = bool(config.get("prefer_hfest", True))
    config["ui"] = bool(config.get("ui", True))
    config["temperature"] = str(config.get("temperature", "0.2"))

    _save_app_config(config)
    return config


# ---------------------------------------------------------------------------
#  Programmatic toolbar icons
# ---------------------------------------------------------------------------

def _make_icon(draw_fn, size: int = 32) -> QIcon:
    """Create a QIcon by drawing onto a transparent QPixmap."""
    pix = QPixmap(size, size)
    pix.fill(Qt.GlobalColor.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    draw_fn(p, size)
    p.end()
    return QIcon(pix)


def _draw_convert_icon(p: QPainter, s: int) -> None:
    """Single right arrow — Convert one file."""
    pen = QPen(QColor("#1f4e79"), 2.8)
    p.setPen(pen)
    cx, cy = s // 2, s // 2
    p.drawLine(cx - 8, cy, cx + 8, cy)
    p.drawLine(cx + 3, cy - 6, cx + 8, cy)
    p.drawLine(cx + 3, cy + 6, cx + 8, cy)


def _draw_convert_selected_icon(p: QPainter, s: int) -> None:
    """Double right arrow — Convert selected files."""
    pen = QPen(QColor("#1f4e79"), 2.4)
    p.setPen(pen)
    cx, cy = s // 2, s // 2
    # first arrow
    p.drawLine(cx - 10, cy, cx + 2, cy)
    p.drawLine(cx - 3, cy - 5, cx + 2, cy)
    p.drawLine(cx - 3, cy + 5, cx + 2, cy)
    # second arrow
    p.drawLine(cx - 2, cy, cx + 10, cy)
    p.drawLine(cx + 5, cy - 5, cx + 10, cy)
    p.drawLine(cx + 5, cy + 5, cx + 10, cy)


def _draw_compile_icon(p: QPainter, s: int) -> None:
    """Hammer shape — Compile."""
    pen = QPen(QColor("#2d7a2d"), 2.6)
    p.setPen(pen)
    # handle (diagonal)
    p.drawLine(s // 2 - 6, s // 2 + 6, s // 2 + 4, s // 2 - 4)
    # head (horizontal bar at top-right)
    pen.setWidth(4)
    p.setPen(pen)
    p.drawLine(s // 2 + 1, s // 2 - 7, s // 2 + 10, s // 2 - 7)


def _draw_fix_icon(p: QPainter, s: int) -> None:
    """Wrench shape — Fix."""
    pen = QPen(QColor("#cc7700"), 2.6)
    p.setPen(pen)
    cx, cy = s // 2, s // 2
    # shaft
    p.drawLine(cx - 7, cy + 7, cx + 3, cy - 3)
    # wrench jaw (V shape at top-right)
    p.drawLine(cx + 3, cy - 3, cx + 8, cy - 8)
    p.drawLine(cx + 3, cy - 3, cx + 8, cy + 2)


def _draw_run_icon(p: QPainter, s: int) -> None:
    """Play triangle — Run (drawn as 3 filled lines)."""
    color = QColor("#2d7a2d")
    p.setPen(QPen(color, 1))
    p.setBrush(color)
    # Fill a triangle by drawing horizontal lines from left edge to right tip
    x0, x1 = int(s * 0.3), int(s * 0.78)
    y_top, y_bot = int(s * 0.2), int(s * 0.8)
    y_mid = (y_top + y_bot) // 2
    for y in range(y_top, y_bot + 1):
        # compute x-right for this scanline (triangle edge)
        if y <= y_mid:
            t = (y - y_top) / max(y_mid - y_top, 1)
        else:
            t = (y_bot - y) / max(y_bot - y_mid, 1)
        xr = int(x0 + (x1 - x0) * t)
        p.drawLine(x0, y, xr, y)


def _draw_search_icon(p: QPainter, s: int) -> None:
    """Magnifying glass — Search."""
    pen = QPen(QColor("#444444"), 2.4)
    p.setPen(pen)
    p.setBrush(Qt.GlobalColor.transparent)
    # lens circle
    cx, cy, r = s * 0.42, s * 0.40, s * 0.24
    p.drawEllipse(int(cx - r), int(cy - r), int(r * 2), int(r * 2))
    # handle
    p.drawLine(int(cx + r * 0.7), int(cy + r * 0.7), int(s * 0.78), int(s * 0.78))


# ---------------------------------------------------------------------------
#  Syntax highlighters for source panels
# ---------------------------------------------------------------------------

_VB6_KEYWORDS = {
    "if", "then", "else", "elseif", "end", "sub", "function", "dim", "set",
    "new", "for", "next", "while", "wend", "do", "loop", "until", "select",
    "case", "exit", "on", "error", "resume", "goto", "call", "with", "me",
    "private", "public", "friend", "property", "get", "let", "type", "enum",
    "const", "as", "byval", "byref", "optional", "not", "and", "or", "xor",
    "mod", "true", "false", "nothing", "is", "like", "to", "each", "in",
    "stop", "debug", "print", "input", "open", "close",
}

_CS_KEYWORDS = {
    "using", "namespace", "class", "struct", "interface", "enum", "record",
    "public", "private", "protected", "internal", "static", "void", "int",
    "string", "bool", "var", "new", "return", "if", "else", "switch", "case",
    "break", "continue", "for", "foreach", "while", "do", "try", "catch",
    "finally", "throw", "async", "await", "Task", "partial", "get", "set",
    "true", "false", "null", "this", "base", "typeof", "default", "in",
    "out", "ref", "is", "as", "operator", "override", "virtual", "sealed",
    "readonly", "const",
}

# Maximum lines/chars before skipping highlighting (performance guard)
_MAX_HIGHLIGHT_LINES = 8000
_MAX_HIGHLIGHT_CHARS = 500_000


def _make_fmt(color: str, bold: bool = False, italic: bool = False) -> QTextCharFormat:
    fmt = QTextCharFormat()
    fmt.setForeground(QColor(color))
    if bold:
        fmt.setFontWeight(QFont.Weight.Bold)
    if italic:
        fmt.setFontItalic(True)
    return fmt


class VB6Highlighter(QSyntaxHighlighter):
    """Syntax highlighter for VB6 source code (light background)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._fmt_kw = _make_fmt("#0000ff", bold=True)
        self._fmt_str = _make_fmt("#800000")
        self._fmt_cmt = _make_fmt("#008000", italic=True)
        self._fmt_hdr = _make_fmt("#555555")
        self._re_str = re.compile(r'"([^"\n]|"")*"')
        self._re_word = re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*\b")

    def highlightBlock(self, text: str) -> None:
        block_num = self.currentBlock().blockNumber()
        doc = self.document()
        if doc and (doc.lineCount() > _MAX_HIGHLIGHT_LINES
                    or doc.characterCount() > _MAX_HIGHLIGHT_CHARS):
            return

        # Header line (first line)
        if block_num == 0:
            self.setFormat(0, len(text), self._fmt_hdr)
            return

        # Full-line comments
        stripped = text.lstrip()
        if stripped.startswith("'") or stripped.lower().startswith("rem "):
            self.setFormat(0, len(text), self._fmt_cmt)
            return

        # Strings
        for m in self._re_str.finditer(text):
            self.setFormat(m.start(), m.end() - m.start(), self._fmt_str)

        # Keywords
        for m in self._re_word.finditer(text):
            if m.group(0).lower() in _VB6_KEYWORDS:
                self.setFormat(m.start(), m.end() - m.start(), self._fmt_kw)


class BlazorHighlighter(QSyntaxHighlighter):
    """Syntax highlighter for Blazor/C# source code (dark background)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._fmt_kw = _make_fmt("#569cd6", bold=True)
        self._fmt_str = _make_fmt("#ce9178")
        self._fmt_cmt = _make_fmt("#6a9955", italic=True)
        self._fmt_razor = _make_fmt("#c586c0")
        self._fmt_num = _make_fmt("#b5cea8")
        self._fmt_htmltag = _make_fmt("#569cd6")
        self._fmt_attr = _make_fmt("#d7ba7d")
        self._fmt_hdr = _make_fmt("#9aa0a6")

        self._re_line_cmt = re.compile(r"//.*")
        self._re_str = re.compile(r'"([^"\n]|\\")*"')
        self._re_num = re.compile(r"\b\d+(?:\.\d+)?\b")
        self._re_razor = re.compile(r"@\w+")
        self._re_htmltag = re.compile(r"</?\s*([A-Za-z0-9_:\-]+)")
        self._re_attr = re.compile(r"\b([A-Za-z_:][A-Za-z0-9_:\-]*)(?==)")
        self._re_word = re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*\b")

    def highlightBlock(self, text: str) -> None:
        block_num = self.currentBlock().blockNumber()
        doc = self.document()
        if doc and (doc.lineCount() > _MAX_HIGHLIGHT_LINES
                    or doc.characterCount() > _MAX_HIGHLIGHT_CHARS):
            return

        # Header line (first line)
        if block_num == 0:
            self.setFormat(0, len(text), self._fmt_hdr)
            return

        # Line comments (// ...)
        for m in self._re_line_cmt.finditer(text):
            self.setFormat(m.start(), m.end() - m.start(), self._fmt_cmt)

        # Strings
        for m in self._re_str.finditer(text):
            self.setFormat(m.start(), m.end() - m.start(), self._fmt_str)

        # Numbers
        for m in self._re_num.finditer(text):
            self.setFormat(m.start(), m.end() - m.start(), self._fmt_num)

        # Razor directives (@page, @inject, etc.)
        for m in self._re_razor.finditer(text):
            self.setFormat(m.start(), m.end() - m.start(), self._fmt_razor)

        # HTML tag names
        for m in self._re_htmltag.finditer(text):
            self.setFormat(m.start(1), m.end(1) - m.start(1), self._fmt_htmltag)

        # HTML attribute names (word=)
        for m in self._re_attr.finditer(text):
            self.setFormat(m.start(1), m.end(1) - m.start(1), self._fmt_attr)

        # C# keywords
        for m in self._re_word.finditer(text):
            if m.group(0) in _CS_KEYWORDS:
                self.setFormat(m.start(), m.end() - m.start(), self._fmt_kw)


# ---------------------------------------------------------------------------
#  Convert Settings Dialog
# ---------------------------------------------------------------------------

class ConvertDialog(QDialog):
    """Modal popup shown before Convert / Convert Selected runs."""

    def __init__(self, parent: "QtMigratorWindow", mode: str, defaults: dict) -> None:
        super().__init__(parent)
        title = "Convert Selected Files" if mode == "selected" else "Convert File"
        self.setWindowTitle(title)
        self.resize(650, 420)

        root = QVBoxLayout(self)
        root.setSpacing(10)

        form = QFormLayout()
        form.setSpacing(8)

        # Model dropdown
        self.model_combo = QComboBox()
        self.model_combo.setEditable(True)
        self.model_combo.addItems(["gpt-5.2-codex", "claude-opus-4-6", "deepseek-r1:8b", "qwen3:8b", "deepseek-r1:32b", "qwen2.5-coder:32b", "qwen3:32b"])
        req_model = (defaults.get("model") or "gpt-5.2-codex").strip()
        idx = self.model_combo.findText(req_model)
        if idx >= 0:
            self.model_combo.setCurrentIndex(idx)
        else:
            self.model_combo.setCurrentText(req_model)
        self.model_combo.setMinimumWidth(220)
        form.addRow("Model:", self.model_combo)

        # Temperature
        self.temp_input = QLineEdit((defaults.get("temperature") or "0.2").strip())
        self.temp_input.setMaximumWidth(100)
        form.addRow("Temperature:", self.temp_input)

        # Context directory + browse
        ctx_row = QHBoxLayout()
        self.context_input = QLineEdit(defaults.get("context_dir") or "")
        ctx_row.addWidget(self.context_input, stretch=1)
        browse_btn = QPushButton("Browse")
        browse_btn.setMaximumWidth(80)
        browse_btn.clicked.connect(self._browse_context)
        ctx_row.addWidget(browse_btn)
        ctx_widget = QWidget()
        ctx_widget.setLayout(ctx_row)
        form.addRow("Context:", ctx_widget)

        # Instructions
        self.instructions_edit = QPlainTextEdit()
        self.instructions_edit.setPlainText(defaults.get("instructions") or "")
        self.instructions_edit.setMinimumHeight(120)
        self.instructions_edit.setStyleSheet("background: #f8f8f8; color: #111111;")
        form.addRow("Instructions:", self.instructions_edit)

        root.addLayout(form)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        convert_btn = QPushButton("Convert")
        convert_btn.setDefault(True)
        convert_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(convert_btn)
        btn_row.addWidget(cancel_btn)
        root.addLayout(btn_row)

    def _browse_context(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self, "Select Context Directory", self.context_input.text().strip()
        )
        if selected:
            self.context_input.setText(selected)

    def selected_model(self) -> str:
        return self.model_combo.currentText().strip() or "gpt-5.2-codex"

    def selected_temperature(self) -> float:
        try:
            return float(self.temp_input.text().strip() or "0.2")
        except ValueError:
            return 0.2

    def context_dir(self) -> str:
        return self.context_input.text().strip()

    def instructions(self) -> str:
        return self.instructions_edit.toPlainText().strip()


class SearchDialog(QDialog):
    def __init__(self, parent: "QtMigratorWindow") -> None:
        super().__init__(parent)
        self.parent_win = parent
        self.setWindowTitle("Search")
        self.resize(980, 560)

        root = QVBoxLayout(self)

        row = QHBoxLayout()
        row.addWidget(QLabel("Find"))
        self.pattern_input = QLineEdit()
        row.addWidget(self.pattern_input, stretch=1)
        self.source_rb = QRadioButton("VB6 Source")
        self.target_rb = QRadioButton(".NET Target")
        self.source_rb.setChecked(True)
        row.addWidget(self.source_rb)
        row.addWidget(self.target_rb)
        self.search_btn = QPushButton("Search")
        self.close_btn = QPushButton("Close")
        row.addWidget(self.search_btn)
        row.addWidget(self.close_btn)
        root.addLayout(row)

        self.results = QTableWidget()
        self.results.setColumnCount(4)
        self.results.setHorizontalHeaderLabels(["File", "Line", "Match", "Snippet"])
        self.results.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.results.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.results.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.results.verticalHeader().setVisible(False)
        self.results.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        root.addWidget(self.results, stretch=1)

        self.search_btn.clicked.connect(self.on_search)
        self.close_btn.clicked.connect(self.reject)
        self.results.itemDoubleClicked.connect(self.on_open_result)
        self.pattern_input.returnPressed.connect(self.on_search)

    def _match_lines(self, lines: List[str], idx1: int) -> str:
        start = max(1, idx1 - 2)
        end = min(len(lines), idx1 + 2)
        rows: List[str] = []
        for i in range(start, end + 1):
            marker = ">" if i == idx1 else " "
            rows.append(f"{marker}{i:5}: {lines[i - 1]}")
        return "\n".join(rows)

    def on_search(self) -> None:
        pattern = self.pattern_input.text().strip()
        if not pattern:
            self.parent_win.set_status("Enter a search pattern.")
            return

        is_target = self.target_rb.isChecked()
        if is_target:
            context_dir = Path(self.parent_win.app_config.get("context_dir", "")).expanduser()
            paths = backend.collect_context_files(context_dir) if context_dir.exists() else []
        else:
            paths = [c.path for c in self.parent_win.components]

        try:
            regex = re.compile(pattern, re.IGNORECASE)
            matcher = lambda text: bool(regex.search(text))
        except Exception:
            needle = pattern.lower()
            matcher = lambda text: needle in text.lower()

        self.results.setRowCount(0)
        total = 0
        for path in paths:
            try:
                text = path.read_text(encoding="utf-8", errors="ignore")
            except Exception:
                try:
                    text = backend.read_text(path)
                except Exception:
                    text = ""
            if not text:
                continue
            lines = text.splitlines()
            for idx1, line in enumerate(lines, start=1):
                if not matcher(line):
                    continue
                row = self.results.rowCount()
                self.results.insertRow(row)
                file_item = QTableWidgetItem(path.name)
                file_item.setData(ROLE_PATH, str(path))
                file_item.setData(ROLE_FLAG, bool(is_target))
                self.results.setItem(row, 0, file_item)
                self.results.setItem(row, 1, QTableWidgetItem(str(idx1)))
                self.results.setItem(row, 2, QTableWidgetItem(line.strip()[:180]))
                self.results.setItem(row, 3, QTableWidgetItem(self._match_lines(lines, idx1)))
                total += 1
                if total >= MAX_SEARCH_MATCHES:
                    break
            if total >= MAX_SEARCH_MATCHES:
                break

        self.results.resizeColumnsToContents()
        if self.results.columnWidth(0) < 180:
            self.results.setColumnWidth(0, 180)
        self.parent_win.set_status(f"Search complete. {total} matches.")

    def on_open_result(self, *_args) -> None:
        row = self.results.currentRow()
        if row < 0:
            return
        file_item = self.results.item(row, 0)
        line_item = self.results.item(row, 1)
        if file_item is None:
            return
        path_raw = file_item.data(ROLE_PATH)
        if not path_raw:
            return
        is_target = bool(file_item.data(ROLE_FLAG))
        line_no = 1
        if line_item is not None:
            try:
                line_no = int(line_item.text().strip())
            except Exception:
                line_no = 1
        self.parent_win.open_search_result(Path(str(path_raw)), is_target, line_no)
        self.accept()


def _parse_vbp_entries(vbp_path: Path) -> List[dict]:
    """Parse a .vbp file and return entries with resolved paths and existence status.

    Each entry: {"type": str, "name": str, "rel_path": str,
                 "abs_path": Path, "found": bool, "vbp": str}
    """
    entries: List[dict] = []
    vbp_dir = vbp_path.parent
    vbp_name = vbp_path.name
    try:
        text = vbp_path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return entries

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        comp_type: str | None = None
        rel_path_raw: str = ""
        name: str = ""

        if line.startswith("Form="):
            comp_type = "Form"
            rel_path_raw = line.split("=", 1)[1].strip().strip('"')
            name = Path(rel_path_raw.replace("\\", "/")).stem

        elif line.startswith("Class="):
            comp_type = "Class"
            body = line.split("=", 1)[1]
            parts = [p.strip() for p in body.split(";")]
            if len(parts) >= 2:
                name = parts[0].strip()
                rel_path_raw = parts[1].strip().strip('"')
            else:
                continue

        elif line.startswith("Module="):
            comp_type = "Module"
            body = line.split("=", 1)[1]
            parts = [p.strip() for p in body.split(";")]
            if len(parts) >= 2:
                name = parts[0].strip()
                rel_path_raw = parts[1].strip().strip('"')
            else:
                continue

        elif line.startswith("UserControl="):
            comp_type = "UserControl"
            body = line.split("=", 1)[1]
            parts = [p.strip() for p in body.split(";")]
            if len(parts) >= 2:
                name = parts[0].strip()
                rel_path_raw = parts[1].strip().strip('"')
            elif len(parts) == 1:
                rel_path_raw = parts[0].strip().strip('"')
                name = Path(rel_path_raw.replace("\\", "/")).stem
            else:
                continue
        else:
            continue

        rel_path = rel_path_raw.replace("\\", "/")
        abs_path = (vbp_dir / rel_path).resolve()
        entries.append({
            "type": comp_type,
            "name": name,
            "rel_path": rel_path_raw,
            "abs_path": abs_path,
            "found": abs_path.exists(),
            "vbp": vbp_name,
        })

    return entries


class MigrationWizardDialog(QDialog):
    """Wizard to scaffold a new .NET Blazor project from VB6 sources."""

    def __init__(self, parent: "QtMigratorWindow") -> None:
        super().__init__(parent)
        self.setWindowTitle("Migration Wizard")
        self.resize(900, 600)
        self._vbp_entries: List[dict] = []

        root = QVBoxLayout(self)
        root.setSpacing(10)

        form = QFormLayout()
        form.setSpacing(8)

        # Project Name
        self._project_name = QLineEdit("NewProject")
        self._project_name.setMinimumWidth(300)
        form.addRow("Project Name:", self._project_name)

        # VB6 Source Dir + Browse + Scan
        vb6_row = QHBoxLayout()
        self._vb6_source_dir = QLineEdit()
        vb6_row.addWidget(self._vb6_source_dir, stretch=1)
        vb6_browse = QPushButton("Browse")
        vb6_browse.setMaximumWidth(80)
        vb6_browse.clicked.connect(self._browse_vb6)
        vb6_row.addWidget(vb6_browse)
        scan_btn = QPushButton("Scan")
        scan_btn.setMaximumWidth(80)
        scan_btn.clicked.connect(self._on_scan)
        vb6_row.addWidget(scan_btn)
        vb6_widget = QWidget()
        vb6_widget.setLayout(vb6_row)
        form.addRow("VB6 Source Dir:", vb6_widget)

        # Target .NET Dir
        net_row = QHBoxLayout()
        self._target_net_dir = QLineEdit()
        net_row.addWidget(self._target_net_dir, stretch=1)
        net_browse = QPushButton("Browse")
        net_browse.setMaximumWidth(80)
        net_browse.clicked.connect(self._browse_target)
        net_row.addWidget(net_browse)
        net_widget = QWidget()
        net_widget.setLayout(net_row)
        hint = QLabel("(will be created if it does not exist)")
        hint.setStyleSheet("color: #666;")
        net_container = QVBoxLayout()
        net_container.setSpacing(2)
        net_container.addWidget(net_widget)
        net_container.addWidget(hint)
        net_wrapper = QWidget()
        net_wrapper.setLayout(net_container)
        form.addRow("Target .NET Dir:", net_wrapper)

        # Static Assets Dir (optional)
        assets_row = QHBoxLayout()
        self._static_assets_dir = QLineEdit()
        assets_row.addWidget(self._static_assets_dir, stretch=1)
        assets_browse = QPushButton("Browse")
        assets_browse.setMaximumWidth(80)
        assets_browse.clicked.connect(self._browse_assets)
        assets_row.addWidget(assets_browse)
        assets_widget = QWidget()
        assets_widget.setLayout(assets_row)
        assets_hint = QLabel("(optional — files copied to wwwroot/ in the new project)")
        assets_hint.setStyleSheet("color: #666;")
        assets_container = QVBoxLayout()
        assets_container.setSpacing(2)
        assets_container.addWidget(assets_widget)
        assets_container.addWidget(assets_hint)
        assets_wrapper = QWidget()
        assets_wrapper.setLayout(assets_container)
        form.addRow("Static Assets:", assets_wrapper)

        root.addLayout(form)

        # VB6 Project Files table (hidden until scan)
        self._file_table = QTableWidget()
        self._file_table.setColumnCount(5)
        self._file_table.setHorizontalHeaderLabels(["Type", "Name", "Path", "Status", ""])
        self._file_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._file_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._file_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._file_table.verticalHeader().setVisible(False)
        self._file_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self._file_table.setVisible(False)
        root.addWidget(self._file_table, stretch=1)

        # Summary label
        self._summary_label = QLabel("")
        self._summary_label.setVisible(False)
        root.addWidget(self._summary_label)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        create_btn = QPushButton("Create Project")
        create_btn.setDefault(True)
        create_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(create_btn)
        btn_row.addWidget(cancel_btn)
        root.addLayout(btn_row)

    # --- Browse handlers ---

    def _browse_vb6(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self, "Select VB6 Source Directory", self._vb6_source_dir.text().strip()
        )
        if selected:
            self._vb6_source_dir.setText(selected)

    def _browse_target(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self, "Select Target .NET Directory", self._target_net_dir.text().strip()
        )
        if selected:
            self._target_net_dir.setText(selected)

    def _browse_assets(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self, "Select Static Assets Directory", self._static_assets_dir.text().strip()
        )
        if selected:
            self._static_assets_dir.setText(selected)

    # --- VBP scanning ---

    def _on_scan(self) -> None:
        """Scan the VB6 source dir for .vbp files and list all referenced components."""
        source_dir = self._vb6_source_dir.text().strip()
        if not source_dir or not Path(source_dir).is_dir():
            QMessageBox.warning(self, "Scan", "Please select a valid VB6 source directory first.")
            return

        root_path = Path(source_dir)
        vbp_files = sorted(root_path.rglob("*.vbp"))
        if not vbp_files:
            QMessageBox.information(self, "Scan", "No .vbp project files found in the selected directory.")
            return

        self._vbp_entries.clear()
        for vbp in vbp_files:
            self._vbp_entries.extend(_parse_vbp_entries(vbp))

        self._refresh_file_table()

    def _refresh_file_table(self) -> None:
        """Rebuild the file table from self._vbp_entries."""
        self._file_table.setRowCount(0)
        for idx, entry in enumerate(self._vbp_entries):
            row = self._file_table.rowCount()
            self._file_table.insertRow(row)
            self._file_table.setItem(row, 0, QTableWidgetItem(entry["type"]))
            self._file_table.setItem(row, 1, QTableWidgetItem(entry["name"]))
            self._file_table.setItem(row, 2, QTableWidgetItem(entry["rel_path"]))

            status_item = QTableWidgetItem("Found" if entry["found"] else "Missing")
            if entry["found"]:
                status_item.setForeground(QColor("#228B22"))
            else:
                status_item.setForeground(QColor("#CC0000"))
            self._file_table.setItem(row, 3, status_item)

            if not entry["found"]:
                locate_btn = QPushButton("Locate...")
                locate_btn.setMaximumWidth(80)
                locate_btn.clicked.connect(lambda checked=False, r=idx: self._on_locate(r))
                self._file_table.setCellWidget(row, 4, locate_btn)

        self._file_table.resizeColumnsToContents()
        self._file_table.setVisible(True)
        self._update_summary()

    def _update_summary(self) -> None:
        total = len(self._vbp_entries)
        found = sum(1 for e in self._vbp_entries if e["found"])
        missing = total - found
        if missing > 0:
            self._summary_label.setText(
                f"<b>{found}</b> of <b>{total}</b> files found. "
                f"<span style='color:#CC0000'><b>{missing}</b> file(s) missing.</span>"
            )
        else:
            self._summary_label.setText(
                f"<b>{found}</b> of <b>{total}</b> files found. All files present."
            )
        self._summary_label.setVisible(True)

    def _on_locate(self, entry_idx: int) -> None:
        """Let the user browse for a missing file and copy it to the expected location."""
        entry = self._vbp_entries[entry_idx]
        ext = Path(entry["rel_path"].replace("\\", "/")).suffix
        ext_filter = f"VB6 files (*{ext})" if ext else "All files (*)"

        selected, _ = QFileDialog.getOpenFileName(
            self,
            f"Locate {entry['name']} ({entry['rel_path']})",
            "",
            ext_filter,
        )
        if not selected:
            return

        # Copy the located file to the expected absolute path
        target = entry["abs_path"]
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(selected, target)
        except Exception as exc:
            QMessageBox.warning(self, "Locate", f"Failed to copy file:\n{exc}")
            return

        entry["found"] = True
        self._refresh_file_table()

    # --- Accessors ---

    def project_name(self) -> str:
        return self._project_name.text().strip() or "NewProject"

    def vb6_source_dir(self) -> str:
        return self._vb6_source_dir.text().strip()

    def target_net_dir(self) -> str:
        return self._target_net_dir.text().strip()

    def static_assets_dir(self) -> str:
        return self._static_assets_dir.text().strip()


# ---------------------------------------------------------------------------
# Template root for scaffolding .NET Blazor projects
# ---------------------------------------------------------------------------
_HOMEFRONT_POC_ROOT = Path("/Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontPOC")

# Directories to copy in bulk (preserving structure)
_TEMPLATE_DIRS = [
    "Components/Hf",
    "Components/Layout",
    "Data",
    "Infrastructure/Logging",
    "wwwroot/lib",
    "wwwroot/js",
]

# Individual files to copy
_TEMPLATE_FILES = [
    "Components/App.razor",
    "Components/_Imports.razor",
    "Components/Routes.razor",
    "Program.cs",
    "appsettings.json",
    "appsettings.Development.json",
    "wwwroot/app.css",
    "wwwroot/favicon.png",
]

_CSPROJ_TEMPLATE = """\
<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <RootNamespace>HomeFront</RootNamespace>
    <AssemblyName>{assembly_name}</AssemblyName>
    <BlazorDisableThrowNavigationException>true</BlazorDisableThrowNavigationException>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Data.SqlClient" Version="5.2.0" />
    <PackageReference Include="Serilog.AspNetCore" Version="9.0.0" />
    <PackageReference Include="Serilog.Settings.Configuration" Version="9.0.0" />
    <PackageReference Include="Serilog.Sinks.Console" Version="6.0.0" />
    <PackageReference Include="Serilog.Sinks.File" Version="7.0.0" />
  </ItemGroup>

</Project>
"""

_HOME_RAZOR_TEMPLATE = """\
@page "/"

<h3>Welcome to {project_name}</h3>
<p>Migration project created successfully.</p>
"""


def _scaffold_dotnet_project(
    target_dir: Path,
    project_name: str,
    static_assets_dir: str | None,
    progress_fn=None,
) -> None:
    """Create a .NET Blazor project at *target_dir* from the HomeFrontPOC template."""

    def _progress(msg: str) -> None:
        if progress_fn:
            progress_fn(msg)

    target_dir = Path(target_dir)
    _progress("Creating .NET project structure...")
    target_dir.mkdir(parents=True, exist_ok=True)

    # Copy template directories
    for rel in _TEMPLATE_DIRS:
        src = _HOMEFRONT_POC_ROOT / rel
        dst = target_dir / rel
        if src.is_dir():
            _progress(f"Copying {rel}...")
            shutil.copytree(src, dst, dirs_exist_ok=True)

    # Copy individual template files
    _progress("Copying project files...")
    for rel in _TEMPLATE_FILES:
        src = _HOMEFRONT_POC_ROOT / rel
        dst = target_dir / rel
        if src.is_file():
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)

    # Generate .csproj
    csproj_path = target_dir / f"{project_name}.csproj"
    csproj_path.write_text(
        _CSPROJ_TEMPLATE.format(assembly_name=project_name),
        encoding="utf-8",
    )

    # Create welcome page
    home_dir = target_dir / "Components" / "Pages"
    home_dir.mkdir(parents=True, exist_ok=True)
    (home_dir / "Home.razor").write_text(
        _HOME_RAZOR_TEMPLATE.format(project_name=project_name),
        encoding="utf-8",
    )

    # Create Migrated output directory
    (target_dir / "Components" / "Pages" / "Migrated").mkdir(parents=True, exist_ok=True)

    # Copy static assets if provided
    if static_assets_dir:
        assets_path = Path(static_assets_dir)
        if assets_path.is_dir():
            _progress("Copying static assets to wwwroot/...")
            shutil.copytree(assets_path, target_dir / "wwwroot", dirs_exist_ok=True)

    _progress(f".NET project '{project_name}' scaffolded.")


class QtMigratorWindow(QMainWindow):
    # Signal to safely invoke callables on the main thread from worker threads.
    _invoke_on_main = Signal(object)

    def __init__(self, runtime: dict, ui_defaults: Dict[str, str], app_config: dict) -> None:
        super().__init__()
        self._invoke_on_main.connect(self._run_on_main)
        self.ui_defaults = ui_defaults
        self.app_config = dict(app_config)

        self.prefer_hfest = str(ui_defaults.get("prefer_hfest", "true")).strip().lower() == "true"
        self.active_project_dir = Path(ui_defaults.get("project_dir", "")).expanduser()
        self.vb6_root_dir = Path(ui_defaults.get("root", "")).expanduser()
        self.output_dir = Path(ui_defaults.get("output_dir", "")).expanduser()
        self.project_path = Path(ui_defaults.get("project", "")).expanduser()
        self.vb6_image_dir = Path(
            ui_defaults.get(
                "vb6_image_dir",
                "/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/wwwroot/images/vb6",
            )
        ).expanduser()
        self.dotnet_image_dir = Path(
            ui_defaults.get(
                "dotnet_image_dir",
                "/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/wwwroot/images/dotnet",
            )
        ).expanduser()
        self.project_root = self.project_path.parent if self.project_path.exists() else self.vb6_root_dir

        self.runtime: dict = {}
        self.components: List = []
        self.edges: Dict[Tuple[str, str], Dict] = {}
        self.roots: List[str] = []
        self.duplicates: set[str] = set()
        self.migration_order: List[str] = []
        self.orphan_set: set[str] = set()
        self.reachable_ids: set[str] = set()
        self.ordered_comp_ids: List[str] = []
        self.id_to_component: Dict[str, object] = {}
        self.comp_id_by_name: Dict[str, str] = {}
        self.comp_id_by_path: Dict[str, str] = {}
        self.adj: Dict[str, List[Tuple[str, bool]]] = {}
        self.file_size_by_id: Dict[str, int] = {}
        self.line_count_by_id: Dict[str, int] = {}

        self.converted_map: Dict[str, Path] = {}
        self.current_comp_id: str | None = None
        self.current_output_path: Path | None = None
        self.last_compile_errors = ""
        self.context_cache = {"dir": "", "signature": "", "blob": ""}
        self.razor_index_cache: Dict[str, Dict[str, List[Path]]] = {}
        self.busy = False
        self.vb6_tree_item_by_comp_id: Dict[str, QTreeWidgetItem] = {}

        self.run_process: subprocess.Popen | None = None
        self.run_log_path: Path | None = None
        self.run_tail_offset = 0
        self.run_tail_buffer = ""
        self.run_log_header_lines: List[str] = []
        self.run_url_announced = ""
        self.run_exit_announced = False
        self.run_restart_count = 0
        self.run_max_restarts = 2
        self.run_browser = backend.BlazorRunBrowser()
        self.run_tail_timer = QTimer(self)
        self.run_tail_timer.setInterval(900)
        self.run_tail_timer.timeout.connect(self._tail_run_log)

        self._apply_runtime(runtime)
        self._init_converted_map()
        self._build_ui()
        self._build_menu()
        self._populate_project_tree()
        self._populate_vb6_tree()
        self._populate_vb6_grid()
        self.set_status(f"Ready - UI: Qt ({QT_BINDING})")

    def _apply_runtime(self, runtime: dict) -> None:
        self.runtime = runtime
        self.components = list(runtime.get("components", []))
        self.edges = dict(runtime.get("edges", {}))
        self.roots = list(runtime.get("roots", []))
        self.duplicates = set(runtime.get("duplicates", set()))
        self.migration_order = list(runtime.get("migration_order", []))
        self.orphan_set = set(runtime.get("orphan_set", set()))

        layout = runtime.get("layout", {}) or {}
        project_dir_raw = layout.get("project_dir")
        if project_dir_raw:
            self.active_project_dir = Path(project_dir_raw).expanduser()
        source_origin = (layout.get("source_origin") or "").strip()
        if source_origin:
            self.vb6_root_dir = Path(source_origin).expanduser()
        elif layout.get("source_dir"):
            self.vb6_root_dir = Path(layout.get("source_dir")).expanduser()
        # Prefer target_origin (the actual .NET output dir) over target_dir (project-local cache)
        target_origin = (layout.get("target_origin") or "").strip()
        if target_origin:
            self.output_dir = Path(target_origin).expanduser()
        elif layout.get("target_dir"):
            self.output_dir = Path(layout.get("target_dir")).expanduser()
        output_project = (layout.get("output_project") or "").strip()
        if output_project:
            self.project_path = Path(output_project).expanduser()
        self.project_root = self.project_path.parent if self.project_path.exists() else self.vb6_root_dir

        self.id_to_component = {f"{c.project}::{c.name}": c for c in self.components}
        self.comp_id_by_name = {}
        for c in self.components:
            self.comp_id_by_name.setdefault(c.name.lower(), f"{c.project}::{c.name}")
        self.comp_id_by_path = {str(c.path.resolve()): f"{c.project}::{c.name}" for c in self.components}

        self.adj = {}
        for (caller, callee), meta in self.edges.items():
            self.adj.setdefault(caller, []).append((callee, bool(meta.get("ambiguous", False))))
        for caller in list(self.adj.keys()):
            self.adj[caller] = sorted(set(self.adj[caller]), key=lambda t: t[0].lower())

        self.reachable_ids = self._compute_reachable_ids()
        self.ordered_comp_ids = self._build_ordered_comp_ids()

        self.file_size_by_id = {}
        self.line_count_by_id = {}
        for comp_id, comp in self.id_to_component.items():
            try:
                self.file_size_by_id[comp_id] = comp.path.stat().st_size
                self.line_count_by_id[comp_id] = backend.count_lines_file(comp.path)
            except Exception:
                self.file_size_by_id[comp_id] = 0
                self.line_count_by_id[comp_id] = 0

        if self.current_comp_id and self.current_comp_id not in self.id_to_component:
            self.current_comp_id = None
            self.current_output_path = None

    def _compute_reachable_ids(self) -> set[str]:
        seed: List[str] = []
        seen_seed: set[str] = set()
        for comp_id in self.roots:
            if comp_id in seen_seed:
                continue
            seen_seed.add(comp_id)
            seed.append(comp_id)
        if "HFSystem::FLogin" in self.id_to_component and "HFSystem::FLogin" not in seen_seed:
            seed.append("HFSystem::FLogin")
        reachable: set[str] = set()
        stack = list(seed)
        while stack:
            node = stack.pop()
            if node in reachable:
                continue
            reachable.add(node)
            for child_id, _ambiguous in self.adj.get(node, []):
                if child_id not in reachable:
                    stack.append(child_id)
        return reachable

    def _build_ordered_comp_ids(self) -> List[str]:
        ordered: List[str] = []
        seen: set[str] = set()
        for comp_id in self.migration_order:
            if comp_id in self.id_to_component and comp_id not in seen:
                seen.add(comp_id)
                ordered.append(comp_id)
        for comp_id in sorted(self.id_to_component.keys(), key=lambda v: v.lower()):
            if comp_id not in seen:
                ordered.append(comp_id)
        return ordered

    def _projects_root_path(self) -> Path:
        raw = (self.ui_defaults.get("projects_root") or "").strip()
        if raw:
            return Path(raw).expanduser().resolve()
        return (Path(__file__).resolve().parent / "projects").resolve()

    def _sync_config_dict(self) -> None:
        self.app_config["projects_root"] = str(self._projects_root_path())
        self.app_config["project_dir"] = str(self.active_project_dir.expanduser().resolve())
        self.app_config["project_name"] = self.active_project_dir.name or "PrescionBuilder"
        self.app_config["root"] = str(self.vb6_root_dir)
        self.app_config["prefer_hfest"] = bool(self.prefer_hfest)
        self.app_config["ui"] = True
        self.app_config["output_project"] = str(self.project_path)
        self.app_config["output_dir"] = str(self.output_dir)
        # model, temperature, context_dir, instructions are kept in app_config
        # and updated by the ConvertDialog popup — no widget reads needed.
        self.app_config["vb6_image_dir"] = str(self.vb6_image_dir)
        self.app_config["dotnet_image_dir"] = str(self.dotnet_image_dir)

    def _persist_config(self) -> None:
        self._sync_config_dict()
        _save_app_config(self.app_config)

    def _invalidate_razor_indexes(self) -> None:
        self.razor_index_cache.clear()

    def _resolve_project_root(self) -> Path | None:
        project = self.project_path
        try:
            if project.suffix.lower() == ".csproj":
                return project.expanduser().resolve().parent
            if project.is_file():
                return project.expanduser().resolve().parent
            if project.is_dir():
                return project.expanduser().resolve()
        except Exception:
            return None
        return None

    def _build_razor_index(self, root_dir: Path, skip_build_dirs: bool) -> Dict[str, List[Path]]:
        root_dir = root_dir.expanduser().resolve()
        cache_key = f"{root_dir}|{1 if skip_build_dirs else 0}"
        cached = self.razor_index_cache.get(cache_key)
        if cached is not None:
            return cached

        index: Dict[str, List[Path]] = {}
        if root_dir.exists():
            skip_names = {"bin", "obj", ".git", ".vs", "node_modules"} if skip_build_dirs else set()
            for path in root_dir.rglob("*.razor"):
                if skip_names and any(part in skip_names for part in path.parts):
                    continue
                index.setdefault(path.stem.lower(), []).append(path)
            for stem in index:
                index[stem].sort(
                    key=lambda p: (
                        0 if "/components/pages/" in p.as_posix().lower() else 1,
                        len(p.as_posix()),
                    )
                )
        self.razor_index_cache[cache_key] = index
        return index

    def _find_existing_razor_for_component(self, comp) -> Path | None:
        comp_name = comp.name.lower()
        candidates: List[Path] = []

        if self.output_dir.exists():
            out_index = self._build_razor_index(self.output_dir, skip_build_dirs=False)
            candidates.extend(out_index.get(comp_name, []))

        project_root = self._resolve_project_root()
        if project_root and project_root.exists():
            proj_index = self._build_razor_index(project_root, skip_build_dirs=True)
            candidates.extend(proj_index.get(comp_name, []))

        deduped: Dict[str, Path] = {}
        for path in candidates:
            try:
                deduped[str(path.resolve())] = path.resolve()
            except Exception:
                deduped[str(path)] = path
        ordered = sorted(
            deduped.values(),
            key=lambda p: (
                0 if "/components/pages/" in p.as_posix().lower() else 1,
                len(p.as_posix()),
            ),
        )
        return ordered[0] if ordered else None

    def _init_converted_map(self) -> None:
        self.converted_map = {}
        self._invalidate_razor_indexes()
        for comp_id, comp in self.id_to_component.items():
            candidate = self._find_existing_razor_for_component(comp)
            if candidate and candidate.exists():
                self.converted_map[comp_id] = candidate

    def _build_ui(self) -> None:
        self.setWindowTitle(f"VB6 Migrator IDE - UI: Qt ({QT_BINDING})")
        self.resize(1920, 1080)

        mono = QFont("Menlo")
        if not mono.exactMatch():
            mono = QFont("Consolas")
        mono.setPointSize(10)

        root = QWidget(self)
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(6, 6, 6, 6)
        root_layout.setSpacing(6)

        status_row = QHBoxLayout()
        self.status_label = QLabel("Ready")
        self.framework_label = QLabel(f"UI: Qt ({QT_BINDING})")
        status_row.addWidget(self.status_label, stretch=1)
        status_row.addWidget(self.framework_label)
        root_layout.addLayout(status_row)

        # Icon-only toolbar with tooltips
        self.toolbar = QToolBar("Main")
        self.toolbar.setMovable(False)
        self.toolbar.setIconSize(QSize(32, 32))
        self.toolbar.setStyleSheet(
            "QToolBar { background: #f4f4f4; border-bottom: 1px solid #d0d0d0; spacing: 2px; padding: 2px; }"
            "QToolBar QToolButton { border: 1px solid transparent; border-radius: 4px; padding: 4px; }"
            "QToolBar QToolButton:hover { background: #e0e0e0; border: 1px solid #c0c0c0; }"
            "QToolBar QToolButton:pressed { background: #d0d0d0; }"
        )

        self.convert_action = self.toolbar.addAction(
            _make_icon(_draw_convert_icon), "Convert")
        self.convert_action.setToolTip("Convert — Convert selected VB6 file to Blazor")
        self.convert_action.triggered.connect(self.on_convert)

        self.convert_selected_action = self.toolbar.addAction(
            _make_icon(_draw_convert_selected_icon), "Convert Selected")
        self.convert_selected_action.setToolTip("Convert Selected — Convert multiple selected VB6 files")
        self.convert_selected_action.triggered.connect(self.on_convert_selected)

        self.toolbar.addSeparator()

        self.compile_action = self.toolbar.addAction(
            _make_icon(_draw_compile_icon), "Compile")
        self.compile_action.setToolTip("Compile — Build the .NET Blazor project")
        self.compile_action.triggered.connect(self.on_compile)

        self.fix_action = self.toolbar.addAction(
            _make_icon(_draw_fix_icon), "Fix")
        self.fix_action.setToolTip("Fix — Fix compile errors using AI")
        self.fix_action.triggered.connect(self.on_fix)

        self.run_action = self.toolbar.addAction(
            _make_icon(_draw_run_icon), "Run")
        self.run_action.setToolTip("Run — Launch the Blazor application")
        self.run_action.triggered.connect(self.on_run)

        self.toolbar.addSeparator()

        self.search_action = self.toolbar.addAction(
            _make_icon(_draw_search_icon), "Search")
        self.search_action.setToolTip("Search — Search across VB6 and Blazor files")
        self.search_action.triggered.connect(self.on_search)

        root_layout.addWidget(self.toolbar)

        main_split = QSplitter(Qt.Orientation.Horizontal)
        root_layout.addWidget(main_split, stretch=1)
        self.setCentralWidget(root)

        # --- Left tab: VB6 Files + Blazor Project ---

        # VB6 file list/tree panel
        vb6_panel = QWidget()
        vb6_layout = QVBoxLayout(vb6_panel)
        vb6_layout.setContentsMargins(4, 4, 4, 4)
        vb6_layout.setSpacing(4)
        vb6_toggle_row = QHBoxLayout()
        self.tree_btn = QPushButton("Tree")
        self.grid_btn = QPushButton("Grid")
        self.tree_btn.setCheckable(True)
        self.grid_btn.setCheckable(True)
        self.tree_btn.setChecked(True)
        mode_group = QButtonGroup(self)
        mode_group.setExclusive(True)
        mode_group.addButton(self.tree_btn, 0)
        mode_group.addButton(self.grid_btn, 1)
        vb6_toggle_row.addWidget(self.tree_btn)
        vb6_toggle_row.addWidget(self.grid_btn)
        vb6_toggle_row.addStretch(1)
        vb6_layout.addLayout(vb6_toggle_row)

        self.vb6_stack = QStackedWidget()
        vb6_layout.addWidget(self.vb6_stack, stretch=1)

        self.vb6_tree = QTreeWidget()
        self.vb6_tree.setHeaderHidden(True)
        self.vb6_tree.itemExpanded.connect(self._on_vb6_tree_item_expanded)
        self.vb6_tree.itemSelectionChanged.connect(self._on_vb6_tree_selection)
        self.vb6_stack.addWidget(self.vb6_tree)

        self.vb6_grid = QTableWidget()
        self.vb6_grid.setColumnCount(7)
        self.vb6_grid.setHorizontalHeaderLabels(
            ["Convertred", "#", "Component", "Type", "Project", "File Size", "Line Count"]
        )
        self.vb6_grid.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.vb6_grid.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.vb6_grid.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.vb6_grid.verticalHeader().setVisible(False)
        self.vb6_grid.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.vb6_grid.setSortingEnabled(True)
        self.vb6_grid.itemSelectionChanged.connect(self._on_vb6_grid_selection)
        self.vb6_stack.addWidget(self.vb6_grid)
        self.tree_btn.clicked.connect(lambda: self.vb6_stack.setCurrentIndex(0))
        self.grid_btn.clicked.connect(lambda: self.vb6_stack.setCurrentIndex(1))

        # Blazor project tree panel
        project_panel = QWidget()
        project_layout = QVBoxLayout(project_panel)
        project_layout.setContentsMargins(4, 4, 4, 4)
        project_layout.setSpacing(4)
        project_header = QHBoxLayout()
        project_refresh_btn = QPushButton("Refresh")
        project_refresh_btn.clicked.connect(self._populate_project_tree)
        project_header.addWidget(project_refresh_btn)
        project_header.addStretch(1)
        project_layout.addLayout(project_header)

        self.project_tree = QTreeWidget()
        self.project_tree.setHeaderHidden(True)
        self.project_tree.itemExpanded.connect(self._on_project_item_expanded)
        self.project_tree.itemClicked.connect(self._on_project_item_clicked)
        project_layout.addWidget(self.project_tree, stretch=1)

        # Left tab control — file navigation
        self.left_tabs = QTabWidget()
        self.left_tabs.addTab(vb6_panel, "VB6 Files")
        self.left_tabs.addTab(project_panel, "Blazor Project")

        # --- Right tab: VB6 Source + .NET Source ---

        # VB6 source panel
        vb_source_panel = QWidget()
        vb_source_layout = QVBoxLayout(vb_source_panel)
        vb_source_layout.setContentsMargins(0, 0, 0, 0)
        vb_source_layout.setSpacing(0)
        self.vb_source = QPlainTextEdit()
        self.vb_source.setReadOnly(True)
        self.vb_source.setFont(mono)
        self.vb_source.setStyleSheet("background: #ffffff; color: #000000;")
        self._vb_highlighter = VB6Highlighter(self.vb_source.document())
        vb_source_layout.addWidget(self.vb_source, stretch=1)

        # .NET source + log output panel
        net_panel = QWidget()
        net_layout = QVBoxLayout(net_panel)
        net_layout.setContentsMargins(0, 0, 0, 0)
        net_layout.setSpacing(0)
        net_split = QSplitter(Qt.Orientation.Vertical)
        net_layout.addWidget(net_split, stretch=1)
        self.net_source = QPlainTextEdit()
        self.net_source.setReadOnly(True)
        self.net_source.setFont(mono)
        self.net_source.setStyleSheet("background: #1e1e1e; color: #d4d4d4;")
        self._blazor_highlighter = BlazorHighlighter(self.net_source.document())
        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setFont(mono)
        self.log_output.setStyleSheet("background: #111111; color: #e0e0e0;")
        net_split.addWidget(self.net_source)
        net_split.addWidget(self.log_output)
        net_split.setSizes([700, 260])

        # Right tab control — source editors
        self.right_tabs = QTabWidget()
        self.right_tabs.addTab(vb_source_panel, "VB6 Source")
        self.right_tabs.addTab(net_panel, ".NET Source")

        # Assemble main splitter: left tabs (15%) | right tabs (85%)
        main_split.addWidget(self.left_tabs)
        main_split.addWidget(self.right_tabs)
        main_split.setSizes([300, 1700])

        # Button signals are connected via toolbar QAction.triggered above.

    def _add_menu_action(self, menu, label: str, slot) -> QAction:
        action = QAction(label, self)
        action.triggered.connect(slot)
        menu.addAction(action)
        return action

    def _build_menu(self) -> None:
        menubar = self.menuBar()

        file_menu = menubar.addMenu("File")
        self._add_menu_action(file_menu, "New Project...", self.on_new_project)
        self._add_menu_action(file_menu, "Migration Wizard...", self.on_migration_wizard)
        self._add_menu_action(file_menu, "Load Existing Project...", self.on_load_project)
        self._add_menu_action(file_menu, "Close Project", self.on_close_project)
        file_menu.addSeparator()
        self._add_menu_action(file_menu, "Exit", self.close)

        project_menu = menubar.addMenu("Project")
        self._add_menu_action(project_menu, "Refresh Stats", self.on_refresh_stats)
        self._add_menu_action(project_menu, "Rebuild Cache", self.on_rebuild_cache)
        self._add_menu_action(project_menu, "Convert Selected", self.on_convert_selected)

        build_menu = menubar.addMenu("Build")
        self._add_menu_action(build_menu, "Compile", self.on_compile)
        self._add_menu_action(build_menu, "Fix", self.on_fix)
        self._add_menu_action(build_menu, "Run", self.on_run)
        self._add_menu_action(build_menu, "Validate In Browser", self.on_validate_in_browser)
        self._add_menu_action(build_menu, "Align Screens", self.on_align_screens)

        tools_menu = menubar.addMenu("Tools")
        self._add_menu_action(tools_menu, "Search", self.on_search)
        self._add_menu_action(tools_menu, "Tag Selected", self.on_tag_selected)

        config_menu = menubar.addMenu("Configuration")
        self._add_menu_action(config_menu, "Settings...", self.open_config_dialog)

        help_menu = menubar.addMenu("Help")
        self._add_menu_action(help_menu, "UI Framework", self.open_framework_message)

    def _reload_project_runtime(self, project_dir: Path, label: str = "Loading project") -> None:
        project_dir = project_dir.expanduser().resolve()
        self.active_project_dir = project_dir

        def work():
            return backend.load_project_runtime(project_dir, prefer_hfest=self.prefer_hfest)

        def success(runtime: dict):
            self._apply_runtime(runtime)
            self._init_converted_map()
            self._populate_project_tree()
            self._populate_vb6_tree()
            self._populate_vb6_grid()
            self.current_comp_id = None
            self.current_output_path = None
            self.vb_source.setPlainText("")
            self.net_source.setPlainText("")
            project_name = runtime.get("layout", {}).get("name", project_dir.name)
            self._persist_config()
            self.set_status(f"Loaded project: {project_name} ({len(self.ordered_comp_ids)} files)")

        def error(exc: Exception):
            self.append_log(f"Load project failed: {exc}")
            self.set_status(f"Load project failed: {exc}")

        self.run_async(label, work, success, error)

    def on_new_project(self) -> None:
        initial = str(self.vb6_root_dir) if self.vb6_root_dir.exists() else str(Path.home())
        source_selected = QFileDialog.getExistingDirectory(self, "Select VB6 Source Folder", initial)
        if not source_selected:
            return
        suggested = backend.sanitize_project_name(Path(source_selected).name or "Project")
        project_name, ok = QInputDialog.getText(self, "New Project", "Project name folder:", QLineEdit.EchoMode.Normal, suggested)
        if not ok:
            self.set_status("New project cancelled.")
            return
        project_name = backend.sanitize_project_name(project_name or suggested)
        projects_root = self._projects_root_path()
        candidate = (projects_root / project_name).resolve()
        overwrite = False
        if candidate.exists() and any(candidate.iterdir()):
            answer = QMessageBox.question(
                self,
                "Overwrite Project",
                f"Project folder already exists:\n{candidate}\n\nOverwrite it?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                self.set_status("New project cancelled.")
                return
            overwrite = True

        def work():
            return backend.create_project_package(
                projects_root=projects_root,
                project_name=project_name,
                vb6_source_root=Path(source_selected),
                output_project=str(self.project_path),
                seed_target_dir=None,
                overwrite=overwrite,
            )

        def success(payload: tuple):
            project_dir, source_count, _target_count = payload
            self.set_status(f"Created project '{project_dir.name}' with {source_count} VB6 files.")
            self._reload_project_runtime(project_dir, label="Loading new project")

        def error(exc: Exception):
            self.append_log(f"New project failed: {exc}")
            self.set_status(f"New project failed: {exc}")

        self.run_async("Create project", work, success, error)

    def on_migration_wizard(self) -> None:
        """Show the Migration Wizard dialog to scaffold a new .NET Blazor project."""
        dlg = MigrationWizardDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            self.set_status("Migration wizard cancelled.")
            return

        project_name = dlg.project_name()
        vb6_source = dlg.vb6_source_dir()
        target_net = dlg.target_net_dir()
        static_assets = dlg.static_assets_dir() or None

        if not vb6_source or not Path(vb6_source).is_dir():
            QMessageBox.warning(self, "Migration Wizard", "Please select a valid VB6 source directory.")
            return
        if not target_net:
            QMessageBox.warning(self, "Migration Wizard", "Please specify a target .NET directory.")
            return

        projects_root = self._projects_root_path()
        safe_name = backend.sanitize_project_name(project_name)
        target_path = Path(target_net)
        csproj_path = target_path / f"{safe_name}.csproj"
        migrated_dir = target_path / "Components" / "Pages" / "Migrated"

        def _progress(msg: str) -> None:
            self._invoke_on_main.emit(lambda m=msg: self.set_status(m))

        def _log(msg: str) -> None:
            self._invoke_on_main.emit(lambda m=msg: self.append_log(m))

        def work():
            # 1. Scaffold the .NET project from HomeFrontPOC template
            _scaffold_dotnet_project(
                target_dir=target_path,
                project_name=safe_name,
                static_assets_dir=static_assets,
                progress_fn=_progress,
            )

            # 2. Create migrator project package (VB6 source cache)
            _progress(f"Creating migrator project with VB6 files from {Path(vb6_source).name}...")
            payload = backend.create_project_package(
                projects_root=projects_root,
                project_name=safe_name,
                vb6_source_root=Path(vb6_source),
                output_project=str(csproj_path),
                seed_target_dir=None,
                overwrite=True,
            )
            return payload

        def success(payload: tuple):
            project_dir, source_count, _target_count = payload

            # Update config to point at the new .NET project
            self.app_config["output_project"] = str(csproj_path)
            self.app_config["output_dir"] = str(migrated_dir)
            self.app_config["context_dir"] = str(target_path)
            self._persist_config()

            _log(f"Migration project '{safe_name}' created with {source_count} VB6 files.")
            _log(f".NET project: {target_path}")
            _log(f"Output: {migrated_dir}")
            self._reload_project_runtime(project_dir, label="Loading migration project")
            self.set_status(f"Migration project '{safe_name}' ready — {source_count} VB6 files")

        def error(exc: Exception):
            self.append_log(f"Migration wizard failed: {exc}")
            self.set_status(f"Migration wizard failed: {exc}")

        self.run_async("Migration Wizard", work, success, error)

    def on_load_project(self) -> None:
        projects_root = self._projects_root_path()
        candidates = backend.list_existing_project_dirs(projects_root, include_default=True)
        if not candidates:
            QMessageBox.information(self, "Load Project", "No existing projects found.")
            return
        labels: List[str] = []
        by_label: Dict[str, Path] = {}
        for path in candidates:
            label = path.name
            if not path.exists():
                label += " (not created)"
            elif not (path / backend.PROJECT_MANIFEST_NAME).exists():
                label += " (no manifest)"
            labels.append(label)
            by_label[label] = path
        selected_label, ok = QInputDialog.getItem(self, "Load Existing Project", "Existing Projects", labels, 0, False)
        if not ok or not selected_label:
            return
        project_dir = by_label.get(str(selected_label))
        if project_dir is None:
            return
        try:
            layout = backend.get_project_layout(project_dir)
            if backend.count_component_files(layout["source_dir"]) == 0:
                raise RuntimeError(f"No VB6 files found in project source: {layout['source_dir']}")
        except Exception as exc:
            QMessageBox.critical(self, "Load Project", f"Invalid project folder:\n{exc}")
            self.set_status(f"Load project failed: {exc}")
            return
        self._reload_project_runtime(layout["project_dir"], label="Loading project")

    def on_close_project(self) -> None:
        answer = QMessageBox.question(
            self,
            "Close Project",
            "Close the current project?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        projects_root = self._projects_root_path()
        try:
            empty_project = backend.ensure_empty_project(projects_root)
        except Exception as exc:
            self.set_status(f"Close project failed: {exc}")
            return
        self._reload_project_runtime(empty_project, label="Closing project")

    def on_refresh_stats(self) -> None:
        if not self.active_project_dir or not self.active_project_dir.exists():
            self.set_status("No active project.")
            return
        self._reload_project_runtime(self.active_project_dir, label="Refreshing stats")

    def on_rebuild_cache(self) -> None:
        if not self.active_project_dir or not self.active_project_dir.exists():
            self.set_status("No active project.")
            return
        try:
            layout = backend.get_project_layout(self.active_project_dir)
            shutil.rmtree(layout["cache_dir"], ignore_errors=True)
            layout["cache_dir"].mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            self.set_status(f"Rebuild cache failed: {exc}")
            return
        self._reload_project_runtime(self.active_project_dir, label="Rebuilding cache")

    def open_config_dialog(self) -> None:
        dlg = QDialog(self)
        dlg.setWindowTitle("Configuration")
        dlg.resize(920, 620)
        root = QVBoxLayout(dlg)

        grid = QVBoxLayout()
        root.addLayout(grid)

        def row_line(label: str, value: str, browse: bool = False, file_mode: bool = False):
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            edit = QLineEdit(value)
            row.addWidget(edit, stretch=1)
            if browse:
                btn = QPushButton("Browse")
                def on_browse():
                    if file_mode:
                        start = edit.text().strip() or str(Path.home())
                        selected, _f = QFileDialog.getOpenFileName(dlg, f"Select {label}", start)
                    else:
                        start = edit.text().strip() or str(Path.home())
                        selected = QFileDialog.getExistingDirectory(dlg, f"Select {label}", start)
                    if selected:
                        edit.setText(selected)
                btn.clicked.connect(on_browse)
                row.addWidget(btn)
            grid.addLayout(row)
            return edit

        root_edit = row_line("VB6 Root", str(self.vb6_root_dir), browse=True)
        prefer_row = QHBoxLayout()
        prefer_row.addWidget(QLabel("Prefer HFEst"))
        prefer_chk = QCheckBox()
        prefer_chk.setChecked(self.prefer_hfest)
        prefer_row.addWidget(prefer_chk)
        prefer_row.addStretch(1)
        grid.addLayout(prefer_row)
        project_edit = row_line("HomeFront.csproj", str(self.project_path), browse=True, file_mode=True)
        output_edit = row_line("Output Dir", str(self.output_dir), browse=True)
        context_edit = row_line("Context Dir", (self.app_config.get("context_dir") or "").strip(), browse=True)
        vb6_img_edit = row_line("VB6 Image Dir", str(self.vb6_image_dir), browse=True)
        net_img_edit = row_line(".NET Image Dir", str(self.dotnet_image_dir), browse=True)

        model_row = QHBoxLayout()
        model_row.addWidget(QLabel("Model"))
        model_combo = QComboBox()
        model_combo.addItems(["gpt-5.2-codex", "claude-opus-4-6", "deepseek-r1:8b", "qwen3:8b", "deepseek-r1:32b", "qwen2.5-coder:32b", "qwen3:32b"])
        model_combo.setCurrentText((self.app_config.get("model") or "gpt-5.2-codex").strip())
        model_row.addWidget(model_combo)
        model_row.addWidget(QLabel("Temp"))
        temp_edit = QLineEdit((self.app_config.get("temperature") or "0.2").strip())
        temp_edit.setMaximumWidth(90)
        model_row.addWidget(temp_edit)
        model_row.addStretch(1)
        grid.addLayout(model_row)

        grid.addWidget(QLabel("Instructions"))
        instr = QPlainTextEdit()
        instr.setPlainText(self.app_config.get("instructions") or "")
        instr.setMinimumHeight(130)
        root.addWidget(instr, stretch=1)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel_btn = QPushButton("Cancel")
        save_btn = QPushButton("Save")
        buttons.addWidget(cancel_btn)
        buttons.addWidget(save_btn)
        root.addLayout(buttons)

        cancel_btn.clicked.connect(dlg.reject)

        def on_save():
            old_prefer = self.prefer_hfest
            self.vb6_root_dir = Path(root_edit.text().strip()).expanduser()
            self.prefer_hfest = prefer_chk.isChecked()
            self.project_path = Path(project_edit.text().strip()).expanduser()
            self.output_dir = Path(output_edit.text().strip()).expanduser()
            self.app_config["context_dir"] = context_edit.text().strip()
            self.vb6_image_dir = Path(vb6_img_edit.text().strip()).expanduser()
            self.dotnet_image_dir = Path(net_img_edit.text().strip()).expanduser()
            self.app_config["model"] = model_combo.currentText().strip() or "gpt-5.2-codex"
            self.app_config["temperature"] = temp_edit.text().strip() or "0.2"
            self.app_config["instructions"] = instr.toPlainText()
            self.project_root = self.project_path.parent if self.project_path.exists() else self.vb6_root_dir
            self.ui_defaults["root"] = str(self.vb6_root_dir)
            self.ui_defaults["project"] = str(self.project_path)
            self.ui_defaults["output_dir"] = str(self.output_dir)
            self.ui_defaults["context_dir"] = self.app_config["context_dir"]
            self.ui_defaults["vb6_image_dir"] = str(self.vb6_image_dir)
            self.ui_defaults["dotnet_image_dir"] = str(self.dotnet_image_dir)
            self.ui_defaults["model"] = self.app_config["model"]
            self.ui_defaults["temperature"] = self.app_config["temperature"]
            self.ui_defaults["prefer_hfest"] = "true" if self.prefer_hfest else "false"
            self._init_converted_map()
            self._populate_project_tree()
            self._populate_vb6_tree()
            self._populate_vb6_grid()
            self._persist_config()
            if self.prefer_hfest != old_prefer and self.active_project_dir.exists():
                self._reload_project_runtime(self.active_project_dir, label="Reloading with updated settings")
            else:
                self.set_status("Configuration updated.")
            dlg.accept()

        save_btn.clicked.connect(on_save)
        dlg.exec()

    def set_status(self, text: str) -> None:
        self.status_label.setText(text)

    def append_log(self, text: str, reset: bool = False) -> None:
        if reset:
            self.log_output.setPlainText(text)
        else:
            existing = self.log_output.toPlainText()
            if existing:
                self.log_output.setPlainText(existing + "\n" + text)
            else:
                self.log_output.setPlainText(text)
        cursor = self.log_output.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.log_output.setTextCursor(cursor)

    def set_busy(self, busy: bool, message: str | None = None) -> None:
        self.busy = busy
        for action in (
            self.convert_action,
            self.convert_selected_action,
            self.compile_action,
            self.fix_action,
            self.run_action,
            self.search_action,
        ):
            action.setEnabled(not busy)
        if message:
            self.set_status(message)

    def _run_on_main(self, fn) -> None:
        """Slot executed on the main thread when _invoke_on_main is emitted."""
        fn()

    def run_async(self, label: str, work_fn, success_fn, error_fn=None) -> None:
        if self.busy:
            self.set_status("Please wait for current task to complete.")
            return
        self.set_busy(True, f"{label}...")

        def on_success(result):
            self.set_busy(False)
            success_fn(result)

        def on_error(exc: Exception):
            self.set_busy(False)
            if error_fn is not None:
                error_fn(exc)
            else:
                msg = f"{label} failed: {exc}"
                self.set_status(msg)
                self.append_log(msg)

        def worker():
            try:
                result = work_fn()
                self._invoke_on_main.emit(lambda r=result: on_success(r))
            except Exception as exc:
                self._invoke_on_main.emit(lambda e=exc: on_error(e))

        threading.Thread(target=worker, daemon=True).start()

    def on_browse_context(self) -> None:
        """Browse context directory (called from menu if needed)."""
        current = (self.app_config.get("context_dir") or "").strip()
        selected = QFileDialog.getExistingDirectory(self, "Select Context Directory", current)
        if selected:
            self.app_config["context_dir"] = selected
            self._persist_config()
            self.set_status(f"Context directory set: {selected}")

    def _project_skip(self, path: Path) -> bool:
        name = path.name
        if name in {"bin", "obj", ".git", ".idea", ".vs", "__pycache__"}:
            return True
        if name == ".DS_Store":
            return True
        return False

    def _project_children(self, dir_path: Path) -> List[Path]:
        try:
            children = [p for p in dir_path.iterdir() if not self._project_skip(p)]
        except Exception:
            return []
        return sorted(children, key=lambda p: (0 if p.is_dir() else 1, p.name.lower()))

    def _add_dummy_if_needed(self, item: QTreeWidgetItem, dir_path: Path) -> None:
        for _child in self._project_children(dir_path):
            item.addChild(QTreeWidgetItem(["..."]))
            break

    def _populate_project_tree(self) -> None:
        self.project_tree.clear()
        if not self.project_root.exists():
            return
        root_item = QTreeWidgetItem([self.project_root.name or str(self.project_root)])
        root_item.setData(0, ROLE_PATH, str(self.project_root))
        root_item.setData(0, ROLE_FLAG, False)
        self.project_tree.addTopLevelItem(root_item)
        self._populate_project_item(root_item)
        root_item.setExpanded(True)

    def _populate_project_item(self, item: QTreeWidgetItem) -> None:
        loaded = bool(item.data(0, ROLE_FLAG))
        if loaded:
            return
        path_raw = item.data(0, ROLE_PATH)
        if not path_raw:
            return
        path = Path(str(path_raw))
        if not path.is_dir():
            return
        item.takeChildren()
        for child_path in self._project_children(path):
            child_item = QTreeWidgetItem([child_path.name])
            child_item.setData(0, ROLE_PATH, str(child_path))
            child_item.setData(0, ROLE_FLAG, False)
            item.addChild(child_item)
            if child_path.is_dir():
                self._add_dummy_if_needed(child_item, child_path)
        item.setData(0, ROLE_FLAG, True)

    def _on_project_item_expanded(self, item: QTreeWidgetItem) -> None:
        self._populate_project_item(item)

    def _on_project_item_clicked(self, item: QTreeWidgetItem, _col: int) -> None:
        path_raw = item.data(0, ROLE_PATH)
        if not path_raw:
            return
        path = Path(str(path_raw))
        if path.is_file():
            self._show_dotnet_file(path)

    def _populate_vb6_tree(self) -> None:
        self.vb6_tree.clear()
        self.vb6_tree_item_by_comp_id.clear()
        seen_root: set[str] = set()
        root_nodes: List[str] = []
        for comp_id in self.roots:
            if comp_id in self.id_to_component and comp_id not in seen_root:
                seen_root.add(comp_id)
                root_nodes.append(comp_id)
        if "HFSystem::FLogin" in self.id_to_component and "HFSystem::FLogin" not in seen_root:
            root_nodes.append("HFSystem::FLogin")

        if root_nodes:
            roots_item = QTreeWidgetItem([f"Call Roots ({len(root_nodes)})"])
            roots_item.setData(0, ROLE_NODE_KIND, "section")
            self.vb6_tree.addTopLevelItem(roots_item)
            for comp_id in root_nodes:
                ambig = comp_id.split("::", 1)[-1] in self.duplicates
                self._add_vb6_component_item(roots_item, comp_id, ambig, [])
            roots_item.setExpanded(True)

        detached = [cid for cid in self.ordered_comp_ids if cid not in self.reachable_ids]
        if detached:
            detached_item = QTreeWidgetItem([f"Detached / Unreachable ({len(detached)})"])
            detached_item.setData(0, ROLE_NODE_KIND, "section")
            self.vb6_tree.addTopLevelItem(detached_item)
            for comp_id in detached:
                ambig = comp_id.split("::", 1)[-1] in self.duplicates
                self._add_vb6_component_item(detached_item, comp_id, ambig, [])

        if not root_nodes and not detached:
            all_item = QTreeWidgetItem([f"All Components ({len(self.ordered_comp_ids)})"])
            all_item.setData(0, ROLE_NODE_KIND, "section")
            self.vb6_tree.addTopLevelItem(all_item)
            for comp_id in self.ordered_comp_ids:
                ambig = comp_id.split("::", 1)[-1] in self.duplicates
                self._add_vb6_component_item(all_item, comp_id, ambig, [])
            all_item.setExpanded(True)

    def _add_vb6_component_item(
        self,
        parent: QTreeWidgetItem,
        comp_id: str,
        ambiguous: bool,
        stack: List[str],
    ) -> QTreeWidgetItem:
        is_cycle = comp_id in stack
        label = comp_id
        if ambiguous:
            label += " [ambig]"
        if is_cycle:
            label += " [cycle]"
        item = QTreeWidgetItem([label])
        item.setData(0, ROLE_COMP_ID, comp_id)
        item.setData(0, ROLE_STACK, list(stack) + [comp_id])
        item.setData(0, ROLE_FLAG, False)
        if comp_id in self.orphan_set:
            item.setForeground(0, QColor("#cc0000"))
        self.vb6_tree_item_by_comp_id.setdefault(comp_id, item)
        parent.addChild(item)
        if not is_cycle and self.adj.get(comp_id):
            item.addChild(QTreeWidgetItem(["..."]))
        else:
            item.setData(0, ROLE_FLAG, True)
        return item

    def _on_vb6_tree_item_expanded(self, item: QTreeWidgetItem) -> None:
        comp_id = item.data(0, ROLE_COMP_ID)
        if not comp_id:
            return
        if bool(item.data(0, ROLE_FLAG)):
            return
        stack_obj = item.data(0, ROLE_STACK)
        stack = list(stack_obj) if isinstance(stack_obj, list) else [str(comp_id)]
        item.takeChildren()
        for child_id, child_ambig in self.adj.get(str(comp_id), []):
            self._add_vb6_component_item(item, child_id, child_ambig, stack)
        item.setData(0, ROLE_FLAG, True)

    def _populate_vb6_grid(self) -> None:
        self.vb6_grid.setSortingEnabled(False)
        self.vb6_grid.blockSignals(True)
        self.vb6_grid.setUpdatesEnabled(False)

        rows: List[Tuple[str, str, int, str, str, str, int, int]] = []
        for idx, comp_id in enumerate(self.ordered_comp_ids, start=1):
            comp = self.id_to_component.get(comp_id)
            if not comp:
                continue
            ext = comp.path.suffix.lower().lstrip(".")
            rows.append(
                (
                    comp_id,
                    "✔" if comp_id in self.converted_map else "",
                    idx,
                    comp.name,
                    ext,
                    comp.project,
                    int(self.file_size_by_id.get(comp_id, 0) or 0),
                    int(self.line_count_by_id.get(comp_id, 0) or 0),
                )
            )

        self.vb6_grid.clearContents()
        self.vb6_grid.setRowCount(len(rows))

        for row, row_data in enumerate(rows):
            comp_id = row_data[0]
            values = row_data[1:]
            for col, value in enumerate(values):
                item = QTableWidgetItem()
                if col in (1, 5, 6):
                    item.setData(Qt.ItemDataRole.DisplayRole, int(value))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                else:
                    item.setText(str(value))
                item.setData(ROLE_COMP_ID, comp_id)
                if comp_id in self.orphan_set:
                    item.setForeground(QColor("#cc0000"))
                self.vb6_grid.setItem(row, col, item)

        self.vb6_grid.resizeColumnsToContents()
        if self.vb6_grid.columnWidth(2) < 220:
            self.vb6_grid.setColumnWidth(2, 220)

        self.vb6_grid.setUpdatesEnabled(True)
        self.vb6_grid.blockSignals(False)
        self.vb6_grid.setSortingEnabled(True)

    def _mark_row_converted(self, comp_id: str) -> None:
        for row in range(self.vb6_grid.rowCount()):
            item = self.vb6_grid.item(row, 0)
            if item is None:
                continue
            row_comp_id = item.data(ROLE_COMP_ID)
            if row_comp_id == comp_id:
                self.vb6_grid.item(row, 0).setText("✔")
                return

    def _on_vb6_tree_selection(self) -> None:
        items = self.vb6_tree.selectedItems()
        if not items:
            return
        comp_id = items[0].data(0, ROLE_COMP_ID)
        if isinstance(comp_id, str) and comp_id:
            self._show_component(str(comp_id))

    def _on_vb6_grid_selection(self) -> None:
        selection_model = self.vb6_grid.selectionModel()
        if selection_model is None:
            return
        rows = selection_model.selectedRows()
        if not rows:
            return
        row = rows[0].row()
        if row < 0 or row >= self.vb6_grid.rowCount():
            return
        item = self.vb6_grid.item(row, 0)
        if not item:
            return
        comp_id = item.data(ROLE_COMP_ID)
        if comp_id:
            self._show_component(str(comp_id))

    def _selected_comp_ids(self, multi: bool) -> List[str]:
        ids: List[str] = []
        in_grid_mode = self.vb6_stack.currentIndex() == 1

        def collect_from_grid() -> List[str]:
            selected: List[str] = []
            rows = self.vb6_grid.selectionModel().selectedRows()
            if not rows:
                return selected
            seen = set()
            for idx in rows:
                item = self.vb6_grid.item(idx.row(), 0)
                if item is None:
                    continue
                comp_id = item.data(ROLE_COMP_ID)
                if isinstance(comp_id, str) and comp_id and comp_id not in seen:
                    seen.add(comp_id)
                    selected.append(comp_id)
            return selected

        def collect_from_tree() -> List[str]:
            selected: List[str] = []
            items = self.vb6_tree.selectedItems()
            if not items:
                return selected
            comp_id = items[0].data(0, ROLE_COMP_ID)
            if isinstance(comp_id, str) and comp_id:
                selected.append(comp_id)
            return selected

        primary = collect_from_grid if in_grid_mode else collect_from_tree
        secondary = collect_from_tree if in_grid_mode else collect_from_grid
        ids = primary()
        if not ids:
            ids = secondary()
        return ids if multi else ids[:1]

    def _show_component(self, comp_id: str) -> None:
        comp = self.id_to_component.get(comp_id)
        if not comp:
            return
        self.current_comp_id = comp_id
        vb_text = backend.read_text(comp.path)
        self.vb_source.setPlainText(f"{comp_id} :: {comp.path}\n\n{vb_text}")
        self.right_tabs.setCurrentIndex(0)

        net_file = self.converted_map.get(comp_id)
        if net_file is None or not net_file.exists():
            candidate = self._find_existing_razor_for_component(comp)
            if candidate and candidate.exists():
                net_file = candidate
                self.converted_map[comp_id] = candidate
        if net_file and net_file.exists():
            self._show_dotnet_file(net_file)
            self._mark_row_converted(comp_id)
        else:
            self.current_output_path = None
            self.net_source.setPlainText("Not converted yet.")
        self.set_status(f"Selected: {comp_id}")

    def _show_vb6_only(self, path: Path) -> None:
        try:
            text = backend.read_text(path)
        except Exception:
            text = ""
        self.vb_source.setPlainText(f"{path.name} :: {path}\n\n{text}")

    def _show_dotnet_file(self, path: Path) -> None:
        binary_suffixes = {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".ico", ".webp", ".svg",
            ".frx", ".ctx", ".dll", ".exe", ".zip", ".pdf",
        }
        header = f"{path.name} :: {path}\n\n"
        if path.suffix.lower() in binary_suffixes:
            self.net_source.setPlainText(header + "[Binary file preview is disabled.]")
            self.current_output_path = None
            return
        try:
            content = path.read_text(encoding="utf-8", errors="ignore")
        except Exception as exc:
            self.net_source.setPlainText(header + f"[Unable to read file: {exc}]")
            self.current_output_path = None
            return
        self.net_source.setPlainText(header + content)
        self.current_output_path = path
        self.right_tabs.setCurrentIndex(1)
        self.set_status(f"Opened .NET file: {path.name}")

    def open_search_result(self, path: Path, is_target: bool, line_no: int) -> None:
        if is_target:
            self._show_dotnet_file(path)
            comp_id = self.comp_id_by_name.get(path.stem.lower())
            if comp_id:
                self._show_component(comp_id)
            self._scroll_editor_to_line(self.net_source, line_no + 2)
            return

        comp_id = self.comp_id_by_path.get(str(path.resolve()))
        if comp_id:
            self._show_component(comp_id)
            self._scroll_editor_to_line(self.vb_source, line_no + 2)
        else:
            self._show_vb6_only(path)
            self.right_tabs.setCurrentIndex(0)
            self._scroll_editor_to_line(self.vb_source, line_no + 2)
            self.net_source.setPlainText("Not converted yet.")

    def _scroll_editor_to_line(self, editor: QPlainTextEdit, line_no: int) -> None:
        doc = editor.document()
        if line_no < 1:
            line_no = 1
        block = doc.findBlockByLineNumber(line_no - 1)
        if not block.isValid():
            return
        cursor = editor.textCursor()
        cursor.setPosition(block.position())
        editor.setTextCursor(cursor)
        editor.centerCursor()

    def _get_context_blob(self, context_dir_raw: str) -> str:
        context_dir = Path(context_dir_raw).expanduser()
        if not context_dir.exists():
            return ""
        files = backend.collect_context_files(context_dir)
        signature = backend.compute_context_signature(files)
        dir_key = str(context_dir.resolve())
        if self.context_cache["dir"] == dir_key and self.context_cache["signature"] == signature:
            return self.context_cache["blob"]
        blob = backend.build_context_blob(context_dir, files)
        self.context_cache = {"dir": dir_key, "signature": signature, "blob": blob}
        return blob

    def _model_and_temp(self) -> Tuple[str, float]:
        model = (self.app_config.get("model") or "gpt-5.2-codex").strip()
        try:
            temperature = float((self.app_config.get("temperature") or "0.2").strip())
        except Exception:
            temperature = 0.2
        return model, temperature

    def _convert_components(self, comp_ids: List[str], model: str, temperature: float,
                           context_dir: str, instructions: str) -> None:
        if not comp_ids:
            self.set_status("Select VB6 item(s) to convert.")
            return
        if not self.project_path.exists():
            self.set_status(f"Project not found: {self.project_path}")
            return

        # Clear the log panel before starting a new conversion
        self.append_log("", reset=True)

        out_dir = self.output_dir
        project_path = self.project_path
        if backend.is_ollama_model(model):
            api_label = "Ollama"
        elif backend.is_claude_model(model):
            api_label = "Claude"
        else:
            api_label = "Codex"
        total = len(comp_ids)

        def _progress(msg: str) -> None:
            """Safely post a status update to the main thread via signal."""
            self._invoke_on_main.emit(lambda m=msg: self.set_status(m))

        def _log(msg: str) -> None:
            """Safely post a log message to the main thread via signal."""
            self._invoke_on_main.emit(lambda m=msg: self.append_log(m))

        def work():
            _progress(f"Loading context for {api_label} conversion...")
            context_blob = self._get_context_blob(context_dir)
            if backend.is_ollama_model(model):
                convert_fn = backend.call_ollama_convert
            elif backend.is_claude_model(model):
                convert_fn = backend.call_claude_convert
            else:
                convert_fn = backend.call_codex_convert
            payloads = []
            if backend.is_ollama_model(model):
                pause = 0
            elif backend.is_claude_model(model):
                pause = 180
            else:
                pause = 5
            for idx, comp_id in enumerate(comp_ids, start=1):
                comp = self.id_to_component.get(comp_id)
                if not comp:
                    continue
                if idx > 1:
                    _progress(f"Waiting {pause}s before {idx}/{total}: {comp.name}...")
                    time.sleep(pause)
                _progress(f"Converting {idx}/{total}: {comp.name}")
                vb6_source = backend.read_text(comp.path)
                out_name = f"{comp.name}.razor"
                result = convert_fn(
                    vb6_source,
                    comp.path,
                    out_name,
                    model,
                    temperature,
                    instructions,
                    context_blob,
                )
                out_dir.mkdir(parents=True, exist_ok=True)
                out_path = out_dir / out_name
                out_path.write_text(result, encoding="utf-8")
                try:
                    backend.ensure_csproj_includes(project_path, out_path)
                except (ValueError, RuntimeError) as csproj_err:
                    _log(f"Warning: Could not add {out_name} to csproj: {csproj_err}")
                payloads.append({"comp_id": comp_id, "out_path": out_path, "result": result})
                _log(f"Converted {idx}/{total}: {comp.name} → {out_name}")
            return payloads

        def success(payloads: List[dict]):
            if not payloads:
                self.set_status("No files converted.")
                return
            self._invalidate_razor_indexes()
            for payload in payloads:
                comp_id = payload["comp_id"]
                out_path = payload["out_path"]
                self.converted_map[comp_id] = out_path
                self._mark_row_converted(comp_id)
            last = payloads[-1]
            self._show_dotnet_file(last["out_path"])
            self.append_log(f"Converted {len(payloads)} file(s).")
            self.set_status(f"Converted {len(payloads)} file(s). Saved to {last['out_path']}")

        def error(exc: Exception):
            msg = f"Conversion failed: {exc}"
            self.append_log(msg)
            self.set_status(msg)

        self.run_async(f"Converting with {api_label}", work, success, error)

    def _show_convert_dialog(self, mode: str) -> None:
        """Show the ConvertDialog popup and run conversion if confirmed."""
        multi = mode == "selected"
        ids = self._selected_comp_ids(multi=multi)
        if not ids:
            self.set_status("Select VB6 item(s) to convert.")
            return

        defaults = {
            "model": self.app_config.get("model", "gpt-5.2-codex"),
            "temperature": self.app_config.get("temperature", "0.2"),
            "context_dir": self.app_config.get("context_dir", ""),
            "instructions": self.app_config.get("instructions", ""),
        }
        dlg = ConvertDialog(self, mode=mode, defaults=defaults)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        # Persist user choices back to config
        model = dlg.selected_model()
        temperature = dlg.selected_temperature()
        context_dir = dlg.context_dir()
        instructions = dlg.instructions()
        self.app_config["model"] = model
        self.app_config["temperature"] = str(temperature)
        self.app_config["context_dir"] = context_dir
        self.app_config["instructions"] = instructions
        self._persist_config()

        self._convert_components(ids, model, temperature, context_dir, instructions)

    def on_convert(self) -> None:
        self._show_convert_dialog(mode="single")

    def on_convert_selected(self) -> None:
        self._show_convert_dialog(mode="selected")

    def on_tag_selected(self) -> None:
        ids = self._selected_comp_ids(multi=True)
        if not ids:
            self.set_status("No selected items to tag.")
            return

        model = (self.app_config.get("model") or "claude-opus-4-6").strip()
        if not backend.is_claude_model(model):
            model = "claude-opus-4-6"
        context_dir = (self.app_config.get("context_dir") or "").strip()

        def work():
            context_blob = self._get_context_blob(context_dir)
            results: List[Tuple[str, str | None]] = []
            for idx, comp_id in enumerate(ids, start=1):
                comp = self.id_to_component.get(comp_id)
                if not comp:
                    results.append((comp_id, "Component not found"))
                    continue
                vb6_source = backend.read_text(comp.path)
                migrated_source = ""
                out_path = self.converted_map.get(comp_id)
                if out_path is None or not out_path.exists():
                    out_path = self._find_existing_razor_for_component(comp)
                if out_path and out_path.exists():
                    migrated_source = out_path.read_text(encoding="utf-8", errors="ignore")
                tagged = backend.call_claude_tag(
                    source_text=vb6_source,
                    source_path=comp.path,
                    model=model,
                    context_blob=context_blob,
                    migrated_source=migrated_source,
                )
                comp.path.write_text(tagged, encoding="latin-1", errors="ignore")
                try:
                    self.file_size_by_id[comp_id] = comp.path.stat().st_size
                    self.line_count_by_id[comp_id] = backend.count_lines_file(comp.path)
                except Exception:
                    pass
                results.append((comp_id, None))
                if idx < len(ids):
                    time.sleep(2)
            return results

        def success(results: List[Tuple[str, str | None]]):
            ok = sum(1 for _, err in results if err is None)
            fail = sum(1 for _, err in results if err is not None)
            self._populate_vb6_grid()
            if self.current_comp_id:
                self._show_component(self.current_comp_id)
            msg = f"Tagging complete: {ok} succeeded"
            if fail:
                first_err = next((f"{cid}: {err}" for cid, err in results if err), "")
                msg += f", {fail} failed"
                if first_err:
                    self.append_log(f"Tag failed: {first_err}")
            self.set_status(msg)

        def error(exc: Exception):
            self.set_status(f"Tag failed: {exc}")
            self.append_log(f"Tag failed: {exc}")

        self.run_async("Tag", work, success, error)

    def _resolve_component_image(self, comp_id: str, is_vb6: bool) -> Path | None:
        comp = self.id_to_component.get(comp_id)
        if not comp:
            return None
        image_dir = self.vb6_image_dir if is_vb6 else self.dotnet_image_dir
        if not image_dir.exists():
            return None
        return backend.find_component_image(image_dir, comp.name)

    def on_align_screens(self) -> None:
        ids = self._selected_comp_ids(multi=False)
        comp_id = ids[0] if ids else (self.current_comp_id or "")
        if not comp_id:
            self.set_status("Select a component to align.")
            return
        comp = self.id_to_component.get(comp_id)
        if not comp:
            self.set_status("Invalid component selection.")
            return

        vb_img = self._resolve_component_image(comp_id, is_vb6=True)
        net_img = self._resolve_component_image(comp_id, is_vb6=False)
        if vb_img is None or not vb_img.exists():
            self.set_status("VB6 screenshot not found. Update VB6 image directory in Configuration.")
            return
        if net_img is None or not net_img.exists():
            self.set_status(".NET screenshot not found. Update .NET image directory in Configuration.")
            return

        out_path = self.converted_map.get(comp_id)
        if out_path is None or not out_path.exists():
            out_path = self._find_existing_razor_for_component(comp)
        if out_path and out_path.exists():
            self.converted_map[comp_id] = out_path
        if out_path is None or not out_path.exists():
            self.set_status(f"Converted file not found for alignment: {out_path}")
            return

        model, temperature = self._model_and_temp()

        def work():
            metrics = backend.run_pixel_pass(vb_img, net_img)
            report = backend.build_alignment_report(vb_img, net_img, metrics)
            blazor_source = out_path.read_text(encoding="utf-8", errors="ignore")
            align_fn = backend.call_claude_align if backend.is_claude_model(model) else backend.call_codex_align
            aligned = align_fn(blazor_source, out_path, report, model, temperature)
            out_path.write_text(aligned, encoding="utf-8")
            return {"out_path": out_path, "aligned": aligned, "metrics": metrics, "report": report}

        def success(payload: dict):
            self.converted_map[comp_id] = payload["out_path"]
            self._mark_row_converted(comp_id)
            self._show_dotnet_file(payload["out_path"])
            metrics = payload.get("metrics", {})
            summary = [
                "Alignment completed.",
                f"VB6 size: {metrics.get('vb6_size')}",
                f".NET size: {metrics.get('dotnet_size')}",
                f"MAE(0-255): {metrics.get('mae_255')}",
                f"MAE(%): {metrics.get('mae_pct')}",
                f"Changed pixels(%): {metrics.get('changed_pct')}",
                "",
                payload.get("report", ""),
            ]
            self.append_log("\n".join(summary).strip(), reset=True)
            self.set_status("Alignment pass completed and Razor file updated.")

        def error(exc: Exception):
            self.append_log(f"Align failed: {exc}", reset=True)
            self.set_status(f"Align failed: {exc}")

        self.run_async("Align Screens", work, success, error)

    def on_compile(self) -> None:
        if not self.project_path.exists():
            self.set_status(f"Project not found: {self.project_path}")
            return

        def work():
            result = subprocess.run(
                ["dotnet", "build", str(self.project_path)],
                cwd=str(self.project_path.parent),
                capture_output=True,
                text=True,
            )
            return {"returncode": result.returncode, "output": (result.stdout or "") + "\n" + (result.stderr or "")}

        def success(payload: dict):
            output = payload.get("output", "")
            if payload.get("returncode") == 0:
                self.last_compile_errors = ""
                self.append_log("Build succeeded.", reset=True)
                self.set_status("Build succeeded.")
                return

            self.last_compile_errors = output
            grouped: Dict[str, List[str]] = {}
            error_regex_full = re.compile(
                r"^(?P<file>.+)\((?P<line>\d+),(?P<col>\d+)\):\s+error\s+(?P<code>[^:]+):\s+(?P<msg>.+)$",
                re.IGNORECASE,
            )
            error_regex_line = re.compile(
                r"^(?P<file>.+)\((?P<line>\d+)\):\s+error\s+(?P<code>[^:]+):\s+(?P<msg>.+)$",
                re.IGNORECASE,
            )
            error_regex_simple = re.compile(
                r"^(?P<file>.+):\s+error\s+(?P<code>[^:]+):\s+(?P<msg>.+)$",
                re.IGNORECASE,
            )
            for line in output.splitlines():
                stripped = line.strip()
                m = error_regex_full.match(stripped) or error_regex_line.match(stripped) or error_regex_simple.match(stripped)
                if m:
                    file_name = Path(m.group("file").strip()).name or "Other"
                    line_no = m.groupdict().get("line")
                    msg = f"{m.group('code').strip()}: {m.group('msg').strip()}"
                    prefix = f"L{line_no}: " if line_no else ""
                    grouped.setdefault(file_name, []).append(prefix + msg)
                elif " error " in stripped.lower():
                    grouped.setdefault("Other", []).append(stripped)

            if grouped:
                chunks: List[str] = []
                for file_name, errs in grouped.items():
                    chunks.append(file_name)
                    chunks.extend([f"  {e}" for e in errs])
                    chunks.append("")
                self.append_log("\n".join(chunks).strip(), reset=True)
            else:
                self.append_log(output.strip() or "Build failed.", reset=True)
            self.set_status("Build failed.")

        def error(exc: Exception):
            msg = f"Compile failed: {exc}"
            self.append_log(msg, reset=True)
            self.set_status(msg)

        self.run_async("Compile", work, success, error)

    def on_fix(self) -> None:
        if self.current_output_path is None or not self.current_output_path.exists():
            self.set_status("No converted file to fix.")
            return
        errors = (self.log_output.toPlainText().strip() or self.last_compile_errors.strip())
        if not errors:
            self.set_status("No compile errors to fix.")
            return

        model, temperature = self._model_and_temp()
        source_path = self.current_output_path

        def work():
            blazor_source = source_path.read_text(encoding="utf-8", errors="ignore")
            fix_fn = backend.call_claude_fix if backend.is_claude_model(model) else backend.call_codex_fix
            fixed = fix_fn(blazor_source, errors, source_path, model, temperature)
            source_path.write_text(fixed, encoding="utf-8")
            return fixed

        def success(fixed: str):
            self._show_dotnet_file(source_path)
            self.append_log("Fix applied.", reset=True)
            self.set_status("Fix applied.")

        def error(exc: Exception):
            msg = f"Fix failed: {exc}"
            self.append_log(msg, reset=True)
            self.set_status(msg)

        self.run_async("Fix", work, success, error)

    def _render_run_log(self) -> str:
        lines: List[str] = []
        if self.run_log_header_lines:
            lines.extend(self.run_log_header_lines)
        if self.run_tail_buffer.strip():
            if lines:
                lines.append("")
            lines.append(self.run_tail_buffer.rstrip())
        if not lines:
            return "Waiting for run output...\n"
        return ("\n".join(lines).rstrip() + "\n")

    def _launch_run_process(self, is_retry: bool = False, retry_reason: str = "", killed_pids: List[int] | None = None) -> None:
        if self.run_process is not None and self.run_process.poll() is None:
            backend._terminate_process(self.run_process)
        self.run_process = None

        log_dir = self.output_dir
        log_dir.mkdir(parents=True, exist_ok=True)
        self.run_log_path = log_dir / "run.log"
        self.run_tail_offset = 0
        self.run_tail_buffer = ""
        self.run_url_announced = ""
        self.run_exit_announced = False
        self.run_browser.reset()

        port = backend._pick_free_local_port()
        requested_url = f"http://127.0.0.1:{port}"
        env = dict(os.environ)
        env["ASPNETCORE_URLS"] = requested_url
        env["ASPNETCORE_ENVIRONMENT"] = "Development"
        env["DOTNET_ENVIRONMENT"] = "Development"

        with self.run_log_path.open("w", encoding="utf-8") as log_file:
            self.run_process = subprocess.Popen(
                ["dotnet", "run", "--no-launch-profile", "--project", str(self.project_path)],
                cwd=str(self.project_path.parent),
                stdout=log_file,
                stderr=subprocess.STDOUT,
                env=env,
                start_new_session=True,
            )

        self.run_log_header_lines = [
            f"Running... PID {self.run_process.pid}",
            f"URL: {requested_url} (starting...)",
            f"Log: {self.run_log_path}",
        ]
        if is_retry:
            note = f"Auto-retry {self.run_restart_count}/{self.run_max_restarts}"
            if retry_reason:
                note += f": {retry_reason}"
            self.run_log_header_lines.insert(0, note)
            if killed_pids:
                self.run_log_header_lines.append(
                    "Killed listener PID(s): " + ", ".join(str(p) for p in killed_pids)
                )

        self.append_log(self._render_run_log(), reset=True)
        self.set_status(f"Program started: {requested_url}")
        if not self.run_tail_timer.isActive():
            self.run_tail_timer.start()

    def _read_new_run_log(self) -> bool:
        if self.run_log_path is None or not self.run_log_path.exists():
            return False
        with self.run_log_path.open("r", encoding="utf-8", errors="ignore") as f:
            f.seek(self.run_tail_offset)
            chunk = f.read()
            self.run_tail_offset = f.tell()
        if not chunk:
            return False
        self.run_tail_buffer += chunk
        if len(self.run_tail_buffer) > 20000:
            self.run_tail_buffer = self.run_tail_buffer[-20000:]
        opened_url = self.run_browser.maybe_open_from_text(chunk)
        if opened_url and opened_url != self.run_url_announced:
            self.run_url_announced = opened_url
            url_line = f"URL: {opened_url}"
            if len(self.run_log_header_lines) >= 2 and self.run_log_header_lines[1].startswith("URL:"):
                self.run_log_header_lines[1] = url_line
            else:
                self.run_log_header_lines.insert(1, url_line)
            self.set_status(f"Opened browser: {opened_url}")
        return True

    def _tail_run_log(self) -> None:
        if self.run_log_path is None:
            self.run_tail_timer.stop()
            return
        try:
            updated = self._read_new_run_log()
            if updated:
                self.append_log(self._render_run_log(), reset=True)
        except Exception:
            pass

        if self.run_process is None:
            self.run_tail_timer.stop()
            return
        if self.run_process.poll() is None:
            return

        # process ended
        try:
            if self._read_new_run_log():
                self.append_log(self._render_run_log(), reset=True)
        except Exception:
            pass

        if not self.run_exit_announced:
            exit_code = self.run_process.poll()
            bind_port = backend._extract_bind_fail_port(self.run_tail_buffer)
            has_bind_error = bind_port > 0 or "address already in use" in self.run_tail_buffer.lower()
            if exit_code not in (0, None) and has_bind_error and self.run_restart_count < self.run_max_restarts:
                self.run_restart_count += 1
                killed_pids = backend._kill_listeners_on_port(bind_port) if bind_port > 0 else []
                self.set_status("Bind error detected; relaunching...")
                self.run_exit_announced = True
                self._launch_run_process(
                    is_retry=True,
                    retry_reason=(f"bind failure on {bind_port}" if bind_port > 0 else "address already in use"),
                    killed_pids=killed_pids,
                )
                return

            self.run_exit_announced = True
            if exit_code == 0:
                self.run_log_header_lines.append("Program exited (code 0).")
                self.set_status("Program exited (code 0).")
            else:
                self.run_log_header_lines.append(f"Program exited with errors (code {exit_code}).")
                self.set_status(f"Program exited with errors (code {exit_code}).")
            self.append_log(self._render_run_log(), reset=True)

        self.run_tail_timer.stop()
        self.run_process = None

    def on_run(self) -> None:
        if not self.project_path.exists():
            self.set_status(f"Project not found: {self.project_path}")
            return
        if self.run_process is not None and self.run_process.poll() is None:
            self.set_status("Program is already running.")
            return
        self.run_restart_count = 0
        self._launch_run_process(is_retry=False)

    def on_validate_in_browser(self) -> None:
        target_url = (self.run_url_announced or self.run_browser.opened_url).strip()
        if not target_url:
            self.set_status("Run the app first so a local browser URL is available.")
            return

        self.set_status(f"Validating in browser: {target_url}")

        def work():
            return backend.run_browser_validation(target_url, headless=False)

        def success(result):
            summary = result.to_text()
            self.append_log(summary, reset=True)
            if result.log_path:
                self.set_status(f"Browser validation passed. Log: {Path(result.log_path).name}")
            else:
                self.set_status("Browser validation passed.")

        def error(exc: Exception):
            result = getattr(exc, "result", None)
            if result is not None:
                detail = result.to_text()
            else:
                detail = f"Browser validation failed.\n{exc}\n"
            self.append_log(detail, reset=True)
            self.set_status(f"Browser validation failed: {exc}")

        self.run_async("Browser Validation", work, success, error)

    def _stop_run_process(self) -> None:
        if self.run_tail_timer.isActive():
            self.run_tail_timer.stop()
        if self.run_process is not None:
            proc = self.run_process
            self.run_process = None
            backend._terminate_process(proc)

    def on_search(self) -> None:
        dlg = SearchDialog(self)
        dlg.pattern_input.setText("")
        dlg.exec()

    def open_framework_message(self) -> None:
        QMessageBox.information(self, "UI Framework", f"Running with Qt ({QT_BINDING}).")

    def closeEvent(self, event) -> None:  # noqa: N802
        self._stop_run_process()
        super().closeEvent(event)


def _ensure_api_keys() -> None:
    """Set fallback API keys if not already present in the environment."""
    if "OPENAI_API_KEY" not in os.environ:
        os.environ["OPENAI_API_KEY"] = "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"
    if "ANTHROPIC_API_KEY" not in os.environ:
        os.environ["ANTHROPIC_API_KEY"] = "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"


def main() -> None:
    _install_exception_logging()
    _ensure_api_keys()
    if len(sys.argv) > 1:
        print(
            "[qt-shell] CLI parameters are ignored. Using settings from migrator_config.json.",
            file=sys.stderr,
        )
    config = _load_app_config()
    prefer_hfest = bool(config.get("prefer_hfest", True))

    projects_root = Path(str(config.get("projects_root") or _default_app_config()["projects_root"])).expanduser().resolve()
    active_project_dir = Path(str(config.get("project_dir") or "")).expanduser().resolve()
    if not active_project_dir.exists():
        preferred = (projects_root / "PrescionBuilder").resolve()
        if preferred.exists():
            active_project_dir = preferred
        else:
            root_dir = Path(str(config.get("root") or "")).expanduser().resolve()
            if root_dir.exists():
                seed_target_dir: Path | None = None
                output_dir_raw = str(config.get("output_dir") or "").strip()
                if output_dir_raw:
                    target_candidate = Path(output_dir_raw).expanduser().resolve()
                    if target_candidate.exists():
                        seed_target_dir = target_candidate
                active_project_dir = backend.ensure_default_project_package(
                    projects_root=projects_root,
                    default_name=str(config.get("project_name") or "PrescionBuilder"),
                    vb6_source_root=root_dir,
                    output_project=str(config.get("output_project") or ""),
                    seed_target_dir=seed_target_dir,
                )
            else:
                raise SystemExit(f"Configured project not found: {active_project_dir}")

    runtime = backend.load_project_runtime(active_project_dir, prefer_hfest=prefer_hfest)

    config["project_dir"] = str(runtime["layout"]["project_dir"])
    config["project_name"] = runtime["layout"]["name"]
    _save_app_config(config)

    if not bool(config.get("ui", True)):
        print(runtime["csv_path"])
        print(runtime["mmd_path"])
        print(runtime["text_path"])
        return

    project_output_csproj = runtime["layout"]["output_project"] or str(config.get("output_project") or "")
    ui_defaults = {
        "root": str(runtime["layout"]["source_dir"]),
        "project": project_output_csproj,
        "output_dir": str(runtime["layout"].get("target_origin") or runtime["layout"]["target_dir"]),
        "project_name": runtime["layout"]["name"],
        "project_dir": str(runtime["layout"]["project_dir"]),
        "projects_root": str(projects_root),
        "context_dir": str(config.get("context_dir") or _default_app_config()["context_dir"]),
        "instructions": str(config.get("instructions") or _default_app_config()["instructions"]),
        "model": str(config.get("model") or "gpt-5.2-codex"),
        "temperature": str(config.get("temperature") or "0.2"),
        "prefer_hfest": "true" if prefer_hfest else "false",
        "vb6_image_dir": str(config.get("vb6_image_dir") or _default_app_config()["vb6_image_dir"]),
        "dotnet_image_dir": str(config.get("dotnet_image_dir") or _default_app_config()["dotnet_image_dir"]),
    }

    app = QApplication(sys.argv)
    win = QtMigratorWindow(runtime, ui_defaults, config)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
