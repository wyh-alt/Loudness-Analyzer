"""
音频响度标准化工具 - 轻量 PyQt6 界面
拖入音频文件/文件夹 -> 自动开始 -> 表格逐首显示响度统计 -> 自动导出 Excel
界面颜色会跟随系统的浅色/深色主题自动切换。
"""

import os
import sys
import ctypes
import tempfile
from pathlib import Path
from datetime import datetime

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QRect
from PyQt6.QtGui import (
    QIcon, QColor, QShortcut, QKeySequence, QPalette, QPen, QPixmap, QPainter,
    QBrush, QFont,
)
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QProgressBar, QFileDialog, QMessageBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QStyledItemDelegate, QStyle, QSplashScreen, QCheckBox, QDoubleSpinBox,
    QComboBox,
)

from core import (
    analyze_file, scan_folder, write_excel, find_tool, build_table_row,
    process_file, CancelToken, ProcCancelled,
    FORMAT_LOSSLESS, FORMAT_LOSSY, TABLE_HEADERS,
)

# PyInstaller onefile 会把资源解压到 sys._MEIPASS；开发模式下按脚本目录取即可
if getattr(sys, "frozen", False):
    APP_DIR = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
else:
    APP_DIR = Path(__file__).resolve().parent
ICON_PATH = APP_DIR / "icon.ico"

# 各数据列的默认宽度，按内容长度估算
DEFAULT_COL_WIDTHS = [205, 90, 80, 55, 80, 90, 90, 90, 90]

# 启动图尺寸/配色：与 make_splash.py 生成的 bootloader PNG 视觉一致
SPLASH_W, SPLASH_H = 360, 148
SPLASH_BG = "#2b2d31"
SPLASH_BORDER = "#43454b"
SPLASH_TITLE_COLOR = "#e6e8ec"
SPLASH_SUB_COLOR = "#a0a4ac"
SPLASH_TITLE = "音频响度标准化工具"

_LIGHT = dict(
    bg="#f5f6f8", surface="#ffffff", border="#d7dbe3",
    text="#1f2430", text_secondary="#6b7280", text_hint="#a3a8b3",
    header_bg="#f0f2f5", header_text="#4b5563", alt_row="#f7f8fa",
    btn_bg="#eef1f6", btn_text="#374151", btn_hover="#e2e6ee",
    progress_bg="#e5e7eb",
    accent="#4c7cf0", accent_hover="#3f6ce0",
    accent_disabled_bg="#c3d0f5", accent_disabled_text="#f0f3ff",
    cancel_hover="#eceef2", cancel_disabled_text="#c3c8d1", cancel_disabled_border="#e5e7eb",
)

_DARK = dict(
    bg="#202124", surface="#2b2d31", border="#43454b",
    text="#e6e8ec", text_secondary="#a0a4ac", text_hint="#75787f",
    header_bg="#313338", header_text="#c7cad1", alt_row="#302f34",
    btn_bg="#3a3c42", btn_text="#e6e8ec", btn_hover="#46484f",
    progress_bg="#3a3c42",
    accent="#5b8cff", accent_hover="#6f99ff",
    accent_disabled_bg="#33406b", accent_disabled_text="#8a93ad",
    cancel_hover="#3a3c42", cancel_disabled_text="#5b5d63", cancel_disabled_border="#43454b",
)


_SVG_ASSET_DIR = Path(tempfile.gettempdir()) / "loudness_analyzer_assets"


def _dump_svg(name: str, svg: str) -> str:
    """把 SVG 落到临时目录并返回可用于 QSS `url(...)` 的正斜杠路径。
    直接内联 data URI 时 Qt QSS 解析不稳定，用文件路径最稳。"""
    _SVG_ASSET_DIR.mkdir(parents=True, exist_ok=True)
    fpath = _SVG_ASSET_DIR / name
    if not fpath.exists() or fpath.read_text(encoding="utf-8") != svg:
        fpath.write_text(svg, encoding="utf-8")
    return str(fpath).replace("\\", "/")


def _arrow_svg(color: str, up: bool) -> str:
    """chevron 箭头（细线折角），比默认的实心小三角更清晰。"""
    if up:
        path = "M2 8L6 4L10 8"
    else:
        path = "M2 4L6 8L10 4"
    svg = (
        f"<svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' "
        f"viewBox='0 0 12 12'>"
        f"<path d='{path}' fill='none' stroke='{color}' "
        f"stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'/>"
        f"</svg>"
    )
    color_slug = color.lstrip("#")
    return _dump_svg(f"arrow_{'up' if up else 'down'}_{color_slug}.svg", svg)


def _check_svg(color: str) -> str:
    """勾选图标"""
    svg = (
        f"<svg xmlns='http://www.w3.org/2000/svg' width='14' height='14' "
        f"viewBox='0 0 14 14'>"
        f"<path d='M3 7.5L6 10.5L11 4.5' fill='none' stroke='{color}' "
        f"stroke-width='2' stroke-linecap='round' stroke-linejoin='round'/>"
        f"</svg>"
    )
    color_slug = color.lstrip("#")
    return _dump_svg(f"check_{color_slug}.svg", svg)


def build_stylesheet(dark: bool) -> str:
    c = _DARK if dark else _LIGHT
    return f"""
QWidget#root {{
    background: {c['bg']};
    color: {c['text']};
    font-family: "Microsoft YaHei UI", "Segoe UI";
}}
QLabel {{
    color: {c['text']};
    background: transparent;
}}
QLineEdit#pathEdit {{
    background: {c['surface']};
    color: {c['text']};
    border: 1px solid {c['border']};
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 13px;
}}
QLineEdit#pathEdit:focus {{
    border: 1px solid {c['accent']};
}}
QLabel#hintLabel {{
    color: {c['text_hint']};
    font-size: 11px;
}}
QLabel#statusLabel {{
    color: {c['text_secondary']};
    font-size: 12px;
}}
QTableWidget {{
    background: {c['surface']};
    color: {c['text']};
    border: 1px solid {c['border']};
    border-radius: 6px;
    gridline-color: {c['border']};
    font-size: 12px;
    alternate-background-color: {c['alt_row']};
    outline: none;
}}
QTableWidget::item {{
    padding: 3px;
    color: {c['text']};
    border: 0px;
    outline: 0;
}}
QTableWidget::item:selected, QTableWidget::item:focus {{
    background: transparent;
    color: {c['text']};
    outline: 0;
    border: 0px;
}}
QHeaderView::section {{
    background: {c['header_bg']};
    color: {c['header_text']};
    border: none;
    border-bottom: 1px solid {c['border']};
    padding: 5px;
    font-weight: 600;
}}
QPushButton#browseBtn {{
    background: {c['btn_bg']};
    color: {c['btn_text']};
    border: 1px solid {c['border']};
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 13px;
}}
QPushButton#browseBtn:hover {{
    background: {c['btn_hover']};
}}
QPushButton#startBtn {{
    background: {c['accent']};
    color: #ffffff;
    border: none;
    border-radius: 6px;
    padding: 7px 20px;
    font-size: 13px;
    font-weight: 600;
}}
QPushButton#startBtn:hover {{
    background: {c['accent_hover']};
}}
QPushButton#startBtn:disabled {{
    background: {c['accent_disabled_bg']};
    color: {c['accent_disabled_text']};
}}
QPushButton#cancelBtn {{
    background: transparent;
    color: {c['text_secondary']};
    border: 1px solid {c['border']};
    border-radius: 6px;
    padding: 7px 16px;
    font-size: 13px;
}}
QPushButton#cancelBtn:hover:!disabled {{
    background: {c['cancel_hover']};
}}
QPushButton#cancelBtn:disabled {{
    color: {c['cancel_disabled_text']};
    border-color: {c['cancel_disabled_border']};
}}
QProgressBar {{
    border: none;
    border-radius: 3px;
    background: {c['progress_bg']};
    max-height: 6px;
}}
QProgressBar::chunk {{
    background: {c['accent']};
    border-radius: 3px;
}}
QMessageBox {{
    background: {c['surface']};
    color: {c['text']};
}}
QCheckBox#normalizeCheck {{
    color: {c['text']};
    spacing: 6px;
    font-size: 13px;
}}
QCheckBox#normalizeCheck::indicator {{
    width: 14px;
    height: 14px;
    border: 1px solid {c['border']};
    border-radius: 3px;
    background: {c['surface']};
}}
QCheckBox#normalizeCheck::indicator:hover {{
    border-color: {c['accent']};
}}
QCheckBox#normalizeCheck::indicator:checked {{
    background: {c['accent']};
    border-color: {c['accent']};
    image: url({_check_svg('#ffffff')});
}}
QWidget#normParams QLabel {{
    color: {c['text_secondary']};
    font-size: 12px;
}}
QDoubleSpinBox#normSpin {{
    background: {c['surface']};
    color: {c['text']};
    border: 1px solid {c['border']};
    border-radius: 4px;
    padding-left: 6px;
    padding-right: 14px;
    font-size: 12px;
}}
QDoubleSpinBox#normSpin:focus {{
    border-color: {c['accent']};
}}
QDoubleSpinBox#normSpin:disabled {{
    color: {c['text_hint']};
}}
QDoubleSpinBox#normSpin::up-button {{
    subcontrol-origin: border;
    subcontrol-position: top right;
    width: 12px;
    border: none;
    background: transparent;
    margin: 1px 1px 0px 0px;
}}
QDoubleSpinBox#normSpin::down-button {{
    subcontrol-origin: border;
    subcontrol-position: bottom right;
    width: 12px;
    border: none;
    background: transparent;
    margin: 0px 1px 1px 0px;
}}
QDoubleSpinBox#normSpin::up-button:hover,
QDoubleSpinBox#normSpin::down-button:hover {{
    background: {c['btn_hover']};
    border-radius: 2px;
}}
QDoubleSpinBox#normSpin::up-button:pressed,
QDoubleSpinBox#normSpin::down-button:pressed {{
    background: {c['accent']};
    border-radius: 2px;
}}
QDoubleSpinBox#normSpin::up-arrow {{
    image: url({_arrow_svg(c['text'], up=True)});
    width: 10px;
    height: 8px;
}}
QDoubleSpinBox#normSpin::down-arrow {{
    image: url({_arrow_svg(c['text'], up=False)});
    width: 10px;
    height: 8px;
}}
QDoubleSpinBox#normSpin::up-arrow:pressed {{
    image: url({_arrow_svg('#ffffff', up=True)});
}}
QDoubleSpinBox#normSpin::down-arrow:pressed {{
    image: url({_arrow_svg('#ffffff', up=False)});
}}
QComboBox#normCombo {{
    background: {c['surface']};
    color: {c['text']};
    border: 1px solid {c['border']};
    border-radius: 4px;
    padding-left: 6px;
    padding-right: 4px;
    font-size: 12px;
}}
QComboBox#normCombo:focus, QComboBox#normCombo:on {{
    border-color: {c['accent']};
}}
QComboBox#normCombo:disabled {{
    color: {c['text_hint']};
}}
QComboBox#normCombo::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: center right;
    width: 16px;
    border: none;
    background: transparent;
}}
QComboBox#normCombo::down-arrow {{
    image: url({_arrow_svg(c['text'], up=False)});
    width: 10px;
    height: 8px;
}}
QComboBox#normCombo::down-arrow:disabled {{
    image: url({_arrow_svg(c['text_hint'], up=False)});
}}
QComboBox#normCombo QAbstractItemView {{
    background: {c['surface']};
    color: {c['text']};
    border: 1px solid {c['border']};
    selection-background-color: {c['accent']};
    selection-color: #ffffff;
    outline: 0;
}}
"""


class DropLineEdit(QLineEdit):
    """支持拖入单个/多个音频文件或文件夹的路径输入框，拖入新的会替换旧路径。
    多路径无法在单行文本框里完整展示，因此内部另存一份 self.paths；
    界面上显示"首个路径 (共 N 项)"的摘要，用户手动编辑文本时自动回落为单路径模式。"""

    dropped = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("pathEdit")
        self.setAcceptDrops(True)
        self.setPlaceholderText("将音频文件或文件夹拖到这里（支持多选）")
        self.paths = []
        # 用户手动编辑（不是程序 setText）时，清空多路径记忆，回到单路径模式
        self.textEdited.connect(lambda _t: self.paths.clear())

    def set_paths(self, paths):
        self.paths = list(paths)
        if not self.paths:
            self.clear()
        elif len(self.paths) == 1:
            self.setText(self.paths[0])
        else:
            first = os.path.basename(self.paths[0].rstrip("/\\")) or self.paths[0]
            self.setText(f"{first} (共 {len(self.paths)} 项)")

    def current_paths(self):
        if self.paths:
            return list(self.paths)
        text = self.text().strip()
        return [text] if text else []

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, e):
        urls = e.mimeData().urls()
        paths = [u.toLocalFile() for u in urls if u.toLocalFile()]
        if paths:
            self.set_paths(paths)
            self.dropped.emit()


class SelectionBorderDelegate(QStyledItemDelegate):
    """Excel 风格：选中的单元格四周画一条稍粗的强调色边框。
    另外支持"当前处理行"整行高亮 —— 在 super().paint() 之前先铺一层半透明底色，
    这样样式表里 ::item 的 alpha 覆盖问题不会影响这里的绘制。"""

    def __init__(self, color, parent=None):
        super().__init__(parent)
        self._color = QColor(color)
        self._current_row = -1

    def set_color(self, color):
        self._color = QColor(color)

    def set_current_row(self, row):
        self._current_row = row

    def paint(self, painter, option, index):
        if self._current_row == index.row() and self._current_row >= 0:
            painter.save()
            bg = QColor(self._color)
            bg.setAlpha(110)
            painter.fillRect(option.rect, bg)
            painter.restore()
        super().paint(painter, option, index)
        if option.state & QStyle.StateFlag.State_Selected:
            painter.save()
            pen = QPen(self._color, 2)
            painter.setPen(pen)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            r = option.rect.adjusted(1, 1, -1, -1)
            painter.drawRect(r)
            painter.restore()


class ResultsTable(QTableWidget):
    """窗口宽度变化时所有列按同一比例缩放；用户手动调整某一列时，最右侧列自动吸收/让出空间；
    首次显示时若默认列宽合计超过可视宽度，则整体压缩到 viewport 内，避免右侧列被裁掉。
    也接受文件/文件夹拖入，转发给外部同一套加载流程。"""

    files_dropped = pyqtSignal(list)  # 拖入的本地路径列表

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._suppress_fit = False
        self._did_initial_fit = False
        self.setAcceptDrops(True)
        # 关掉表格内部的 InternalMove 拖拽，避免与外部文件拖入冲突
        self.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            super().dragEnterEvent(e)

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            super().dragMoveEvent(e)

    def dropEvent(self, e):
        urls = e.mimeData().urls() if e.mimeData().hasUrls() else []
        paths = [u.toLocalFile() for u in urls if u.toLocalFile()]
        if paths:
            e.acceptProposedAction()
            self.files_dropped.emit(paths)
        else:
            super().dropEvent(e)

    def showEvent(self, event):
        super().showEvent(event)
        if not self._did_initial_fit and self.columnCount() > 0:
            self._did_initial_fit = True
            self._fit_all_to_viewport()

    def _fit_all_to_viewport(self):
        header = self.horizontalHeader()
        n = self.columnCount()
        if n == 0:
            return
        total = sum(header.sectionSize(c) for c in range(n))
        viewport_w = self.viewport().width()
        if total <= 0 or viewport_w <= 0 or total == viewport_w:
            return
        min_w = header.minimumSectionSize()
        ratio = viewport_w / total
        self._suppress_fit = True
        for c in range(n):
            cur = header.sectionSize(c)
            header.resizeSection(c, max(min_w, int(cur * ratio)))
        # 取整误差补进最后一列，只是把右边界贴住 viewport，不再单独拉宽它
        diff = viewport_w - sum(header.sectionSize(c) for c in range(n))
        if diff != 0:
            last = n - 1
            header.resizeSection(last, max(min_w, header.sectionSize(last) + diff))
        self._suppress_fit = False

    def _fit_last_to_viewport(self):
        if self.columnCount() < 2:
            return
        header = self.horizontalHeader()
        last = self.columnCount() - 1
        others = sum(header.sectionSize(c) for c in range(self.columnCount()) if c != last)
        available = self.viewport().width() - others
        min_w = header.minimumSectionSize()
        self._suppress_fit = True
        header.resizeSection(last, max(min_w, available))
        self._suppress_fit = False

    def on_section_resized(self, idx, old, new):
        if self._suppress_fit:
            return
        if idx == self.columnCount() - 1:
            return
        self._fit_last_to_viewport()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.columnCount() < 2:
            return
        old_w = event.oldSize().width()
        new_w = event.size().width()
        if old_w <= 0 or new_w <= 0 or old_w == new_w:
            return
        header = self.horizontalHeader()
        ratio = new_w / old_w
        min_w = header.minimumSectionSize()
        self._suppress_fit = True
        # 所有列（含最后一列）都按同一比例缩放，缩窄窗口时响度范围列同步变窄
        for col in range(self.columnCount()):
            cur = header.sectionSize(col)
            header.resizeSection(col, max(min_w, round(cur * ratio)))
        self._suppress_fit = False
        # 缩放取整可能与 viewport 差几个像素，仅把误差补进最后一列，避免右侧留白/溢出
        diff = self.viewport().width() - sum(
            header.sectionSize(c) for c in range(self.columnCount())
        )
        if diff != 0:
            self._suppress_fit = True
            last = self.columnCount() - 1
            header.resizeSection(last, max(min_w, header.sectionSize(last) + diff))
            self._suppress_fit = False


class AnalyzeWorker(QThread):
    """给已列出的 rows 逐首跑 analyze_file，把响度指标回填到 row["data"]。
    取消时会立刻 kill 正在跑的 ffmpeg（通过 CancelToken），不等文件跑完。"""
    progress = pyqtSignal(int, int, str)
    row_updated = pyqtSignal(int, dict)
    finished_ok = pyqtSignal(int, int)  # total, error_count
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(self, rows, ffmpeg, ffprobe):
        super().__init__()
        self.rows = rows
        self.ffmpeg = ffmpeg
        self.ffprobe = ffprobe
        self._token = CancelToken()

    def cancel(self):
        self._token.cancel()

    def run(self):
        total = len(self.rows)
        errors = 0
        for idx, row in enumerate(self.rows):
            if self._token.cancelled:
                self.cancelled.emit()
                return
            path = Path(row["dir"]) / row["name"]
            self.progress.emit(idx + 1, total, str(path))
            try:
                data = analyze_file(self.ffmpeg, self.ffprobe, path, cancel_token=self._token)
                new_row = {"name": path.name, "dir": str(path.parent), "data": data}
            except ProcCancelled:
                self.cancelled.emit()
                return
            except Exception as e:
                errors += 1
                new_row = {"name": path.name, "dir": str(path.parent), "error": str(e)}
            self.row_updated.emit(idx, new_row)
        self.finished_ok.emit(total, errors)


class ProcessWorker(QThread):
    """逐首另存处理：可选响度标准化 + 可选格式转换。
    源文件不会被改动，处理产物写到 out_dir，处理后重扫产物拿到新指标回填表格。"""
    progress = pyqtSignal(int, int, str)
    row_updated = pyqtSignal(int, dict)
    # out_dir, total, processed, skipped, errors
    finished_ok = pyqtSignal(str, int, int, int, int)
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(
        self, rows, out_dir, ffmpeg, ffprobe,
        *, normalize_loudness, target_i, target_tp, tolerance_lu,
        format_config,
    ):
        super().__init__()
        self.rows = rows
        self.out_dir = out_dir
        self.ffmpeg = ffmpeg
        self.ffprobe = ffprobe
        self.normalize_loudness = normalize_loudness
        self.target_i = target_i
        self.target_tp = target_tp
        self.tolerance_lu = tolerance_lu
        self.format_config = format_config
        self._token = CancelToken()

    def cancel(self):
        self._token.cancel()

    def run(self):
        total = len(self.rows)
        processed = 0
        skipped = 0  # 符合要求、直接原字节复制到输出目录的
        errors = 0

        for idx, row in enumerate(self.rows):
            if self._token.cancelled:
                self.cancelled.emit()
                return

            src_path = Path(row["dir"]) / row["name"]
            self.progress.emit(idx + 1, total, str(src_path))

            data = row.get("data") or {}
            measured_lufs = data.get("lufs_i")
            measured_tp = data.get("true_peak_db")

            try:
                dst, loudnorm_applied, converted = process_file(
                    self.ffmpeg, self.ffprobe, src_path, self.out_dir,
                    normalize_loudness=self.normalize_loudness,
                    target_i=self.target_i,
                    target_tp=self.target_tp,
                    tolerance_lu=self.tolerance_lu,
                    measured_lufs=measured_lufs,
                    measured_tp=measured_tp,
                    format_config=self.format_config,
                    cancel_token=self._token,
                )
                if loudnorm_applied or converted:
                    processed += 1
                else:
                    # loudnorm 被容差跳过 + 无格式转换 = 直接 copy，视为"符合要求无需处理"
                    skipped += 1
                # 处理后的文件重新走一遍分析，把新指标回填表格
                new_data = analyze_file(self.ffmpeg, self.ffprobe, dst, cancel_token=self._token)
                new_row = {
                    "name": dst.name,
                    "dir": str(dst.parent),
                    "data": new_data,
                }
            except ProcCancelled:
                self.cancelled.emit()
                return
            except Exception as e:
                errors += 1
                new_row = {"name": src_path.name, "dir": str(src_path.parent), "error": str(e)}

            self.row_updated.emit(idx, new_row)

        self.finished_ok.emit(str(self.out_dir), total, processed, skipped, errors)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("root")
        self.setWindowTitle("音频响度标准化工具")
        self.resize(1000, 650)
        self.setMinimumSize(1000, 650)
        if ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(ICON_PATH)))

        self.ffmpeg = find_tool("ffmpeg")
        self.ffprobe = find_tool("ffprobe")
        self.worker = None
        self._accent_color = _LIGHT["accent"]  # 由 set_theme() 跟随主题更新，用于选中单元格边框

        # 表格数据：每行是 {name, dir} / {name, dir, data} / {name, dir, error}
        self._loaded_rows = []
        self._highlighted_row = -1  # 处理中高亮的行索引

        layout = QVBoxLayout(self)
        # 底部边距缩小，让按钮行更贴近窗口下沿
        layout.setContentsMargins(20, 20, 20, 10)
        layout.setSpacing(10)

        path_row = QHBoxLayout()
        self.path_edit = DropLineEdit(self)
        self.browse_btn = QPushButton("浏览...", self)
        self.browse_btn.setObjectName("browseBtn")
        path_row.addWidget(self.path_edit, 1)
        path_row.addWidget(self.browse_btn)
        layout.addLayout(path_row)

        hint = QLabel("支持 WAV / MP3 / FLAC / M4A 等常见音频格式", self)
        hint.setObjectName("hintLabel")
        layout.addWidget(hint)

        self.table = ResultsTable(0, len(TABLE_HEADERS), self)
        self.table.setHorizontalHeaderLabels(TABLE_HEADERS)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        # Excel 风格：按单元格选中，支持鼠标拖拽出矩形选区
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        # Qt 默认会给"当前单元格"画一个虚线焦点框，样式表压不住，直接关掉表格的键盘焦点；
        # Ctrl+A/Ctrl+C 走 WindowShortcut，不受影响
        self.table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        # 默认按"整格"滚动，横向拖会一格一跳看着卡；改成按像素滚动就顺滑了
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self._selection_delegate = SelectionBorderDelegate(self._accent_color, self.table)
        self.table.setItemDelegate(self._selection_delegate)
        header = self.table.horizontalHeader()
        header.setMinimumSectionSize(40)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        for col, width in enumerate(DEFAULT_COL_WIDTHS):
            header.resizeSection(col, width)
        header.sectionResized.connect(self.table.on_section_resized)
        layout.addWidget(self.table, 1)

        # 表格设了 NoFocus，绑在 self.table 上的快捷键不会触发；挂到窗口上，靠焦点分发
        copy_sc = QShortcut(QKeySequence.StandardKey.Copy, self, activated=self.on_copy_shortcut)
        copy_sc.setContext(Qt.ShortcutContext.WindowShortcut)
        selectall_sc = QShortcut(QKeySequence.StandardKey.SelectAll, self, activated=self.on_select_all_shortcut)
        selectall_sc.setContext(Qt.ShortcutContext.WindowShortcut)

        # 进度条挪到"格式标准化"下方（下一段的末尾 addWidget）
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setTextVisible(False)

        SPIN_W = 100  # 数值+单位（"-12.0 LUFS"/"-1.0 dBTP"）+ 右侧上下按钮全部完整可见
        COMBO_W = 100

        # 响度标准化行：复选框 + 目标响度 / 容差 / 最高实际峰值
        norm_row = QHBoxLayout()
        norm_row.setSpacing(16)

        self.normalize_cb = QCheckBox("响度标准化", self)
        self.normalize_cb.setObjectName("normalizeCheck")
        self.normalize_cb.setChecked(True)
        norm_row.addWidget(self.normalize_cb)

        self.norm_params_widget = QWidget(self)
        self.norm_params_widget.setObjectName("normParams")
        params_row = QHBoxLayout(self.norm_params_widget)
        params_row.setContentsMargins(0, 0, 0, 0)
        params_row.setSpacing(6)

        params_row.addWidget(QLabel("目标响度:", self))
        self.target_i_spin = QDoubleSpinBox(self)
        self.target_i_spin.setObjectName("normSpin")
        self.target_i_spin.setDecimals(1)
        self.target_i_spin.setSingleStep(0.5)
        self.target_i_spin.setRange(-70.0, -5.0)
        self.target_i_spin.setValue(-12.0)
        self.target_i_spin.setSuffix(" LUFS")
        self.target_i_spin.setKeyboardTracking(False)
        self.target_i_spin.setFixedWidth(SPIN_W)
        params_row.addWidget(self.target_i_spin)

        params_row.addSpacing(16)
        params_row.addWidget(QLabel("容差:", self))
        self.tolerance_spin = QDoubleSpinBox(self)
        self.tolerance_spin.setObjectName("normSpin")
        self.tolerance_spin.setDecimals(1)
        self.tolerance_spin.setSingleStep(0.1)
        self.tolerance_spin.setRange(0.0, 20.0)
        self.tolerance_spin.setValue(1.0)
        self.tolerance_spin.setSuffix(" LU")
        self.tolerance_spin.setKeyboardTracking(False)
        self.tolerance_spin.setFixedWidth(SPIN_W)
        params_row.addWidget(self.tolerance_spin)

        params_row.addSpacing(16)
        params_row.addWidget(QLabel("最高实际峰值电平:", self))
        self.max_tp_spin = QDoubleSpinBox(self)
        self.max_tp_spin.setObjectName("normSpin")
        self.max_tp_spin.setDecimals(1)
        self.max_tp_spin.setSingleStep(0.5)
        self.max_tp_spin.setRange(-9.0, 0.0)
        self.max_tp_spin.setValue(-1.0)
        self.max_tp_spin.setSuffix(" dBTP")
        self.max_tp_spin.setKeyboardTracking(False)
        self.max_tp_spin.setFixedWidth(SPIN_W)
        params_row.addWidget(self.max_tp_spin)

        self.norm_params_widget.setEnabled(self.normalize_cb.isChecked())
        norm_row.addWidget(self.norm_params_widget)
        norm_row.addStretch(1)
        layout.addLayout(norm_row)

        # 格式标准化行：复选框 + 音频格式 / 采样率 / 位深(仅无损) 或 码率(仅有损) / 声道
        fmt_row = QHBoxLayout()
        fmt_row.setSpacing(16)

        self.format_cb = QCheckBox("格式标准化", self)
        self.format_cb.setObjectName("normalizeCheck")
        self.format_cb.setChecked(False)
        fmt_row.addWidget(self.format_cb)

        self.fmt_params_widget = QWidget(self)
        self.fmt_params_widget.setObjectName("normParams")
        fmt_params = QHBoxLayout(self.fmt_params_widget)
        fmt_params.setContentsMargins(0, 0, 0, 0)
        fmt_params.setSpacing(6)

        fmt_params.addWidget(QLabel("音频格式:", self))
        self.format_combo = QComboBox(self)
        self.format_combo.setObjectName("normCombo")
        for label, ext in [(".wav", ".wav"), (".mp3", ".mp3"), (".m4a", ".m4a"), (".flac", ".flac")]:
            self.format_combo.addItem(label, ext)
        self.format_combo.setFixedWidth(COMBO_W)
        fmt_params.addWidget(self.format_combo)

        fmt_params.addSpacing(16)
        fmt_params.addWidget(QLabel("采样率:", self))
        self.sr_combo = QComboBox(self)
        self.sr_combo.setObjectName("normCombo")
        for label, val in [("44100 Hz", 44100), ("48000 Hz", 48000)]:
            self.sr_combo.addItem(label, val)
        self.sr_combo.setFixedWidth(COMBO_W)
        fmt_params.addWidget(self.sr_combo)

        # 位深（仅无损时显示）
        fmt_params.addSpacing(16)
        self.bit_depth_label = QLabel("位深度:", self)
        fmt_params.addWidget(self.bit_depth_label)
        self.bit_depth_combo = QComboBox(self)
        self.bit_depth_combo.setObjectName("normCombo")
        for label, val in [("16 Bit", 16), ("24 Bit", 24), ("32 Bit", 32)]:
            self.bit_depth_combo.addItem(label, val)
        self.bit_depth_combo.setFixedWidth(COMBO_W)
        fmt_params.addWidget(self.bit_depth_combo)

        # 码率（仅有损时显示）
        self.bit_rate_label = QLabel("比特率:", self)
        fmt_params.addWidget(self.bit_rate_label)
        self.bit_rate_combo = QComboBox(self)
        self.bit_rate_combo.setObjectName("normCombo")
        for label, val in [("320 kbps", 320_000), ("256 kbps", 256_000),
                           ("192 kbps", 192_000), ("128 kbps", 128_000), ("64 kbps", 64_000)]:
            self.bit_rate_combo.addItem(label, val)
        self.bit_rate_combo.setFixedWidth(COMBO_W)
        fmt_params.addWidget(self.bit_rate_combo)

        fmt_params.addSpacing(16)
        fmt_params.addWidget(QLabel("声道:", self))
        self.channels_combo = QComboBox(self)
        self.channels_combo.setObjectName("normCombo")
        for label, val in [("立体声", 2), ("单声道", 1)]:
            self.channels_combo.addItem(label, val)
        self.channels_combo.setFixedWidth(COMBO_W)
        fmt_params.addWidget(self.channels_combo)

        self.fmt_params_widget.setEnabled(self.format_cb.isChecked())
        fmt_row.addWidget(self.fmt_params_widget)
        fmt_row.addStretch(1)
        layout.addLayout(fmt_row)

        # 进度条：格式标准化下方，按钮行上方
        layout.addWidget(self.progress_bar)

        # 状态文本与主按钮同一行，压缩纵向空间
        btn_row = QHBoxLayout()
        self.status_label = QLabel("就绪", self)
        self.status_label.setObjectName("statusLabel")
        btn_row.addWidget(self.status_label)
        btn_row.addStretch(1)
        # 导出与取消：次级按钮（灰底），响度统计与开始处理：主强调按钮（蓝底）
        self.export_btn = QPushButton("导出表格", self)
        self.export_btn.setObjectName("cancelBtn")
        self.export_btn.setEnabled(False)  # 只有检测/处理完成后才允许导出
        self.cancel_btn = QPushButton("取消", self)
        self.cancel_btn.setObjectName("cancelBtn")
        self.cancel_btn.setEnabled(False)
        self.measure_btn = QPushButton("响度统计", self)
        self.measure_btn.setObjectName("startBtn")
        self.measure_btn.setEnabled(False)
        self.start_btn = QPushButton("开始处理", self)
        self.start_btn.setObjectName("startBtn")
        self.start_btn.setEnabled(False)
        btn_row.addWidget(self.export_btn)
        btn_row.addWidget(self.cancel_btn)
        btn_row.addWidget(self.measure_btn)
        btn_row.addWidget(self.start_btn)
        layout.addLayout(btn_row)

        self.browse_btn.clicked.connect(self.on_browse)
        self.start_btn.clicked.connect(self.on_start)
        self.cancel_btn.clicked.connect(self.on_cancel)
        self.export_btn.clicked.connect(self.on_export)
        self.measure_btn.clicked.connect(self.on_measure)
        self.path_edit.dropped.connect(self.on_path_dropped)
        self.table.files_dropped.connect(self.on_table_dropped)
        self.normalize_cb.toggled.connect(self.on_normalize_toggled)
        self.format_cb.toggled.connect(self.on_format_toggled)
        self.format_combo.currentIndexChanged.connect(self._sync_format_specific_visibility)

        # 初始应用一次"位深/码率随格式切换"的联动
        self._sync_format_specific_visibility()

        if not self.ffmpeg or not self.ffprobe:
            self.start_btn.setEnabled(False)
            self.measure_btn.setEnabled(False)
            self.status_label.setText("缺少依赖：未检测到 ffmpeg / ffprobe，请安装后加入系统 PATH")

    def on_browse(self):
        folder = QFileDialog.getExistingDirectory(self, "选择包含音频文件的文件夹")
        if folder:
            self.path_edit.set_paths([folder])
            self.on_path_dropped()

    def on_table_dropped(self, paths):
        """表格区域拖入 —— 走跟拖到路径框一样的加载流程"""
        if self.worker and self.worker.isRunning():
            return
        self.path_edit.set_paths(paths)
        self.on_path_dropped()

    # ---------------- 参数区交互 ----------------

    def on_normalize_toggled(self, checked):
        self.norm_params_widget.setEnabled(checked)

    def on_format_toggled(self, checked):
        self.fmt_params_widget.setEnabled(checked)

    def _sync_format_specific_visibility(self, *_):
        """格式切换：无损（wav/flac）显位深、隐码率；有损（mp3/m4a）反之"""
        ext = self.format_combo.currentData()
        is_lossless = ext in FORMAT_LOSSLESS
        self.bit_depth_label.setVisible(is_lossless)
        self.bit_depth_combo.setVisible(is_lossless)
        self.bit_rate_label.setVisible(not is_lossless)
        self.bit_rate_combo.setVisible(not is_lossless)

    def _current_format_config(self):
        """把格式参数装成 process_file 用的 dict；未勾选时返回 None"""
        if not self.format_cb.isChecked():
            return None
        ext = self.format_combo.currentData()
        cfg = {
            "ext": ext,
            "sample_rate": self.sr_combo.currentData(),
            "channels": self.channels_combo.currentData(),
        }
        if ext in FORMAT_LOSSLESS:
            cfg["bit_depth"] = self.bit_depth_combo.currentData()
        else:
            cfg["bit_rate"] = self.bit_rate_combo.currentData()
        return cfg

    # ---------------- 拖入 / 响度统计 / 开始处理 ----------------

    def on_path_dropped(self):
        """拖入或点浏览：只把音频路径列进表格，不做检测"""
        if self.worker and self.worker.isRunning():
            return
        paths = self.path_edit.current_paths()
        if not paths:
            return
        missing = [p for p in paths if not os.path.exists(p)]
        if missing:
            QMessageBox.warning(self, "错误", "以下路径不存在：\n" + "\n".join(missing))
            return
        try:
            files = scan_folder(paths)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"扫描失败: {e}")
            return
        if not files:
            QMessageBox.warning(self, "提示", "未找到可识别的音频文件。")
            return

        self.table.setRowCount(0)
        self._loaded_rows = []
        for p in files:
            row = {"name": p.name, "dir": str(p.parent)}
            self._loaded_rows.append(row)
            r = self.table.rowCount()
            self.table.insertRow(r)
            self._fill_row(r, build_table_row(row))
        self.progress_bar.setValue(0)
        self.status_label.setText(
            f"已列出 {len(files)} 个文件；点「响度统计」检测响度，或直接「开始处理」（处理时自动检测）"
        )
        # 拖入阶段还没做检测，不允许导出（表格只有文件名）
        self.export_btn.setEnabled(False)
        self.measure_btn.setEnabled(bool(self.ffmpeg and self.ffprobe))
        self.start_btn.setEnabled(bool(self.ffmpeg and self.ffprobe))

    def on_measure(self):
        """响度统计：对当前 loaded_rows 逐首跑 analyze_file"""
        if self.worker and self.worker.isRunning():
            return
        if not self._loaded_rows:
            QMessageBox.warning(self, "提示", "请先拖入音频。")
            return
        self.progress_bar.setValue(0)
        self.measure_btn.setEnabled(False)
        self.start_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.status_label.setText("准备检测响度...")
        self.worker = AnalyzeWorker(list(self._loaded_rows), self.ffmpeg, self.ffprobe)
        self.worker.progress.connect(self.on_progress)
        self.worker.row_updated.connect(self.on_row_updated)
        self.worker.finished_ok.connect(self.on_measure_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.cancelled.connect(self.on_cancelled)
        self.worker.start()

    def on_start(self):
        """开始处理：弹目录选择框 → 逐首 process_file 另存到该目录"""
        if self.worker and self.worker.isRunning():
            return
        if not self._loaded_rows:
            QMessageBox.warning(self, "提示", "请先拖入音频。")
            return
        do_loudness = self.normalize_cb.isChecked()
        format_cfg = self._current_format_config()
        if not do_loudness and not format_cfg:
            QMessageBox.information(
                self, "提示",
                "响度标准化和格式标准化都未启用，音频无需处理。",
            )
            return

        # 弹出输出目录选择
        first = self.path_edit.current_paths()[0] if self.path_edit.current_paths() else self._loaded_rows[0]["dir"]
        default_dir = first if os.path.isdir(first) else os.path.dirname(first)
        out_dir = QFileDialog.getExistingDirectory(
            self, "选择处理后音频的输出目录", default_dir,
        )
        if not out_dir:
            return

        self.progress_bar.setValue(0)
        self.start_btn.setEnabled(False)
        self.measure_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.status_label.setText("准备处理...")
        self.worker = ProcessWorker(
            list(self._loaded_rows), out_dir, self.ffmpeg, self.ffprobe,
            normalize_loudness=do_loudness,
            target_i=self.target_i_spin.value(),
            target_tp=self.max_tp_spin.value(),
            tolerance_lu=self.tolerance_spin.value(),
            format_config=format_cfg,
        )
        self.worker.progress.connect(self.on_progress)
        self.worker.row_updated.connect(self.on_row_updated)
        self.worker.finished_ok.connect(self.on_process_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.cancelled.connect(self.on_cancelled)
        self.worker.start()

    def _ask_excel_path(self, first_path):
        default_dir = first_path if os.path.isdir(first_path) else os.path.dirname(first_path)
        default_name = f"音频响度统计_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out_path, _ = QFileDialog.getSaveFileName(
            self, "选择 Excel 保存位置",
            os.path.join(default_dir, default_name),
            "Excel 文件 (*.xlsx)",
        )
        return out_path or None

    def on_cancel(self):
        if self.worker:
            self.worker.cancel()
        self.cancel_btn.setEnabled(False)
        self.status_label.setText("正在取消...")

    def on_progress(self, idx, total, name):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(idx)
        self.status_label.setText(f"正在处理 ({idx}/{total})：{name}")
        # 高亮当前行 + 滚动到可视区，便于用户直观追踪进度
        self._highlight_row(idx - 1)

    def _highlight_row(self, row_idx):
        """通过 delegate 直接绘制整行高亮 + 滚动到该行。row_idx=-1 清除高亮。
        走 delegate.paint() 是因为 QSS 里 ::item 会吞掉 setBackground 的 alpha。"""
        if row_idx is None or row_idx < 0 or row_idx >= self.table.rowCount():
            self._selection_delegate.set_current_row(-1)
            self._highlighted_row = -1
            self.table.viewport().update()
            return
        self._selection_delegate.set_current_row(row_idx)
        self._highlighted_row = row_idx
        self.table.viewport().update()
        anchor = self.table.item(row_idx, 0)
        if anchor is not None:
            self.table.scrollToItem(anchor, QAbstractItemView.ScrollHint.PositionAtCenter)

    def on_copy_shortcut(self):
        """焦点在路径输入框时复制路径文本；否则复制表格里已选中的单元格。
        无选中就什么也不做，避免把无关内容塞进剪贴板。"""
        fw = QApplication.focusWidget()
        if isinstance(fw, QLineEdit):
            fw.copy()
            return
        indexes = self.table.selectedIndexes()
        if not indexes:
            return
        rows = sorted(set(i.row() for i in indexes))
        cols = sorted(set(i.column() for i in indexes))
        lines = []
        for r in rows:
            cells = [self.table.item(r, c).text() if self.table.item(r, c) else "" for c in cols]
            lines.append("\t".join(cells))
        QApplication.clipboard().setText("\n".join(lines))

    def on_select_all_shortcut(self):
        fw = QApplication.focusWidget()
        if isinstance(fw, QLineEdit):
            fw.selectAll()
            return
        self.table.selectAll()

    def _fill_row(self, r, cells):
        for col, value in enumerate(cells):
            item = QTableWidgetItem(str(value))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(r, col, item)

    def on_row_updated(self, idx, row):
        if 0 <= idx < len(self._loaded_rows):
            self._loaded_rows[idx] = row
        if 0 <= idx < self.table.rowCount():
            self._fill_row(idx, build_table_row(row))

    def _reset_buttons(self):
        has_rows = bool(self._loaded_rows)
        deps_ok = bool(self.ffmpeg and self.ffprobe)
        # 只要有可导出的响度数据就启用导出（检测完 or 处理完的行）
        has_data = any(r.get("data") or r.get("error") for r in self._loaded_rows)
        self.start_btn.setEnabled(has_rows and deps_ok)
        self.measure_btn.setEnabled(has_rows and deps_ok)
        self.cancel_btn.setEnabled(False)
        self.export_btn.setEnabled(has_data)
        # 一轮结束，清除进度高亮
        self._highlight_row(-1)

    def on_measure_finished(self, total, error_count):
        self._reset_buttons()
        self.status_label.setText(
            f"响度检测完成：共 {total} 个文件，{error_count} 个失败"
        )

    def on_process_finished(self, out_dir, total, processed, skipped, error_count):
        self._reset_buttons()
        self.status_label.setText(
            f"处理完成：共 {total}，实际处理 {processed}，符合要求无需处理 {skipped}，失败 {error_count}"
        )
        reply = QMessageBox.question(
            self, "处理完成",
            f"共 {total} 个文件\n"
            f"  · 实际处理：{processed}\n"
            f"  · 符合要求无需处理：{skipped}\n"
            f"  · 失败：{error_count}\n\n"
            f"处理后的音频已保存到：\n{out_dir}\n\n是否打开该目录？",
        )
        if reply == QMessageBox.StandardButton.Yes:
            os.startfile(out_dir)

    def on_export(self):
        if not self._loaded_rows:
            QMessageBox.warning(self, "提示", "当前表格为空，请先拖入音频。")
            return
        with_data = [r for r in self._loaded_rows if r.get("data") or r.get("error")]
        if not with_data:
            QMessageBox.warning(
                self, "提示",
                "当前表格里没有响度数据可导出。请先点「响度统计」检测，或直接「开始处理」。",
            )
            return
        paths = self.path_edit.current_paths()
        first = paths[0] if paths else self._loaded_rows[0]["dir"]
        out_path = self._ask_excel_path(first)
        if not out_path:
            return
        try:
            write_excel(with_data, out_path)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"写入 Excel 失败：{e}")
            return
        self.status_label.setText(f"已导出到 {out_path}")
        reply = QMessageBox.question(
            self, "导出完成",
            f"表格已保存到：\n{out_path}\n\n是否打开所在文件夹？",
        )
        if reply == QMessageBox.StandardButton.Yes:
            os.startfile(os.path.dirname(out_path))

    def on_failed(self, message):
        self._reset_buttons()
        self.status_label.setText("出错")
        QMessageBox.critical(self, "错误", message)

    def on_cancelled(self):
        self._reset_buttons()
        self.status_label.setText("已取消")

    def set_theme(self, dark: bool):
        self._accent_color = (_DARK if dark else _LIGHT)["accent"]
        if hasattr(self, "_selection_delegate"):
            self._selection_delegate.set_color(self._accent_color)
            self.table.viewport().update()


def apply_theme(app, window, scheme=None):
    scheme = scheme or app.styleHints().colorScheme()
    is_dark = scheme == Qt.ColorScheme.Dark
    app.setStyleSheet(build_stylesheet(is_dark))
    window.set_theme(is_dark)

    # 系统强调色（如橙色）会通过 QPalette.Highlight 在原生控件绘制里露出一点边角，
    # 样式表盖不住；把 Highlight 直接改成跟正常背景/文字一样的颜色，从根上消掉这个残留色块
    c = _DARK if is_dark else _LIGHT
    palette = app.palette()
    palette.setColor(QPalette.ColorRole.Highlight, QColor(c["surface"]))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(c["text"]))
    app.setPalette(palette)


def _make_splash_pixmap(subtitle_text: str) -> QPixmap:
    """按与 assets/splash.png 相同的规格重绘一份 Qt 用的启动图，只把副标题换掉。"""
    pix = QPixmap(SPLASH_W, SPLASH_H)
    pix.fill(Qt.GlobalColor.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    p.setBrush(QBrush(QColor(SPLASH_BG)))
    p.setPen(QPen(QColor(SPLASH_BORDER), 1))
    p.drawRoundedRect(0, 0, SPLASH_W - 1, SPLASH_H - 1, 10, 10)

    title_font = QFont("Microsoft YaHei UI", 14)
    title_font.setBold(True)
    p.setFont(title_font)
    p.setPen(QColor(SPLASH_TITLE_COLOR))
    p.drawText(QRect(0, 38, SPLASH_W, 32), Qt.AlignmentFlag.AlignHCenter, SPLASH_TITLE)

    sub_font = QFont("Microsoft YaHei UI", 9)
    p.setFont(sub_font)
    p.setPen(QColor(SPLASH_SUB_COLOR))
    p.drawText(QRect(0, 86, SPLASH_W, 22), Qt.AlignmentFlag.AlignHCenter, subtitle_text)
    p.end()
    return pix


def _close_pyi_splash():
    """打包模式下关掉 bootloader splash；开发模式下 import 失败直接跳过。"""
    try:
        import pyi_splash  # type: ignore[import-not-found]
    except ImportError:
        return
    try:
        if pyi_splash.is_alive():
            pyi_splash.close()
    except Exception:
        pass


def main():
    try:
        myappid = "audiotools.loudnessstats.batch.v1"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    # Windows 原生 style 的表格选中态会露出系统强调色的原生焦点框，样式表压不住；
    # 换成 Qt 自带的 Fusion style，完全按样式表/调色板绘制，不会有原生主题的残留视觉
    app.setStyle("Fusion")
    if ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(ICON_PATH)))

    # Qt splash 立刻上屏，与打包模式下的 bootloader splash 做无缝交接；开发模式下作为唯一 splash
    qt_splash = QSplashScreen(
        _make_splash_pixmap("正在加载界面…"),
        Qt.WindowType.WindowStaysOnTopHint,
    )
    qt_splash.show()
    app.processEvents()
    _close_pyi_splash()

    window = MainWindow()
    apply_theme(app, window)
    app.styleHints().colorSchemeChanged.connect(lambda scheme: apply_theme(app, window, scheme))

    window.show()
    qt_splash.finish(window)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
