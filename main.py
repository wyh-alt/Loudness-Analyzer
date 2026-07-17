"""
音频响度标准化工具 - 轻量 PyQt6 界面
拖入音频文件/文件夹 -> 自动开始 -> 表格逐首显示响度统计 -> 自动导出 Excel
界面颜色会跟随系统的浅色/深色主题自动切换。
"""

import os
import sys
import ctypes
import shutil
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
)

from core import (
    analyze_file, scan_folder, write_excel, find_tool, build_table_row,
    normalize_file, restore_backup, TABLE_HEADERS,
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
    """Excel 风格：选中的单元格四周画一条稍粗的强调色边框，其它单元格照常绘制。"""

    def __init__(self, color, parent=None):
        super().__init__(parent)
        self._color = QColor(color)

    def set_color(self, color):
        self._color = QColor(color)

    def paint(self, painter, option, index):
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
    首次显示时若默认列宽合计超过可视宽度，则整体压缩到 viewport 内，避免右侧列被裁掉。"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._suppress_fit = False
        self._did_initial_fit = False

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


class ScanWorker(QThread):
    progress = pyqtSignal(int, int, str)
    row_ready = pyqtSignal(dict)
    finished_ok = pyqtSignal(str, int, int)
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(self, folders, out_path, ffmpeg, ffprobe):
        super().__init__()
        self.folders = folders
        # out_path 为 None 时仅扫描并显示到表格，不落 Excel（响度标准化模式的初始载入）
        self.out_path = out_path
        self.ffmpeg = ffmpeg
        self.ffprobe = ffprobe
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        try:
            files = scan_folder(self.folders)
        except Exception as e:
            self.failed.emit(f"扫描失败: {e}")
            return

        total = len(files)
        if total == 0:
            self.failed.emit("未找到可识别的音频文件，请检查选择的文件/文件夹。")
            return

        rows = []
        error_count = 0
        for idx, path in enumerate(files, start=1):
            if self._cancel:
                self.cancelled.emit()
                return
            self.progress.emit(idx, total, str(path))
            try:
                data = analyze_file(self.ffmpeg, self.ffprobe, path)
                row = {"name": path.name, "dir": str(path.parent), "data": data}
            except Exception as e:
                error_count += 1
                row = {"name": path.name, "dir": str(path.parent), "error": str(e)}
            rows.append(row)
            self.row_ready.emit(row)

        if self.out_path:
            try:
                write_excel(rows, self.out_path)
            except Exception as e:
                self.failed.emit(f"写入 Excel 失败: {e}")
                return

        self.finished_ok.emit(self.out_path or "", total, error_count)


class NormalizeWorker(QThread):
    progress = pyqtSignal(int, int, str)
    row_updated = pyqtSignal(int, dict)
    backup_registered = pyqtSignal(str, str)  # original_path, backup_path
    finished_ok = pyqtSignal(str, int, int, int)  # out_path, total, processed, errors
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(
        self, rows, out_path, ffmpeg, ffprobe,
        target_i, target_tp, tolerance_lu, backup_dir,
    ):
        super().__init__()
        self.rows = rows
        self.out_path = out_path
        self.ffmpeg = ffmpeg
        self.ffprobe = ffprobe
        self.target_i = target_i
        self.target_tp = target_tp
        self.tolerance_lu = tolerance_lu
        self.backup_dir = backup_dir
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        total = len(self.rows)
        processed = 0
        error_count = 0
        result_rows = []

        for idx, row in enumerate(self.rows):
            if self._cancel:
                self.cancelled.emit()
                return

            path = Path(row["dir"]) / row["name"]
            self.progress.emit(idx + 1, total, str(path))

            if row.get("error"):
                # 之前扫描就失败的直接保留原错误信息，不再尝试标准化
                result_rows.append(row)
                continue

            data = row.get("data", {}) or {}
            measured_lufs = data.get("lufs_i")
            measured_tp = data.get("true_peak_db")
            try:
                backup, was_processed = normalize_file(
                    self.ffmpeg, self.ffprobe, path,
                    self.target_i, self.target_tp, self.tolerance_lu,
                    backup_dir=self.backup_dir,
                    measured_lufs=measured_lufs,
                    measured_tp=measured_tp,
                )
                self.backup_registered.emit(str(path), str(backup))
                if was_processed:
                    processed += 1
                    new_data = analyze_file(self.ffmpeg, self.ffprobe, path)
                    new_row = {"name": path.name, "dir": str(path.parent), "data": new_data}
                else:
                    new_row = row
            except Exception as e:
                error_count += 1
                new_row = {"name": path.name, "dir": str(path.parent), "error": str(e)}

            result_rows.append(new_row)
            self.row_updated.emit(idx, new_row)

        # xlsx 由主界面"导出表格"按钮触发，worker 不再直接落盘
        if self.out_path:
            try:
                write_excel(result_rows, self.out_path)
            except Exception as e:
                self.failed.emit(f"写入 Excel 失败: {e}")
                return

        self.finished_ok.emit(self.out_path or "", total, processed, error_count)


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

        # 响度标准化相关状态
        self._loaded_rows = []          # 与表格顺序一一对应的分析结果
        self._backups = {}              # 已备份的原文件 {原路径: 备份路径}
        self._backup_dir = None         # 本轮标准化用的临时备份目录
        self._norm_state = "idle"       # "idle" 或 "processed"（决定按钮是"开始处理"还是"撤销处理"）

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
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

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setTextVisible(False)
        layout.addWidget(self.progress_bar)

        # 响度标准化设置区（复选框 + 参数同排，节省纵向空间）
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

        SPIN_W = 100  # 数值+单位（"-12.0 LUFS"/"-1.0 dBTP"）+ 右侧上下按钮全部完整可见

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

        # 状态文本单独占一行（在标准化设置行上方），把纵向节奏拉开
        status_row = QHBoxLayout()
        self.status_label = QLabel("就绪", self)
        self.status_label.setObjectName("statusLabel")
        status_row.addWidget(self.status_label)
        status_row.addStretch(1)
        layout.addLayout(status_row)

        # 响度标准化行与"导出表格 / 取消 / 开始处理"按钮同排
        self.export_btn = QPushButton("导出表格", self)
        self.export_btn.setObjectName("browseBtn")
        self.export_btn.setEnabled(False)  # 表格空时禁用；有数据后自动启用
        self.cancel_btn = QPushButton("取消", self)
        self.cancel_btn.setObjectName("cancelBtn")
        self.cancel_btn.setEnabled(False)
        self.start_btn = QPushButton("开始处理", self)
        self.start_btn.setObjectName("startBtn")
        norm_row.addWidget(self.export_btn)
        norm_row.addWidget(self.cancel_btn)
        norm_row.addWidget(self.start_btn)
        layout.addLayout(norm_row)

        self.browse_btn.clicked.connect(self.on_browse)
        self.start_btn.clicked.connect(self.on_start)
        self.cancel_btn.clicked.connect(self.on_cancel)
        self.export_btn.clicked.connect(self.on_export)
        self.path_edit.dropped.connect(self.on_path_dropped)
        self.normalize_cb.toggled.connect(self.on_normalize_toggled)
        self.target_i_spin.valueChanged.connect(self.on_norm_param_changed)
        self.tolerance_spin.valueChanged.connect(self.on_norm_param_changed)
        self.max_tp_spin.valueChanged.connect(self.on_norm_param_changed)

        self._update_start_button()

        if not self.ffmpeg or not self.ffprobe:
            self.start_btn.setEnabled(False)
            self.status_label.setText("缺少依赖：未检测到 ffmpeg / ffprobe，请安装后加入系统 PATH")

    def on_browse(self):
        folder = QFileDialog.getExistingDirectory(self, "选择包含音频文件的文件夹")
        if folder:
            self.path_edit.set_paths([folder])

    # ---------------- 响度标准化设置区交互 ----------------

    def on_normalize_toggled(self, checked):
        # 参数区始终显示，仅根据是否勾选决定能否编辑（未勾选时整体置灰不可交互）
        self.norm_params_widget.setEnabled(checked)
        # 切换模式后按钮语义变化；不改变已有备份，避免用户误操作丢失撤销能力
        self._update_start_button()

    def on_norm_param_changed(self, _value):
        # 只有当已经完成标准化（存在可撤销状态）时，参数变更才需要重置回"开始处理"
        if self._norm_state == "processed":
            self._discard_backups()
            self._norm_state = "idle"
            self._update_start_button()
            self.status_label.setText("参数已调整，撤销状态已失效")

    def _update_start_button(self):
        if self.normalize_cb.isChecked() and self._norm_state == "processed":
            self.start_btn.setText("撤销处理")
        else:
            self.start_btn.setText("开始处理")

    def _discard_backups(self):
        """清空当前的备份记录并删除临时备份目录（不还原文件）"""
        self._backups.clear()
        if self._backup_dir and os.path.isdir(self._backup_dir):
            shutil.rmtree(self._backup_dir, ignore_errors=True)
        self._backup_dir = None

    def on_path_dropped(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.start_btn.isEnabled():
            return
        # 新素材载入意味着旧备份对应的路径可能不再是"当前处理集"，直接丢弃
        if self._backups:
            self._discard_backups()
        self._norm_state = "idle"
        self._update_start_button()

        # 拖入只做检测 + 表格显示；xlsx 由"导出表格"按钮触发，不再自动落盘
        self._start_scan(out_path=None)

    def on_start(self):
        # 撤销模式：还原备份并回到"开始处理"状态
        if self.normalize_cb.isChecked() and self._norm_state == "processed":
            self._start_undo()
            return
        # 标准化模式：对已载入的文件做批量标准化
        if self.normalize_cb.isChecked():
            self._start_normalize()
            return
        # 未勾选响度标准化：xlsx 在拖入时就已经落盘，按钮不再触发第二次生成
        QMessageBox.information(
            self, "提示",
            "当前未选择响度标准化，音频无需处理。",
        )

    def _ask_excel_path(self, first_path):
        default_dir = first_path if os.path.isdir(first_path) else os.path.dirname(first_path)
        default_name = f"音频响度统计_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out_path, _ = QFileDialog.getSaveFileName(
            self, "选择 Excel 保存位置",
            os.path.join(default_dir, default_name),
            "Excel 文件 (*.xlsx)",
        )
        return out_path or None

    def _start_scan(self, out_path):
        paths = self.path_edit.current_paths()
        if not paths:
            QMessageBox.warning(self, "提示", "请先拖入或选择一个音频文件/文件夹。")
            return
        missing = [p for p in paths if not os.path.exists(p)]
        if missing:
            QMessageBox.warning(self, "错误", "以下路径不存在：\n" + "\n".join(missing))
            return

        self.table.setRowCount(0)
        self._loaded_rows = []
        self.progress_bar.setValue(0)
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.export_btn.setEnabled(False)
        self.status_label.setText("准备扫描...")
        self.worker = ScanWorker(paths, out_path, self.ffmpeg, self.ffprobe)
        self.worker.progress.connect(self.on_progress)
        self.worker.row_ready.connect(self.on_row_ready)
        self.worker.finished_ok.connect(self.on_scan_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.cancelled.connect(self.on_cancelled)
        self.worker.start()

    def _start_normalize(self):
        if not self._loaded_rows:
            QMessageBox.warning(self, "提示", "请先拖入音频，等待载入分析完成后再开始处理。")
            return
        # 只统计能处理的行（错误行会跳过）
        processable = [r for r in self._loaded_rows if not r.get("error")]
        if not processable:
            QMessageBox.warning(self, "提示", "载入的音频都无法分析，无法进行标准化。")
            return

        # 每轮标准化用一个独立的临时目录存放备份，撤销后整个删掉
        self._backup_dir = tempfile.mkdtemp(prefix="loudness_norm_")
        self._backups = {}

        self.progress_bar.setValue(0)
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.export_btn.setEnabled(False)
        self.status_label.setText("准备标准化...")
        # xlsx 不再在处理阶段自动生成，交给"导出表格"按钮
        self.worker = NormalizeWorker(
            list(self._loaded_rows), None, self.ffmpeg, self.ffprobe,
            target_i=self.target_i_spin.value(),
            target_tp=self.max_tp_spin.value(),
            tolerance_lu=self.tolerance_spin.value(),
            backup_dir=self._backup_dir,
        )
        self.worker.progress.connect(self.on_progress)
        self.worker.row_updated.connect(self.on_row_updated)
        self.worker.backup_registered.connect(self.on_backup_registered)
        self.worker.finished_ok.connect(self.on_normalize_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.cancelled.connect(self.on_cancelled)
        self.worker.start()

    def _start_undo(self):
        if not self._backups:
            self._norm_state = "idle"
            self._update_start_button()
            return
        self.start_btn.setEnabled(False)
        self.status_label.setText("正在还原...")
        QApplication.processEvents()

        errors = []
        total = len(self._backups)
        for i, (original, backup) in enumerate(list(self._backups.items()), start=1):
            self.status_label.setText(f"正在还原 ({i}/{total})：{os.path.basename(original)}")
            QApplication.processEvents()
            try:
                if os.path.exists(backup):
                    restore_backup(backup, original)
            except Exception as e:
                errors.append(f"{original}: {e}")

        self._discard_backups()
        self._norm_state = "idle"
        # 还原后重新扫描一次，让表格里的响度指标回到原始值
        self.status_label.setText("还原完成，正在重新分析...")
        self._update_start_button()
        self.start_btn.setEnabled(True)
        if errors:
            QMessageBox.warning(
                self, "部分文件还原失败",
                "以下文件还原时出错：\n" + "\n".join(errors[:20]),
            )
        self._start_scan(out_path=None)

    def on_cancel(self):
        if self.worker:
            self.worker.cancel()
        self.cancel_btn.setEnabled(False)
        self.status_label.setText("正在取消...")

    def on_progress(self, idx, total, name):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(idx)
        self.status_label.setText(f"正在处理 ({idx}/{total})：{name}")

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

    def on_row_ready(self, row):
        self._loaded_rows.append(row)
        r = self.table.rowCount()
        self.table.insertRow(r)
        self._fill_row(r, build_table_row(row))
        self.table.scrollToBottom()

    def on_row_updated(self, idx, row):
        if 0 <= idx < len(self._loaded_rows):
            self._loaded_rows[idx] = row
        if 0 <= idx < self.table.rowCount():
            self._fill_row(idx, build_table_row(row))

    def on_backup_registered(self, original, backup):
        self._backups[original] = backup

    def _reset_buttons(self):
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        # 有可导出内容时启用导出按钮
        self.export_btn.setEnabled(bool(self._loaded_rows))
        self._update_start_button()

    def on_scan_finished(self, out_path, total, error_count):
        self._reset_buttons()
        self.status_label.setText(
            f"已载入 {total} 个文件，{error_count} 个失败；可点击「导出表格」保存 xlsx"
        )

    def on_normalize_finished(self, out_path, total, processed, error_count):
        # 只要有备份就允许撤销（即使全部因容差被跳过，撤销至少能把备份目录清掉）
        self._norm_state = "processed" if self._backups else "idle"
        self._reset_buttons()
        self.status_label.setText(
            f"标准化完成：共 {total} 个文件，实际处理 {processed}，失败 {error_count}；"
            f"可点击「导出表格」保存新 xlsx"
        )

    def on_export(self):
        if not self._loaded_rows:
            QMessageBox.warning(self, "提示", "当前表格为空，请先拖入音频。")
            return
        paths = self.path_edit.current_paths()
        first = paths[0] if paths else self._loaded_rows[0]["dir"]
        out_path = self._ask_excel_path(first)
        if not out_path:
            return
        try:
            write_excel(list(self._loaded_rows), out_path)
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
        # 标准化过程中出错但已经产生备份的话，保留撤销能力（例如仅 Excel 写入失败的情况）
        if self._backups:
            self._norm_state = "processed"
        self._reset_buttons()
        self.status_label.setText("出错")
        QMessageBox.critical(self, "错误", message)

    def on_cancelled(self):
        # 用户取消标准化中途：已产生的备份先保留，允许通过撤销把已改过的文件还原
        if self._backups:
            self._norm_state = "processed"
        self._reset_buttons()
        self.status_label.setText("已取消")

    def closeEvent(self, event):
        # 关闭窗口时清理临时备份目录；不主动还原已被标准化的文件
        self._discard_backups()
        super().closeEvent(event)

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
