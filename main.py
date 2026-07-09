"""
批量音频响度统计工具 - 轻量 PyQt6 界面
拖入音频文件/文件夹 -> 点击开始 -> 表格逐首显示响度统计 -> 自动导出 Excel
界面颜色会跟随系统的浅色/深色主题自动切换。
"""

import os
import sys
import ctypes
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
    QStyledItemDelegate, QStyle, QSplashScreen,
)

from core import analyze_file, scan_folder, write_excel, find_tool, build_table_row, TABLE_HEADERS

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
SPLASH_TITLE = "批量音频响度统计工具"

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
"""


class DropLineEdit(QLineEdit):
    """支持拖入单个音频文件或文件夹的路径输入框，拖入新的会替换旧路径"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("pathEdit")
        self.setAcceptDrops(True)
        self.setPlaceholderText("将音频文件或文件夹拖到这里")

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, e):
        urls = e.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if path:
                self.setText(path)


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

        try:
            write_excel(rows, self.out_path)
        except Exception as e:
            self.failed.emit(f"写入 Excel 失败: {e}")
            return

        self.finished_ok.emit(self.out_path, total, error_count)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("root")
        self.setWindowTitle("批量音频响度统计工具")
        self.resize(900, 520)
        self.setMinimumSize(900, 380)
        if ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(ICON_PATH)))

        self.ffmpeg = find_tool("ffmpeg")
        self.ffprobe = find_tool("ffprobe")
        self.worker = None
        self._accent_color = _LIGHT["accent"]  # 由 set_theme() 跟随主题更新，用于选中单元格边框

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

        action_row = QHBoxLayout()
        self.status_label = QLabel("就绪", self)
        self.status_label.setObjectName("statusLabel")
        action_row.addWidget(self.status_label)
        action_row.addStretch(1)
        self.cancel_btn = QPushButton("取消", self)
        self.cancel_btn.setObjectName("cancelBtn")
        self.cancel_btn.setEnabled(False)
        self.start_btn = QPushButton("开始", self)
        self.start_btn.setObjectName("startBtn")
        action_row.addWidget(self.cancel_btn)
        action_row.addWidget(self.start_btn)
        layout.addLayout(action_row)

        self.browse_btn.clicked.connect(self.on_browse)
        self.start_btn.clicked.connect(self.on_start)
        self.cancel_btn.clicked.connect(self.on_cancel)

        if not self.ffmpeg or not self.ffprobe:
            self.start_btn.setEnabled(False)
            self.status_label.setText("缺少依赖：未检测到 ffmpeg / ffprobe，请安装后加入系统 PATH")

    def on_browse(self):
        folder = QFileDialog.getExistingDirectory(self, "选择包含音频文件的文件夹")
        if folder:
            self.path_edit.setText(folder)

    def on_start(self):
        path = self.path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "提示", "请先拖入或选择一个音频文件/文件夹。")
            return
        if not os.path.exists(path):
            QMessageBox.warning(self, "错误", "路径不存在，请重新选择。")
            return

        default_dir = path if os.path.isdir(path) else os.path.dirname(path)
        default_name = f"音频响度统计_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out_path, _ = QFileDialog.getSaveFileName(
            self, "选择 Excel 保存位置",
            os.path.join(default_dir, default_name),
            "Excel 文件 (*.xlsx)",
        )
        if not out_path:
            return

        self.table.setRowCount(0)
        self.progress_bar.setValue(0)
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.status_label.setText("准备扫描...")

        self.worker = ScanWorker([path], out_path, self.ffmpeg, self.ffprobe)
        self.worker.progress.connect(self.on_progress)
        self.worker.row_ready.connect(self.on_row_ready)
        self.worker.finished_ok.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.cancelled.connect(self.on_cancelled)
        self.worker.start()

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

    def on_row_ready(self, row):
        cells = build_table_row(row)
        r = self.table.rowCount()
        self.table.insertRow(r)
        for col, value in enumerate(cells):
            item = QTableWidgetItem(str(value))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(r, col, item)
        self.table.scrollToBottom()

    def _reset_buttons(self):
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)

    def on_finished(self, out_path, total, error_count):
        self._reset_buttons()
        self.status_label.setText(f"完成：共 {total} 个文件，{error_count} 个失败")
        reply = QMessageBox.question(
            self, "处理完成",
            f"共 {total} 个文件，{error_count} 个失败。\n结果已保存到：\n{out_path}\n\n是否打开所在文件夹？",
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
