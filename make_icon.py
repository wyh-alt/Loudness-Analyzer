"""生成喇叭 + 声波风格的应用图标，覆盖 icon.ico 里全部尺寸档位。
Windows Explorer / 任务栏 / 启动菜单会根据显示比例挑不同尺寸档，所以要一次导多档。

依赖：PyQt6（渲染 SVG）+ Pillow（拼多档 ICO）。项目已装。
"""

from io import BytesIO
from pathlib import Path

from PyQt6.QtCore import QByteArray, Qt
from PyQt6.QtGui import QImage, QPainter, QPixmap
from PyQt6.QtSvg import QSvgRenderer
from PyQt6.QtWidgets import QApplication
from PIL import Image

APP_DIR = Path(__file__).resolve().parent

# 主色跟 UI accent 保持一致；透明背景，任何主题下都不会跟系统颜色打架
BODY = "#4c7cf0"
WAVE = "#4c7cf0"

# 扬声器：磁铁矩形 + 号筒梯形，右侧三道弧形声波（近→远逐渐减细）
ICON_SVG = f"""<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 256 256'>
  <!-- 扬声器主体（磁铁+号筒），一体化路径 -->
  <path d='M36 96 L90 96 L150 44 L150 212 L90 160 L36 160 Z'
        fill='{BODY}' stroke='{BODY}' stroke-width='8' stroke-linejoin='round'/>
  <!-- 三条声波弧线，右侧向外扩散 -->
  <path d='M172 100 Q192 128 172 156' fill='none' stroke='{WAVE}'
        stroke-width='12' stroke-linecap='round'/>
  <path d='M200 80  Q232 128 200 176' fill='none' stroke='{WAVE}'
        stroke-width='12' stroke-linecap='round'/>
  <path d='M228 60  Q272 128 228 196' fill='none' stroke='{WAVE}'
        stroke-width='12' stroke-linecap='round'/>
</svg>"""


def _render(size: int) -> Image.Image:
    """把 SVG 渲染到指定像素尺寸的透明 PNG"""
    renderer = QSvgRenderer(QByteArray(ICON_SVG.encode("utf-8")))
    if not renderer.isValid():
        raise RuntimeError("SVG parse failed")
    img = QImage(size, size, QImage.Format.Format_ARGB32)
    img.fill(Qt.GlobalColor.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
    renderer.render(p)
    p.end()
    buf = BytesIO()
    # 通过 QImage.save 转到 PNG，再交给 Pillow —— 这样 alpha 通道最稳
    from PyQt6.QtCore import QBuffer
    qbuf = QBuffer()
    qbuf.open(QBuffer.OpenModeFlag.ReadWrite)
    img.save(qbuf, "PNG")
    buf.write(qbuf.data().data())
    buf.seek(0)
    return Image.open(buf).convert("RGBA")


def main():
    app = QApplication.instance() or QApplication([])
    _ = app
    sizes = [16, 24, 32, 48, 64, 128, 256]
    imgs = [_render(s) for s in sizes]
    # 用最大档作为基准图，sizes 告诉 Pillow 需要嵌入哪些子档
    base = imgs[-1]
    out_path = APP_DIR / "icon.ico"
    base.save(out_path, format="ICO", sizes=[(s, s) for s in sizes])
    print(f"wrote {out_path}  sizes={sizes}")
    # 顺带保存 256 版 PNG 便于预览
    imgs[-1].save(APP_DIR / "icon_preview.png")


if __name__ == "__main__":
    main()
