# -*- mode: python ; coding: utf-8 -*-
"""LoudnessAnalyzer 单文件打包 spec。

关键点：
- ffmpeg.exe / ffprobe.exe 通过 build.bat 预置到项目根 bin/ 下，再被 binaries 打进 exe，
  运行时通过 sys._MEIPASS 定位（见 core.find_tool）。
- assets/splash.png 由 make_splash.py 生成，Splash() 让 bootloader 在解释器启动前就显示，
  Qt splash 在 main() 里接力，避免单文件 exe 解压/初始化的黑屏空档。
"""

from pathlib import Path

APP_NAME = "LoudnessAnalyzer"
SPEC_DIR = Path(SPECPATH)
BIN_DIR = SPEC_DIR / "bin"
ASSETS_DIR = SPEC_DIR / "assets"
ICON_PATH = SPEC_DIR / "icon.ico"
SPLASH_PATH = ASSETS_DIR / "splash.png"

binaries = []
for tool in ("ffmpeg.exe", "ffprobe.exe"):
    p = BIN_DIR / tool
    if not p.exists():
        raise SystemExit(
            f"Missing bundled tool: {p}\n"
            f"Run build.bat first — it copies ffmpeg/ffprobe from PATH into bin/."
        )
    binaries.append((str(p), "."))

datas = []
if ICON_PATH.exists():
    datas.append((str(ICON_PATH), "."))

if not SPLASH_PATH.exists():
    raise SystemExit(
        f"Missing splash asset: {SPLASH_PATH}\n"
        f"Run: python make_splash.py"
    )


a = Analysis(
    ["main.py"],
    pathex=[str(SPEC_DIR)],
    binaries=binaries,
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

splash = Splash(
    str(SPLASH_PATH),
    binaries=a.binaries,
    datas=a.datas,
    # 文案已经烘焙在 PNG 里（"正在启动…"），不需要运行时更新
    text_pos=None,
    text_size=10,
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    splash,
    splash.binaries,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(ICON_PATH)] if ICON_PATH.exists() else None,
)
