@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo === [1/6] Installing Python dependencies ===
python -m pip install -r requirements-build.txt
if errorlevel 1 goto :err

echo.
echo === [2/6] Preparing bundled ffmpeg / ffprobe ===
python _locate_tool.py ffmpeg
if errorlevel 1 goto :err
python _locate_tool.py ffprobe
if errorlevel 1 goto :err

echo.
echo === [3/6] Generating splash image ===
python make_splash.py
if errorlevel 1 goto :err

echo.
echo === [4/6] Generating app icon ===
python make_icon.py
if errorlevel 1 goto :err

echo.
echo === [5/6] Cleaning old build / dist artefacts ===
REM 彻底清掉旧的 build 缓存和 dist 产物，避免 PyInstaller 残留旧图标资源
if exist build (
    rmdir /s /q build
)
if exist dist (
    rmdir /s /q dist
)

echo.
echo === [6/6] Running PyInstaller ===
python -m PyInstaller --clean --noconfirm Loudness-Analyzer.spec
if errorlevel 1 goto :err

echo.
echo Build complete. Output: dist\LoudnessAnalyzer.exe
echo.
echo NOTE: Windows Explorer caches icons per exe path. If the taskbar / Explorer
echo       still shows the old icon, run one of:
echo         ie4uinit.exe -show
echo         del /f /q %%LOCALAPPDATA%%\IconCache.db  (then restart Explorer)
echo       Or simply rename the exe / launch from a new path.
REM 主动尝试刷新一下（大多数 Windows 版本上 ie4uinit 存在且可用）
where ie4uinit >nul 2>&1 && ie4uinit.exe -show
goto :eof


:err
echo.
echo Build failed.
exit /b 1
