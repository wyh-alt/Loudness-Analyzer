@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo === [1/5] Installing Python dependencies ===
python -m pip install -r requirements-build.txt
if errorlevel 1 goto :err

echo.
echo === [2/5] Preparing bundled ffmpeg / ffprobe ===
python _locate_tool.py ffmpeg
if errorlevel 1 goto :err
python _locate_tool.py ffprobe
if errorlevel 1 goto :err

echo.
echo === [3/5] Generating splash image ===
python make_splash.py
if errorlevel 1 goto :err

echo.
echo === [4/5] Generating app icon ===
python make_icon.py
if errorlevel 1 goto :err

echo.
echo === [5/5] Running PyInstaller ===
python -m PyInstaller --clean --noconfirm Loudness-Analyzer.spec
if errorlevel 1 goto :err

echo.
echo Build complete. Output: dist\LoudnessAnalyzer.exe
goto :eof


:err
echo.
echo Build failed.
exit /b 1
