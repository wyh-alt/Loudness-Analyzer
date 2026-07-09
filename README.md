# Loudness Analyzer / 批量音频响度统计工具

拖入音频文件或整个文件夹，逐首统计峰值、真实峰值、LUFS-I/S/M、LRA、削波次数等指标，并自动导出 Excel 明细。界面跟随系统深色/浅色主题。

## 直接使用

在 [Releases](https://github.com/wyh-alt/Loudness-Analyzer/releases) 下载最新的 `LoudnessAnalyzer.exe`，双击即可运行。ffmpeg / ffprobe 已经打包进 exe，全新的 Windows 电脑无需额外安装依赖。

## 支持的格式

WAV、MP3、FLAC、M4A、AAC、OGG、OPUS、WMA、AIFF、APE、DSF、WV。

## 开发

```powershell
python -m pip install -r requirements.txt
python main.py
```

## 从源码打包

```powershell
build.bat
```

`build.bat` 会：

1. 安装 `requirements-build.txt` 的依赖（包含 PyInstaller + Pillow）。
2. 从 PATH 自动定位 `ffmpeg.exe` / `ffprobe.exe` 拷贝到 `bin/`。
3. 用 Pillow 生成 `assets/splash.png`（bootloader 启动图）。
4. 调 PyInstaller 输出单文件 `dist\LoudnessAnalyzer.exe`。

前置：本机需要装有 ffmpeg，`ffmpeg` 与 `ffprobe` 在 PATH 中可用。

## 启动图（Splash）设计

- 打包模式：PyInstaller `Splash()` 在解释器启动前展示 `assets/splash.png`（"正在启动…"）。
- Qt 阶段：`main()` 里立刻创建视觉一致的 `QSplashScreen`（"正在加载界面…"），再关闭 bootloader splash；主窗口 `show()` 后 `finish(window)` 关闭。
- 视觉：360×148、深色圆角背景、居中标题+居中副标题，Fluent 深色风格。

## 许可

MIT
