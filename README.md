# Loudness Analyzer / 音频响度标准化工具

拖入音频文件或整个文件夹，逐首检测 峰值 / 真实峰值 / LUFS-I/S/M / LRA / 削波次数，并可**批量做响度标准化 + 格式标准化**。原文件不会被改动，处理产物另存到用户指定目录。响度指标可一键导出成 Excel 明细表。界面跟随系统深色/浅色主题。

![icon](icon_preview.png)

## 直接使用

在 [Releases](https://github.com/wyh-alt/Loudness-Analyzer/releases) 下载最新的 `LoudnessAnalyzer.exe`，双击即可运行。ffmpeg / ffprobe 已经打包进 exe，全新的 Windows 电脑无需额外安装依赖。

## 支持的输入格式

WAV、MP3、FLAC、M4A、AAC、OGG、OPUS、WMA、AIFF、APE、DSF、WV。

## 三个主按钮

| 按钮 | 作用 |
|---|---|
| **响度统计** | 对表格里已列出的文件逐首跑响度分析，把 LUFS-I/S/M / LRA / TP 等指标填到表格 |
| **开始处理** | 弹目录选择框 → 逐首按当前设置做响度/格式标准化 → 另存到该目录，表格实时刷新为处理后的指标 |
| **导出表格** | 弹保存路径 → 把当前表格状态导出为 xlsx |

## 典型工作流

1. **拖入**音频文件或文件夹到输入框（也可点"浏览…"选目录）。程序**只列出文件**到表格，不做检测（快速预览）。
2. 想看响度就点 **"响度统计"**；不想看就直接下一步。
3. 勾选并配置 **"响度标准化"** 和/或 **"格式标准化"**（默认响度勾选、格式不勾选）。
4. 点 **"开始处理"**，选输出目录 → 程序把每首都处理后另存到该目录，同时把新指标填回表格。
5. 想要 xlsx 记录点 **"导出表格"** 随时保存。

## 响度标准化

- **目标响度**（LUFS-I）：-70 ~ -5，默认 -12
- **容差**（LU）：0 ~ 20，默认 1
- **最高实际峰值电平**（dBTP）：-9 ~ 0，默认 -1

**跳过条件**：源 LUFS-I 在 `目标 ± 容差` 且 True Peak ≤ 目标 dBTP —— 两者都合规就跳过 loudnorm，直接复制原字节到输出目录。任一超标都会走一遍 loudnorm 二次通过（线性模式，源 LRA > 20 时自动落到 dynamic）。

## 格式标准化

- **音频格式**：.wav / .mp3 / .m4a / .flac
- **采样率**：44100 Hz / 48000 Hz
- **位深度**（仅无损格式显示）：16 / 24 / 32 Bit
- **比特率**（仅有损格式显示）：320 / 256 / 192 / 128 / 64 kbps
- **声道**：立体声 / 单声道

未勾选时按源格式保留输出（wav 保留位深、mp3 保留码率等）；勾选后按上述下拉严格重编码。

## 开发

```powershell
python -m pip install -r requirements.txt
python main.py
```

需要本机已装 ffmpeg，`ffmpeg` 和 `ffprobe` 在 PATH 中可用。

## 从源码打包

```powershell
build.bat
```

`build.bat` 会：

1. 安装 `requirements-build.txt` 的依赖（PyInstaller + Pillow）
2. 从 PATH 自动定位 `ffmpeg.exe` / `ffprobe.exe` 拷贝到 `bin/`
3. 用 Pillow 生成 `assets/splash.png`（bootloader 启动图）
4. 用 PyQt6 + Pillow 从 SVG 生成多档 `icon.ico`
5. 调 PyInstaller 输出单文件 `dist\LoudnessAnalyzer.exe`

## 目录结构速览

```
main.py              # PyQt6 界面、workers（AnalyzeWorker / ProcessWorker）
core.py              # ffprobe/ffmpeg 调用、响度指标、loudnorm、process_file、Excel 导出
make_icon.py         # 从 SVG 生成 icon.ico
make_splash.py       # 生成 splash.png
build.bat            # 一键打包脚本
Loudness-Analyzer.spec  # PyInstaller spec
```

## 许可

MIT
