# Loudness Analyzer / 音频响度标准化工具

拖入音频文件或整个文件夹，逐首统计 峰值 / 真实峰值 / LUFS-I/S/M / LRA / 削波次数，并可**批量做响度标准化**（loudnorm 二次通过，保留原格式规格）。所有响度指标能一键导出成 Excel 明细表。界面跟随系统深色/浅色主题。

![icon](icon_preview.png)

## 直接使用

在 [Releases](https://github.com/wyh-alt/Loudness-Analyzer/releases) 下载最新的 `LoudnessAnalyzer.exe`，双击即可运行。ffmpeg / ffprobe 已经打包进 exe，全新的 Windows 电脑无需额外安装依赖。

## 支持的格式

WAV、MP3、FLAC、M4A、AAC、OGG、OPUS、WMA、AIFF、APE、DSF、WV。

## 使用流程

1. **拖入** 音频文件或文件夹到输入框（也可以点"浏览…"选目录）。程序自动逐首检测响度并把结果写入下方表格。
2. **导出表格**：点右下角"导出表格"→ 选保存位置 → 生成 xlsx 明细。什么时候点都行，导出的是当前表格状态。
3. **响度标准化**（默认已勾选）：调好目标响度 / 容差 / 最高实际峰值电平，点"开始处理"。程序会：
   - 对每个文件调用 ffmpeg loudnorm 二次通过做线性响度归一化
   - **保留原格式规格**：采样率、声道、位深、码率、样本格式一致（详见下节）
   - 原地覆盖源文件，同时把原始字节完整备份到临时目录
   - 表格实时刷新为处理后的新指标
4. **撤销处理**：点"撤销处理"，从备份**字节级还原**原音频（不是反向处理，削波音频也能完整回到原样）。拖入新素材或调整参数后按钮会重置，撤销能力随之释放。

## 响度标准化的关键规则

- **跳过条件**：源 LUFS-I 已在 `目标 ± 容差` 且 True Peak ≤ 目标 dBTP —— 两者都合规才跳过重编码；任一超标都会走一遍 loudnorm。
- **格式保留**：
  - 采样率 用源采样率
  - WAV 位深 保留 16/24/32，32-bit float 源用 `pcm_f32le`
  - FLAC 位深 24-bit 源用 `-sample_fmt s32` 保精度
  - MP3 / AAC / OGG / OPUS / WMA 码率 用 ffprobe 读到的源码率 `-b:a`
  - 声道数、AIFF 大小端 一并保留
- **算法**：loudnorm `linear=true`，第一遍测量 → 第二遍带 measured_I / TP / LRA / thresh / offset 应用线性缩放；源 LRA 超过 target LRA(20) 时 loudnorm 自动落到 dynamic 模式。

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

1. 安装 `requirements-build.txt` 的依赖（PyInstaller + Pillow）。
2. 从 PATH 自动定位 `ffmpeg.exe` / `ffprobe.exe` 拷贝到 `bin/`。
3. 用 Pillow 生成 `assets/splash.png`（bootloader 启动图）。
4. 用 PyQt6 + Pillow 从 SVG 生成多档 `icon.ico`。
5. 调 PyInstaller 输出单文件 `dist\LoudnessAnalyzer.exe`。

## 启动图（Splash）设计

- 打包模式：PyInstaller `Splash()` 在解释器启动前展示 `assets/splash.png`（"正在启动…"）。
- Qt 阶段：`main()` 里立刻创建视觉一致的 `QSplashScreen`（"正在加载界面…"），再关闭 bootloader splash；主窗口 `show()` 后 `finish(window)` 关闭。
- 视觉：360×148、深色圆角背景、居中标题+居中副标题，Fluent 深色风格。

## 目录结构速览

```
main.py              # PyQt6 界面、workers、按钮状态机
core.py              # ffprobe/ffmpeg 调用、响度指标、loudnorm 二次通过、Excel 导出
make_icon.py         # 从 SVG 生成 icon.ico（多档）
make_splash.py       # 生成 splash.png
build.bat            # 一键打包脚本
Loudness-Analyzer.spec  # PyInstaller spec
```

## 许可

MIT
