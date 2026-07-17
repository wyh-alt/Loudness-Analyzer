# Loudness Analyzer / 音频响度标准化工具

一款批量处理音频响度和格式的桌面工具：拖入文件/文件夹 → 可选检测响度 → 按目标参数**响度标准化**和/或**格式标准化** → 处理后音频另存到指定目录，源文件不动。所有响度指标可一键导出 Excel 明细。界面跟随系统深色/浅色主题。

![icon](icon_preview.png)

## 直接使用

在 [Releases](https://github.com/wyh-alt/Loudness-Analyzer/releases) 下载最新的 `LoudnessAnalyzer.exe`，双击即可运行。ffmpeg / ffprobe 已经打包进 exe，全新的 Windows 电脑无需额外安装依赖。

## 支持的输入格式

WAV、MP3、FLAC、M4A、AAC、OGG、OPUS、WMA、AIFF、APE、DSF、WV。

## 四个主按钮

| 按钮 | 作用 |
|---|---|
| **导出表格** | 弹保存路径 → 把当前表格状态导出为 xlsx（检测完 or 处理完的行才可导出） |
| **取消** | 立即终止正在跑的响度统计或音频处理（会 kill 当前 ffmpeg，不等它自然结束） |
| **响度统计** | 对表格里已列出的文件逐首跑响度分析，把 LUFS-I/S/M / LRA / TP 等指标填到表格 |
| **开始处理** | 弹目录选择框 → 按当前设置对每首做响度/格式标准化 → 另存到该目录，表格实时刷新为处理后的指标 |

## 典型工作流

1. **拖入**音频文件或文件夹（拖到路径框或**表格区域**都可以；也可点"浏览…"选目录）。程序**只列出文件**到表格，不做检测（快速预览）。
2. 想看响度就点 **"响度统计"**；不想看直接下一步。
3. 勾选并配置 **"响度标准化"** 和/或 **"格式标准化"**（默认响度勾选、格式不勾选）。
4. 点 **"开始处理"** → 选输出目录 → 每首都会处理后另存到该目录（源文件不变），表格里实时刷新为处理后的指标；当前处理行整行会高亮 + 自动滚动到视图中间。
5. 处理完弹汇总：**总数 / 实际处理 / 符合要求无需处理 / 失败**，并询问是否打开输出目录。
6. 想要 xlsx 记录点 **"导出表格"** 随时保存。

## 响度标准化

- **目标响度**（LUFS-I）：-70 ~ -5，默认 -12
- **容差**（LU）：0 ~ 20，默认 1
- **最高实际峰值电平**（dBTP）：-9 ~ 0，默认 -1

**跳过条件**：源 LUFS-I 在 `目标 ± 容差` 且 True Peak ≤ 目标 dBTP —— 两者都合规就跳过 loudnorm，直接把源字节复制到输出目录（在汇总里计入"符合要求无需处理"）。

**处理管线**（任一项超标时启用）：

```
[原始] → loudnorm 第一遍（measure，print_format=json）
      → 解析 input_i / input_lra / input_tp / input_thresh / target_offset
      → loudnorm 第二遍（linear=true + 全部 measured_* + offset）  → 纯静态增益
      → alimiter=limit=<TP linear>:level=disabled:asc=0             → 真峰兜底
      → 编码到目标格式
```

- **两遍法 + linear=true**：保证响度调整是纯静态增益，动态范围（LRA）不被压缩
- **alimiter 兜底**：即便 loudnorm 因源 LRA > 20 静默回退到 dynamic，或线性缩放后 inter-sample 峰值超过目标 TP，alimiter 也能把峰值压回目标 —— `level=disabled` / `asc=0` 关闭自动电平控制，只做纯粹的峰值限制，不额外压动态、也不把响度顶上去

## 格式标准化

- **音频格式**：.wav / .mp3 / .m4a / .flac
- **采样率**：44100 Hz / 48000 Hz
- **位深度**（仅无损格式显示）：16 / 24 / 32 Bit
- **比特率**（仅有损格式显示）：320 / 256 / 192 / 128 / 64 kbps
- **声道**：立体声 / 单声道

未勾选时按源格式保留输出（wav 保留位深、mp3 保留码率等）；勾选后按上述下拉严格重编码。可以单独使用（只改格式不动响度），也可以叠加响度标准化。

## 交互细节

- **拖入区域**：路径输入框、表格区域都接受文件/文件夹拖入
- **表格高亮 + 自动滚动**：处理过程中当前文件所在行整行会用主强调色高亮，表格自动滚动让它保持在视图中间
- **立即取消**：`取消`按钮直接 kill 正在跑的 ffmpeg 子进程，一般 20ms 以内响应
- **响度数值精度**：所有 LUFS / dBTP / dBFS / LU 数值统一两位小数补零显示
- **Excel 表头**：单位换行到第二行显示，除文件名列外所有列居中

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
main.py                 # PyQt6 界面、AnalyzeWorker / ProcessWorker、按钮状态机
core.py                 # ffprobe/ffmpeg 调用、响度指标、process_file、CancelToken、Excel 导出
make_icon.py            # 从 SVG 生成 icon.ico（多档）
make_splash.py          # 生成 splash.png
build.bat               # 一键打包脚本
Loudness-Analyzer.spec  # PyInstaller spec
```

## 许可

MIT
