"""
批量音频响度/统计 - 核心分析逻辑（与 UI 无关）
拖入文件/文件夹 -> 扫描常见音频格式 -> 计算 Peak/RMS/LUFS/LRA/True Peak/DC Offset/削波等 -> 导出 Excel
依赖：系统需安装 ffmpeg / ffprobe 并在 PATH 中可用。
"""

import os
import re
import sys
import json
import shutil
import subprocess
from pathlib import Path

import numpy as np

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

AUDIO_EXTS = {
    ".wav", ".mp3", ".flac", ".m4a", ".aac", ".ogg", ".oga", ".opus",
    ".wma", ".aiff", ".aif", ".ape", ".dsf", ".wv",
}

# 没有真实位深概念的有损编码（ffprobe 的 sample_fmt 只是解码缓冲区格式，非原始位深）
LOSSY_CODECS = {"mp3", "aac", "vorbis", "wmav1", "wmav2", "wmapro", "opus"}

CLIP_THRESHOLD = 0.999      # 判定接近满幅的采样阈值
CLIP_MIN_CONSECUTIVE = 3    # 连续多少个采样点超阈值才算真正削波（区分单点脉冲噪声）
READ_CHUNK_FRAMES = 1_000_000  # 每次从 ffmpeg 管道读取的帧数（每帧含全部声道）

# Audition "响度(旧版)" 的官方算法未公开，社区反推：50ms 分帧 RMS(dB)，
# 过滤掉 -60dB 以下的静音帧后，对剩余帧的 dB 值取算术平均。这里按同样思路实现，可能与 Audition 数值有 1-2 dB 偏差。
LEGACY_LOUDNESS_WINDOW_MS = 50
LEGACY_LOUDNESS_SILENCE_DB = -60.0

# Windows 下隐藏 ffmpeg/ffprobe 弹出的控制台窗口
_CREATIONFLAGS = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


def find_tool(name):
    # PyInstaller 打包后，ffmpeg/ffprobe 会被解压到 sys._MEIPASS 目录，优先用打进 exe 里的那份，
    # 避免依赖用户系统的 PATH 配置
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", ""))
        candidate = base / f"{name}.exe"
        if candidate.exists():
            return str(candidate)
    return shutil.which(name)


def db(value):
    """线性幅度 -> dBFS，静音返回 -inf"""
    if value is None or value <= 0:
        return float("-inf")
    return 20.0 * np.log10(value)


def fmt_db(value, digits=1, unit=""):
    if value is None:
        return "N/A"
    if value == float("-inf"):
        return f"-inf {unit}" if unit else "-inf"
    text = str(round(float(value), digits))
    return f"{text} {unit}" if unit else text


def fmt_seconds(seconds):
    if seconds is None:
        return "N/A"
    m, s = divmod(seconds, 60)
    h, m = divmod(int(m), 60)
    if h:
        return f"{h}:{m:02d}:{s:05.2f}"
    return f"{m}:{s:05.2f}"


def probe_file(ffprobe, path):
    """用 ffprobe 读取容器/流层面的元数据"""
    cmd = [
        ffprobe, "-v", "quiet", "-print_format", "json",
        "-show_format", "-show_streams", str(path),
    ]
    result = subprocess.run(
        cmd, capture_output=True, creationflags=_CREATIONFLAGS
    )
    if result.returncode != 0:
        raise RuntimeError(f"ffprobe 失败: {result.stderr.decode(errors='ignore')[:300]}")

    info = json.loads(result.stdout.decode(errors="ignore"))

    audio_stream = None
    for s in info.get("streams", []):
        if s.get("codec_type") == "audio":
            audio_stream = s
            break
    if audio_stream is None:
        raise RuntimeError("未找到音频流")

    sample_rate = int(audio_stream.get("sample_rate", 0) or 0)
    channels = int(audio_stream.get("channels", 0) or 0)
    codec_name = audio_stream.get("codec_name", "")

    bit_depth = audio_stream.get("bits_per_raw_sample") or audio_stream.get("bits_per_sample")
    if codec_name in LOSSY_CODECS:
        # 有损编码没有真实位深，ffprobe 报出的 sample_fmt 只是解码缓冲区格式，非原始位深
        bit_depth = None
    elif not bit_depth or int(bit_depth) == 0:
        sample_fmt = audio_stream.get("sample_fmt", "")
        fallback = {
            "u8": 8, "u8p": 8,
            "s16": 16, "s16p": 16,
            "s32": 32, "s32p": 32,
            "flt": 32, "fltp": 32,
            "dbl": 64, "dblp": 64,
        }
        bit_depth = fallback.get(sample_fmt)

    duration = audio_stream.get("duration") or info.get("format", {}).get("duration")
    duration = float(duration) if duration else None

    return {
        "sample_rate": sample_rate,
        "channels": channels,
        "bit_depth": int(bit_depth) if bit_depth else None,
        "duration": duration,
        "codec_name": codec_name,
    }


def analyze_basic_stats(ffmpeg, path, channels):
    """将音频解码为 float32 PCM，流式计算 Peak / RMS / DC Offset / 削波，避免整段读入内存"""
    if channels <= 0:
        channels = 1
    bytes_per_frame = 4 * channels

    cmd = [
        ffmpeg, "-v", "error", "-i", str(path),
        "-f", "f32le", "-acodec", "pcm_f32le", "pipe:1",
    ]
    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        creationflags=_CREATIONFLAGS,
    )

    count = 0
    sum_ch = np.zeros(channels, dtype=np.float64)
    sumsq_ch = np.zeros(channels, dtype=np.float64)
    peak = 0.0
    clip_count = 0
    tail = np.zeros((0, channels), dtype=np.float32)
    leftover = b""

    read_size = READ_CHUNK_FRAMES * bytes_per_frame
    while True:
        chunk = proc.stdout.read(read_size)
        if not chunk:
            break
        data = leftover + chunk
        usable_len = (len(data) // bytes_per_frame) * bytes_per_frame
        leftover = data[usable_len:]
        if usable_len == 0:
            continue

        arr = np.frombuffer(data[:usable_len], dtype="<f4").reshape(-1, channels)
        count += arr.shape[0]
        sum_ch += arr.sum(axis=0, dtype=np.float64)
        sumsq_ch += (arr.astype(np.float64) ** 2).sum(axis=0)
        chunk_peak = np.max(np.abs(arr)) if arr.size else 0.0
        if chunk_peak > peak:
            peak = float(chunk_peak)

        # 统计削波次数：每一段连续 >= CLIP_MIN_CONSECUTIVE 个满幅采样点算一次削波事件
        check_arr = np.concatenate([tail, arr], axis=0)
        mask = np.abs(check_arr) >= CLIP_THRESHOLD
        if mask.any():
            for c in range(channels):
                col = mask[:, c].view(np.int8)
                if not col.any():
                    continue
                padded = np.concatenate(([0], col, [0]))
                edges = np.flatnonzero(np.diff(padded))
                run_lengths = edges[1::2] - edges[0::2]
                clip_count += int(np.sum(run_lengths >= CLIP_MIN_CONSECUTIVE))
        tail_len = CLIP_MIN_CONSECUTIVE - 1
        tail = arr[-tail_len:] if arr.shape[0] >= tail_len else arr

    proc.stdout.close()
    stderr_data = proc.stderr.read()
    proc.wait()
    if proc.returncode != 0 and count == 0:
        raise RuntimeError(f"ffmpeg 解码失败: {stderr_data.decode(errors='ignore')[:300]}")

    if count == 0:
        return {
            "peak_db": float("-inf"),
            "rms_db": [float("-inf")] * channels,
            "dc_offset": 0.0,
            "clip_count": 0,
        }

    mean_ch = sum_ch / count
    rms_ch = np.sqrt(sumsq_ch / count)

    dc_idx = int(np.argmax(np.abs(mean_ch)))
    dc_offset = float(mean_ch[dc_idx])

    return {
        "peak_db": db(peak),
        "rms_db": [db(v) for v in rms_ch],
        "dc_offset": dc_offset,
        "clip_count": clip_count,
    }


_FRAME_RE = re.compile(r"\bM:\s*(-?inf|-?\d+\.?\d*)\s+S:\s*(-?inf|-?\d+\.?\d*)")


def _parse_loudness_value(text):
    return float("-inf") if "inf" in text else float(text)


def analyze_loudness(ffmpeg, path):
    """调用 ffmpeg ebur128 滤镜计算 LUFS-I / LUFS-S(max) / LUFS-M(max) / LRA / True Peak"""
    cmd = [
        ffmpeg, "-nostats", "-i", str(path),
        "-filter_complex", "ebur128=peak=true:framelog=info",
        "-f", "null", "-",
    ]
    result = subprocess.run(
        cmd, capture_output=True, creationflags=_CREATIONFLAGS
    )
    text = result.stderr.decode(errors="ignore")

    m_vals, s_vals = [], []
    for m, s in _FRAME_RE.findall(text):
        m_vals.append(_parse_loudness_value(m))
        s_vals.append(_parse_loudness_value(s))

    def extract(label, pattern):
        m = re.search(label + r".*?" + pattern, text, re.DOTALL)
        return float(m.group(1)) if m else None

    lufs_i = extract(r"Integrated loudness:", r"I:\s*(-?\d+\.?\d*)")
    lra = extract(r"Loudness range:", r"LRA:\s*(-?\d+\.?\d*)")
    true_peak = extract(r"True peak:", r"Peak:\s*(-?\d+\.?\d*)")

    return {
        "lufs_i": lufs_i,
        "lufs_s_max": max(s_vals) if s_vals else None,
        "lufs_m_max": max(m_vals) if m_vals else None,
        "lra": lra,
        "true_peak_db": true_peak,
    }


def analyze_file(ffmpeg, ffprobe, path):
    meta = probe_file(ffprobe, path)
    basic = analyze_basic_stats(ffmpeg, path, meta["channels"])
    loud = analyze_loudness(ffmpeg, path)

    rms = basic["rms_db"]
    rms_l = rms[0] if len(rms) >= 1 else None
    rms_r = rms[1] if len(rms) >= 2 else None

    return {
        "format": path.suffix.upper().lstrip("."),
        "sample_rate": meta["sample_rate"],
        "bit_depth": meta["bit_depth"],
        "channels": meta["channels"],
        "duration": meta["duration"],
        "peak_db": basic["peak_db"],
        "rms_l_db": rms_l,
        "rms_r_db": rms_r,
        "dc_offset": basic["dc_offset"],
        "clip_count": basic["clip_count"],
        "lufs_i": loud["lufs_i"],
        "lufs_s_max": loud["lufs_s_max"],
        "lufs_m_max": loud["lufs_m_max"],
        "lra": loud["lra"],
        "true_peak_db": loud["true_peak_db"],
    }


def scan_folder(root_folders):
    files = []
    for folder in root_folders:
        folder = Path(folder)
        if folder.is_file():
            if folder.suffix.lower() in AUDIO_EXTS:
                files.append(folder)
            continue
        for dirpath, _, filenames in os.walk(folder):
            for name in filenames:
                if Path(name).suffix.lower() in AUDIO_EXTS:
                    files.append(Path(dirpath) / name)
    return sorted(set(files))


# Excel 导出：完整明细列（还原早期版本的全部字段），表头全部译为中文
HEADERS = [
    "序号", "文件名", "采样率(Hz)", "位深", "声道数", "时长",
    "峰值(dBFS)", "左声道RMS(dBFS)", "右声道RMS(dBFS)",
    "平均响度(LUFS)", "短期最大响度(LUFS)", "瞬时最大响度(LUFS)", "响度范围(LU)", "真实峰值(dBFS)",
    "直流偏移", "是否削波", "备注",
]

# UI 表格：精简列，额外带上采样率/位深方便快速核对文件规格
TABLE_HEADERS = [
    "文件", "时长", "采样率", "位深", "峰值",
    "瞬时最大", "短期最大", "平均响度", "响度范围",
]


def build_excel_row(i, row):
    """把一条分析结果转换成 Excel 明细行（列顺序对应 HEADERS）"""
    if row.get("error"):
        return [
            i, row["name"], "", "", "", "",
            "", "", "", "", "", "", "", "", "", "", row["error"],
        ]

    d = row["data"]
    return [
        i, row["name"],
        d["sample_rate"] or "N/A",
        d["bit_depth"] if d["bit_depth"] else "N/A",
        d["channels"] or "N/A",
        fmt_seconds(d["duration"]),
        fmt_db(d["peak_db"]),
        fmt_db(d["rms_l_db"]),
        fmt_db(d["rms_r_db"]) if d["rms_r_db"] is not None else "-",
        fmt_db(d["lufs_i"]),
        fmt_db(d["lufs_s_max"]),
        fmt_db(d["lufs_m_max"]),
        fmt_db(d["lra"]),
        fmt_db(d["true_peak_db"]),
        round(d["dc_offset"], 6),
        "是" if d["clip_count"] > 0 else "否",
        "",
    ]


def build_table_row(row):
    """把一条分析结果转换成 UI 表格显示行（列顺序对应 TABLE_HEADERS）"""
    if row.get("error"):
        return [row["name"], "", "", "", "", "", "", "", row["error"]]

    d = row["data"]
    return [
        row["name"],
        fmt_seconds(d["duration"]),
        d["sample_rate"] or "N/A",
        d["bit_depth"] if d["bit_depth"] else "N/A",
        fmt_db(d["peak_db"], unit="dBFS"),
        fmt_db(d["lufs_m_max"], unit="LUFS"),
        fmt_db(d["lufs_s_max"], unit="LUFS"),
        fmt_db(d["lufs_i"], unit="LUFS"),
        fmt_db(d["lra"], unit="LU"),
    ]


def write_excel(rows, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "音频响度统计"
    ws.append(HEADERS)

    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"

    clip_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    error_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    for i, row in enumerate(rows, start=1):
        ws.append(build_excel_row(i, row))
        if row.get("error"):
            for cell in ws[ws.max_row]:
                cell.fill = error_fill
        elif row["data"]["clip_count"] > 0:
            for cell in ws[ws.max_row]:
                cell.fill = clip_fill

    widths = [6, 28, 10, 10, 8, 10, 11, 14, 14, 13, 15, 15, 11, 12, 10, 9, 30]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    wb.save(out_path)
