"""定位并复制 ffmpeg/ffprobe 到 bin/，跳过 chocolatey / scoop 的小体积 shim。
仅供 build.bat 内部调用：python _locate_tool.py <tool_name>
"""

import os
import shutil
import sys
from pathlib import Path

MIN_SIZE = 10 * 1024 * 1024  # 10MB, shim 通常几百 KB


def all_matches(name: str):
    """遍历 PATH 找出所有同名可执行文件，按 PATH 顺序返回。"""
    pathext = os.environ.get("PATHEXT", ".EXE").split(os.pathsep)
    seen = set()
    for d in os.environ.get("PATH", "").split(os.pathsep):
        if not d:
            continue
        for ext in [""] + pathext:
            p = Path(d) / f"{name}{ext}"
            try:
                if p.is_file():
                    key = str(p.resolve()).lower()
                    if key not in seen:
                        seen.add(key)
                        yield p
            except OSError:
                continue


def main():
    if len(sys.argv) < 2:
        print("usage: _locate_tool.py <tool_name>", file=sys.stderr)
        sys.exit(2)
    tool = sys.argv[1]
    base = tool[:-4] if tool.lower().endswith(".exe") else tool

    dst = Path(__file__).resolve().parent / "bin" / f"{base}.exe"
    dst.parent.mkdir(exist_ok=True)
    if dst.exists():
        print(f"[ok] {dst.name} already present")
        return

    for candidate in all_matches(base):
        size = candidate.stat().st_size
        if size < MIN_SIZE:
            print(f"[skip] {candidate} ({size} bytes, likely a shim)")
            continue
        shutil.copy2(candidate, dst)
        print(f"[ok] copied {candidate} -> {dst}")
        return

    print(
        f"ERROR: real {base}.exe not found in PATH. "
        f"Install ffmpeg from https://www.gyan.dev/ffmpeg/builds/ and add its bin/ to PATH.",
        file=sys.stderr,
    )
    sys.exit(1)


if __name__ == "__main__":
    main()
