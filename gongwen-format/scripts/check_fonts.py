#!/usr/bin/env python3
"""Check optional local and installed gongwen fonts."""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
FONT_DIR = SKILL_DIR / "assets" / "fonts"
FONT_FILES = [
    "FangZhengXiaoBiaoSong.TTF",
    "FangSong_GB2312.ttf",
    "KaiTi_GB2312.ttf",
]


def main() -> None:
    print(f"Optional font dir: {FONT_DIR}")
    for name in FONT_FILES:
        path = FONT_DIR / name
        print(f"{name}: {'OK' if path.exists() else 'MISSING'}")

    fc_match = shutil.which("fc-match")
    if not fc_match:
        print("fc-match not found; skip system font matching check.")
        return

    for query in ["方正小标宋", "仿宋_GB2312", "楷体_GB2312", "FangSong_GB2312", "KaiTi_GB2312"]:
        result = subprocess.run([fc_match, query], capture_output=True, text=True, check=False)
        print(f"fc-match {query}: {result.stdout.strip()}")


if __name__ == "__main__":
    main()
