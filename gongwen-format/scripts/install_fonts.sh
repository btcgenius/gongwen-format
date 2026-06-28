#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
FONT_DIR="$SKILL_DIR/assets/fonts"
FONT_FILES=(
  "FangZhengXiaoBiaoSong.TTF"
  "FangSong_GB2312.ttf"
  "KaiTi_GB2312.ttf"
)

if [[ ! -d "$FONT_DIR" ]]; then
  echo "Optional font directory not found: $FONT_DIR" >&2
  exit 1
fi

for font in "${FONT_FILES[@]}"; do
  if [[ ! -f "$FONT_DIR/$font" ]]; then
    echo "Missing optional font file: $FONT_DIR/$font" >&2
    echo "Font files are not bundled in the public package. Put legally licensed font files in assets/fonts first." >&2
    exit 1
  fi
done

SYSTEM_INSTALL=0
if [[ "${1:-}" == "--system" ]]; then
  SYSTEM_INSTALL=1
fi

OS_NAME="$(uname -s)"

if [[ "$OS_NAME" == "Darwin" ]]; then
  TARGET_DIR="$HOME/Library/Fonts"
elif [[ "$SYSTEM_INSTALL" == "1" ]]; then
  TARGET_DIR="/usr/local/share/fonts/gongwen-format"
else
  TARGET_DIR="$HOME/.local/share/fonts/gongwen-format"
fi

mkdir -p "$TARGET_DIR"

for font in "${FONT_FILES[@]}"; do
  cp "$FONT_DIR/$font" "$TARGET_DIR"/
done

if command -v fc-cache >/dev/null 2>&1; then
  fc-cache -f "$TARGET_DIR" >/dev/null
fi

echo "Installed optional gongwen fonts to: $TARGET_DIR"
echo "Fonts:"
ls -1 "$TARGET_DIR" | grep -E 'FangZheng|FangSong|KaiTi' || true
