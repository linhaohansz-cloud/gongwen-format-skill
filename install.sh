#!/usr/bin/env bash
# 国企公文格式 skill —— macOS / Linux 自安装脚本
# 用法：
#   ./install.sh "https://你的直链/国企公文格式-skill.zip"   # 从直链下载
#   ./install.sh /path/to/国企公文格式-skill.zip             # 从本地文件
set -euo pipefail

DEST="$HOME/.workbuddy/skills/国企公文格式"
TMP="$(mktemp -d)"

if [ $# -lt 1 ]; then
  echo "用法: $0 <zip直链或本地路径>"
  exit 1
fi

SRC_ARG="$1"
if [[ "$SRC_ARG" == http* ]]; then
  echo "正在下载: $SRC_ARG"
  curl -L "$SRC_ARG" -o "$TMP/skill.zip"
else
  cp "$SRC_ARG" "$TMP/skill.zip"
fi

cd "$TMP"
unzip -o -q skill.zip

SRC_DIR="$(find "$TMP" -name SKILL.md | head -1 | xargs dirname)"
mkdir -p "$DEST"
cp -r "$SRC_DIR"/. "$DEST"/

echo ""
echo "✅ 已安装到: $DEST"
echo "下次直接对 AI 说：用「国企公文格式」skill 即可。"
