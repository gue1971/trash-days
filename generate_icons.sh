#!/bin/bash
# ./generate_icons.sh

SOURCE="source.png"
ICONS_DIR="icons"

if [ ! -f "$SOURCE" ]; then
  echo "エラー: $SOURCE がないで。"
  exit 1
fi

if [ ! -d "$ICONS_DIR" ]; then
  mkdir -p "$ICONS_DIR"
fi

echo "--- アイコン生成開始 ---"
convert "$SOURCE" -resize 512x512 "$ICONS_DIR/icon-512.png"
convert "$SOURCE" -resize 192x192 "$ICONS_DIR/icon-192.png"
convert "$SOURCE" -resize 180x180 "apple-touch-icon.png"

# favicon.ico (32px + 48px)
convert "$SOURCE" -resize 48x48 tmp_48.png
convert "$SOURCE" -resize 32x32 tmp_32.png
convert tmp_48.png tmp_32.png -colors 256 favicon.ico
rm tmp_*.png

echo "✅ 完了！"