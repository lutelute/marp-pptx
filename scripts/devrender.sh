#!/usr/bin/env bash
# Dev render helper: MD -> PPTX -> PDF -> PNG montages for fast visual review.
# Usage: scripts/devrender.sh [input.md] [palette] [outdir]
set -euo pipefail
cd "$(dirname "$0")/.."
IN="${1:-example.md}"
PALETTE="${2:-}"
OUT="${3:-/tmp/devrender}"
MATH="${4:-png}"   # png renders correctly under LibreOffice for visual checks
SOFFICE="/Applications/LibreOffice.app/Contents/MacOS/soffice"
rm -rf "$OUT" && mkdir -p "$OUT"
PAL_ARG=""
[ -n "$PALETTE" ] && PAL_ARG="-p $PALETTE"
PYTHONPATH=src .venv/bin/python -m marp_pptx convert "$IN" $PAL_ARG --math "$MATH" -o "$OUT/out.pptx"
"$SOFFICE" --headless --convert-to pdf --outdir "$OUT" "$OUT/out.pptx" >/dev/null 2>&1
pdftoppm -png -r 90 "$OUT/out.pdf" "$OUT/slide" >/dev/null 2>&1
# 9-up montages
i=1; n=$(ls "$OUT"/slide-*.png | wc -l | tr -d ' ')
sheet=1
files=( $(ls "$OUT"/slide-*.png) )
total=${#files[@]}
idx=0
while [ $idx -lt $total ]; do
  chunk=( "${files[@]:$idx:9}" )
  montage "${chunk[@]}" -tile 3x3 -geometry 460x+5+5 -background '#888' "$OUT/sheet$sheet.png" 2>/dev/null || true
  idx=$((idx+9)); sheet=$((sheet+1))
done
echo "Rendered $total slides -> $OUT (sheets: $((sheet-1)))"
ls "$OUT"/sheet*.png
