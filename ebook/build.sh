#!/usr/bin/env bash
# Rigenera EPUB e PDF da ../GUIDA_WEB_DESIGN_BUSINESS_RU.md
# Requisiti: pandoc, node + playwright (Chromium)
set -euo pipefail
cd "$(dirname "$0")"
python3 - <<'PY'
src = open('../GUIDA_WEB_DESIGN_BUSINESS_RU.md', encoding='utf-8').read()
lines = src.split('\n')
src = '\n'.join(lines[1:])                      # titolo -> metadata.yaml
a, b = src.index('## Содержание'), src.index('# ЧАСТЬ 1')
src = src[:a] + src[b:]                         # indice generato da pandoc
open('build.md', 'w', encoding='utf-8').write(src)
PY
pandoc metadata.yaml build.md -o Guida_Web_Business_RU.epub --toc --toc-depth=2 --split-level=1 --css=style.css
pandoc metadata.yaml build.md -o print.html -s --toc --toc-depth=2 --css=style.css --css=print.css --embed-resources
node make_pdf.js "$PWD/print.html" "$PWD/Guida_Web_Business_RU.pdf"
rm -f build.md print.html
echo "OK: Guida_Web_Business_RU.epub, Guida_Web_Business_RU.pdf"
