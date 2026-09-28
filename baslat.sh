#!/usr/bin/env bash
# Finans 101 lab — tek komutla başlat: ortamı kur, kitapları üret ve kontrol et, görüntüleyiciyi aç.
#   ./baslat.sh            → http://127.0.0.1:8765  (durdurmak: Ctrl+C)
# Yalnız bu bilgisayarda (127.0.0.1) dinler; dışarıya açılmaz.
set -e
cd "$(dirname "$0")"
unset VIRTUAL_ENV
./kontrol.sh
echo
echo "Görüntüleyici: http://127.0.0.1:${PORT:-8765}/   (Ctrl+C ile durdurun)"
exec uv run --frozen python -m http.server "${PORT:-8765}" --bind 127.0.0.1 --directory goruntuleyici/site
