#!/usr/bin/env bash
# Finans 101 — lab kontrolü. Hiçbir yere yayın yapmaz, ağa bir şey göndermez; yalnız bu bilgisayarda çalışır.
#
#   ./kontrol.sh            kitapları üret → formülleri çalıştır → bağımsız yeniden hesapla → negatif test
#                           → işlev uyumluluğu → grafikler → görüntüleyici            → GEÇTİ/KALDI
#   ./kontrol.sh --ekran    + bütün ders ekranlarını yeniden çek (ekran/<ders>/*.png)
set -u
cd "$(dirname "$0")"
EKRAN=0
for a in "$@"; do
  case "$a" in
    --ekran) EKRAN=1 ;;
    *) echo "bilinmeyen seçenek: $a"; exit 2 ;;
  esac
done
unset VIRTUAL_ENV
HATA=0
sonuc() { printf '%-34s %s\n' "$1" "$2"; }
adim() {  # adim "ad" komut...
  local ad="$1"; shift
  local cikti
  cikti=$("$@" 2>&1); local kod=$?
  if [ $kod -eq 0 ]; then sonuc "$ad" "GEÇTİ"; else sonuc "$ad" "KALDI"; HATA=1; fi
  echo "$cikti" | grep -v -iE 'warning|^\s*$' | sed 's/^/    /'
}

echo "python $(uv run --frozen python -c 'import sys;print(sys.version.split()[0])') · uv $(uv --version | cut -d' ' -f2)"
adim "ortam (uv sync --frozen)"        uv sync --frozen --quiet
adim "1 kitapları üret"               uv run --frozen python kitaplar/uret.py
adim "2 formülleri çalıştır"          uv run --frozen python kitaplar/hesapla.py
adim "3 bağımsız yeniden hesap"       uv run --frozen python kitaplar/kontrol_degerler.py --sessiz
adim "4 negatif test"                 uv run --frozen python kitaplar/negatif_test.py
adim "5 işlev uyumluluğu"             uv run --frozen python - <<'PY'
# Excel + Google E-Tablolar'da ortak, kursta anlatılan işlevler dışında işlev yok; her kitap açılışta yeniden hesaplar
import re, sys
from pathlib import Path
from openpyxl import load_workbook
IZIN = {"SUM", "SUMIF", "COUNTIF", "IF", "IFERROR", "AND", "ROUND", "ROUNDUP", "MIN", "MAX", "AVERAGE", "STDEV.P", "PMT", "IPMT", "FV"}
kod = 0
for p in sorted(Path("calisma-kitaplari").glob("*.xlsx")):
    wb = load_workbook(p)
    bulunan = set()
    for ws in wb.worksheets:
        for r in ws.iter_rows():
            for c in r:
                if isinstance(c.value, str) and c.value.startswith("="):
                    bulunan |= set(re.findall(r"([A-Z][A-Z0-9.]*)\(", re.sub(r'"[^"]*"', "", c.value)))
    fazla = bulunan - IZIN
    tam = bool(wb.calculation.fullCalcOnLoad)
    print(f"{p.name:32} işlev {len(bulunan):2} · izin dışı {sorted(fazla) or 'yok'} · açılışta yeniden hesap {'evet' if tam else 'HAYIR'}")
    if fazla or not tam:
        kod = 1
sys.exit(kod)
PY
adim "6 grafikler"                    uv run --frozen python grafikler/ciz.py
adim "7 görüntüleyici"                uv run --frozen python goruntuleyici/olustur.py
if [ "$EKRAN" = 1 ]; then
  adim "8 ders ekranları"             uv run --frozen python araclar/ekranlari_uret.py
fi
n=$(ls ekran/*/*.png 2>/dev/null | wc -l)
echo
echo "ekran/ altında $n görüntü · calisma-kitaplari/ altında $(ls calisma-kitaplari/*.xlsx | wc -l) kitap"
if [ "$HATA" = 0 ]; then echo "SONUÇ: GEÇTİ"; else echo "SONUÇ: KALDI"; fi
exit $HATA
