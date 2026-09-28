"""Bütün ders ekranlarını tek seferde üretir (görüntüleyici görüntüleri + gerçek grafikler).

    uv run python araclar/ekranlari_uret.py

Ön koşul: kitaplar/uret.py → kitaplar/hesapla.py → goruntuleyici/olustur.py → grafikler/ciz.py
(./kontrol.sh hepsini sırayla çalıştırır). Çıktı: ekran/<ders>/<ad>.png — liste SITE-HARITASI.md'de.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ekran import cek, tarayici_ac  # noqa: E402
from ortak import LAB, ekran_yolu  # noqa: E402

H1, H2, H3, H4 = "hafta1-paranin-temelleri", "hafta2-borc-ve-birikim", "hafta3-isletme-finansi", "hafta4-karar-vermek"

# (ders, ad, kitap, sayfa, hücre, formül mü, satır aralığı, ölçek)
CEKIMLER = [
    ("1.1", "gelir-gider", H1, 1, "H9", False, None, 1.0),
    ("1.1", "gelir-gider-formuller", H1, 1, "H6", True, None, 1.0),
    ("1.2", "butce", H1, 2, "D15", False, None, 1.0),
    ("1.2", "butce-formuller", H1, 2, "D6", True, None, 1.0),
    ("1.3", "nakit-akisi", H1, 3, "H12", False, None, 0.8),
    ("1.3", "nakit-akisi-formuller", H1, 3, "C6", True, None, 0.8),
    ("1.4", "acil-durum-fonu", H1, 4, "C11", False, "1-17,30-37", 1.0),
    ("2.1", "bilesik-faiz", H2, 1, "C20", False, None, 1.0),
    ("2.1", "bilesik-faiz-formuller", H2, 1, "C10", True, None, 1.0),
    ("2.2", "kredi-karti", H2, 2, "G13", False, None, 1.0),
    ("2.2", "kredi-plani", H2, 3, "C8", False, None, 1.0),
    ("2.2", "kredi-plani-formuller", H2, 3, "C8", True, None, 1.0),
    ("2.3", "birikim-hedefi", H2, 4, "C9", False, "1-18,32-37", 1.0),
    ("2.4", "enflasyon", H2, 5, "C17", False, None, 1.0),
    ("3.1", "kar-ve-nakit", H3, 1, "E18", False, None, 0.9),
    ("3.2", "basabas", H3, 2, "C10", False, None, 1.0),
    ("3.2", "basabas-formuller", H3, 2, "C10", True, "1-16", 1.0),
    ("3.3", "fiyatlama", H3, 3, "C11", False, None, 1.0),
    ("3.4", "nakit-akis-tablosu", H3, 4, "E21", False, None, 0.78),
    ("4.1", "risk-ve-getiri", H4, 1, "D25", False, None, 1.0),
    ("4.2", "cesitlendirme", H4, 2, "F20", False, None, 1.0),
    ("4.3", "hedef-plani", H4, 3, "F18", False, None, 1.0),
    ("4.4", "bitirme-sablonu", "bitirme-sablonu", 1, "C18", False, None, 1.0),
    ("4.4", "bitirme-ornek", "bitirme-ornek", 1, "C18", False, None, 1.0),
    ("4.4", "bitirme-ornek-butce", "bitirme-ornek", 2, "C28", False, None, 1.0),
    ("4.4", "bitirme-ornek-varsayimlar", "bitirme-ornek", 5, "C6", False, None, 1.0),
]
GRAFIKLER = {"1.3": ["1.3-kisisel-bakiye"], "2.1": ["2.1-basit-bilesik"], "2.2": ["2.2-kredi-taksit"],
             "2.3": ["2.3-birikim"], "2.4": ["2.4-alim-gucu"], "3.1": ["3.1-kar-nakit"], "3.2": ["3.2-basabas"],
             "3.4": ["3.4-isletme-kasa"], "4.1": ["4.1-senaryolar"], "4.2": ["4.2-cesitlendirme"], "4.3": ["4.3-hedefler"]}


def main() -> None:
    with tarayici_ac() as t:
        for ders, ad, kitap, sayfa, hucre, formul, satir, olcek in CEKIMLER:
            cek(t, ders, ad, kitap, sayfa, hucre, formul, satir, False, olcek)
    for ders, adlar in GRAFIKLER.items():
        for ad in adlar:
            hedef = ekran_yolu(ders, "grafik-" + ad.split("-", 1)[1])
            shutil.copy(LAB / "grafikler" / f"{ad}.png", hedef)
            print(hedef.relative_to(LAB))


if __name__ == "__main__":
    main()
