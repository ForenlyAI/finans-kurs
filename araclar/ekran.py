"""Gerçek çalışma kitabının görüntüleyicideki ekran görüntüsü (Playwright + Chromium, 1280×800).

Kullanım (lab klasöründen; önce ./kontrol.sh ya da goruntuleyici/olustur.py çalışmış olmalı):
    uv run python araclar/ekran.py <ders> <ad> <kitap> [--sayfa N] [--hucre C8] [--formul] [--satir 5-20] [--olcek 0.9] [--tam]
    uv run python araclar/ekran.py <ders> <ad> --png grafikler/x.png      # hazır grafiği ders klasörüne kopyalar

Örnek:
    uv run python araclar/ekran.py 2.2 kredi-plani hafta2-borc-ve-birikim --sayfa 3 --hucre C8

Görüntüleyici statik bir HTML'dir (goruntuleyici/site/); sunucu gerekmez, dosya doğrudan açılır.
Yazı tipi: sistemdeki DejaVu Sans / DejaVu Sans Mono (Türkçe karakterleri tam).
"""

from __future__ import annotations

import argparse
import contextlib
import shutil
import sys
from pathlib import Path
from urllib.parse import urlencode

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ortak import LAB, MASAUSTU, SITE, ekran_yolu  # noqa: E402


@contextlib.contextmanager
def tarayici_ac():
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        t = p.chromium.launch()
        try:
            yield t
        finally:
            t.close()


def adres(kitap: str, sayfa: int = 1, hucre: str | None = None, formul: bool = False, satir: str | None = None,
          olcek: float = 1.0) -> str:
    q = {"sayfa": sayfa}
    if olcek != 1.0:
        q["olcek"] = olcek
    if hucre:
        q["hucre"] = hucre
    if formul:
        q["formul"] = 1
    if satir:
        q["satir"] = satir
    return (SITE / f"{kitap}.html").as_uri() + "?" + urlencode(q)


def cek(tarayici, ders: str, ad: str, kitap: str, sayfa: int = 1, hucre: str | None = None,
        formul: bool = False, satir: str | None = None, tam: bool = False, olcek: float = 1.0) -> Path:
    s = tarayici.new_page(viewport=MASAUSTU, locale="tr-TR", reduced_motion="reduce")
    s.goto(adres(kitap, sayfa, hucre, formul, satir, olcek))
    s.wait_for_selector("body[data-hazir='1']")
    s.evaluate("document.fonts.ready.then(() => true)")
    if hucre:  # seçili hücre görünür alanda olsun
        s.evaluate(f"document.querySelector('section.acik td[data-a=\"{hucre.upper()}\"]')?.scrollIntoView({{block:'nearest', inline:'nearest'}})")
    s.wait_for_timeout(120)
    yol = ekran_yolu(ders, ad)
    s.screenshot(path=str(yol), full_page=tam)
    s.close()
    print(yol.relative_to(LAB))
    return yol


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("ders")
    ap.add_argument("ad")
    ap.add_argument("kitap", nargs="?")
    ap.add_argument("--sayfa", type=int, default=1)
    ap.add_argument("--hucre")
    ap.add_argument("--formul", action="store_true")
    ap.add_argument("--satir")
    ap.add_argument("--tam", action="store_true")
    ap.add_argument("--olcek", type=float, default=1.0)
    ap.add_argument("--png")
    a = ap.parse_args()
    if a.png:
        hedef = ekran_yolu(a.ders, a.ad)
        shutil.copy(LAB / a.png, hedef)
        print(hedef.relative_to(LAB))
        return
    with tarayici_ac() as t:
        cek(t, a.ders, a.ad, a.kitap, a.sayfa, a.hucre, a.formul, a.satir, a.tam, a.olcek)


if __name__ == "__main__":
    main()
