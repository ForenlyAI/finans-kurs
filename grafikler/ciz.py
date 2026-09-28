"""Grafikleri GERÇEK hesaplanmış değerlerden çizer (hesap/<kitap>.json → grafikler/*.png, 1280×720).

Kullanım (lab klasöründen, hesapla.py'den sonra):
    uv run python grafikler/ciz.py

Sayılar elle yazılmaz: her nokta çalışma kitabındaki bir formül hücresinin değeridir (hücre adresi koddadır).
Renkler: kategorik sıra mavi → turuncu → yeşil-mavi (dataviz doğrulayıcısından geçti; yeşil-mavi düşük
kontrastlı olduğu için her seri doğrudan etiketlidir). Tek eksen; ızgara soluk.
"""

from __future__ import annotations

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.ticker import FuncFormatter  # noqa: E402

LAB = Path(__file__).resolve().parent.parent
HEDEF = Path(__file__).resolve().parent
MAVI, TURUNCU, YESIL = "#2a78d6", "#eb6834", "#1baf7a"
YAZI, SOLUK, ZEMIN = "#0b0b0b", "#52514e", "#fcfcfb"
AYLAR = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"]
AY = [chr(c) for c in range(ord("B"), ord("N"))]

plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 13, "axes.edgecolor": "#c9c8c2",
                     "axes.labelcolor": SOLUK, "xtick.color": SOLUK, "ytick.color": SOLUK,
                     "axes.spines.top": False, "axes.spines.right": False, "figure.facecolor": ZEMIN,
                     "axes.facecolor": ZEMIN, "axes.grid": True, "grid.color": "#e6e5df", "grid.linewidth": 0.8})


def v(kitap: str, sayfa: str, hucre: str):
    return VERI[kitap][sayfa][hucre]["v"]


def tl(x, _=None):
    s = f"{abs(x):,.0f}".replace(",", ".")
    return ("−" if x < 0 else "") + s + " TL"


def yuzde(x, _=None):
    return "%" + f"{x * 100:.0f}"


def yuzde1(x):
    return "%" + f"{x * 100:.1f}".replace(".", ",")


def yeni(baslik: str, alt: str):
    fig, ax = plt.subplots(figsize=(12.8, 7.2), dpi=100)
    fig.subplots_adjust(left=0.1, right=0.95, top=0.83, bottom=0.12)
    fig.text(0.1, 0.93, baslik, fontsize=20, fontweight="bold", color=YAZI)
    fig.text(0.1, 0.885, alt, fontsize=12, color=SOLUK)
    fig.text(0.95, 0.02, "ÖRNEK VERİ · Finans 101 lab · kaynak: calisma-kitaplari/", fontsize=10, color=SOLUK, ha="right")
    ax.set_axisbelow(True)
    return fig, ax


def sag_bosluk(ax, oran=0.3):
    a, b = ax.get_xlim()
    veri_son = max(max(line.get_xdata()) for line in ax.get_lines() if len(line.get_xdata()))
    ax.set_xticks([x for x in ax.get_xticks() if a <= x <= veri_son])
    ax.set_xlim(a, b + (b - a) * oran)


def kaydet(fig, ad: str):
    yol = HEDEF / f"{ad}.png"
    fig.savefig(yol, facecolor=ZEMIN)
    plt.close(fig)
    print(yol.relative_to(LAB))


def cubuk_bakiye(kitap, sayfa, satir, baslik, alt, ad):
    deg = [v(kitap, sayfa, f"{c}{satir}") for c in AY]
    fig, ax = yeni(baslik, alt)
    renkler = [MAVI if x >= 0 else TURUNCU for x in deg]
    ax.bar(AYLAR, deg, color=renkler, width=0.62, edgecolor=ZEMIN, linewidth=2)
    ax.axhline(0, color=SOLUK, linewidth=1)
    for i, x in enumerate(deg):
        if x < 0:
            ax.annotate(f"AÇIK\n{tl(x)}", (i, x), xytext=(0, -8), textcoords="offset points", ha="center", va="top",
                        fontsize=11, color=YAZI)
    ax.annotate(tl(deg[-1]), (11, deg[-1]), xytext=(0, 6), textcoords="offset points", ha="center", fontsize=11, color=YAZI)
    ax.yaxis.set_major_formatter(FuncFormatter(tl))
    ax.set_ylim(min(deg) * 2.3 if min(deg) < 0 else 0, max(deg) * 1.15)
    ax.grid(axis="x", visible=False)
    kaydet(fig, ad)


def main():
    global VERI
    VERI = {}
    for p in (LAB / "hesap").glob("*.json"):
        VERI[p.stem] = json.loads(p.read_text(encoding="utf-8"))["sayfalar"]
    h1, h2, h3, h4 = "hafta1-paranin-temelleri", "hafta2-borc-ve-birikim", "hafta3-isletme-finansi", "hafta4-karar-vermek"

    # 1.3
    cubuk_bakiye(h1, "Nakit Akışı", 12, "Ay sonu bakiyesi: yıl artıda, iki ay eksi",
                 "Deniz Örnek (uydurma) · 12 aylık kişisel nakit akışı · Nakit Akışı!B12:M12", "1.3-kisisel-bakiye")

    # 2.1
    yil = list(range(11))
    basit = [v(h2, "Bileşik Faiz", f"B{10 + y}") for y in yil]
    bil = [v(h2, "Bileşik Faiz", f"C{10 + y}") for y in yil]
    fig, ax = yeni("Basit faiz düz çizgi, bileşik faiz eğri", f"100.000 TL · yıllık %{v(h2, 'Bileşik Faiz', 'C6') * 100:.0f} ÖRNEK oran · Bileşik Faiz!B10:C20")
    ax.plot(yil, basit, color=MAVI, linewidth=2.5, marker="o", markersize=7)
    ax.plot(yil, bil, color=TURUNCU, linewidth=2.5, marker="o", markersize=7)
    ax.annotate(f"basit: {tl(basit[-1])}", (10, basit[-1]), xytext=(10, 0), textcoords="offset points", ha="left", va="center", color=YAZI)
    ax.annotate(f"bileşik: {tl(bil[-1])}", (10, bil[-1]), xytext=(10, 0), textcoords="offset points", ha="left", va="center", color=YAZI)
    ax.yaxis.set_major_formatter(FuncFormatter(tl))
    ax.set_xlabel("yıl")
    ax.set_xticks(yil)
    ax.legend(["basit faiz", "bileşik faiz"], frameon=False, loc="upper left")
    sag_bosluk(ax)
    kaydet(fig, "2.1-basit-bilesik")

    # 2.2
    faiz = [v(h2, "Kredi Planı", f"D{13 + i}") for i in range(12)]
    ana = [v(h2, "Kredi Planı", f"E{13 + i}") for i in range(12)]
    fig, ax = yeni("Taksit aynı, içindeki faiz her ay azalır",
                   f"50.000 TL kredi · aylık %3 ÖRNEK · 12 ay · toplam faiz {tl(v(h2, 'Kredi Planı', 'C10'))} · Kredi Planı!D13:E24")
    x = list(range(1, 13))
    ax.bar(x, ana, color=MAVI, width=0.62, label="anapara", edgecolor=ZEMIN, linewidth=2)
    ax.bar(x, faiz, bottom=ana, color=TURUNCU, width=0.62, label="faiz", edgecolor=ZEMIN, linewidth=2)
    ax.annotate(f"faiz {tl(faiz[0])}", (1, ana[0] + faiz[0]), xytext=(0, 6), textcoords="offset points", ha="center", fontsize=11)
    ax.annotate(f"faiz {tl(faiz[-1])}", (12, ana[-1] + faiz[-1]), xytext=(0, 6), textcoords="offset points", ha="center", fontsize=11)
    ax.yaxis.set_major_formatter(FuncFormatter(tl))
    ax.set_xticks(x)
    ax.set_xlabel("ay")
    ax.set_ylim(0, (ana[0] + faiz[0]) * 1.15)
    ax.grid(axis="x", visible=False)
    ax.legend(frameon=False, loc="upper left", ncol=2)
    kaydet(fig, "2.2-kredi-taksit")

    # 2.3
    son = [v(h2, "Birikim Hedefi", f"E{14 + i}") for i in range(24)]
    katki = v(h2, "Birikim Hedefi", "C9")
    mevcut = v(h2, "Birikim Hedefi", "C8")
    konan = [mevcut + katki * (i + 1) for i in range(24)]
    fig, ax = yeni("Hedefe 24 ayda: katkı + getiri",
                   f"hedef 120.000 TL · aylık katkı {tl(katki)} · aylık %1,5 ÖRNEK getiri · Birikim Hedefi!E14:E37")
    x = list(range(1, 25))
    ax.plot(x, son, color=MAVI, linewidth=2.5)
    ax.plot(x, konan, color=TURUNCU, linewidth=2.5)
    ax.annotate(f"ay sonu bakiye {tl(son[-1])}", (24, son[-1]), xytext=(10, 0), textcoords="offset points", ha="left", va="center")
    ax.annotate(f"cepten konan {tl(konan[-1])}", (24, konan[-1]), xytext=(10, 0), textcoords="offset points", ha="left", va="center")
    ax.yaxis.set_major_formatter(FuncFormatter(tl))
    ax.set_xlabel("ay")
    ax.legend(["bakiye", "cepten konan toplam"], frameon=False, loc="upper left")
    sag_bosluk(ax)
    kaydet(fig, "2.3-birikim")

    # 2.4
    yil = list(range(6))
    fiyat = [v(h2, "Enflasyon", f"B{10 + y}") for y in yil]
    guc = [v(h2, "Enflasyon", f"C{10 + y}") for y in yil]
    fig, ax = yeni("Aynı sepet pahalanır, aynı para küçülür",
                   f"bugün 10.000 TL'lik sepet · yıllık %{v(h2, 'Enflasyon', 'C6') * 100:.0f} ÖRNEK enflasyon · Enflasyon!B10:C15")
    ax.plot(yil, fiyat, color=TURUNCU, linewidth=2.5, marker="o", markersize=7)
    ax.plot(yil, guc, color=MAVI, linewidth=2.5, marker="o", markersize=7)
    ax.annotate(f"sepetin fiyatı {tl(fiyat[-1])}", (5, fiyat[-1]), xytext=(10, 0), textcoords="offset points", ha="left", va="center")
    ax.annotate(f"10.000 TL'nin alım gücü {tl(guc[-1])}", (5, guc[-1]), xytext=(10, 0), textcoords="offset points", ha="left", va="center")
    ax.yaxis.set_major_formatter(FuncFormatter(tl))
    ax.set_xticks(yil)
    ax.set_xlabel("yıl")
    ax.set_ylim(0, max(fiyat) * 1.12)
    ax.legend(["aynı sepetin fiyatı", "paranın alım gücü (bugünkü TL)"], frameon=False, loc="upper left")
    sag_bosluk(ax)
    kaydet(fig, "2.4-alim-gucu")

    # 3.1
    alti = AY[:6]
    kar = [v(h3, "Kâr ve Nakit", f"{c}16") for c in alti]
    nakit = [v(h3, "Kâr ve Nakit", f"{c}17") for c in alti]
    fig, ax = yeni("Kâr artıyor, kasa hâlâ eksi", "Örnek Mum Atölyesi (uydurma) · birikimli kâr ve birikimli nakit · Kâr ve Nakit!B16:G17")
    ax.plot(AYLAR[:6], kar, color=MAVI, linewidth=2.5, marker="o", markersize=8)
    ax.plot(AYLAR[:6], nakit, color=TURUNCU, linewidth=2.5, marker="o", markersize=8)
    ax.axhline(0, color=SOLUK, linewidth=1)
    ax.annotate(f"birikimli kâr {tl(kar[-1])}", (5, kar[-1]), xytext=(10, 0), textcoords="offset points", ha="left", va="center")
    ax.annotate(f"birikimli nakit {tl(nakit[-1])}", (5, nakit[-1]), xytext=(10, 0), textcoords="offset points", ha="left", va="center")
    ax.yaxis.set_major_formatter(FuncFormatter(tl))
    ax.legend(["birikimli kâr", "birikimli nakit"], frameon=False, loc="upper left")
    ax.set_xlim(-0.3, 7.2)
    kaydet(fig, "3.1-kar-nakit")

    # 3.2
    adet = [v(h3, "Başabaş", f"A{17 + i}") for i in range(11)]
    ciro = [v(h3, "Başabaş", f"B{17 + i}") for i in range(11)]
    mal = [v(h3, "Başabaş", f"E{17 + i}") for i in range(11)]
    bb = int(v(h3, "Başabaş", "C10"))
    fig, ax = yeni(f"Başabaş: ayda {bb} mum", "Örnek Mum Atölyesi (uydurma) · fiyat 250 TL · değişken maliyet 130 TL · sabit gider 30.000 TL · Başabaş!A17:F27")
    ax.plot(adet, ciro, color=MAVI, linewidth=2.5)
    ax.plot(adet, mal, color=TURUNCU, linewidth=2.5)
    ax.axvline(bb, color=SOLUK, linewidth=1, linestyle="--")
    ax.annotate(f"başabaş {bb} adet · {tl(v(h3, 'Başabaş', 'C11'))}", (bb, bb * 250), xytext=(10, -30), textcoords="offset points")
    ax.annotate("ciro", (500, ciro[-1]), xytext=(10, 0), textcoords="offset points", ha="left", va="center")
    ax.annotate("toplam maliyet", (500, mal[-1]), xytext=(10, 0), textcoords="offset points", ha="left", va="center")
    ax.yaxis.set_major_formatter(FuncFormatter(tl))
    ax.set_xlabel("aylık satış (adet)")
    ax.legend(["ciro", "toplam maliyet"], frameon=False, loc="upper left")
    sag_bosluk(ax)
    kaydet(fig, "3.2-basabas")

    # 3.4
    cubuk_bakiye(h3, "Nakit Akış Tablosu", 21, "Ekipman alınan ay kasa eksiye düşüyor",
                 "Örnek Mum Atölyesi (uydurma) · ay sonu kasa · Nakit Akış Tablosu!B21:M21", "3.4-isletme-kasa")

    # 4.1
    fig, ax = yeni("Ortalaması en yüksek senaryo en çok kazandırmadı",
                   "UYDURMA senaryolar — ürün ya da piyasa verisi değildir · yatırım tavsiyesi değildir · Risk ve Getiri!E9:G18")
    x = list(range(0, 11))
    for col, renk, ad, gs in (("E", MAVI, "A", "B"), ("F", TURUNCU, "B", "C"), ("G", YESIL, "C", "D")):
        yol = [v(h4, "Risk ve Getiri", "C5")] + [v(h4, "Risk ve Getiri", f"{col}{9 + i}") for i in range(10)]
        ax.plot(x, yol, color=renk, linewidth=2.5)
        ort = v(h4, "Risk ve Getiri", f"{gs}20")
        bil = v(h4, "Risk ve Getiri", f"{gs}25")
        ax.annotate(f"{ad}: ortalama {yuzde1(ort)}, bileşik {yuzde1(bil)}", (10, yol[-1]), xytext=(8, 0),
                    textcoords="offset points", va="center", fontsize=11)
    ax.set_xlim(0, 13.8)
    ax.yaxis.set_major_formatter(FuncFormatter(tl))
    ax.set_xlabel("yıl")
    ax.set_xticks(range(11))
    ax.legend(["Senaryo A", "Senaryo B", "Senaryo C"], frameon=False, loc="upper left")
    kaydet(fig, "4.1-senaryolar")

    # 4.2
    agirlik = ["%0 X", "%25 X", "%50 X", "%75 X", "%100 X"]
    std = [v(h4, "Çeşitlendirme", f"{c}20") for c in "DEFGH"]
    ort = [v(h4, "Çeşitlendirme", f"{c}19") for c in "DEFGH"]
    fig, ax = yeni("Karışım, iki varlıktan da az dalgalanıyor",
                   "UYDURMA iki varlık X ve Y · dalgalanma = yıllık getirinin standart sapması · Çeşitlendirme!D19:H20")
    ax.bar(agirlik, std, color=MAVI, width=0.6, edgecolor=ZEMIN, linewidth=2)
    for i, (s_, o) in enumerate(zip(std, ort)):
        ax.annotate(f"dalgalanma {yuzde1(s_)}\nortalama {yuzde1(o)}", (i, s_), xytext=(0, 6),
                    textcoords="offset points", ha="center", fontsize=11)
    ax.yaxis.set_major_formatter(FuncFormatter(yuzde))
    ax.set_ylim(0, max(std) * 1.3)
    ax.set_xlabel("X'in payı (kalanı Y)")
    ax.grid(axis="x", visible=False)
    kaydet(fig, "4.2-cesitlendirme")

    # 4.3
    hedef = [v(h4, "Hedef Planı", f"A{10 + i}") for i in range(4)]
    katki = [v(h4, "Hedef Planı", f"F{10 + i}") for i in range(4)]
    dahil = [v(h4, "Hedef Planı", f"G{10 + i}") for i in range(4)]
    ayir = v(h4, "Hedef Planı", "C7")
    fig, ax = yeni(f"Plana dahil hedefler: {tl(v(h4, 'Hedef Planı', 'F15'))} / ay · ayrılabilen {tl(ayir)}",
                   "Deniz Örnek (uydurma) · hedef başına gereken aylık katkı · Hedef Planı!F10:G13")
    renk = [MAVI if d == "Evet" else "#c9c8c2" for d in dahil]
    ax.barh(hedef[::-1], katki[::-1], color=renk[::-1], height=0.55, edgecolor=ZEMIN, linewidth=2)
    for i, (k_, d) in enumerate(zip(katki[::-1], dahil[::-1])):
        ax.annotate(f"{tl(k_)} · {'plana dahil' if d == 'Evet' else 'sonraya'}", (k_, i), xytext=(6, 0),
                    textcoords="offset points", va="center", fontsize=11)
    ax.xaxis.set_major_formatter(FuncFormatter(tl))
    ax.set_xlim(0, max(katki) * 1.45)
    ax.grid(axis="y", visible=False)
    fig.subplots_adjust(left=0.27)
    kaydet(fig, "4.3-hedefler")


if __name__ == "__main__":
    main()
