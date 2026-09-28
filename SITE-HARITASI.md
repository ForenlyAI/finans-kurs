# Finans 101 Lab — Dosya ve Hücre Haritası

İçerik dersleri bu dosyayı **sözleşme** olarak kullanır: ders metninde geçen dosya adı, sayfa (sekme) adı,
hücre adresi ve değer burada yazdığı gibidir. Bir ad ya da adres değişirse önce `kitaplar/uret.py`, sonra
`kitaplar/kontrol_degerler.py` + `kitaplar/capalar.json`, sonra bu dosya, en son ders metni güncellenir;
`./kontrol.sh` GEÇTİ demeden ekran çekilmez.

Kök: `lab/calisma-kitaplari/`. Değerler 2026-09-25 `./kontrol.sh` çıktısıdır (hesap/*.json). Bütün kişi,
işletme, tutar ve oranlar **ÖRNEK**tir. Ortak düzen: A1 başlık · A2 ÖRNEK/uydurma notu · A3 renk açıklaması ·
sarı = girdi, beyaz = formül.

## hafta1-paranin-temelleri.xlsx — Deniz Örnek (uydurma kişi)

| Ders | Sayfa | Girdi | Anahtar hücre (değer) | Kullanılan işlevler |
|---|---|---|---|---|
| 1.1 | `Gelir-Gider` | A6:E25 — 20 işlem (gün, açıklama, kategori, tür Gelir/Gider/Birikim, tutar) | H6 toplam gelir **45.500 TL** · H7 toplam gider **36.600 TL** · H8 birikim **4.000 TL** · H9 kalan **4.900 TL** · H13:H21 kategori toplamları · H23 "EVET" (kategoriler = gider) | SUMIF, SUM, IF, ROUND, IFERROR |
| 1.2 | `Bütçe` | B6:B14 zorunlu mu (Evet/Hayır) · C6:C14 plan · C20 birikim planı 4.000 | D6:D14 gerçekleşen (Gelir-Gider'den SUMIF) · C15 plan **33.700** / D15 gerçek **36.600** · E15 **−2.900** · G = "AŞILDI" (Market, Dışarıda yemek, Abonelik, Giyim, Diğer, Faturalar) · D17 zorunlu **30.050** · D21 kalan **4.900** | SUMIF, SUM, IF, IFERROR |
| 1.3 | `Nakit Akışı` | B6 açılış 2.000 · satır 8 düzensiz gelir (Kasım 5.000, Aralık 15.000) · satır 10 dönemsel gider (Mart 14.000, Temmuz 25.000, Eylül 12.000, Aralık 6.000) | B12:M12 kapanış; H12 Temmuz **−2.700** ve J12 Eylül **−4.900** "AÇIK" · M12 **23.800** · B15 en düşük **−4.900** · B16 **2** ay açık | SUM, MIN, COUNTIF, IF |
| 1.4 | `Acil Durum Fonu` | C6 hedef ay 3 (öğrenci seçer) · C8 mevcut 12.000 · C10 aylık 4.000 | C5 zorunlu gider **30.050** · C7 hedef **90.150** · C9 eksik **78.150** · C11 **20 ay** · A14:E37 24 aylık tablo, 20. ayda "EVET" | MIN, MAX, IF, ROUNDUP |

## hafta2-borc-ve-birikim.xlsx — Deniz Örnek

| Ders | Sayfa | Girdi | Anahtar hücre (değer) | İşlevler |
|---|---|---|---|---|
| 2.1 | `Bileşik Faiz` | C5 100.000 · C6 yıllık %25 ÖRNEK · C7 10 yıl | B10:D20 yıl yıl basit / bileşik / fark · B20 basit **350.000** · C20 bileşik **931.323** · C22 FV kontrolü = C20 | FV |
| 2.2 | `Kredi Kartı` | C5 dönem borcu 20.000 · C6 akdi %4 ÖRNEK · C7 gecikme %4,3 ÖRNEK · C8 asgari oran %20 · B11:B14 ödeme oranları | G11:G14 bu ayın faizi: tamamı **0** · yarısı **400** · asgari **640** · asgarinin altı **774** (gecikme) | IF |
| 2.2 | `Kredi Planı` | C5 50.000 · C6 aylık %3 ÖRNEK · C7 12 ay | C8 taksit **5.023,10** · C9 toplam **60.277,25** · C10 toplam faiz **10.277,25** · A13:F24 ödeme planı (1. ay faiz 1.500,00 → 12. ay 146,30) · C27 IPMT kontrolü **1.500,00** | PMT, IPMT, SUM |
| 2.3 | `Birikim Hedefi` | C5 hedef 120.000 · C6 24 ay · C7 aylık %1,5 ÖRNEK · C8 mevcut 10.000 | C9 aylık katkı **3.691,65** · C10 getirisiz **4.583,33** · C11 getirinin katkısı **21.400,37** · E37 **120.000,00** | PMT |
| 2.4 | `Enflasyon` | C5 sepet 10.000 · C6 enflasyon %20 ÖRNEK · C7 birikim faizi %22 ÖRNEK | B15 5 yıl sonra sepet **24.883** · C15 alım gücü **4.019** · C17 reel getiri **%1,7** · C18 hedefin 2 yıl sonraki karşılığı **172.800** · C19 yeni aylık katkı **5.535,64** | PMT |

## hafta3-isletme-finansi.xlsx — Örnek Mum Atölyesi (uydurma işletme)

| Ders | Sayfa | Girdi | Anahtar hücre (değer) | İşlevler |
|---|---|---|---|---|
| 3.1 | `Kâr ve Nakit` | C5 peşin oran %40 (kalanı 2 ay sonra) · B9:G9 satış · B10:G10 gider | satır 11 kâr, 15 nakit akışı · G16 birikimli kâr **98.000** · G17 birikimli nakit **−16.000** · D18:G18 "EVET" (kâr artıda, nakit ekside) · C20 başlangıçta gereken nakit **60.000** | SUM, MIN, IF, AND |
| 3.2 | `Başabaş` | C5 sabit gider 30.000 · C6 fiyat 250 · C7 değişken maliyet 130 · C12 beklenen 400 adet | C8 katkı payı **120** · C9 **%48** · C10 başabaş **250 adet** · C11 **62.500 TL** · C13 güvenlik payı **150** · C14 beklenen kâr **18.000** · A17:F27 0–500 adet tablosu | ROUNDUP |
| 3.3 | `Fiyatlama` | C10 maliyet üstü oran %40 ÖRNEK · A15:A17, A19 fiyat seçenekleri (A18 = C11) | C8 birim sabit pay **75** · C9 tam maliyet **205** · C11 maliyet artı fiyat **287** · C12 marj **%28,6** · C15:C19 başabaş adedi (200 TL → 429 … 320 TL → 158) | ROUNDUP, IFERROR |
| 3.4 | `Nakit Akış Tablosu` | B10:M10 satış adedi · C7 peşin %50 · E17 Nisan ekipman 40.000 · B20 açılış 50.000 | satır 21 kapanış; E21 Nisan **−14.400**, F21 Mayıs **−6.150** "KASA AÇIĞI" · M21 **83.800** · C25 **2** ay | SUM, MIN, COUNTIF, IF |

## hafta4-karar-vermek.xlsx — UYDURMA senaryolar (yatırım tavsiyesi değildir)

| Ders | Sayfa | Girdi | Anahtar hücre (değer) | İşlevler |
|---|---|---|---|---|
| 4.1 | `Risk ve Getiri` | C5 10.000 · C6 enflasyon %20 ÖRNEK · B9:D18 üç senaryonun 10 yıllık uydurma getirisi | ortalama B20/C20/D20 **%21,3 / %20,7 / %30,5** · dalgalanma B23:D23 **%0,9 / %13,5 / %56,7** · 10 yıl sonu B24:D24 **68.943 / 61.420 / 45.156** · bileşik B25:D25 **%21,3 / %19,9 / %16,3** · reel B26:D26 **%1,1 / −%0,1 / −%3,1** | AVERAGE, MIN, MAX, STDEV.P |
| 4.2 | `Çeşitlendirme` | D6:H6 X payı %0/25/50/75/100 · B8:C17 X ve Y uydurma getirileri | D20:H20 dalgalanma **%12,5 / %4,9 / %3,0 / %10,4 / %18,0** · D19:H19 ortalama %11,0 … %12,5 · D22:H22 zararlı yıl 2 / 0 / 0 / 2 / 3 | AVERAGE, STDEV.P, MIN, COUNTIF |
| 4.3 | `Hedef Planı` | C5 enflasyon %20 · C6 aylık getiri %1 · C7 ayrılabilen 9.000 · A10:D13 dört hedef · G10:G13 plana dahil mi | E10:E13 o günkü tutar (72.000 / 36.000 / 59.154 / 29.102) · F10:F13 aylık katkı (5.677 / 2.839 / 3.016 / 2.782) · F15 dahil toplam **8.516** · F18 **"PLAN YETİYOR"** (dördü birden dahil edilirse 14.313 → "PLAN YETMİYOR") | PMT, IFERROR, SUMIF, IF |

## bitirme-sablonu.xlsx / bitirme-ornek.xlsx (4.4)

| Sayfa | İçerik |
|---|---|
| `Başla` | B8:B15 ölçütler Ç1–Ç8 · C8:C15 Evet/Hayır (girdi) · C17 `COUNTIF` sayısı · C18 "TESLİME HAZIR" / "EKSİK VAR" |
| `Bütçe` | A6:D15 on kategori satırı (şablonda boş) · C16/D16 toplam · C18 gelir · C19 birikim · C20 kalan · C21 zorunlu gider · acil durum fonu: C24 kaç ay (girdi) · C25 hedef · C26 mevcut · C27 aylık · C28 kaç ayda (örnekte **20**) |
| `Nakit Akışı` | 1.3 ile aynı düzen; satır 7/9/11 Bütçe'ye bağlı |
| `Hedefler` | 4.3 ile aynı düzen (A10:G13) |
| `Varsayımlar` | A6:D11 varsayım · değer · kaynak (ya da "ÖRNEK — kaynak yok") · tarih |

Örnek dosyada C18 = "TESLİME HAZIR", şablonda "EKSİK VAR". Şablonda boş girdilerle hiçbir hücre hata vermez (Ç1).

## Ekranlar (ekran/<ders>/) — 2026-09-25

| Ders | Görüntü (1280×800 görüntüleyici; `grafik-*` 1280×720) |
|---|---|
| 1.1 | gelir-gider · gelir-gider-formuller |
| 1.2 | butce · butce-formuller |
| 1.3 | nakit-akisi · nakit-akisi-formuller · grafik-kisisel-bakiye |
| 1.4 | acil-durum-fonu |
| 2.1 | bilesik-faiz · bilesik-faiz-formuller · grafik-basit-bilesik |
| 2.2 | kredi-karti · kredi-plani · kredi-plani-formuller · grafik-kredi-taksit |
| 2.3 | birikim-hedefi · grafik-birikim |
| 2.4 | enflasyon · grafik-alim-gucu |
| 3.1 | kar-ve-nakit · grafik-kar-nakit |
| 3.2 | basabas · basabas-formuller · grafik-basabas |
| 3.3 | fiyatlama |
| 3.4 | nakit-akis-tablosu · grafik-isletme-kasa |
| 4.1 | risk-ve-getiri · grafik-senaryolar |
| 4.2 | cesitlendirme · grafik-cesitlendirme |
| 4.3 | hedef-plani · grafik-hedefler |
| 4.4 | bitirme-sablonu · bitirme-ornek · bitirme-ornek-butce · bitirme-ornek-varsayimlar · kontrol-degerler (terminal) · negatif-test (terminal) |

Excel ya da Google E-Tablolar'ın kendi arayüzü ekranda **gösterilmez** (kurs malı olmayan arayüz, sürüm
farkı); menü adımları resmî yardım sayfası bağlantısıyla anlatılır.
