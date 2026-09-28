# Finans 101 — Lab

Bu klasörde kursun bütün uygulamalarında kullanılan **gerçek elektronik tablo dosyaları** (`.xlsx`),
formülleri çalıştırıp her değeri **bağımsız olarak yeniden hesaplayan** kontrol, dosyayı olduğu gibi
gösteren **yerel görüntüleyici** ve gerçek sayılardan çizilen **grafikler** var.

> **ÖRNEK VERİ.** Deniz Örnek (kişi) ve Örnek Mum Atölyesi (işletme) uydurmadır. Bütün tutarlar ve
> oranlar (faiz, enflasyon, getiri) **örnek değerdir; bugünkü piyasa bilgisi değildir.** Hafta 4
> senaryoları ve X/Y varlıkları uydurmadır; ürün, fon, hisse ya da geçmiş piyasa verisi değildir.
> Bu lab **yatırım tavsiyesi değildir**, hukuki ya da mali danışmanlık da değildir.

## 1. Öğrenci için: dosyayı açın

`calisma-kitaplari/` içindeki dosyaları **Google E-Tablolar** (Dosya → İçe aktar → Yükle) ya da
**Excel** ile açın. Sarı hücreler **girdi**dir (değiştirin), beyaz hücreler **formül**dür (elle yazmayın).
Dosyalar açılışta bütün formülleri yeniden hesaplayacak şekilde kaydedilmiştir.

| Dosya | Hafta | Sayfalar |
|---|---|---|
| `hafta1-paranin-temelleri.xlsx` | 1 | Gelir-Gider · Bütçe · Nakit Akışı · Acil Durum Fonu |
| `hafta2-borc-ve-birikim.xlsx` | 2 | Bileşik Faiz · Kredi Kartı · Kredi Planı · Birikim Hedefi · Enflasyon |
| `hafta3-isletme-finansi.xlsx` | 3 | Kâr ve Nakit · Başabaş · Fiyatlama · Nakit Akış Tablosu |
| `hafta4-karar-vermek.xlsx` | 4 | Risk ve Getiri · Çeşitlendirme · Hedef Planı |
| `bitirme-sablonu.xlsx` | 4.4 | Başla (Ç1–Ç8 listesi) · Bütçe · Nakit Akışı · Hedefler · Varsayımlar — boş |
| `bitirme-ornek.xlsx` | 4.4 | aynı şablon, Deniz Örnek verisiyle dolu |

**Türkçe Excel ve işlev adları:** dosyada formüller İngilizce işlev adlarıyla saklanır (dosya biçimi böyledir).
Türkçe Excel aynı formülü Türkçe adla gösterir: `SUM`→`TOPLA`, `SUMIF`→`ETOPLA`, `COUNTIF`→`EĞERSAY`,
`IF`→`EĞER`, `IFERROR`→`EĞERHATA`, `AND`→`VE`, `ROUND`→`YUVARLA`, `ROUNDUP`→`YUKARIYUVARLA`, `MIN`→`MİN`,
`MAX`→`MAK`, `AVERAGE`→`ORTALAMA`, `STDEV.P`→`STDSAPMA.P`, `PMT`→`DEVRESEL_ÖDEME`, `IPMT`→`FAİZTUTARI`,
`FV`→`GD` (Microsoft Destek tr-tr sayfaları, 2026-09-25). Google E-Tablolar varsayılan olarak İngilizce işlev
adlarını kullanır; Dosya → Ayarlar'dan "Her zaman İngilizce işlev adlarını kullan" kapatılırsa Türkçe gösterir.
Bağımsız değişken ayracı bölge ayarına göre virgül ya da noktalı virgül olabilir.

## 2. Tek komutla başlat (kurs ekibi)

Gerekli: [uv](https://docs.astral.sh/uv/) (Python'u ve paketleri kendisi kurar) ve Playwright'ın Chromium'u (yalnız ekran çekimi için).

```bash
./baslat.sh      # kurar, üretir, kontrol eder, görüntüleyiciyi açar → http://127.0.0.1:8765  (Ctrl+C ile durur)
```

Yalnız kontrol:

```bash
./kontrol.sh             # üret → formülleri çalıştır → bağımsız yeniden hesap → negatif test → işlev uyumluluğu → grafikler → görüntüleyici
./kontrol.sh --ekran     # + bütün ders ekranlarını yeniden çek (ekran/<ders>/*.png)
```

Sürümler `pyproject.toml` + `uv.lock` ile sabittir: Python 3.12, openpyxl 3.1.5, formulas 1.3.0,
matplotlib 3.10.6, playwright 1.63.0.

## 3. Hat: dosyadan ekrana

```
kitaplar/uret.py            → calisma-kitaplari/*.xlsx       formüllü gerçek dosyalar (değer hesaplamaz)
kitaplar/hesapla.py         → hesap/*.json                   formül motoru (formulas) her formülü çalıştırır; #HATA varsa KALDI
kitaplar/kontrol_degerler.py                                 1157 formül hücresinin HER BİRİNİ formülsüz, düz Python'la yeniden
                                                             hesaplar; kontrolsüz formül, fark ya da çıpa hatası = KALDI
kitaplar/capalar.json                                        kodun DIŞINDA elle hesaplanmış çıpa değerleri + satır etiketleri
kitaplar/negatif_test.py                                     6 bilinçli bozma (yanlış oran, yanlış bağlantı, elle sayı, girdi
                                                             permütasyonu, eksik aralık) — hepsini yakalamazsa KALDI
goruntuleyici/olustur.py    → goruntuleyici/site/*.html      gerçek dosyanın hücreleri, biçimleri, formülleri + hesaplanan değerler
grafikler/ciz.py            → grafikler/*.png                hesap/*.json'daki değerlerden (hücre adresi grafiğin alt başlığında)
araclar/ekranlari_uret.py   → ekran/<ders>/*.png             görüntüleyici 1280×800 + grafik kopyaları
araclar/terminal.py         → ekran/<ders>/*.png + .txt      gerçek komut çıktısı, 740×462 (Website Yapımı 101 aracıyla aynı)
```

**Neden iki hesap?** `hesapla.py` dosyadaki formülleri Excel'in hesap kurallarıyla çalıştırır;
`kontrol_degerler.py` aynı sonuçları formül metnine hiç bakmadan kendi Python kodu ve kapalı formüllerle
(ör. taksit = anapara × oran / (1 − (1 + oran)^−vade)) yeniden bulur. İkisi 0,000001 içinde aynı olmalı.
Negatif test, kontrolün gerçekten hata yakaladığını kanıtlar (ilk sürümde tolerans 0,005'ti; bir yılı düşürülmüş
standart sapma bu toleransın altında kaldığı için KAÇTI — tolerans 0,000001'e indirildi).

**Bilinen sınır:** makinede Excel ya da LibreOffice yok. Dosyaların gerçek Excel / Google E-Tablolar'da açılıp
aynı değerleri verdiği bu lab'da denenmedi (ACIK-SORULAR Q9). Kullanılan işlevler iki üründe de belgeli
olanlarla sınırlıdır (`kontrol.sh` adım 5).

## 4. Görüntüleyici

`goruntuleyici/site/index.html` — sunucusuz da açılır. Adres seçenekleri: `?sayfa=3&hucre=C8`
(seçili hücre, formül çubuğunda formülü ve Türkçe Excel karşılığı), `&formul=1` (hücrelerde formül metni),
`&satir=1-17,30-37` (yalnız bu satırlar), `&olcek=0.8` (geniş tabloları sığdırmak için).

Bu sayfa bir Excel ya da Google E-Tablolar ekranı **değildir** ve öyle görünmeye çalışmaz: başlığında
"Finans 101 · Lab görüntüleyici" ve dosya yolu yazar. Excel/Sheets menüleri derste resmî belge bağlantısıyla anlatılır.

## 5. Klasörler

```
lab/
├── calisma-kitaplari/   ← ÖĞRENCİ DOSYALARI (.xlsx)
├── kitaplar/            uret.py, hesapla.py, kontrol_degerler.py, capalar.json, negatif_test.py
├── hesap/               formül motoru çıktısı (JSON, üretilir)
├── goruntuleyici/       olustur.py + site/ (üretilir)
├── grafikler/           ciz.py + *.png (üretilir)
├── araclar/             ekran.py, terminal.py, ekranlari_uret.py, ortak.py, fonts/
├── ekran/<ders>/        ders ekranları (.png; terminal çekimlerinde .txt ham çıktı)
├── kontrol.sh, baslat.sh
├── pyproject.toml, uv.lock, .python-version
├── README.md            bu dosya
└── SITE-HARITASI.md     dosya, sayfa, hücre sözleşmesi (dersler buna dayanır)
```
