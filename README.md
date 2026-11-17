# Selanik – 100 Dairelik Gayrimenkul Yatırım Araştırması (2026)

**Ana teslimat:** [`Selanik_100_Daire_Yatirim_Analizi_2026.xlsx`](Selanik_100_Daire_Yatirim_Analizi_2026.xlsx)

Veri kontrol tarihi: **17.11.2026**

## İçerik

| Sekme | Açıklama |
|---|---|
| YÖNETİCİ PANELİ | KPI'lar, en iyi 10 fırsat, bölgesel karşılaştırma, grafikler |
| 100 DAİRE ANA LİSTESİ | 103 sütun: genel bilgiler, konum, özellikler, maliyet ve gelir, puanlar (filtre, sabit başlık, renk kodları, tıklanabilir ilan ve harita bağlantıları) |
| EN İYİ 20 YATIRIM | Genel puana göre ilk 20 (neden seçildiği ve en büyük risk) ve salt net getiriye göre ilk 20 |
| KİRA EMSALLERİ | 598 kiralık emsal (her daire için 5–6 adet), tıklanabilir |
| MAHALLE ANALİZİ | Satış ve kira €/m², brüt getiri, kira çarpanı, talep ve boş kalma riski |
| SENARYO ANALİZİ | Kötümser, baz ve iyimser senaryo varsayımları (düzenlenebilir) ve 100 dairenin üç senaryodaki sonuçları |
| 5 VE 10 YILLIK NAKİT AKIŞI | Seçilen daire için yıl yıl tablo, 100 daire için 5 ve 10 yıllık IRR |
| RİSK VE EKSİK VERİLER | 14 risk işareti (ağırlıklar düzenlenebilir), risk puanı, eksik veriler |
| KAYNAKLAR | Tüm ilan ve mükerrer ilan bağlantıları, veri ve vergi kaynakları |
| PARAMETRELER | Tüm hesaplama varsayımları (sarı hücreler düzenlenebilir) |
| ANALİZ SORULARI | Talimatnamedeki 10 sorunun cevabı |
| VERGİ-HUKUK NOTLARI | Yabancı yatırımcı için vergi notları; avukat ve mühendis kontrol listesi |
| YEDEK İLANLAR | Ana listeye alınmayan 390 tekil ilan ve alınmama nedeni |

## Yöntem (özet)

- **İlan kaynağı:** spitimou.net. Bu site Yunan emlak ofislerinin ilanlarını toplar. Spitogatos, xe.gr, indomio, plot.gr, properstar, remax.gr ve century21global'e bu ortamdan erişilemedi (bot koruması veya HTTP 403/500 hatası).
- **Satılık ilanlar:** Selanik Belediyesi, daire, 2 ve daha fazla yatak odası, 45–110 m², yenilenmiş → 2.657 ilan. Bunlardan 984 ilanın detayı çekildi. Mükerrer ilanlar (m², kat, 250 m içinde konum ve ±%10 fiyat eşleşmesi) birleştirilince 498 tekil daire kaldı. Bu dairelerden 100'ü seçildi.
- **Seçilen 100 dairenin tamamı şu kriterleri karşılıyor:**
  - Öncelikli merkez bölgede
  - İlan metninde 2 yatak odası teyitli
  - 2019 veya sonrasında yenilenmiş
  - 45–110 m²
  - Asansörlü
  - Kullanıma hazır
- **Kira emsalleri:** Her daire için en az 5 emsal kullanıldı. Emsaller 900 m içinde, ±%30 m² aralığında ve 2 yatak odalı. Tüm emsaller ilan (istenen) kiralarıdır. Sözleşme kirasına geçiş için %5 indirim uygulanır.
- **Mesafeler:**
  - Metro istasyonu koordinatları Wikipedia'dan alındı.
  - Ana caddeler elle tanımlanmış yaklaşık güzergâhlardır (±100–150 m hata payı).
  - Denize mesafe portaldan alındı.

`kaynak_kod/` klasörü veri çekme, seçim ve Excel oluşturma betiklerini içerir. `veri/` klasörü modelin girdi verisini içerir.

> Bu çalışma yatırım, hukuk veya vergi tavsiyesi değildir. Satın almadan önce ilanın güncelliğini, fiyatını ve hukuki durumunu emlak ofisi, Yunan avukat ve mühendis aracılığıyla teyit edin.
