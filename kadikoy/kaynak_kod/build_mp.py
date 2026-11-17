import json, sys, datetime, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as GL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, ScatterChart, Reference, Series

DATA = json.load(open(sys.argv[1])); OUT = sys.argv[2]
NOTES = json.load(open(sys.argv[3])) if len(sys.argv) > 3 else {}
R = DATA['recs']; MARKET = DATA['market']; EXTRAS = DATA['extras']
COMPS = [dict(no=r['no'], **c) for r in R for c in r.get('comps', [])]
CHECK = datetime.date(2026, 11, 17)
N = len(R)

NAVY = '1F3864'; NAVY2 = '2E5597'; GREY = 'F2F2F2'; GREEN_D = '375623'; YEL = 'FFF2CC'
F = 'Arial'
fT = Font(name=F, size=16, bold=True, color='FFFFFF'); fH = Font(name=F, size=9, bold=True, color='FFFFFF')
fN = Font(name=F, size=9, color='000000'); fIn = Font(name=F, size=9, color='0000FF')
fLink = Font(name=F, size=9, color='0563C1', underline='single'); fX = Font(name=F, size=9, color='008000')
fB = Font(name=F, size=10, bold=True, color=NAVY); fSm = Font(name=F, size=8, italic=True, color='595959')
fBold = Font(name=F, size=9, bold=True)
fillT = PatternFill('solid', fgColor=NAVY); fillH = PatternFill('solid', fgColor=NAVY2)
fillG = PatternFill('solid', fgColor=GREY); fillIn = PatternFill('solid', fgColor=YEL); fillGr = PatternFill('solid', fgColor='E2EFDA')
thin = Side(style='thin', color='BFBFBF'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
TL = '#,##0 "₺";(#,##0 "₺");"-"'; TL2 = '#,##0.0 "₺";(#,##0.0 "₺");"-"'; EUR = '#,##0 "€";(#,##0 "€");"-"'
PCT = '0.00%;(0.00%);"-"'; PCT1 = '0.0%;(0.0%);"-"'; NUM = '#,##0;(#,##0);"-"'; NUM1 = '0.0;(0.0);"-"'; DATE = 'dd.mm.yyyy'
WRAP = Alignment(wrap_text=True, vertical='top'); CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

wb = Workbook()
SN = dict(panel='YÖNETİCİ PANELİ', main='DAİRE ANA LİSTESİ', top='EN İYİ 20 YATIRIM', comp='KİRA EMSALLERİ', mah='MAHALLE ANALİZİ',
          scen='SENARYO ANALİZİ', cf='5 VE 10 YILLIK NAKİT AKIŞI', risk='RİSK VE EKSİK VERİLER', src='KAYNAKLAR', par='PARAMETRELER',
          qa='ANALİZ SORULARI', tax='VERGİ-HUKUK NOTLARI', ext='YEDEK İLANLAR')
ws = {}
for i, k in enumerate(SN):
    ws[k] = wb.active if i == 0 else wb.create_sheet()
    ws[k].title = SN[k]; ws[k].sheet_view.showGridLines = False
def q(k): return "'" + SN[k] + "'!"
def title(sh, text, sub, ncols=12):
    sh.row_dimensions[1].height = 30
    for c in range(1, ncols + 1): sh.cell(1, c).fill = fillT; sh.cell(2, c).fill = fillT
    sh['A1'] = text; sh['A1'].font = fT; sh['A1'].alignment = Alignment(vertical='center')
    sh['A2'] = sub; sh['A2'].font = Font(name=F, size=9, italic=True, color='FFFFFF')
def hdr(sh, row, col, text, width=None):
    c = sh.cell(row, col, text); c.font = fH; c.fill = fillH; c.alignment = CENTER; c.border = BOX
    if width: sh.column_dimensions[GL(col)].width = width
    return c
def put(sh, row, col, val, fmt=None, font=None, fill=None, align=None, border=True):
    c = sh.cell(row, col, val)
    c.font = font or (fN if (isinstance(val, str) and val.startswith('=')) else fIn)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if align: c.alignment = align
    if border: c.border = BOX
    return c
def link(sh, row, col, url, text):
    c = sh.cell(row, col, text); c.hyperlink = url; c.font = fLink; c.border = BOX; return c
def dname(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
def taxf(T, cpi='1'):
    """Konut kira vergisi: istisna + götürü gider + 2026 tarifesi; dilimler CPI ile endekslenir."""
    M = f'MAX(0,({T})/({cpi})-P_ISTISNA*P_ISTISNA_UYGULA)*(1-P_GOTURU)'
    prog = (f'(MIN({M},P_V1_LIM)*P_V1+MAX(0,MIN({M},P_V2_LIM)-P_V1_LIM)*P_V2+MAX(0,MIN({M},P_V3_LIM)-P_V2_LIM)*P_V3'
            f'+MAX(0,MIN({M},P_V4_LIM)-P_V3_LIM)*P_V4+MAX(0,{M}-P_V4_LIM)*P_V5)')
    return f'IF(P_VERGI_YONTEM=1,{prog},{M}*P_VERGI_SABIT)*({cpi})'

# ================= PARAMETRELER =================
P = ws['par']
title(P, 'PARAMETRELER – Tüm hesaplama varsayımları (TL)', 'Mavi yazı + sarı zemin = değiştirilebilir girdi. Siyah = formül. Kontrol tarihi: 17.11.2026', 6)
for i, (h, w) in enumerate([('Kategori', 18), ('Parametre', 56), ('Değer', 14), ('Birim', 12), ('Açıklama / Kaynak', 95)], 1): hdr(P, 4, i, h, w)
PARAMS = [
 ('Genel', 'P_SENARYO', 'Aktif senaryo (1=Kötümser, 2=Baz, 3=İyimser)', 2, 'seçim', 'Ana liste, puanlama ve sıralama bu senaryoya göre hesaplanır.', NUM),
 ('Genel', 'P_KONTROL', 'Veri kontrol tarihi', CHECK, 'tarih', 'İlanların emlakjet.com’dan çekildiği tarih.', DATE),
 ('Genel', 'P_EURTRY', 'EUR/TRY kuru (TCMB döviz satış, 17.11.2026)', 55.4446, '₺/€', 'Kaynak: tcmb.gov.tr/kurlar/today.xml (17.11.2026). Euro karşılıkları için.', '0.0000'),
 ('Genel', 'P_EUR_ENF', 'Euro bölgesi enflasyon varsayımı (yıllık)', 0.02, '%', 'Kur projeksiyonu (satın alma gücü paritesi) için.', PCT1),
 ('Satın alma', 'P_TAPU', 'Tapu harcı – alıcının ödediği oran', 0.02, '%', 'Toplam %4 (alıcı %2 + satıcı %2). Satıcı payını da üstlenirseniz 0,04 yazın. Matrah: beyan edilen bedel ile emlak vergi değerinden yüksek olanı; gerçek bedelin altında beyan cezaya tabidir.', PCT1),
 ('Satın alma', 'P_DONER', 'Döner sermaye ücreti', 2534, '₺', '2026 tutarı kaynaklarda 2.534 ₺ (2.227 + 307); büyükşehir katsayısıyla farklı olabilir – tapu müdürlüğünden teyit edin.', TL),
 ('Satın alma', 'P_KOMISYON', 'Alıcı emlakçı komisyonu (KDV dahil)', 0.024, '%', '%2 + %20 KDV (piyasa uygulaması).', PCT1),
 ('Satın alma', 'P_SABIT', 'Sabit masraflar (ekspertiz, avukat/tapu kontrolü, DASK ilk poliçe, abonelik devirleri)', 25000, '₺', 'Varsayım.', TL),
 ('Satın alma', 'P_HAZIRLIK', 'Kiraya hazırlık / küçük onarım rezervi', 20000, '₺', 'Yeni bina olduğu için düşük tutuldu. Varsayım.', TL),
 ('Satın alma', 'P_ESYA_UYGULA', 'Eşyasız daireler eşyalansın mı? (1=Evet, 0=Hayır)', 0, 'seçim', 'Kadıköy’de uzun dönem kiralık stokun büyük kısmı eşyasız; kira emsalleri de çoğunlukla eşyasız. Varsayılan Hayır.', NUM),
 ('Satın alma', 'P_ESYA', 'Eşya paketi maliyeti', 250000, '₺', 'Orta segment mobilya + beyaz eşya. Varsayım.', TL),
 ('Kira', 'P_NET_ORAN', 'Net m² bilinmiyorsa net/brüt oranı', 0.80, 'x', 'Kira hesabı net m² üzerinden yapılır (brüt/net oranı binadan binaya çok değişir).', '0.00'),
 ('Kira', 'P_ZEMIN_IND', 'Zemin / bahçe / giriş katı kira indirimi', 0.10, '%', 'Emsaller farklı katlardan oluşur; zemin seviyesindeki daireler (güvenlik, ışık, gürültü) genellikle daha düşük kiralanır. Varsayım.', '0.0%'),
 ('Kira', 'P_ILAN_INDIRIM', 'İlan kirası → sözleşme kirası indirimi', 0.05, '%', 'Emsaller İSTENEN kiralardır; gerçekleşmiş sözleşme kirası değildir.', PCT1),
 ('Kira', 'P_KIRACI_AY', 'Kiracı bulma ücreti (ay kira, KDV dahil)', 1.2, 'ay', '1 aylık kira + %20 KDV (piyasa uygulaması).', '0.00'),
 ('Kira', 'P_SOZLESME', 'Ortalama kiracı kalış süresi', 3, 'yıl', 'Varsayım.', NUM),
 ('Gider', 'P_BAKIM', 'Bakım-onarım (yıllık brüt kiranın %’si)', 0.03, '%', 'Yeni bina. Varsayım.', PCT1),
 ('Gider', 'P_YONETIM', 'Kira yönetim ücreti (tahsil edilen kiranın %’si)', 0.0, '%', 'Türkiye’de uzun dönem kirada genellikle mal sahibi yönetir; profesyonel yönetim için %5–10 yazın.', PCT1),
 ('Gider', 'P_SIGORTA', 'DASK + konut sigortası', 60, '₺/m²/yıl', 'Varsayım; DASK zorunludur.', TL),
 ('Gider', 'P_EV_ORAN', 'Emlak vergisi oranı (büyükşehir konut)', 0.002, 'binde', 'Binde 2 (büyükşehir konut). Kaynak: hesapkurdu, tahanci.av.tr.', '0.000%'),
 ('Gider', 'P_EV_DEGER', 'Emlak vergi değeri / piyasa değeri oranı', 0.35, 'x', 'Vergi değeri genellikle piyasanın altındadır; daire bazında belediyeden teyit edin. Varsayım.', '0.00'),
 ('Gider', 'P_AIDAT_PAY', 'Mal sahibinin üstlendiği aidat payı (dolu aylarda; demirbaş vb.)', 0.10, '%', 'Boş aylarda aidatın tamamı mal sahibine aittir.', PCT1),
 ('Gider', 'P_AIDAT_VARS', 'Aidat bilinmiyorsa kullanılacak tutar', 2500, '₺/ay', 'Asansörlü yeni bina için varsayım (site içinde daha yüksek).', TL),
 ('Vergi', 'P_VERGI_YONTEM', 'Kira vergisi yöntemi (1=Kademeli – bu daire tek kira geliri; 2=Sabit marjinal oran)', 1, 'seçim', 'Ücret vb. başka geliriniz varsa kira bunlarla birleşir ve üst dilime çıkar → 2 seçip marjinal oranı girin.', NUM),
 ('Vergi', 'P_ISTISNA_UYGULA', 'Konut kira geliri istisnası uygulansın mı? (1=Evet, 0=Hayır)', 1, 'seçim', 'Ticari/zirai/mesleki kazancı olanlar ve belirli gelir sınırını aşanlar yararlanamaz; istisna kişi başına bir kez uygulanır (birden çok daire için toplamda).', NUM),
 ('Vergi', 'P_ISTISNA', '2026 konut kira geliri istisnası', 58000, '₺', 'GVK Genel Tebliği Seri No: 332; KPMG.', TL),
 ('Vergi', 'P_GOTURU', 'Götürü gider oranı', 0.15, '%', 'Götürü gider yöntemi (gerçek gider yerine %15).', PCT1),
 ('Vergi', 'P_V1_LIM', 'Dilim 1 üst sınırı', 190000, '₺', '2026 ücret dışı gelir tarifesi (GVK 103; Seri No: 332).', TL),
 ('Vergi', 'P_V1', 'Dilim 1 oranı', 0.15, '%', '', PCT1),
 ('Vergi', 'P_V2_LIM', 'Dilim 2 üst sınırı', 400000, '₺', '', TL),
 ('Vergi', 'P_V2', 'Dilim 2 oranı', 0.20, '%', '', PCT1),
 ('Vergi', 'P_V3_LIM', 'Dilim 3 üst sınırı', 1000000, '₺', '', TL),
 ('Vergi', 'P_V3', 'Dilim 3 oranı', 0.27, '%', '', PCT1),
 ('Vergi', 'P_V4_LIM', 'Dilim 4 üst sınırı', 5300000, '₺', '', TL),
 ('Vergi', 'P_V4', 'Dilim 4 oranı', 0.35, '%', '', PCT1),
 ('Vergi', 'P_V5', 'Dilim 5 oranı (üstü)', 0.40, '%', '', PCT1),
 ('Vergi', 'P_VERGI_SABIT', 'Sabit marjinal oran (yöntem 2)', 0.27, '%', '', PCT1),
 ('Puanlama', 'W_GETIRI', 'Ağırlık – Net kira getirisi', 0.40, '%', 'Talimatnamedeki ağırlıklar.', PCT1),
 ('Puanlama', 'W_KONUM', 'Ağırlık – Konum ve kiracı talebi', 0.20, '%', '', PCT1),
 ('Puanlama', 'W_FIYAT', 'Ağırlık – Fiyatın emsallere uygunluğu', 0.15, '%', '', PCT1),
 ('Puanlama', 'W_KALITE', 'Ağırlık – Bina ve daire kalitesi', 0.10, '%', '', PCT1),
 ('Puanlama', 'W_RISK', 'Ağırlık – Boş kalma / kiracı riski', 0.10, '%', '', PCT1),
 ('Puanlama', 'W_LIKIT', 'Ağırlık – Gelecekte satış kolaylığı', 0.05, '%', '', PCT1),
 ('Puanlama', 'P_NY_MIN', 'Net getiri puanı: 0 puan eşiği', 0.02, '%', '', PCT),
 ('Puanlama', 'P_NY_MAX', 'Net getiri puanı: 100 puan eşiği', 0.05, '%', '', PCT),
 ('Puanlama', 'P_VERI_CEZA', 'Yetersiz veri çarpanı (emsal <3)', 0.85, 'x', 'Yetersiz veriye sahip daireler sıralamada yukarı çıkamasın diye.', '0.00'),
 ('Puanlama', 'P_ILAN_GUN', 'İlan “eski” sayılma eşiği (gün)', 60, 'gün', '', NUM),
 ('Pazarlık', 'P_PAZARLIK', 'Standart pazarlık payı', 0.05, '%', 'İstanbul’da ilan–satış farkı için varsayım.', PCT1),
 ('Pazarlık', 'P_PAZARLIK_MAX', 'Azami önerilen pazarlık', 0.12, '%', '', PCT1),
 ('Pazarlık', 'P_HEDEF', 'Hedef net getiri (azami fiyat hesabı için)', 0.045, '%', '', PCT1),
]
PROW = {}
for i, (cat, nm, lab, val, unit, note, fmt) in enumerate(PARAMS):
    rr = 5 + i
    put(P, rr, 1, cat, font=fN); put(P, rr, 2, lab, font=fN, align=WRAP)
    c = put(P, rr, 3, val, fmt=fmt)
    if not (isinstance(val, str) and val.startswith('=')): c.fill = fillIn
    put(P, rr, 4, unit, font=fN); put(P, rr, 5, note, font=fSm, align=WRAP)
    dname(nm, f"{q('par')}$C${rr}"); PROW[nm] = rr
for nm, lst in [('P_SENARYO', '"1,2,3"'), ('P_VERGI_YONTEM', '"1,2"'), ('P_ESYA_UYGULA', '"0,1"'), ('P_ISTISNA_UYGULA', '"0,1"')]:
    dv = DataValidation(type='list', formula1=lst); P.add_data_validation(dv); dv.add(f'C{PROW[nm]}')
rr = 5 + len(PARAMS) + 1
put(P, rr, 2, 'Ağırlık toplamı (100% olmalı)', font=fB); put(P, rr, 3, '=W_GETIRI+W_KONUM+W_FIYAT+W_KALITE+W_RISK+W_LIKIT', fmt=PCT1)
put(P, rr + 1, 2, 'Satın alma masrafları toplamı (% fiyat, sabitler hariç)', font=fB); put(P, rr + 1, 3, '=P_TAPU+P_KOMISYON', fmt=PCT1)
dname('P_MASRAF_PCT', f"{q('par')}$C${rr + 1}")
put(P, rr + 2, 2, 'Yıllık emlak vergisi (% piyasa değeri)', font=fB); put(P, rr + 2, 3, '=P_EV_ORAN*P_EV_DEGER', fmt='0.000%')
dname('P_EV', f"{q('par')}$C${rr + 2}")
P.freeze_panes = 'A5'

# ================= SENARYO PARAMETRELERİ =================
S = ws['scen']
title(S, 'SENARYO ANALİZİ – Kötümser / Baz / İyimser', 'Sarı hücreler değiştirilebilir. Aktif senaryo PARAMETRELER’de seçilir; tüm sayfalar otomatik güncellenir.', 30)
for i, (h, w) in enumerate([('Senaryo varsayımı', 50), ('Birim', 9), ('Kötümser', 12), ('Baz', 12), ('İyimser', 12), ('AKTİF', 12), ('Açıklama', 75)], 1): hdr(S, 4, i, h, w)
SCN = [
 ('KIRA', 'Kira seviyesi (emsal ortalamasına çarpan)', 'x', (0.92, 1.00, 1.06), '0.00', ''),
 ('BOS', 'Yıllık boş kalma (ay)', 'ay', (1.5, 0.75, 0.25), '0.00', 'Kadıköy’ün lüks semtlerinde talep güçlü; kiracı değişiminde 2–6 hafta.'),
 ('IND', 'Alım fiyatında pazarlık indirimi', '%', (0.02, 0.05, 0.08), PCT1, 'İlan fiyatına göre.'),
 ('MAS', 'İşletme gideri çarpanı (bakım + sigorta)', 'x', (1.15, 1.00, 0.90), '0.00', ''),
 ('ENF1', '1. yıl TÜFE (yıllık)', '%', (0.30, 0.25, 0.22), PCT1, 'Ekim 2026 yıllık TÜFE %29,77; 12 aylık ortalama %31,23 (TÜİK, 03.11.2026). Dezenflasyon varsayımı.'),
 ('ENFU', 'Uzun dönem TÜFE (5. yıldan itibaren)', '%', (0.18, 0.12, 0.09), PCT1, '1. yıldan 5. yıla doğrusal yakınsama.'),
 ('KREEL', 'Kira artışı – TÜFE farkı (yıllık)', '%', (-0.03, 0.0, 0.01), PCT1, 'TBK 344: yenilemede artış 12 aylık TÜFE ortalamasını aşamaz; 5 yıl sonra rayiç tespit mümkündür.'),
 ('DREEL', 'Mülk değerinde reel değişim (yıllık)', '%', (-0.03, 0.01, 0.03), PCT1, 'Nominal değer artışı = (1+TÜFE)×(1+reel değişim)–1.'),
 ('KURREEL', 'TL’nin euroya karşı reel değişimi (yıllık)', '%', (-0.05, 0.0, 0.02), PCT1, 'Negatif = TL reel değer kaybı. Kur artışı = (1+TÜFE)/(1+Euro enfl.)/(1+reel değişim)–1.'),
 ('CIKIS', 'Çıkışta satış maliyeti (satıcı tapu harcı %2 + komisyon %2,4)', '%', (0.05, 0.044, 0.044), PCT1, '5 yıldan uzun tutulan konutta değer artış kazancı vergisi yoktur (GVK mük. 80); model satışın 5 yıldan sonra yapıldığını varsayar.'),
]
for i, (k, lab, unit, vals, fmt, note) in enumerate(SCN):
    rr = 5 + i
    put(S, rr, 1, lab, font=fN); put(S, rr, 2, unit, font=fN)
    for j, v in enumerate(vals): put(S, rr, 3 + j, v, fmt=fmt, fill=fillIn)
    put(S, rr, 6, f'=INDEX(C{rr}:E{rr},1,P_SENARYO)', fmt=fmt, fill=fillGr); put(S, rr, 7, note, font=fSm, align=WRAP)
    for j, s in enumerate(['K', 'B', 'I', 'A']): dname(f'S{s}_{k}', f"{q('scen')}${'CDEF'[j]}${rr}")
put(S, 16, 1, 'Aktif senaryo adı', font=fB); put(S, 16, 6, '=CHOOSE(P_SENARYO,"Kötümser","Baz","İyimser")', font=Font(name=F, size=10, bold=True, color=GREEN_D))

# ================= MAIN LIST =================
M = ws['main']; FR = 5; LR = FR + N - 1
COLS = [
 ('no', 'Daire No', 8, 'A. GENEL BİLGİLER'), ('title', 'İlan başlığı', 34, ''), ('url', 'İlan URL', 10, ''), ('plat', 'Platform', 16, ''), ('adv', 'Emlak ofisi', 20, ''),
 ('contact', 'Danışman / ilan no', 22, ''), ('upd', 'Son güncelleme', 11, ''), ('first', 'İlk yayın', 11, ''), ('ver', 'Doğrulama tarihi', 11, ''), ('status', 'İlan durumu', 18, ''), ('photo', 'Fotoğraf', 10, ''),
 ('mah', 'Mahalle', 13, 'B. KONUM'), ('zone', 'Bölge sınıfı', 12, ''), ('addr', 'Tam adres', 10, ''), ('maps', 'Google Maps', 10, ''),
 ('dcenter', 'Kadıköy çarşı/iskeleye mesafe (m)', 10, ''), ('rail', 'En yakın raylı sistem', 20, ''), ('drail', 'Raylı sisteme mesafe (m)', 9, ''),
 ('dbagdat', 'Bağdat Caddesi’ne mesafe (m)', 10, ''), ('dsahil', 'Sahile mesafe (m, yaklaşık)', 9, ''), ('feats', 'Yakınlık özellikleri (ilan)', 24, ''),
 ('adv_txt', 'Kiralama avantajları', 30, ''), ('dis_txt', 'Dezavantajlar', 30, ''),
 ('price', 'Satış fiyatı (₺)', 13, 'C. DAİRE ÖZELLİKLERİ'), ('price_eur', 'Satış fiyatı (€)', 10, ''), ('net', 'Net m²', 7, ''), ('sqm', 'Brüt m²', 7, ''), ('rsqm', 'Kira hesabı m² (net)', 7, ''), ('ppsm', '₺/brüt m²', 10, ''),
 ('room', 'Oda tipi', 7, ''), ('baths', 'Banyo', 6, ''), ('age', 'Bina yaşı (ilan)', 11, ''), ('ybuilt', 'Yapım yılı (yaklaşık)', 9, ''), ('floor', 'Kat', 10, ''), ('floorn', 'Kat (sayı)', 6, ''),
 ('tf', 'Binadaki kat sayısı', 7, ''), ('elev', 'Asansör', 7, ''), ('balc', 'Balkon', 7, ''), ('heat', 'Isıtma', 14, ''), ('furn', 'Eşya', 9, ''), ('park', 'Otopark', 14, ''),
 ('site', 'Site içinde', 7, ''), ('kredi', 'Krediye uygunluk', 11, ''), ('tapu', 'Tapu durumu', 13, ''), ('fees', 'Aidat (₺/ay, ilan)', 9, ''), ('usage', 'Kullanım durumu', 13, ''),
 ('rentnow', 'Mevcut kira (₺/ay, ilan metninden; varsa 1. yıl kira)', 11, ''), ('rclaim', 'İlanda beyan edilen kira potansiyeli (₺/ay, kontrol amaçlı)', 11, ''), ('legal', 'Hukuki / teknik risk notları', 30, ''), ('dups', 'Aynı daire – diğer ilanlar', 26, ''),
 ('ncomp', 'Kira emsal sayısı', 7, 'D. MALİYETLER VE GELİRLER (TL)'), ('compavg', 'Emsal ort. ilan kirası (₺/net m²)', 10, ''), ('complink', 'Kira emsalleri', 10, ''),
 ('basem2', 'Baz kira ₺/net m² (ilan indirimi sonrası)', 10, ''), ('rk', 'Aylık kira – Kötümser', 10, ''), ('rb', 'Aylık kira – Baz', 10, ''), ('ri', 'Aylık kira – İyimser', 10, ''),
 ('ra', 'Aylık kira – AKTİF', 10, ''), ('gross', 'Yıllık brüt kira', 11, ''), ('vac', 'Boş kalma (ay/yıl)', 7, ''), ('coll', 'Tahsil edilen yıllık kira', 11, ''),
 ('maint', 'Bakım', 9, ''), ('mgmt', 'Yönetim + kiracı bulma', 10, ''), ('ptax', 'Emlak vergisi', 9, ''), ('ins', 'DASK + konut sigortası', 9, ''), ('feeown', 'Mal sahibi aidat payı', 9, ''),
 ('opex', 'Toplam işletme gideri', 11, ''), ('buy', 'Alım bedeli (pazarlık sonrası)', 13, ''), ('ttax', 'Tapu harcı (alıcı)', 11, ''), ('notary', 'Döner sermaye + sabit masraflar', 11, ''),
 ('comm', 'Emlakçı komisyonu', 11, ''), ('furncost', 'Eşya + hazırlık', 11, ''), ('tinv', 'TOPLAM YATIRIM (₺)', 13, ''), ('tinv_eur', 'Toplam yatırım (€)', 10, ''),
 ('noi', 'Vergi öncesi net işletme geliri', 11, ''), ('rtax', 'Kira geliri vergisi', 10, ''), ('ncf', 'Vergi sonrası net kira (₺/yıl)', 11, ''), ('ncf_eur', 'Vergi sonrası net kira (€/yıl)', 9, ''),
 ('gy', 'Brüt kira getirisi (ilan fiyatına)', 9, ''), ('ny', 'NET KİRA GETİRİSİ', 9, ''), ('payback', 'Geri dönüş (yıl, sabit kira)', 8, ''), ('mult', 'Kira çarpanı (yıl)', 8, ''),
 ('exitppsm', 'Emsal satış ₺/m² (aynı mahalle, ≤15 yaş ilan medyanı)', 11, ''), ('exitn', 'Emsal satış n', 6, ''), ('pratio', 'Fiyat / emsal oranı', 8, ''),
 ('nego', 'Pazarlık önerilen alım fiyatı (₺)', 13, ''), ('maxp', 'Hedef net getiri için azami fiyat (₺)', 13, ''),
 ('risk', 'Risk puanı (0–100)', 7, ''), ('risklvl', 'Yatırım riski', 9, ''), ('mainrisk', 'Risk işaretleri (ağırlık sırasıyla)', 32, ''),
 ('s1', 'Puan: Net getiri', 7, 'E. YATIRIM PUANI'), ('s2', 'Puan: Konum & talep', 7, ''), ('s3', 'Puan: Fiyat uygunluğu', 7, ''), ('s4', 'Puan: Kalite', 7, ''),
 ('s5', 'Puan: Boş kalma/kiracı riski', 7, ''), ('s6', 'Puan: Satış kolaylığı', 7, ''), ('dataok', 'Veri yeterliliği', 9, ''), ('score', 'GENEL YATIRIM PUANI', 9, ''),
 ('helper', 'Sıralama yardımcı', 6, ''), ('rank', 'SIRA', 6, ''), ('nyrank', 'Net getiri sırası', 6, ''),
]
C = {k: GL(i) for i, (k, h, w, g) in enumerate(COLS, 1)}
CI = {k: i for i, (k, h, w, g) in enumerate(COLS, 1)}
title(M, f'ÇOK PORTALLI DAİRE LİSTESİ – Kadıköy lüks mahalleler, 2010+ 1+1 / 2+1 ({N} tekil daire · Emlakjet + Century 21 + Turyap)',
      'Mavi = ilan verisi · Siyah = formül · Mesafeler kuş uçuşu; Bağdat Cd. ve sahil güzergâhı yaklaşık · Kira emsalleri ilan (istenen) kiralarıdır', len(COLS))
gfill = PatternFill('solid', fgColor='8EA9DB')
for i, (k, h, w, g) in enumerate(COLS, 1):
    c = M.cell(3, i, g or None); c.fill = gfill; c.font = Font(name=F, size=9, bold=True, color=NAVY)
    hdr(M, 4, i, h, w)
M.row_dimensions[4].height = 62
def ref(k, r): return f'{C[k]}{r}'
MAH_TXT = {
 'Caferağa': ('Moda: Kadıköy çarşı, vapur iskelesi, M4 Kadıköy ve Moda sahiline yürüme; genç profesyonel, öğrenci ve yabancı kiracı talebi çok güçlü, 1+1’e talep yüksek.', 'Gece hayatı gürültüsü, otopark sıkıntısı, dar parseller; eski yapı stoku içinde yeni binalar sınırlı.'),
 'Kozyatağı': ('M4 Kozyatağı, E-5 ve TEM bağlantısı, iş merkezleri ve hastaneler; beyaz yaka kiracı talebi.', 'E-5 ve yoğun trafik gürültüsü; sahile görece uzak.'),
 'Caddebostan': ('Bağdat Caddesi + Caddebostan sahili; en prestijli adreslerden, üst gelir grubu kiracı.', 'En yüksek m² fiyatları → düşük getiri; yüksek aidat.'),
 'Suadiye': ('Bağdat Caddesi, Marmaray Suadiye, sahil; kentsel dönüşümle geniş yeni bina stoku.', 'Yoğun inşaat ve yeni arz → kira rekabeti; inşaat gürültüsü.'),
 'Fenerbahçe': ('Fenerbahçe parkı, marina, sahil; sakin ve prestijli.', 'Kadıköy’ün en yüksek m² fiyatlarından; raylı sisteme mesafe; getiri düşük.'),
 'Feneryolu': ('Marmaray Feneryolu, Bağdat Caddesi başlangıcı, Kadıköy merkeze yakın.', 'Dar sokaklar, otopark sıkıntısı, minibüs yolu trafiği.'),
 'Göztepe': ('Marmaray ve M4 Göztepe, Marmara Üniv. Göztepe kampüsü, parklar.', 'E-5 tarafında gürültü; minibüs yolu trafiği.'),
 'Erenköy': ('Marmaray Erenköy, köklü yerleşim, okullar, Bağdat Caddesi yakın.', 'Dönüşüm inşaatları; bazı sokaklarda otopark.'),
 'Zühtüpaşa': ('Kadıköy merkez ve Söğütlüçeşme (Marmaray + Metrobüs) yakın.', 'Söğütlüçeşme–Fikirtepe trafiği, karışık doku.'),
 'Koşuyolu': ('Sakin doku, Acıbadem metro ve hastanelere yakın.', 'Raylı sisteme mesafe; az ilan, düşük likidite.'),
}
DEMAND = {'Caferağa': (92, 90, 'Düşük'), 'Kozyatağı': (88, 82, 'Düşük'), 'Caddebostan': (88, 90, 'Düşük'), 'Suadiye': (85, 85, 'Düşük–Orta (yeni arz)'),
          'Fenerbahçe': (80, 85, 'Orta'), 'Feneryolu': (85, 80, 'Düşük'), 'Göztepe': (85, 82, 'Düşük'), 'Erenköy': (84, 82, 'Düşük'), 'Zühtüpaşa': (80, 72, 'Orta'), 'Koşuyolu': (75, 72, 'Orta')}
MAHS = [m for m in DEMAND]
CR = f"{q('comp')}$A$6:$A$3000"; CV = f"{q('comp')}$K$6:$K$3000"
MI = lambda r: f"MATCH({ref('mah', r)},{q('mah')}$A$6:$A$15,0)"
for idx, rec in enumerate(R):
    r = FR + idx
    put(M, r, CI['no'], rec['no'], font=Font(name=F, size=9, bold=True, color=NAVY))
    put(M, r, CI['title'], rec['title'], align=WRAP)
    link(M, r, CI['url'], rec['url'], 'İlanı aç ↗'); put(M, r, CI['plat'], rec['source'] + (f" (+{rec['nsrc'] - 1} portal)" if rec.get('nsrc', 1) > 1 else ''))
    put(M, r, CI['adv'], rec['adv']); put(M, r, CI['contact'], f"{rec['agent']} | İlan no: {rec['ilan_no']} | Tel: ilan sayfasından", align=WRAP)
    put(M, r, CI['upd'], datetime.date.fromisoformat(rec['updated']) if rec['updated'] else 'Bilinmiyor', fmt=DATE)
    put(M, r, CI['first'], datetime.date.fromisoformat(rec['created']) if rec['created'] else 'Bilinmiyor', fmt=DATE)
    put(M, r, CI['ver'], CHECK, fmt=DATE)
    put(M, r, CI['status'], f'=IF(ISNUMBER({ref("upd", r)}),IF(P_KONTROL-{ref("upd", r)}<=P_ILAN_GUN,"Yayında (17.11.2026 erişildi)","Yayında ama güncelleme eski – teyit"),"Yayında (17.11.2026 erişildi; güncelleme tarihi yok)")')
    if rec['images']: link(M, r, CI['photo'], rec['images'][0], 'Fotoğraf ↗')
    else: put(M, r, CI['photo'], '-')
    put(M, r, CI['mah'], rec['mah']); put(M, r, CI['zone'], 'Öncelikli lüks')
    put(M, r, CI['addr'], 'İlanda yok' if rec.get('precise', True) else 'İlanda yok – konum mahalle merkezi (yaklaşık)', align=WRAP)
    link(M, r, CI['maps'], f"https://www.google.com/maps?q={rec['lat']:.6f},{rec['lng']:.6f}", 'Harita ↗')
    put(M, r, CI['dcenter'], rec['dcenter'], fmt=NUM); put(M, r, CI['rail'], rec['rail']); put(M, r, CI['drail'], rec['drail'], fmt=NUM)
    put(M, r, CI['dbagdat'], rec['bag_txt'] if rec['bag_txt'] is not None else rec['dbagdat'], fmt=NUM)
    if rec['bag_txt'] is not None: M.cell(r, CI['dbagdat']).comment = None
    put(M, r, CI['dsahil'], rec['dsahil'], fmt=NUM)
    put(M, r, CI['feats'], ', '.join(rec['feats']) or '-', align=WRAP)
    a, d = MAH_TXT[rec['mah']]; put(M, r, CI['adv_txt'], a, align=WRAP); put(M, r, CI['dis_txt'], d, align=WRAP)
    put(M, r, CI['price'], rec['price'], fmt=TL); put(M, r, CI['price_eur'], f'={ref("price", r)}/P_EURTRY', fmt=EUR)
    put(M, r, CI['net'], rec['net'] or 'Bilinmiyor', fmt=NUM); put(M, r, CI['sqm'], rec['brut'], fmt=NUM)
    put(M, r, CI['rsqm'], f'=IF(ISNUMBER({ref("net", r)}),{ref("net", r)},{ref("sqm", r)}*P_NET_ORAN)', fmt=NUM)
    put(M, r, CI['ppsm'], f'={ref("price", r)}/{ref("sqm", r)}', fmt=TL)
    put(M, r, CI['room'], rec['room']); put(M, r, CI['baths'], rec['baths'] or 'Bilinmiyor')
    put(M, r, CI['age'], rec['age']); put(M, r, CI['ybuilt'], f"{2026 - (rec['age_hi'] or 0)}–2026" if (rec['age_hi'] or 0) > 0 else '2025–2026', align=CENTER)
    put(M, r, CI['floor'], rec['floor']); put(M, r, CI['floorn'], rec['floorn'] if rec['floorn'] is not None else 0, fmt=NUM1)
    put(M, r, CI['tf'], rec['tf']); put(M, r, CI['elev'], rec['elev'] or 'Belirtilmemiş'); put(M, r, CI['balc'], rec['balkon'] or 'Belirtilmemiş')
    put(M, r, CI['heat'], rec['heat']); put(M, r, CI['furn'], rec['esya'] or 'Belirtilmemiş'); put(M, r, CI['park'], rec['park'] or 'Belirtilmemiş')
    put(M, r, CI['site'], rec['site']); put(M, r, CI['kredi'], rec['kredi']); put(M, r, CI['tapu'], rec['tapu'])
    put(M, r, CI['fees'], rec['aidat'] if rec['aidat'] else '', fmt=TL); put(M, r, CI['usage'], rec['usage'])
    put(M, r, CI['rentnow'], rec['rent_now'] if rec['rent_now'] else '', fmt=TL)
    put(M, r, CI['rclaim'], rec.get('rent_claim') or '', fmt=TL)
    lg = []
    if rec['tapu'] != 'Kat Mülkiyeti': lg.append(f"Tapu: {rec['tapu']} → iskân/kat mülkiyeti teyidi")
    if (rec['age_hi'] or 0) == 0: lg.append('Sıfır bina: iskân, yapı denetim, KDV durumu ve teslim tarihi teyit edilmeli')
    elif (rec['age_hi'] or 0) <= 8: lg.append('TBDY 2018 sonrası olabilir – ruhsat tarihi kontrol')
    else: lg.append('2007 yönetmeliği dönemi (2010–2015) – yapı denetim raporu ve beton sınıfı kontrol')
    if rec['kredi'] and 'Uygun Değil' in rec['kredi']: lg.append('Krediye uygun değil → hukuki eksik olabilir')
    if rec['usage'] == 'Kiracı Oturuyor': lg.append('Kiracılı – sözleşme ve kira tutarı incelenmeli')
    lg.append('TAKBİS tapu kaydı, ipotek/haciz/şerh, yönetim planı, DASK')
    put(M, r, CI['legal'], '; '.join(lg), align=WRAP)
    put(M, r, CI['dups'], '; '.join(f"{x['adv']}: {x['price']:,.0f} ₺".replace(',', '.') for x in rec['dups']) if rec['dups'] else 'Tekrarı bulunamadı', align=WRAP)
    # D
    put(M, r, CI['ncomp'], f'=COUNTIF({CR},{ref("no", r)})', fmt=NUM)
    put(M, r, CI['compavg'], f"=IFERROR(AVERAGEIF({CR},{ref('no', r)},{CV}),INDEX({q('mah')}$K$6:$K$15,{MI(r)})/P_NET_ORAN)", fmt=TL2)
    first = next((i for i, x in enumerate(COMPS) if x['no'] == rec['no']), None)
    c = M.cell(r, CI['complink'], 'Emsaller ↗' if first is not None else 'Emsal yok'); c.font = fLink; c.border = BOX
    if first is not None: c.hyperlink = f"#'{SN['comp']}'!A{6 + first}"
    put(M, r, CI['basem2'], f'={ref("compavg", r)}*(1-P_ILAN_INDIRIM)', fmt=TL2)
    LK = ref('rentnow', r)
    for k_, sc in (('rk', 'SK'), ('rb', 'SB'), ('ri', 'SI')):
        put(M, r, CI[k_], f'=IF(ISNUMBER({LK}),{LK},ROUND({ref("basem2", r)}*{ref("rsqm", r)}*{sc}_KIRA*(1-IF({ref("floorn", r)}<1,P_ZEMIN_IND,0)),-2))', fmt=TL)
    put(M, r, CI['ra'], f'=CHOOSE(P_SENARYO,{ref("rk", r)},{ref("rb", r)},{ref("ri", r)})', fmt=TL, fill=fillGr)
    put(M, r, CI['gross'], f'={ref("ra", r)}*12', fmt=TL); put(M, r, CI['vac'], '=SA_BOS', fmt='0.00')
    put(M, r, CI['coll'], f'={ref("ra", r)}*(12-{ref("vac", r)})', fmt=TL)
    put(M, r, CI['maint'], f'={ref("gross", r)}*P_BAKIM*SA_MAS', fmt=TL)
    put(M, r, CI['mgmt'], f'={ref("coll", r)}*P_YONETIM+{ref("ra", r)}*P_KIRACI_AY/P_SOZLESME', fmt=TL)
    put(M, r, CI['ptax'], f'={ref("price", r)}*P_EV', fmt=TL)
    put(M, r, CI['ins'], f'={ref("sqm", r)}*P_SIGORTA*SA_MAS', fmt=TL)
    put(M, r, CI['feeown'], f'=IF({ref("fees", r)}="",P_AIDAT_VARS,{ref("fees", r)})*({ref("vac", r)}+(12-{ref("vac", r)})*P_AIDAT_PAY)', fmt=TL)
    put(M, r, CI['opex'], f'=SUM({ref("maint", r)}:{ref("feeown", r)})', fmt=TL)
    put(M, r, CI['buy'], f'={ref("price", r)}*(1-SA_IND)', fmt=TL)
    put(M, r, CI['ttax'], f'={ref("buy", r)}*P_TAPU', fmt=TL)
    put(M, r, CI['notary'], '=P_DONER+P_SABIT', fmt=TL)
    put(M, r, CI['comm'], f'={ref("buy", r)}*P_KOMISYON', fmt=TL)
    put(M, r, CI['furncost'], f'=IF(AND(P_ESYA_UYGULA=1,{ref("furn", r)}<>"Eşyalı"),P_ESYA,0)+P_HAZIRLIK', fmt=TL)
    put(M, r, CI['tinv'], f'=SUM({ref("buy", r)}:{ref("furncost", r)})', fmt=TL, font=fBold)
    put(M, r, CI['tinv_eur'], f'={ref("tinv", r)}/P_EURTRY', fmt=EUR)
    put(M, r, CI['noi'], f'={ref("coll", r)}-{ref("opex", r)}', fmt=TL)
    put(M, r, CI['rtax'], '=' + taxf(ref('coll', r)), fmt=TL)
    put(M, r, CI['ncf'], f'={ref("noi", r)}-{ref("rtax", r)}', fmt=TL, font=fBold)
    put(M, r, CI['ncf_eur'], f'={ref("ncf", r)}/P_EURTRY', fmt=EUR)
    put(M, r, CI['gy'], f'={ref("gross", r)}/{ref("price", r)}', fmt=PCT)
    put(M, r, CI['ny'], f'={ref("ncf", r)}/{ref("tinv", r)}', fmt=PCT, font=fBold)
    put(M, r, CI['payback'], f'=IF({ref("ncf", r)}>0,{ref("tinv", r)}/{ref("ncf", r)},"Negatif")', fmt=NUM1)
    put(M, r, CI['mult'], f'={ref("price", r)}/{ref("gross", r)}', fmt=NUM1)
    put(M, r, CI['exitppsm'], rec['exit_ppsm'] or '', fmt=TL); put(M, r, CI['exitn'], rec['exit_n'] or '', fmt=NUM)
    put(M, r, CI['pratio'], f"=IF({ref('exitppsm', r)}=\"\",{ref('ppsm', r)}/INDEX({q('mah')}$I$6:$I$15,{MI(r)}),{ref('ppsm', r)}/{ref('exitppsm', r)})", fmt='0.00')
    put(M, r, CI['nego'], f'=ROUND(MAX({ref("price", r)}*(1-P_PAZARLIK_MAX),MIN({ref("price", r)}*(1-P_PAZARLIK),IF({ref("pratio", r)}>1,{ref("price", r)}/{ref("pratio", r)},{ref("price", r)}*(1-P_PAZARLIK)))),-4)', fmt=TL)
    put(M, r, CI['maxp'], f'=IF({ref("ncf", r)}>0,ROUND(({ref("ncf", r)}/P_HEDEF-P_DONER-P_SABIT-{ref("furncost", r)})/(1+P_MASRAF_PCT),-4),0)', fmt=TL)
    put(M, r, CI['risk'], f"={q('risk')}R{6 + idx}", fmt=NUM, font=fX)
    put(M, r, CI['risklvl'], f'=IF({ref("risk", r)}>=35,"Yüksek",IF({ref("risk", r)}>=18,"Orta","Düşük"))')
    put(M, r, CI['mainrisk'], f"={q('risk')}S{6 + idx}", font=fX, align=WRAP)
    put(M, r, CI['s1'], f'=MAX(0,MIN(100,({ref("ny", r)}-P_NY_MIN)/(P_NY_MAX-P_NY_MIN)*100))', fmt=NUM)
    put(M, r, CI['s2'], f"=0.5*INDEX({q('mah')}$O$6:$O$15,{MI(r)})+0.3*MAX(0,100-{ref('drail', r)}/12)+0.2*MAX(0,100-MIN({ref('dbagdat', r)},{ref('dsahil', r)})/15)", fmt=NUM)
    put(M, r, CI['s3'], f'=MAX(0,MIN(100,(1.25-{ref("pratio", r)})/0.5*100))', fmt=NUM)
    put(M, r, CI['s4'], (f'=MIN(100,IF(LEFT({ref("age", r)},1)="0",40,IF(OR({ref("age", r)}="1",{ref("age", r)}="2",{ref("age", r)}="3",{ref("age", r)}="4",{ref("age", r)}="5"),35,IF({ref("age", r)}="6-10",28,20)))'
                         f'+IF({ref("elev", r)}="Var",15,0)+IF({ref("site", r)}="Evet",10,0)+IF(ISNUMBER(SEARCH("Kapalı",{ref("park", r)})),15,IF(ISNUMBER(SEARCH("Açık",{ref("park", r)})),8,0))'
                         f'+IF({ref("kredi", r)}="Krediye Uygun",10,0)+IF({ref("balc", r)}="Var",5,0)+IF({ref("tapu", r)}="Kat Mülkiyeti",5,0))'), fmt=NUM)
    put(M, r, CI['s5'], f'=100-{ref("risk", r)}', fmt=NUM)
    put(M, r, CI['s6'], (f"=MIN(100,0.5*INDEX({q('mah')}$P$6:$P$15,{MI(r)})+IF(OR(AND({ref('room', r)}=\"1+1\",{ref('sqm', r)}>=45,{ref('sqm', r)}<=75),AND({ref('room', r)}=\"2+1\",{ref('sqm', r)}>=75,{ref('sqm', r)}<=115)),30,15)"
                         f"+IF({ref('kredi', r)}=\"Krediye Uygun\",20,0))"), fmt=NUM)
    put(M, r, CI['dataok'], f'=IF({ref("ncomp", r)}<3,"Yetersiz","Yeterli")')
    put(M, r, CI['score'], f'=ROUND((W_GETIRI*{ref("s1", r)}+W_KONUM*{ref("s2", r)}+W_FIYAT*{ref("s3", r)}+W_KALITE*{ref("s4", r)}+W_RISK*{ref("s5", r)}+W_LIKIT*{ref("s6", r)})*IF({ref("dataok", r)}="Yetersiz",P_VERI_CEZA,1),1)', fmt=NUM1, font=fBold)
    put(M, r, CI['helper'], f'={ref("score", r)}+ROW()/10000000', fmt='0.0000')
    put(M, r, CI['rank'], f'=RANK({ref("helper", r)},${C["helper"]}${FR}:${C["helper"]}${LR})', fmt=NUM, font=Font(name=F, size=9, bold=True, color=NAVY))
    put(M, r, CI['nyrank'], f'=RANK({ref("ny", r)},${C["ny"]}${FR}:${C["ny"]}${LR})+COUNTIF(${C["ny"]}${FR}:{C["ny"]}{r},{ref("ny", r)})-1', fmt=NUM)
    if idx % 2:
        for cc in range(1, len(COLS) + 1):
            cell = M.cell(r, cc)
            if cell.fill is None or cell.fill.fgColor.rgb in (None, '00000000'): cell.fill = fillG
    M.row_dimensions[r].height = 30
M.freeze_panes = 'C5'; M.auto_filter.ref = f'A4:{GL(len(COLS))}{LR}'
rng = lambda k: f'{C[k]}{FR}:{C[k]}{LR}'
for k in ('ny', 'gy'):
    M.conditional_formatting.add(rng(k), ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
M.conditional_formatting.add(rng('risk'), ColorScaleRule(start_type='min', start_color='63BE7B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='F8696B'))
M.conditional_formatting.add(rng('score'), DataBarRule(start_type='num', start_value=0, end_type='num', end_value=100, color='2E7D32'))
M.conditional_formatting.add(rng('pratio'), CellIsRule(operator='greaterThan', formula=['1.15'], fill=PatternFill('solid', fgColor='F4CCCC')))
M.conditional_formatting.add(rng('pratio'), CellIsRule(operator='lessThan', formula=['0.9'], fill=PatternFill('solid', fgColor='C6EFCE')))
M.conditional_formatting.add(rng('risklvl'), CellIsRule(operator='equal', formula=['"Yüksek"'], fill=PatternFill('solid', fgColor='F4CCCC')))
M.conditional_formatting.add(rng('rank'), CellIsRule(operator='lessThanOrEqual', formula=['10'], fill=PatternFill('solid', fgColor='C6EFCE')))
M.conditional_formatting.add(rng('tapu'), CellIsRule(operator='notEqual', formula=['"Kat Mülkiyeti"'], fill=PatternFill('solid', fgColor='FCE4D6')))

# ================= KİRA EMSALLERİ =================
K = ws['comp']
title(K, 'KİRA EMSALLERİ – Emlakjet + Century 21 + Turyap kiralık ilanlarından', 'UYARI: Emsal kiralar ilanlarda İSTENEN aylık kiradır; gerçekleşmiş sözleşme kirası değildir. Seçim: ≤2,5 km, aynı oda tipi, ±%35 brüt m², yeni binalar öncelikli, aykırı değerler çıkarıldı.', 16)
put(K, 3, 1, f'Kiralık ilan havuzu: emlakjet.com, Kadıköy (tüm mahalleler), 1+1/2+1 → {DATA["rentpool"]} tekil ilan. Mesafe kuş uçuşudur.', font=fSm, border=False)
KH = [('Daire No', 9), ('Emsal #', 6), ('Kiralık ilan', 11), ('Emsal mahallesi', 14), ('Oda', 6), ('Brüt m²', 7), ('Net m²', 7), ('Kat', 10), ('Bina yaşı', 11), ('Aylık ilan kirası (₺)', 11),
      ('Kira ₺/net m²', 10), ('Eşya', 9), ('Site', 6), ('Mesafe (m)', 8), ('İlan güncelleme', 11), ('Satılık daire mahallesi', 14), ('Kaynak portal', 14)]
for i, (h, w) in enumerate(KH, 1): hdr(K, 5, i, h, w)
MR = f"{q('main')}$A${FR}:$A${LR}"
for i, cp in enumerate(COMPS):
    r = 6 + i
    put(K, r, 1, cp['no'], font=Font(name=F, size=9, bold=True, color=NAVY)); put(K, r, 2, cp['k']); link(K, r, 3, cp['url'], 'İlan ↗')
    put(K, r, 4, cp['mah']); put(K, r, 5, cp['room']); put(K, r, 6, cp['brut'], fmt=NUM); put(K, r, 7, cp['net'] or '', fmt=NUM); put(K, r, 8, cp['floor']); put(K, r, 9, cp['age'])
    put(K, r, 10, cp['rent'], fmt=TL); put(K, r, 11, f'=J{r}/IF(ISNUMBER(G{r}),G{r},F{r}*P_NET_ORAN)', fmt=TL2); put(K, r, 12, cp['esya'] or '-'); put(K, r, 13, cp['site'] or '-'); put(K, r, 14, cp['dist'], fmt=NUM)
    put(K, r, 15, datetime.date.fromisoformat(cp['upd']) if cp['upd'] else '-', fmt=DATE)
    put(K, r, 16, f"=INDEX({q('main')}${C['mah']}${FR}:${C['mah']}${LR},MATCH(A{r},{MR},0))", font=fX)
    put(K, r, 17, (cp.get('src') or '') + ('' if cp.get('precise', True) else ' (konum: mahalle merkezi)'))
KL = 6 + max(len(COMPS), 1) - 1
K.freeze_panes = 'A6'; K.auto_filter.ref = f'A5:Q{KL}'
K.conditional_formatting.add(f'K6:K{KL}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))

# ================= MAHALLE =================
H = ws['mah']
title(H, 'MAHALLE ANALİZİ – Satış ve kira ₺/m², brüt getiri, talep ve boş kalma riski', 'Sol blok: örneklemden formülle. Sağ blok: emlakjet mahalle piyasa göstergeleri (Kasım 2026, tüm daire ilanları – eski binalar dahil). Talep/likidite puanları analist değerlendirmesidir (sarı).', 22)
MH = [('Mahalle', 14), ('Örneklemdeki daire', 9), ('Ort. satış fiyatı (₺)', 13), ('Ort. ₺/m² (örneklem)', 11), ('Ort. emsal kira ₺/net m²', 10), ('Brüt getiri (örneklem)', 9), ('Kira çarpanı (yıl)', 8),
      ('Ort. NET getiri', 9), ('Piyasa ort. ₺/m² (emlakjet, tüm daireler)', 12), ('Satılık ilan (n)', 7), ('Piyasa ort. kira ₺/m² (emlakjet, brüt)', 11), ('Kiralık ilan (n)', 7),
      ('Piyasa brüt getirisi', 9), ('Emlakjet amortisman (yıl)', 9), ('Talep puanı', 8), ('Likidite puanı', 8), ('Boş kalma riski', 14), ('Avantajlar', 44), ('Dezavantajlar', 40), ('Emlakjet yıl sonu fiyat beklentisi', 10)]
for i, (h, w) in enumerate(MH, 1): hdr(H, 5, i, h, w)
MAINR = lambda k: f"{q('main')}${C[k]}${FR}:${C[k]}${LR}"
for i, m in enumerate(MAHS):
    r = 6 + i; mk = MARKET.get(m, {})
    put(H, r, 1, m, font=Font(name=F, size=9, bold=True, color=NAVY))
    put(H, r, 2, f"=COUNTIF({MAINR('mah')},A{r})", fmt=NUM)
    put(H, r, 3, f"=IFERROR(AVERAGEIF({MAINR('mah')},A{r},{MAINR('price')}),\"-\")", fmt=TL)
    put(H, r, 4, f"=IFERROR(AVERAGEIF({MAINR('mah')},A{r},{MAINR('ppsm')}),\"-\")", fmt=TL)
    put(H, r, 5, f"=IFERROR(AVERAGEIF({q('comp')}$P$6:$P${KL},A{r},{q('comp')}$K$6:$K${KL}),K{r}/P_NET_ORAN)", fmt=TL2)
    put(H, r, 6, f"=IFERROR(AVERAGEIF({MAINR('mah')},A{r},{MAINR('gy')}),\"-\")", fmt=PCT); put(H, r, 7, f"=IFERROR(AVERAGEIF({MAINR('mah')},A{r},{MAINR('mult')}),\"-\")", fmt=NUM1)
    put(H, r, 8, f"=IFERROR(AVERAGEIF({MAINR('mah')},A{r},{MAINR('ny')}),\"-\")", fmt=PCT)
    put(H, r, 9, mk.get('sale_m2'), fmt=TL); put(H, r, 10, mk.get('n_sale'), fmt=NUM); put(H, r, 11, mk.get('rent_m2'), fmt=TL2); put(H, r, 12, mk.get('n_rent') or '', fmt=NUM)
    put(H, r, 13, f'=IFERROR(K{r}*12/I{r},"-")', fmt=PCT); put(H, r, 14, mk.get('amort'), fmt=NUM)
    d1, d2, d3 = DEMAND[m]; put(H, r, 15, d1, fill=fillIn); put(H, r, 16, d2, fill=fillIn); put(H, r, 17, d3, fill=fillIn)
    a, d = MAH_TXT[m]; put(H, r, 18, a, align=WRAP); put(H, r, 19, d, align=WRAP); put(H, r, 20, (mk.get('exp') or 0) / 100, fmt=PCT1)
    H.row_dimensions[r].height = 44
HL = 6 + len(MAHS) - 1
for col in ('F', 'H', 'M'):
    H.conditional_formatting.add(f'{col}6:{col}{HL}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
H.freeze_panes = 'B6'
put(H, HL + 2, 1, 'Not: Emlakjet mahalle göstergeleri sitedeki tüm satılık/kiralık daire ilanlarının ortalamasıdır (bina yaşı ve oda tipi ayrımı yok). “Amortisman” = ortalama fiyat / yıllık ortalama kira. Kaynak: emlakjet.com mahalle sayfaları, 17.11.2026.', font=fSm, border=False)

# ================= RİSK =================
RS = ws['risk']
title(RS, 'RİSK VE EKSİK VERİLER – Daire bazında risk işaretleri (1 = risk var)', '3. satırdaki ağırlıklar (sarı) değiştirilebilir. Risk puanı = Σ(işaret × ağırlık), en fazla 100.', 22)
mq = lambda k, r: f'{q("main")}{C[k]}{r}'
FLAGS = [
 ('Tapu kat mülkiyeti değil (kat irtifakı / tapu kaydı yok / intifa / arsa tapulu)', 15, lambda r: f'=IF({mq("tapu", r)}="Kat Mülkiyeti",0,1)'),
 ('Kira emsali < 3', 12, lambda r: f'=IF({mq("ncomp", r)}<3,1,0)'),
 ('Eski ilan / güncelleme tarihi yok', 10, lambda r: f'=IF(ISNUMBER({mq("upd", r)}),IF(P_KONTROL-{mq("upd", r)}>P_ILAN_GUN,1,0),1)'),
 ('Krediye uygun değil / bilinmiyor', 8, lambda r: f'=IF({mq("kredi", r)}="Krediye Uygun",0,1)'),
 ('Kiracılı (sözleşme şartları bilinmiyor)', 8, lambda r: f'=IF({mq("usage", r)}="Kiracı Oturuyor",1,0)'),
 ('Fiyat emsalin %15+ üstünde', 8, lambda r: f'=IF({mq("pratio", r)}>1.15,1,0)'),
 ('Zemin / bahçe / giriş / bodrum katı', 6, lambda r: f'=IF({mq("floorn", r)}<1,1,0)'),
 ('Asansör yok / belirtilmemiş', 6, lambda r: f'=IF({mq("elev", r)}="Var",0,1)'),
 ('Sıfır bina – iskân / KDV / teslim teyidi', 5, lambda r: f'=IF(LEFT({mq("age", r)},1)="0",1,0)'),
 ('2010–2015 yapımı (TBDY 2018 öncesi)', 5, lambda r: f'=IF({mq("age", r)}="11-15",1,0)'),
 ('Otopark yok / belirtilmemiş', 4, lambda r: f'=IF(OR({mq("park", r)}="Yok",{mq("park", r)}="Belirtilmemiş"),1,0)'),
 ('Raylı sisteme > 1.000 m', 3, lambda r: f'=IF({mq("drail", r)}>1000,1,0)'),
 ('Aynı daire farklı fiyatlarla birden çok ilanda', 3, lambda r: f'=IF({mq("dups", r)}="Tekrarı bulunamadı",0,1)'),
 ('Aidat bilgisi yok', 2, lambda r: f'=IF({mq("fees", r)}="",1,0)'),
]
hdr(RS, 5, 1, 'Daire No', 9); hdr(RS, 5, 2, 'Mahalle', 13); hdr(RS, 5, 3, 'Fiyat (₺)', 12)
for j, (nm, w, fn) in enumerate(FLAGS): hdr(RS, 5, 4 + j, nm, 11); put(RS, 3, 4 + j, w, fill=fillIn, align=CENTER)
put(RS, 3, 3, 'Ağırlık →', font=fB, border=False)
RC = 4 + len(FLAGS); assert GL(RC) == 'R'
hdr(RS, 5, RC, 'RİSK PUANI', 9); hdr(RS, 5, RC + 1, 'Risk işaretleri (ağırlık sırasıyla)', 45); hdr(RS, 5, RC + 2, 'Eksik / teyit edilmesi gereken veriler', 50)
RS.row_dimensions[5].height = 80
last = GL(RC - 1)
for idx in range(N):
    r = 6 + idx; mr = FR + idx
    put(RS, r, 1, f"={mq('no', mr)}", font=fX); put(RS, r, 2, f"={mq('mah', mr)}", font=fX); put(RS, r, 3, f"={mq('price', mr)}", fmt=TL, font=fX)
    for j, (nm, w, fn) in enumerate(FLAGS): put(RS, r, 4 + j, fn(mr), fmt='0', align=CENTER)
    put(RS, r, RC, f'=MIN(100,SUMPRODUCT(D{r}:{last}{r},$D$3:${last}$3))', fmt=NUM, font=fBold)
    order = sorted(range(len(FLAGS)), key=lambda j: -FLAGS[j][1])
    parts = '&'.join(f'IF({GL(4 + j)}{r}=1,{GL(4 + j)}$5&"; ","")' for j in order)
    put(RS, r, RC + 1, f'=IF(R{r}=0,"Belirgin risk yok",{parts})', align=WRAP)
    put(RS, r, RC + 2, (f'=IF({mq("fees", mr)}="","Aidat; ","")&IF({mq("net", mr)}="Bilinmiyor","Net m²; ","")&"Tam adres; "&IF({mq("tapu", mr)}<>"Kat Mülkiyeti","İskân/kat mülkiyeti; ","")'
                        f'&IF({mq("park", mr)}="Belirtilmemiş","Otopark; ","")&IF(LEFT({mq("status", mr)},6)<>"Yayınd","İlan güncelliği; ","")&"Ruhsat tarihi, yapı denetim raporu, DASK poliçesi (her daire)"'), align=WRAP)
RL = 6 + N - 1
RS.freeze_panes = 'D6'; RS.auto_filter.ref = f'A5:{GL(RC + 2)}{RL}'
RS.conditional_formatting.add(f'D6:{last}{RL}', CellIsRule(operator='equal', formula=['1'], fill=PatternFill('solid', fgColor='F4CCCC')))
RS.conditional_formatting.add(f'R6:R{RL}', ColorScaleRule(start_type='min', start_color='63BE7B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='F8696B'))
r = RL + 2; put(RS, r, 1, 'Risk özeti', font=fB, border=False)
for j, (nm, w, fn) in enumerate(FLAGS): put(RS, r + 1 + j, 2, nm, font=fN); put(RS, r + 1 + j, 3, f'=SUM({GL(4 + j)}6:{GL(4 + j)}{RL})', fmt=NUM)

# ================= SENARYO per apartment =================
hdr(S, 18, 1, 'Daire bazında üç senaryo (1. yıl, aynı formül zinciri)'); S.merge_cells('A18:G18')
blocks = [('Kötümser', 'SK'), ('Baz', 'SB'), ('İyimser', 'SI')]
sub = ['Aylık kira (₺)', 'Tahsil edilen yıllık kira', 'İşletme gideri', 'Kira vergisi', 'Vergi sonrası net', 'Toplam yatırım', 'Net getiri']
hdr(S, 20, 1, 'Daire No'); hdr(S, 20, 2, 'Mahalle'); hdr(S, 20, 3, 'İlan fiyatı (₺)')
col = 4
for bn, pf in blocks:
    c = S.cell(19, col, bn); c.font = fH; c.fill = fillT; c.alignment = CENTER
    S.merge_cells(start_row=19, start_column=col, end_row=19, end_column=col + 6)
    for s_ in sub: hdr(S, 20, col, s_, 11); col += 1
hdr(S, 20, col, 'Fark: İyimser – Kötümser', 11); S.row_dimensions[20].height = 40
for idx in range(N):
    r = 21 + idx; mr = FR + idx; m = lambda k: mq(k, mr)
    put(S, r, 1, f"={m('no')}", font=fX); put(S, r, 2, f"={m('mah')}", font=fX); put(S, r, 3, f"={m('price')}", fmt=TL, font=fX)
    col = 4
    for bn, pf in blocks:
        L = [GL(col + i) for i in range(7)]
        put(S, r, col, f"=IF(ISNUMBER({m('rentnow')}),{m('rentnow')},ROUND({m('basem2')}*{m('rsqm')}*{pf}_KIRA*(1-IF({m('floorn')}<1,P_ZEMIN_IND,0)),-2))", fmt=TL)
        put(S, r, col + 1, f'={L[0]}{r}*(12-{pf}_BOS)', fmt=TL)
        put(S, r, col + 2, f"={L[0]}{r}*12*P_BAKIM*{pf}_MAS+{L[1]}{r}*P_YONETIM+{L[0]}{r}*P_KIRACI_AY/P_SOZLESME+{m('price')}*P_EV+{m('sqm')}*P_SIGORTA*{pf}_MAS+IF({m('fees')}=\"\",P_AIDAT_VARS,{m('fees')})*({pf}_BOS+(12-{pf}_BOS)*P_AIDAT_PAY)", fmt=TL)
        put(S, r, col + 3, '=' + taxf(f'{L[1]}{r}'), fmt=TL)
        put(S, r, col + 4, f'={L[1]}{r}-{L[2]}{r}-{L[3]}{r}', fmt=TL)
        put(S, r, col + 5, f"={m('price')}*(1-{pf}_IND)*(1+P_MASRAF_PCT)+P_DONER+P_SABIT+{m('furncost')}", fmt=TL)
        put(S, r, col + 6, f'={L[4]}{r}/{L[5]}{r}', fmt=PCT, font=fBold)
        col += 7
    put(S, r, col, f'={GL(col - 1)}{r}-{GL(col - 8)}{r}', fmt=PCT)
SL = 21 + N - 1
for c0 in (10, 17, 24):
    S.conditional_formatting.add(f'{GL(c0)}21:{GL(c0)}{SL}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
S.freeze_panes = 'D21'
put(S, 17, 1, 'Portföy ortalaması net getiri (Kötümser / Baz / İyimser) →', font=fB, border=False)
for j, c0 in enumerate((10, 17, 24)): put(S, 17, 3 + j, f'=AVERAGE({GL(c0)}21:{GL(c0)}{SL})', fmt=PCT, font=Font(name=F, size=10, bold=True, color=GREEN_D))
dname('SCN_AVG_K', f"{q('scen')}$C$17"); dname('SCN_AVG_B', f"{q('scen')}$D$17"); dname('SCN_AVG_I', f"{q('scen')}$E$17")

# ================= NAKİT AKIŞI =================
CF = ws['cf']
title(CF, '5 VE 10 YILLIK NAKİT AKIŞI – Enflasyon, kira artışı (TBK 344), değer artışı, kur ve IRR (TL ve €)',
      'Makro yol (satır 5–8) aktif senaryodan gelir. Üst blok: seçilen daire; alt blok: tüm daireler. TL nominal IRR yüksek enflasyon nedeniyle yanıltıcı olabilir – € bazlı IRR ve reel getiri ile birlikte okuyun.', 30)
CF.column_dimensions['A'].width = 36; CF.column_dimensions['B'].width = 14
hdr(CF, 4, 1, 'Makro yol (aktif senaryo)')
for y in range(11): hdr(CF, 4, 2 + y, f'Yıl {y}', 12)
labs = ['TÜFE (yıllık)', 'Fiyat endeksi (kümülatif)', 'Kira endeksi (yıl başı kirası / 1. yıl)', 'Mülk değer endeksi', 'EUR/TRY kuru (yıl sonu)']
for i, l in enumerate(labs): put(CF, 5 + i, 1, l, font=fN)
for y in range(11):
    cl = GL(2 + y); pc = GL(1 + y)
    if y == 0:
        put(CF, 5, 2, 0, fmt=PCT1, font=fN); put(CF, 6, 2, 1, fmt='0.000', font=fN); put(CF, 7, 2, '', font=fN); put(CF, 8, 2, 1, fmt='0.000', font=fN); put(CF, 9, 2, '=P_EURTRY', fmt='0.00')
        continue
    put(CF, 5, 2 + y, f'=SA_ENFU+(SA_ENF1-SA_ENFU)*MAX(0,(5-{y})/4)', fmt=PCT1)
    put(CF, 6, 2 + y, f'={pc}6*(1+{cl}5)', fmt='0.000')
    put(CF, 7, 2 + y, '=1' if y == 1 else f'={pc}7*(1+{pc}5+SA_KREEL)', fmt='0.000')
    put(CF, 8, 2 + y, f'={pc}8*(1+{cl}5)*(1+SA_DREEL)', fmt='0.000')
    put(CF, 9, 2 + y, f'={pc}9*(1+{cl}5)/(1+P_EUR_ENF)/(1+SA_KURREEL)', fmt='0.00')
put(CF, 11, 1, 'Seçili daire →', font=fB, border=False); put(CF, 11, 2, R[0]['no'], fill=fillIn)
dvn = DataValidation(type='list', formula1=f"={q('main')}$A${FR}:$A${LR}"); CF.add_data_validation(dvn); dvn.add('B11')
def mlook(k): return f"INDEX({q('main')}${C[k]}${FR}:${C[k]}${LR},MATCH($B$11,{q('main')}$A${FR}:$A${LR},0))"
info = [('Mahalle', 'mah', None), ('İlan fiyatı', 'price', TL), ('Alım bedeli', 'buy', TL), ('Toplam yatırım', 'tinv', TL), ('1. yıl tahsil edilen kira', 'coll', TL), ('1. yıl işletme gideri', 'opex', TL)]
for i, (lab, k, fmt) in enumerate(info): put(CF, 12 + i, 1, lab, font=fN); put(CF, 12 + i, 2, f'={mlook(k)}', fmt=fmt, font=fX)
# rows: 12 mah,13 price,14 buy,15 tinv,16 coll,17 opex
hdr(CF, 19, 1, 'Kalem (₺)')
for y in range(11): hdr(CF, 19, 2 + y, f'Yıl {y}')
rows_cf = ['Tahsil edilen kira', 'İşletme giderleri', 'Kira geliri vergisi', 'Vergi sonrası net nakit akışı', 'Kümülatif net kira', 'Mülk değeri (yıl sonu)',
           '5. yılda çıkış: net satış geliri', 'Nakit akışı – 5 yıl (₺)', '10. yılda çıkış: net satış geliri', 'Nakit akışı – 10 yıl (₺)', 'Nakit akışı – 5 yıl (€)', 'Nakit akışı – 10 yıl (€)']
for i, lab in enumerate(rows_cf): put(CF, 20 + i, 1, lab, font=fBold if 'Nakit akışı' in lab else fN)
for y in range(11):
    cl = GL(2 + y); pc = GL(1 + y)
    if y == 0:
        for rr_ in (20, 21, 22, 23, 24, 26, 28): put(CF, rr_, 2, 0, fmt=TL, font=fN)
        put(CF, 25, 2, '=B14', fmt=TL); put(CF, 27, 2, '=-B15', fmt=TL); put(CF, 29, 2, '=-B15', fmt=TL)
        put(CF, 30, 2, '=B27/B9', fmt=EUR); put(CF, 31, 2, '=B29/B9', fmt=EUR)
        continue
    put(CF, 20, 2 + y, f'=$B$16*{cl}7', fmt=TL)
    put(CF, 21, 2 + y, f'=$B$17*{pc}6', fmt=TL)
    put(CF, 22, 2 + y, '=' + taxf(f'{cl}20', f'{pc}6'), fmt=TL)
    put(CF, 23, 2 + y, f'={cl}20-{cl}21-{cl}22', fmt=TL)
    put(CF, 24, 2 + y, f'={pc}24+{cl}23', fmt=TL)
    put(CF, 25, 2 + y, f'=$B$14*{cl}8', fmt=TL)
    put(CF, 26, 2 + y, f'=IF({y}=5,{cl}25*(1-SA_CIKIS),0)', fmt=TL)
    put(CF, 27, 2 + y, f'=IF({y}<=5,{cl}23+{cl}26,0)', fmt=TL)
    put(CF, 28, 2 + y, f'=IF({y}=10,{cl}25*(1-SA_CIKIS),0)', fmt=TL)
    put(CF, 29, 2 + y, f'={cl}23+{cl}28', fmt=TL)
    put(CF, 30, 2 + y, f'={cl}27/{cl}9', fmt=EUR); put(CF, 31, 2 + y, f'={cl}29/{cl}9', fmt=EUR)
for rr_ in (23, 27, 29, 30, 31):
    for y in range(11): CF.cell(rr_, 2 + y).fill = fillGr
put(CF, 33, 1, 'Özet', font=fB, border=False)
summ = [('5 yıl – IRR (TL nominal)', '=IRR(B27:G27)', PCT), ('5 yıl – IRR (€)', '=IRR(B30:G30)', PCT), ('5 yıl – reel IRR (TÜFE’den arındırılmış, yaklaşık)', '=(1+B34)/(G6^(1/5))-1', PCT),
        ('10 yıl – IRR (TL nominal)', '=IRR(B29:L29)', PCT), ('10 yıl – IRR (€)', '=IRR(B31:L31)', PCT), ('10 yıl – reel IRR (yaklaşık)', '=(1+B37)/(L6^(1/10))-1', PCT),
        ('10 yıl – kümülatif net kira (₺)', '=L24', TL), ('10 yıl – toplam kazanç (€)', '=SUM(B31:L31)', EUR)]
for i, (lab, f_, fmt) in enumerate(summ): put(CF, 34 + i, 1, lab, font=fN); put(CF, 34 + i, 2, f_, fmt=fmt, font=fBold)
G0 = 46
put(CF, G0 - 1, 1, 'Tüm daireler – 5 ve 10 yıllık tutma (aktif senaryo, € bazında nakit akışları)', font=fB, border=False)
heads = ['Daire No', 'Mahalle', 'Toplam yatırım (€)'] + [f'5Y € Yıl {y}' for y in range(6)] + [f'10Y € Yıl {y}' for y in range(11)] + ['5Y IRR (€)', '10Y IRR (€)', '5Y IRR (TL, yaklaşık)', '10Y IRR (TL, yaklaşık)', '10Y toplam getiri (€)']
for i, h in enumerate(heads): hdr(CF, G0, 1 + i, h, None if i < 2 else 10)
CF.row_dimensions[G0].height = 30
for idx in range(N):
    r = G0 + 1 + idx; mr = FR + idx; m = lambda k: mq(k, mr)
    put(CF, r, 1, f"={m('no')}", font=fX); put(CF, r, 2, f"={m('mah')}", font=fX); put(CF, r, 3, f"={m('tinv')}/$B$9", fmt=EUR, font=fX)
    def ncf(y):
        cl = GL(2 + y); pc = GL(1 + y)
        return f"({m('coll')}*${cl}$7-{m('opex')}*${pc}$6-{taxf(m('coll') + '*$' + cl + '$7', '$' + pc + '$6')})"
    def exitv(y):
        cl = GL(2 + y); return f"{m('buy')}*${cl}$8*(1-SA_CIKIS)"
    for y in range(6):
        cl = GL(2 + y)
        put(CF, r, 4 + y, f'=-C{r}' if y == 0 else f'=({ncf(y)}' + (f'+{exitv(5)}' if y == 5 else '') + f')/${cl}$9', fmt=EUR)
    for y in range(11):
        cl = GL(2 + y)
        put(CF, r, 10 + y, f'=-C{r}' if y == 0 else f'=({ncf(y)}' + (f'+{exitv(10)}' if y == 10 else '') + f')/${cl}$9', fmt=EUR)
    put(CF, r, 21, f'=IRR(D{r}:I{r})', fmt=PCT, font=fBold); put(CF, r, 22, f'=IRR(J{r}:T{r})', fmt=PCT, font=fBold)
    put(CF, r, 23, f'=(1+U{r})*($G$9/$B$9)^(1/5)-1', fmt=PCT); put(CF, r, 24, f'=(1+V{r})*($L$9/$B$9)^(1/10)-1', fmt=PCT)
    put(CF, r, 25, f'=SUM(J{r}:T{r})/C{r}', fmt=PCT1)
GE = G0 + N
for cc in ('U', 'V'):
    CF.conditional_formatting.add(f'{cc}{G0 + 1}:{cc}{GE}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
put(CF, 43, 1, 'Portföy ortalaması: 5Y / 10Y IRR (€)', font=fB); put(CF, 43, 2, f'=AVERAGE(U{G0 + 1}:U{GE})', fmt=PCT); put(CF, 43, 3, f'=AVERAGE(V{G0 + 1}:V{GE})', fmt=PCT)
put(CF, 44, 1, 'Portföy ortalaması: 5Y / 10Y IRR (TL)', font=fB); put(CF, 44, 2, f'=AVERAGE(W{G0 + 1}:W{GE})', fmt=PCT); put(CF, 44, 3, f'=AVERAGE(X{G0 + 1}:X{GE})', fmt=PCT)
CF.freeze_panes = 'B5'

# ================= EN İYİ 20 =================
TP = ws['top']
title(TP, 'EN İYİ 20 YATIRIM – Genel yatırım puanına göre', 'Sıralama INDEX/MATCH formülleriyle çalışır; parametre veya senaryo değişince güncellenir.', 20)
TH = [('Sıra', 6), ('Daire No', 9), ('İlan', 8), ('Mahalle', 13), ('Oda', 6), ('Fiyat (₺)', 13), ('Fiyat (€)', 10), ('Brüt m²', 7), ('Aylık kira – aktif (₺)', 11), ('Toplam yatırım (₺)', 13),
      ('Net nakit (₺/yıl)', 11), ('Brüt getiri', 8), ('NET GETİRİ', 9), ('Risk', 6), ('GENEL PUAN', 8), ('10Y IRR (€)', 8), ('Pazarlık önerilen fiyat (₺)', 13), ('Neden seçildi', 48), ('Risk işaretleri', 34)]
for i, (h, w) in enumerate(TH, 1): hdr(TP, 4, i, h, w)
TP.row_dimensions[4].height = 40
TOPN = min(20, N)
for k in range(TOPN):
    r = 5 + k
    put(TP, r, 1, k + 1, font=Font(name=F, size=10, bold=True, color=NAVY), align=CENTER)
    MT = f"MATCH($A{r},{q('main')}${C['rank']}${FR}:${C['rank']}${LR},0)"
    I = lambda kk: f"INDEX({q('main')}${C[kk]}${FR}:${C[kk]}${LR},{MT})"
    put(TP, r, 2, f'={I("no")}', font=Font(name=F, size=9, bold=True, color='008000'))
    put(TP, r, 3, f'=HYPERLINK(INDEX({q("src")}$C$7:$C${6 + N},MATCH(B{r},{q("src")}$A$7:$A${6 + N},0)),"İlan ↗")', font=fLink)
    for j, (kk, fmt) in enumerate([('mah', None), ('room', None), ('price', TL), ('price_eur', EUR), ('sqm', NUM), ('ra', TL), ('tinv', TL), ('ncf', TL), ('gy', PCT), ('ny', PCT), ('risk', NUM), ('score', NUM1)]):
        put(TP, r, 4 + j, f'={I(kk)}', fmt=fmt, font=fX)
    put(TP, r, 16, f"=INDEX({q('cf')}$V${G0 + 1}:$V${GE},{MT})", fmt=PCT, font=fX)
    put(TP, r, 17, f'={I("nego")}', fmt=TL, font=fX)
    why = (f'=IF({I("ny")}>=AVERAGE({q("main")}{C["ny"]}{FR}:{C["ny"]}{LR}),"Net getiri ortalamanın üstünde (%"&ROUND({I("ny")}*100,2)&"); ","")'
           f'&IF({I("pratio")}<0.95,"₺/m² emsal medyanın altında ("&ROUND({I("pratio")},2)&"x); ","")'
           f'&IF({I("drail")}<=600,{I("rail")}&" "&{I("drail")}&" m; ","")'
           f'&IF({I("dbagdat")}<=300,"Bağdat Cd. yakın; ","")&IF(LEFT({I("age")},1)="0","Sıfır bina; ","")'
           f'&IF({I("tapu")}="Kat Mülkiyeti","Kat mülkiyetli; ","")&IF({I("risk")}<18,"Düşük risk; ","")&"Kira "&{I("ncomp")}&" emsalle destekli"')
    put(TP, r, 18, why, align=WRAP); put(TP, r, 19, f'={I("mainrisk")}', font=fX, align=WRAP)
    TP.row_dimensions[r].height = 42
TP.conditional_formatting.add('M5:M24', ColorScaleRule(start_type='min', start_color='FFEB84', end_type='max', end_color='63BE7B'))
TP.conditional_formatting.add('O5:O24', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=100, color='2E7D32'))
TP.freeze_panes = 'C5'
put(TP, 27, 1, 'Ek tablo: Yalnızca NET KİRA GETİRİSİNE göre en yüksek 20 daire', font=fB, border=False)
for i, h in enumerate(['Net getiri sırası', 'Daire No', 'Mahalle', 'Oda', 'Fiyat (₺)', 'Net getiri', 'Genel puan sırası', 'Risk', 'Veri'], 1): hdr(TP, 28, i, h)
for k in range(TOPN):
    r = 29 + k; put(TP, r, 1, k + 1, align=CENTER)
    MT = f"MATCH($A{r},{q('main')}${C['nyrank']}${FR}:${C['nyrank']}${LR},0)"
    I = lambda kk: f"INDEX({q('main')}${C[kk]}${FR}:${C[kk]}${LR},{MT})"
    for j, (kk, fmt) in enumerate([('no', None), ('mah', None), ('room', None), ('price', TL), ('ny', PCT), ('rank', NUM), ('risk', NUM), ('dataok', None)]):
        put(TP, r, 2 + j, f'={I(kk)}', fmt=fmt, font=fX)

# ================= KAYNAKLAR =================
SR = ws['src']
title(SR, 'KAYNAKLAR – İlan bağlantıları, kira emsalleri, veri kaynakları ve kontrol tarihleri', 'Tüm bağlantılar tıklanabilir. Kontrol tarihi: 17.11.2026.', 9)
put(SR, 4, 1, 'A) Satılık ilanlar (ana liste)', font=fB, border=False)
for i, (h, w) in enumerate([('Daire No', 9), ('İlan no', 11), ('İlan URL', 70), ('Emlak ofisi', 26), ('Danışman', 26), ('İlk yayın', 11), ('Son güncelleme', 11), ('Kontrol', 11)], 1): hdr(SR, 6, i, h, w)
for idx, rec in enumerate(R):
    r = 7 + idx
    put(SR, r, 1, rec['no']); put(SR, r, 2, str(rec['ilan_no'])); c = put(SR, r, 3, rec['url']); c.hyperlink = rec['url']; c.font = fLink
    put(SR, r, 4, rec['adv']); put(SR, r, 5, rec['agent']); put(SR, r, 6, rec['created'] or '-'); put(SR, r, 7, rec['updated'] or '-'); put(SR, r, 8, CHECK, fmt=DATE)
r = 7 + N + 1
put(SR, r, 1, 'B) Aynı dairenin diğer ilanları (mükerrer)', font=fB, border=False); r += 1
for i, h in enumerate(['Daire No', '', 'İlan URL', 'Emlak ofisi', 'Fiyat (₺)'], 1): hdr(SR, r, i, h)
for rec in R:
    for x in rec['dups']:
        r += 1; put(SR, r, 1, rec['no']); c = put(SR, r, 3, x['url']); c.hyperlink = x['url']; c.font = fLink; put(SR, r, 4, x['adv'] or '-'); put(SR, r, 5, x['price'], fmt=TL)
r += 2
put(SR, r, 1, f'C) Kira emsalleri: KİRA EMSALLERİ sayfasında tıklanabilir ({len(COMPS)} satır).', font=fB, border=False); r += 2
put(SR, r, 1, 'D) Veri ve resmî/yarı resmî kaynaklar', font=fB, border=False); r += 1
SRCS = [
 ('İlan verisi', 'emlakjet.com – Kadıköy Caferağa, Kozyatağı, Caddebostan, Suadiye, Fenerbahçe, Feneryolu, Göztepe, Erenköy, Zühtüpaşa, Koşuyolu satılık daire + Kadıköy satılık residence listeleri ve ilan detay sayfaları.', 'https://www.emlakjet.com/satilik-daire/istanbul-kadikoy', '17.11.2026'),
 ('Kira verisi', 'emlakjet.com – Kadıköy tüm kiralık daire + residence ilanları (1+1/2+1).', 'https://www.emlakjet.com/kiralik-daire/istanbul-kadikoy', '17.11.2026'),
 ('Piyasa', 'emlakjet.com mahalle piyasa göstergeleri (ortalama m² fiyatı, kira, amortisman, beklenti).', 'https://www.emlakjet.com/satilik-daire/istanbul-kadikoy-caferaga-mahallesi', '17.11.2026'),
 ('Kur', 'TCMB gösterge niteliğindeki döviz kurları (17.11.2026).', 'https://www.tcmb.gov.tr/kurlar/today.xml', '17.11.2026'),
 ('Enflasyon', 'TÜİK TÜFE Ekim 2026: yıllık %29,77, 12 aylık ortalama %31,23 (basın).', 'https://www.alomaliye.com/2026/11/03/enflasyon-rakamlari-tufe-ekim-2026/', '17.11.2026'),
 ('Vergi', '2026 gelir vergisi tarifesi ve konut kira istisnası (58.000 ₺) – KPMG.', 'https://kpmgvergi.com/yayinlar/mali-bultenler/vergi/2026-yili-gelir-vergisi-dilimleri-ile-dikkate-alinacak-bazi-had-ve-tutarlar-belirlendi/3351', '17.11.2026'),
 ('Vergi', 'PwC – 2026 gelir vergisi tarifesi.', 'https://www.pwc.com.tr/tr/hizmetlerimiz/vergi/vergi-portali/pratik-bilgiler/oranlar/gelir-vergisi-tarifesi.html', '17.11.2026'),
 ('Vergi', 'GİB – Gelir İdaresi Başkanlığı (kira geliri rehberi, tebliğler).', 'https://www.gib.gov.tr', 'Teyit için'),
 ('Harç', 'Tapu harcı %4 (alıcı %2 + satıcı %2), döner sermaye 2026.', 'https://www.tahanci.av.tr/tapu-harci/', '17.11.2026'),
 ('Konum', 'Raylı sistem koordinatları – Wikipedia (Marmaray Suadiye, Bostancı, Feneryolu, Göztepe, Söğütlüçeşme; M4 Göztepe, Yenisahra, Kozyatağı, Bostancı) ve bir GitHub veri seti (Ayrılık Çeşmesi, Acıbadem, Ünalan). Erenköy Marmaray ve Kadıköy M4 yaklaşık.', 'https://en.wikipedia.org/wiki/M4_(Istanbul_Metro)', '17.11.2026'),
 ('Konum', 'Bağdat Caddesi ve sahil güzergâhları elle yaklaşık koordinatlarla tanımlandı (±150–250 m).', '', '17.11.2026'),
 ('İlan verisi', 'Century 21 Türkiye – tüm satılık/kiralık konut listesi tarandı; Kadıköy ilanlarının detay sayfaları (koordinat yok → mahalle merkezi).', 'https://www.century21.com.tr/satilik/konut', '17.11.2026'),
 ('İlan verisi', 'Turyap – portföy arama formu (İstanbul / Kadıköy, satılık ve kiralık daire) ve detay sayfaları.', 'https://www.turyap.com.tr/Portfoyler.aspx', '17.11.2026'),
 ('Erişilemeyen', 'sahibinden.com ve hepsiemlak.com (bot koruması, HTTP 403 – aşılmadı), remax.com.tr ve coldwellbanker.com.tr (bağlantı yok), Realty World (mahalle filtresi JS ile yükleniyor). Arama motoru dizinindeki Sahibinden kayıtları 2020 tarihli olduğu için kullanılmadı.', '', '17.11.2026'),
]
for i, h in enumerate(['Tür', 'Açıklama', 'Bağlantı', 'Kontrol'], 1): hdr(SR, r, i, h)
for t_, d_, u_, k_ in SRCS:
    r += 1; put(SR, r, 1, t_); put(SR, r, 2, d_, align=WRAP)
    if u_: c = put(SR, r, 3, u_); c.hyperlink = u_; c.font = fLink
    else: put(SR, r, 3, '-')
    put(SR, r, 4, k_); SR.row_dimensions[r].height = 30
SR.column_dimensions['B'].width = 60

# ================= YEDEK =================
X = ws['ext']
title(X, 'YEDEK İLANLAR – Ana listeye alınmayan lüks mahalle 1+1/2+1 ilanları', 'Bu ilanlar modele DAHİL DEĞİLDİR. “Neden” sütunu kriter dışı kalma sebebini gösterir.', 10)
for i, (h, w) in enumerate([('#', 5), ('İlan', 8), ('Mahalle', 13), ('Oda', 6), ('Fiyat (₺)', 13), ('Brüt m²', 7), ('₺/m²', 10), ('Bina yaşı', 14), ('Kat', 10), ('Emlak ofisi', 24), ('Güncelleme', 11), ('Neden ana listede değil', 44)], 1): hdr(X, 4, i, h, w)
for i, e in enumerate(EXTRAS):
    r = 5 + i
    put(X, r, 1, i + 1); link(X, r, 2, e['url'], 'İlan ↗'); put(X, r, 3, e['mah']); put(X, r, 4, e['room']); put(X, r, 5, e['price'], fmt=TL); put(X, r, 6, e['brut'] or '', fmt=NUM)
    put(X, r, 7, f'=IFERROR(E{r}/F{r},"-")', fmt=TL); put(X, r, 8, e['age']); put(X, r, 9, e['floor']); put(X, r, 10, e['adv'] or '-'); put(X, r, 11, e['upd']); put(X, r, 12, e['why'], align=WRAP)
X.freeze_panes = 'A5'; X.auto_filter.ref = f'A4:L{4 + max(len(EXTRAS), 1)}'

# ================= VERGİ-HUKUK =================
V = ws['tax']
title(V, 'SATIN ALMA, VERGİ VE HUKUKİ KONTROL NOTLARI (Türkiye vergi mukimi bireysel yatırımcı)', 'Genel bilgidir; hukuki/vergisel tavsiye değildir. Kontrol: 17.11.2026', 3)
V.column_dimensions['A'].width = 34; V.column_dimensions['B'].width = 120
TXT = [
 ('1. SATIN ALMA MALİYETLERİ', ''),
 ('Tapu harcı', 'Toplam %4: alıcı %2 + satıcı %2 (uygulamada pazarlık konusu). Matrah: beyan edilen satış bedeli ile emlak vergi değerinden yüksek olanı. Gerçek bedelin altında beyan vergi ziyaı cezasına yol açar ve alıcı açısından ileride değer artış kazancı vergisini büyütür.'),
 ('Döner sermaye', '2026 için kaynaklarda ≈2.534 ₺; bölge katsayısına göre değişir.'),
 ('KDV (sıfır bina)', 'Müteahhitten ilk satışta KDV uygulanır (net alan, arsa birim değeri ve yapı ruhsat tarihine göre %1/%10/%20). İlan fiyatının KDV dahil olup olmadığını yazılı teyit edin. İkinci el satışta KDV yoktur.'),
 ('Emlakçı', 'Alıcı komisyonu %2 + KDV (piyasa uygulaması); yer gösterme belgesi imzalatılır.'),
 ('2. YILLIK VERGİLER VE GİDERLER', ''),
 ('Emlak vergisi', 'Büyükşehirde konut için binde 2; matrah emlak vergi değeridir. Mayıs ve Kasım’da iki taksit.'),
 ('DASK', 'Zorunlu deprem sigortası; tapu işlemi ve abonelikler için gerekir.'),
 ('3. KİRA GELİRİNİN VERGİLENDİRİLMESİ', ''),
 ('İstisna', '2026 konut kira geliri istisnası 58.000 ₺ (kişi başına, tüm konut kiralarına toplamda bir kez). Ticari/zirai/mesleki kazancı olanlar ve belirli gelir sınırını aşanlar yararlanamaz.'),
 ('Gider yöntemi', 'Götürü gider (%15) veya gerçek gider. Modelde götürü gider kullanılmıştır.'),
 ('Tarife (2026, ücret dışı)', '190.000 ₺’ye kadar %15 · 400.000 ₺’ye kadar %20 · 1.000.000 ₺’ye kadar %27 · 5.300.000 ₺’ye kadar %35 · üzeri %40. Ücret vb. gelirlerle birleşince marjinal oran yükselir.'),
 ('Beyan', 'Yıllık gelir vergisi beyannamesi Mart ayında; ödeme Mart ve Temmuz’da iki taksit. Kira tahsilatının banka yoluyla yapılması zorunludur (≥ belirli tutar).'),
 ('Kira artış sınırı', 'TBK 344: yenilenen dönemde kira artışı önceki 12 ayın TÜFE ortalamasını aşamaz; 5 yıl sonra (veya 5 yılı aşan sözleşmede) hâkim rayice göre tespit edebilir.'),
 ('Değer artış kazancı', '5 yıl elde tutulduktan sonra satılan konutta değer artış kazancı vergisi yoktur. 5 yıl dolmadan satışta, ÜFE endekslemesi sonrası kazanç vergilenir.'),
 ('4. AVUKATIN KONTROL ETMESİ GEREKENLER', ''),
 ('Tapu', 'TAKBİS kaydı: malik, hisse durumu, ipotek, haciz, şerh, beyanlar; kat mülkiyeti mi kat irtifakı mı? “Tapu kaydı yok”, “intifa hakkı” ve “arsa tapulu” ilanlar yüksek risklidir.'),
 ('İskân ve ruhsat', 'Yapı kullanma izin belgesi (iskân), yapı ruhsatı tarihi (deprem yönetmeliği dönemi), onaylı projeye uygunluk; kaçak kat/çekme kat/bodrumdan dönüştürme.'),
 ('Kiracı', 'Kiracılı satışta sözleşme, depozito, tahliye taahhüdü; yeni malik 6 ay içinde ihtiyaç nedeniyle tahliye davası koşulları.'),
 ('Site/apartman', 'Yönetim planı, aidat borcu, planlanan büyük onarım/dönüşüm kararları.'),
 ('5. MÜHENDİSİN KONTROL ETMESİ GEREKENLER', ''),
 ('Statik', 'Yapı denetim raporları, beton sınıfı (C30 ve üstü), zemin etüdü; 2010–2018 arası binalarda 2007 yönetmeliği, 2019 sonrası TBDY 2018.'),
 ('Tesisat ve yalıtım', 'Isı/ses yalıtımı, su yalıtımı (bodrum/çatı), doğalgaz ve elektrik projesi, asansör ruhsatı.'),
 ('Ölçüm', 'Net/brüt m² (ilan değerleri beyan), otopark ve depo hakları.'),
]
r = 4
for a_, b_ in TXT:
    if not b_:
        r += 1; c = V.cell(r, 1, a_); c.font = Font(name=F, size=11, bold=True, color='FFFFFF'); c.fill = fillH; V.cell(r, 2).fill = fillH
    else:
        put(V, r, 1, a_, font=fBold); put(V, r, 2, b_, font=fN, align=WRAP); V.row_dimensions[r].height = max(18, 14 * (len(b_) // 120 + 1))
    r += 1

# ================= PANEL =================
D = ws['panel']
title(D, f'KADIKÖY LÜKS MAHALLELER – ÇOK PORTALLI DAİRE YATIRIM ANALİZİ ({N} daire)', 'Kaynaklar: Emlakjet + Century 21 + Turyap · 2010 ve sonrası 1+1 / 2+1 · Kontrol tarihi 17.11.2026 · Tüm göstergeler formülle güncellenir', 12)
for cc, w in zip('ABCDEFGHIJK', [26, 15, 26, 15, 26, 15, 26, 15, 11, 11, 4]): D.column_dimensions[cc].width = w
put(D, 3, 1, 'Aktif senaryo:', font=fB, border=False); put(D, 3, 2, f"={q('scen')}F16", font=Font(name=F, size=11, bold=True, color=GREEN_D), border=False)
put(D, 3, 3, 'EUR/TRY:', font=fB, border=False); put(D, 3, 4, '=P_EURTRY', fmt='0.00', border=False)
MRN = lambda k: f"{q('main')}${C[k]}${FR}:${C[k]}${LR}"
KPIS = [('Araştırılan daire', f"=COUNTA({MRN('no')})", NUM), ('Ortalama satış fiyatı (₺)', f"=AVERAGE({MRN('price')})", TL), ('Ortalama satış fiyatı (€)', f"=AVERAGE({MRN('price_eur')})", EUR),
        ('Ortalama ₺/m²', f"=AVERAGE({MRN('ppsm')})", TL), ('Ortalama aylık kira (₺)', f"=AVERAGE({MRN('ra')})", TL), ('Ortalama brüt getiri', f"=AVERAGE({MRN('gy')})", PCT),
        ('Ortalama NET getiri', f"=AVERAGE({MRN('ny')})", PCT), ('En yüksek net getiri', f"=MAX({MRN('ny')})", PCT), ('En yüksek net getirili daire', f"=INDEX({MRN('no')},MATCH(MAX({MRN('ny')}),{MRN('ny')},0))", None),
        ('Ortalama 10Y IRR (€)', f"={q('cf')}C43", PCT), ('Ortalama 10Y IRR (TL)', f"={q('cf')}C44", PCT), ('Yüksek riskli daire', f"=COUNTIF({MRN('risklvl')},\"Yüksek\")", NUM)]
kfill = PatternFill('solid', fgColor='DDEBF7')
for i, (lab, f_, fmt) in enumerate(KPIS):
    row = 5 + (i // 4) * 3; colx = 1 + (i % 4) * 2
    c = D.cell(row, colx, lab); c.font = Font(name=F, size=8, bold=True, color='595959'); c.fill = kfill; D.cell(row, colx + 1).fill = kfill
    c = D.cell(row + 1, colx, f_); c.font = Font(name=F, size=15, bold=True, color=NAVY); c.fill = kfill
    if fmt: c.number_format = fmt
    D.cell(row + 1, colx + 1).fill = kfill; D.merge_cells(start_row=row + 1, start_column=colx, end_row=row + 1, end_column=colx + 1)
r0 = 15
put(D, r0, 1, 'EN İYİ 10 FIRSAT', font=Font(name=F, size=11, bold=True, color=NAVY), border=False)
for i, h in enumerate(['Sıra', 'Daire No', 'Mahalle', 'Oda', 'Fiyat (₺)', 'Kira (₺/ay)', 'Net getiri', 'Risk', 'Puan', '10Y IRR (€)'], 1): hdr(D, r0 + 1, i, h)
for k in range(min(10, N)):
    r = r0 + 2 + k; tr = 5 + k; put(D, r, 1, k + 1, align=CENTER, font=fN)
    for j, (col_, fmt) in enumerate([('B', None), ('D', None), ('E', None), ('F', TL), ('I', TL), ('M', PCT), ('N', NUM), ('O', NUM1), ('P', PCT)]):
        put(D, r, 2 + j, f"={q('top')}{col_}{tr}", fmt=fmt, font=fX)
r1 = r0 + 14
put(D, r1, 1, 'MAHALLE KARŞILAŞTIRMASI', font=Font(name=F, size=11, bold=True, color=NAVY), border=False)
for i, h in enumerate(['Mahalle', 'Daire', 'Ort. ₺/m²', 'Emsal kira ₺/m²', 'Brüt getiri', 'Kira çarpanı', 'Ort. NET getiri', 'Piyasa brüt getirisi', 'Talep'], 1): hdr(D, r1 + 1, i, h)
for k in range(len(MAHS)):
    r = r1 + 2 + k; hr = 6 + k
    for j, (col_, fmt) in enumerate([('A', None), ('B', NUM), ('D', TL), ('E', TL2), ('F', PCT), ('G', NUM1), ('H', PCT), ('M', PCT), ('O', NUM)]):
        put(D, r, 1 + j, f"={q('mah')}{col_}{hr}", fmt=fmt, font=fX)
RE = r1 + 1 + len(MAHS)
r2 = RE + 2
put(D, r2, 1, 'SENARYO ÖZETİ – ortalama net getiri', font=Font(name=F, size=11, bold=True, color=NAVY), border=False)
for j, (lab, nm) in enumerate([('Kötümser', 'SCN_AVG_K'), ('Baz', 'SCN_AVG_B'), ('İyimser', 'SCN_AVG_I')]):
    hdr(D, r2 + 1, 1 + j, lab); put(D, r2 + 2, 1 + j, f'={nm}', fmt=PCT, font=Font(name=F, size=11, bold=True, color=NAVY))
put(D, r2 + 4, 1, 'Not: Kira tahminleri yakın çevredeki kiralık İLAN fiyatlarına dayanır; %5 sözleşme indirimi ve boş kalma varsayımı uygulanmıştır. Net getiri 1. yıl TL nakit akışına göredir; enflasyon ortamında toplam getiri için 5 VE 10 YILLIK NAKİT AKIŞI sayfasındaki € IRR’ye bakın.', font=fSm, border=False)
def bar(title_, col, cats_col, r_from, r_to, anchor, color, horiz=True, fmt=None):
    ch = BarChart(); ch.type = 'bar' if horiz else 'col'; ch.title = title_; ch.roundedCorners = False
    ch.add_data(Reference(D, min_col=col, min_row=r_from - 1, max_row=r_to), titles_from_data=True); ch.set_categories(Reference(D, min_col=cats_col, min_row=r_from, max_row=r_to))
    ch.legend = None; ch.height = 9; ch.width = 16; ch.x_axis.delete = False; ch.y_axis.delete = False; ch.title.overlay = False
    if fmt: ch.y_axis.number_format = fmt
    ch.series[0].graphicalProperties.solidFill = color; D.add_chart(ch, anchor)
bar('Mahallelere göre ortalama NET getiri', 7, 1, r1 + 2, RE, 'L4', '2E7D32', True, '0.0%')
bar('Mahallelere göre ortalama ₺/m²', 3, 1, r1 + 2, RE, 'L23', NAVY, False)
sc = ScatterChart(); sc.title = 'Satış fiyatı (₺) – Net getiri'; sc.style = 13; sc.roundedCorners = False
se = Series(Reference(M, min_col=CI['ny'], min_row=FR, max_row=LR), Reference(M, min_col=CI['price'], min_row=FR, max_row=LR), title='Daireler')
se.marker.symbol = 'circle'; se.marker.size = 6; se.graphicalProperties.line.noFill = True
se.marker.graphicalProperties.solidFill = NAVY2; se.marker.graphicalProperties.line.solidFill = NAVY2
sc.series.append(se); sc.x_axis.title = 'Satış fiyatı (₺)'; sc.y_axis.title = 'Net getiri'; sc.y_axis.number_format = '0.0%'; sc.x_axis.number_format = '#,##0'
sc.legend = None; sc.height = 9; sc.width = 16; sc.x_axis.delete = False; sc.y_axis.delete = False; sc.title.overlay = False; sc.x_axis.title.overlay = False; sc.y_axis.title.overlay = False
D.add_chart(sc, 'L42')
bar('En iyi 10 – genel yatırım puanı', 9, 2, r0 + 2, r0 + 1 + min(10, N), 'L61', '548235', True)

# ================= ANALİZ SORULARI =================
A = ws['qa']
title(A, 'ANALİZ SORULARI – Talimatnamedeki 10 soru (Kadıköy, çok portallı)', 'Tablolar formülle güncellenir. Yazılı yorumlar BAZ senaryo ve 17.11.2026 verisiyle hazırlanmıştır.', 8)
A.column_dimensions['A'].width = 40
for cc in 'BCDEFGH': A.column_dimensions[cc].width = 15
r = 4
def qhead(t):
    global r
    c = A.cell(r, 1, t); c.font = Font(name=F, size=10, bold=True, color='FFFFFF')
    for cc in range(1, 9): A.cell(r, cc).fill = fillH
    r += 1
def qtext(t):
    global r
    if not t: return
    c = A.cell(r, 1, t); c.font = fN; c.alignment = WRAP; A.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    A.row_dimensions[r].height = max(15, 13 * (len(t) // 135 + 1)); r += 1
qhead('1. En yüksek net kira getirisi sağlayan 20 daire hangileri?'); qtext(NOTES.get('q1', 'EN İYİ 20 YATIRIM sayfasının alt tablosunda formülle listelenir.'))
qhead('2. Aynı kira gelirini daha düşük satın alma maliyetiyle sağlayan daireler var mı?'); qtext(NOTES.get('q2'))
qhead('3. Hangi mahallelerde kira çarpanı daha avantajlı?')
for i, h in enumerate(['Mahalle', 'Kira çarpanı (örneklem)', 'Brüt getiri', 'Ort. net getiri', 'Emlakjet amortisman'], 1): hdr(A, r, i, h)
r += 1
for k in range(len(MAHS)):
    hr = 6 + k
    for j, (col_, fmt) in enumerate([('A', None), ('G', NUM1), ('F', PCT), ('H', PCT), ('N', NUM)]): put(A, r, 1 + j, f"={q('mah')}{col_}{hr}", fmt=fmt, font=fX)
    r += 1
qtext(NOTES.get('q3'))
qhead('4. Bağdat Caddesi’ne yakın olmanın fiyat primini kira geliri karşılıyor mu?')
for i, h in enumerate(['Grup', 'Daire', 'Ort. ₺/m²', 'Ort. emsal kira ₺/m²', 'Ort. brüt getiri', 'Ort. net getiri'], 1): hdr(A, r, i, h)
r += 1
for lab, cond in [('Bağdat Cd.’ye ≤300 m', '"<=300"'), ('Bağdat Cd.’ye >300 m', '">300"')]:
    put(A, r, 1, lab, font=fN); put(A, r, 2, f"=COUNTIF({MRN('dbagdat')},{cond})", fmt=NUM)
    for j, (k_, fmt) in enumerate([('ppsm', TL), ('compavg', TL2), ('gy', PCT), ('ny', PCT)]): put(A, r, 3 + j, f"=IFERROR(AVERAGEIF({MRN('dbagdat')},{cond},{MRN(k_)}),\"-\")", fmt=fmt)
    r += 1
qtext(NOTES.get('q4'))
qhead('5. Hangi daireler gereğinden pahalı?'); qtext(NOTES.get('q5', 'Ana listede “Fiyat / emsal oranı” > 1,15 olan hücreler kırmızıdır.'))
qhead('6. Hangi dairelerde pazarlık potansiyeli daha yüksek?'); qtext(NOTES.get('q6'))
qhead('7. Hangi binalarda teknik ve hukuki riskler daha yüksek?'); qtext(NOTES.get('q7'))
qhead('8. Düzenli kira geliri için hangi 10 daire tercih edilmeli?'); qtext(NOTES.get('q8'))
qhead('9. Aynı sermayeyle bir büyük daire mi, birkaç küçük daire mi? (1+1 / 2+1)')
for i, h in enumerate(['Grup', 'Daire', 'Ort. fiyat (₺)', 'Ort. toplam yatırım (₺)', 'Ort. aylık kira (₺)', 'Ort. net getiri', 'Ort. risk'], 1): hdr(A, r, i, h)
r += 1
for lab, cnd in [('1+1', f'{MRN("room")},"1+1"'), ('2+1', f'{MRN("room")},"2+1"'), ('Brüt ≤70 m²', f'{MRN("sqm")},"<=70"'), ('Brüt 71–100 m²', f'{MRN("sqm")},">70",{MRN("sqm")},"<=100"'), ('Brüt >100 m²', f'{MRN("sqm")},">100"')]:
    put(A, r, 1, lab, font=fN); put(A, r, 2, f'=COUNTIFS({cnd})', fmt=NUM)
    for j, (k_, fmt) in enumerate([('price', TL), ('tinv', TL), ('ra', TL), ('ny', PCT), ('risk', NUM)]): put(A, r, 3 + j, f'=IFERROR(AVERAGEIFS({MRN(k_)},{cnd}),"-")', fmt=fmt)
    r += 1
qtext(NOTES.get('q9'))
qhead('10. Beş ve on yıllık elde tutma senaryolarında toplam getiri nasıl değişiyor?')
for i, h in enumerate(['Gösterge', 'Değer'], 1): hdr(A, r, i, h)
r += 1
for lab, f_, fmt in [('Ortalama 5 yıllık IRR (€)', f"={q('cf')}B43", PCT), ('Ortalama 10 yıllık IRR (€)', f"={q('cf')}C43", PCT), ('Ortalama 5 yıllık IRR (TL)', f"={q('cf')}B44", PCT), ('Ortalama 10 yıllık IRR (TL)', f"={q('cf')}C44", PCT)]:
    put(A, r, 1, lab, font=fN); put(A, r, 2, f_, fmt=fmt); r += 1
qtext(NOTES.get('q10'))

for sh in wb.worksheets:
    sh.sheet_properties.pageSetUpPr.fitToPage = True; sh.page_setup.orientation = 'landscape'; sh.page_setup.fitToWidth = 1; sh.page_setup.fitToHeight = 0
for k, color in [('panel', NAVY), ('main', NAVY2), ('top', '548235'), ('par', 'BF8F00'), ('scen', 'BF8F00'), ('risk', 'C00000')]: ws[k].sheet_properties.tabColor = color
wb.save(OUT); print('saved', OUT, N)
