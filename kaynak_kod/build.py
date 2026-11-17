import json, sys, datetime, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as GL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, ScatterChart, Reference, Series
from openpyxl.comments import Comment

DATA = json.load(open(sys.argv[1]))
OUT = sys.argv[2]
NOTES = json.load(open(sys.argv[3])) if len(sys.argv) > 3 else None
R = DATA['recs']; COMPS = DATA['comps']; MARKET = DATA['market']; EXTRAS = DATA['extras']
CHECK = datetime.date(2026, 11, 17)

NAVY = '1F3864'; NAVY2 = '2E5597'; GREY = 'F2F2F2'; GREY2 = 'D9D9D9'; GREEN = 'E2EFDA'; GREEN_D = '375623'; YEL = 'FFF2CC'
F = 'Arial'
fT = Font(name=F, size=16, bold=True, color='FFFFFF')
fH = Font(name=F, size=9, bold=True, color='FFFFFF')
fN = Font(name=F, size=9, color='000000')
fIn = Font(name=F, size=9, color='0000FF')
fLink = Font(name=F, size=9, color='0563C1', underline='single')
fX = Font(name=F, size=9, color='008000')
fB = Font(name=F, size=10, bold=True, color=NAVY)
fSm = Font(name=F, size=8, italic=True, color='595959')
fillT = PatternFill('solid', fgColor=NAVY); fillH = PatternFill('solid', fgColor=NAVY2)
fillG = PatternFill('solid', fgColor=GREY); fillIn = PatternFill('solid', fgColor=YEL); fillGr = PatternFill('solid', fgColor=GREEN)
thin = Side(style='thin', color='BFBFBF'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
EUR = '#,##0 "€";(#,##0 "€");"-"'
EUR2 = '#,##0.00 "€";(#,##0.00 "€");"-"'
PCT = '0.00%;(0.00%);"-"'
PCT1 = '0.0%;(0.0%);"-"'
NUM = '#,##0;(#,##0);"-"'
NUM1 = '0.0;(0.0);"-"'
DATE = 'dd.mm.yyyy'
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

wb = Workbook()
SN = dict(panel='YÖNETİCİ PANELİ', main='100 DAİRE ANA LİSTESİ', top='EN İYİ 20 YATIRIM', comp='KİRA EMSALLERİ',
          mah='MAHALLE ANALİZİ', scen='SENARYO ANALİZİ', cf='5 VE 10 YILLIK NAKİT AKIŞI', risk='RİSK VE EKSİK VERİLER',
          src='KAYNAKLAR', par='PARAMETRELER', qa='ANALİZ SORULARI', tax='VERGİ-HUKUK NOTLARI', ext='YEDEK İLANLAR')
ws = {}
first = True
for k in ['panel', 'main', 'top', 'comp', 'mah', 'scen', 'cf', 'risk', 'src', 'par', 'qa', 'tax', 'ext']:
    if first:
        ws[k] = wb.active; ws[k].title = SN[k]; first = False
    else:
        ws[k] = wb.create_sheet(SN[k])
    ws[k].sheet_view.showGridLines = False
def q(k):  # quoted sheet ref
    return "'" + SN[k] + "'!"

def title(sh, text, sub, width_cols=12):
    sh.row_dimensions[1].height = 30
    for c in range(1, width_cols + 1):
        sh.cell(1, c).fill = fillT; sh.cell(2, c).fill = fillT
    sh['A1'] = text; sh['A1'].font = fT; sh['A1'].alignment = Alignment(vertical='center')
    sh['A2'] = sub; sh['A2'].font = Font(name=F, size=9, italic=True, color='FFFFFF')

def hdr(sh, row, col, text, width=None, fill=fillH):
    c = sh.cell(row, col, text); c.font = fH; c.fill = fill; c.alignment = CENTER; c.border = BOX
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
    c = sh.cell(row, col, text); c.hyperlink = url; c.font = fLink; c.border = BOX
    return c

def dname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# =====================================================================================
# PARAMETRELER
# =====================================================================================
P = ws['par']
title(P, 'PARAMETRELER – Tüm hesaplama varsayımları',
      'Mavi yazı + sarı zemin = değiştirilebilir girdi. Siyah = formül. Tüm sayfalar bu hücrelere bağlıdır. Kontrol tarihi: 17.11.2026', 6)
for i, (h, w) in enumerate([('Kategori', 22), ('Parametre', 52), ('Değer', 14), ('Birim', 12), ('Açıklama / Kaynak', 95)], 1):
    hdr(P, 4, i, h, w)
PARAMS = [
 ('Genel', 'P_SENARYO', 'Aktif senaryo (1=Kötümser, 2=Baz, 3=İyimser)', 2, 'seçim', 'Ana liste, puanlama ve sıralama bu senaryoya göre hesaplanır. SENARYO ANALİZİ sayfasında üç senaryo yan yana gösterilir.', NUM),
 ('Genel', 'P_KONTROL', 'Veri kontrol tarihi', CHECK, 'tarih', 'İlanların spitimou.net API üzerinden çekildiği / kontrol edildiği tarih.', DATE),
 ('Satın alma', 'P_MUKIM', 'Alıcı profili (1=AB/AEA vergi mukimi veya 01.07.2027 öncesi sözleşme; 2=AB/AEA dışı vergi mukimi, 01.07.2027 ve sonrası)', 1, 'seçim', 'Yeni yasa (raporlara göre Kanun 5355/2026, FEK A 188, 13.11.2026): AB/AEA dışı VERGİ MUKİMLERİ için konut devir vergisi %3 → %15 (01.07.2027 sonrası sözleşmeler). Uyruk değil vergi mukimliği esas; avukatla teyit edin.', NUM),
 ('Satın alma', 'P_DEVIR1', 'Devir vergisi – standart (%3 + %3 belediye payı)', 0.0309, '%', 'Kaynak: globalpropertyguide.com/europe/greece/taxes-and-costs; vergi, satış bedeli ile objektif (vergi) değerden YÜKSEK olana uygulanır.', PCT),
 ('Satın alma', 'P_DEVIR2', 'Devir vergisi – AB/AEA dışı vergi mukimi (01.07.2027+)', 0.1545, '%', '%15 + belediye payı ≈ %15,45 (basın: greekcitytimes.com 17.11.2026, imidaily.com). Resmî FEK metniyle teyit edilmelidir.', PCT),
 ('Satın alma', 'P_DEVIR', 'Uygulanan devir vergisi oranı', '=IF(P_MUKIM=2,P_DEVIR2,P_DEVIR1)', '%', 'Formül – alıcı profiline göre seçilir.', PCT),
 ('Satın alma', 'P_NOTER', 'Noter ücreti (KDV dahil)', 0.010, '% fiyat', 'Piyasa aralığı %0,7–1,0 + KDV (globalpropertyguide, elxis.com). Varsayım.', PCT),
 ('Satın alma', 'P_TAPU', 'Tapu/Kadastro (Ktimatologio) tescil ve harçlar', 0.006, '% fiyat', 'Kaynaklarda ≈%0,5–0,6. Varsayım.', PCT),
 ('Satın alma', 'P_AVUKAT', 'Avukat ücreti (KDV dahil)', 0.0124, '% fiyat', '%1 + %24 KDV. Pazarlıkla düşebilir. Varsayım.', PCT),
 ('Satın alma', 'P_KOMISYON', 'Alıcı emlakçı komisyonu (KDV dahil)', 0.0248, '% fiyat', 'Piyasa uygulaması %2 + %24 KDV. Varsayım.', PCT),
 ('Satın alma', 'P_SABIT', 'Sabit masraflar (mühendis kontrolü, tercüme, vekâlet, AFM, banka)', 1500, '€', 'Varsayım; mühendis (μηχανικός) teknik/imar kontrolü dahil.', EUR),
 ('Satın alma', 'P_HAZIRLIK', 'Kiraya hazırlık / küçük onarım rezervi (her daire)', 1500, '€', 'Boya, küçük tamirat, kilit, temizlik. Varsayım.', EUR),
 ('Satın alma', 'P_ESYA_UYGULA', 'Eşyası belirtilmemiş daireler eşyalansın mı? (1=Evet, 0=Hayır)', 1, 'seçim', 'Kira emsalleri ağırlıkla eşyalı/yenilenmiş dairelerden oluşur; tutarlılık için varsayılan Evet.', NUM),
 ('Satın alma', 'P_ESYA', 'Eşya paketi maliyeti (2+1)', 4500, '€', 'Orta segment mobilya + beyaz eşya. Varsayım.', EUR),
 ('Satın alma', 'P_TADILAT_DEVAM', 'İlanda “tadilat devam ediyor” ise ek rezerv', 3000, '€', 'Varsayım.', EUR),
 ('Kira', 'P_ILAN_INDIRIM', 'İlan kirası → gerçekleşen sözleşme kirası indirimi', 0.05, '%', 'Emsaller İSTENEN (ilan) kiralarıdır; gerçekleşmiş sözleşme kirası değildir. Varsayım.', PCT1),
 ('Kira', 'P_KIRACI_AY', 'Kiracı bulma ücreti (ay kira, KDV dahil)', 1.24, 'ay', 'Uygulamada 1 aylık kira + %24 KDV. Varsayım.', '0.00'),
 ('Kira', 'P_SOZLESME', 'Ortalama kiracı kalış süresi', 3, 'yıl', 'Yunanistan’da konut kira sözleşmesi asgari 3 yıl (Medeni Kanun md. 612A). Varsayım.', NUM),
 ('Gider', 'P_BAKIM', 'Bakım-onarım (yıllık brüt kiranın %’si)', 0.05, '%', 'Yenilenmiş daire için. Varsayım.', PCT1),
 ('Gider', 'P_YONETIM', 'Kira yönetim ücreti (tahsil edilen kiranın %’si, KDV dahil)', 0.0992, '%', '%8 + KDV. Kendiniz yönetirseniz 0 yazın. Varsayım.', PCT),
 ('Gider', 'P_SIGORTA', 'Konut sigortası', 2.5, '€/m²/yıl', 'Bina + mal sahibi sorumluluk. Varsayım.', EUR2),
 ('Gider', 'P_ENFIA', 'ENFIA emlak vergisi (tahmini)', 4.0, '€/m²/yıl', 'ENFIA asıl vergi bölge değerine göre ≈€2,9–4,5/m² (bölge €1.500–2.500 bandı; hellascalculator/elxis). Daire bazında AADE’den teyit gerekir.', EUR2),
 ('Gider', 'P_AIDAT_PAY', 'Mal sahibinin üstlendiği aidat payı (dolu aylarda)', 0.20, '%', 'Büyük onarım/rezerv kalemleri mal sahibine aittir. Boş aylarda aidatın tamamı mal sahibine aittir.', PCT1),
 ('Gider', 'P_AIDAT_VARS', 'Aidat bilinmiyorsa kullanılacak tutar', 40, '€/ay', 'Asansörlü merkez binası için varsayım.', EUR),
 ('Vergi', 'P_VERGI_YONTEM', 'Kira vergisi yöntemi (1=Kademeli – bu daire tek kira geliri; 2=Sabit marjinal oran)', 1, 'seçim', 'Birden çok daire alınırsa toplam kira geliri üst dilimlere çıkar; o durumda 2’yi seçip marjinal oranı girin.', NUM),
 ('Vergi', 'P_GOTURU', 'Götürü gider indirimi', 0.05, '%', 'Brüt kiradan %5 götürü indirim (gerçek giderler indirilemez). Kaynak: Yunan Gelir Vergisi Kanunu md.40; basın 2026.', PCT1),
 ('Vergi', 'P_V1_LIM', 'Dilim 1 üst sınırı', 12000, '€', 'Kanun 5246/2025 – 2026 gelirleri için kira geliri tarifesi.', EUR),
 ('Vergi', 'P_V1', 'Dilim 1 oranı', 0.15, '%', '', PCT1),
 ('Vergi', 'P_V2_LIM', 'Dilim 2 üst sınırı', 24000, '€', 'Yeni ara dilim (2026).', EUR),
 ('Vergi', 'P_V2', 'Dilim 2 oranı', 0.25, '%', '2025’e kadar %35 idi.', PCT1),
 ('Vergi', 'P_V3_LIM', 'Dilim 3 üst sınırı', 36000, '€', 'Kaynaklarda 35.000/36.000 farkı var – AADE’den teyit edin.', EUR),
 ('Vergi', 'P_V3', 'Dilim 3 oranı', 0.35, '%', '', PCT1),
 ('Vergi', 'P_V4', 'Dilim 4 oranı (üstü)', 0.45, '%', '', PCT1),
 ('Vergi', 'P_VERGI_SABIT', 'Sabit marjinal oran (yöntem 2)', 0.25, '%', 'Yöntem 2 seçilirse kullanılır.', PCT1),
 ('Puanlama', 'W_GETIRI', 'Ağırlık – Net kira getirisi', 0.40, '%', 'Talimatnamedeki ağırlıklar.', PCT1),
 ('Puanlama', 'W_KONUM', 'Ağırlık – Konum ve kiracı talebi', 0.20, '%', '', PCT1),
 ('Puanlama', 'W_FIYAT', 'Ağırlık – Fiyatın emsallere uygunluğu', 0.15, '%', '', PCT1),
 ('Puanlama', 'W_KALITE', 'Ağırlık – Bina ve daire kalitesi', 0.10, '%', '', PCT1),
 ('Puanlama', 'W_RISK', 'Ağırlık – Boş kalma / kiracı riski', 0.10, '%', '', PCT1),
 ('Puanlama', 'W_LIKIT', 'Ağırlık – Gelecekte satış kolaylığı', 0.05, '%', '', PCT1),
 ('Puanlama', 'P_NY_MIN', 'Net getiri puanı: 0 puan eşiği', 0.015, '%', 'Bu net getiri ve altı 0 puan.', PCT),
 ('Puanlama', 'P_NY_MAX', 'Net getiri puanı: 100 puan eşiği', 0.050, '%', 'Bu net getiri ve üstü 100 puan.', PCT),
 ('Puanlama', 'P_VERI_CEZA', 'Yetersiz veri çarpanı (emsal <3 veya 2+1 teyitsiz)', 0.85, 'x', 'Yetersiz veriye sahip daireler sıralamada yukarı çıkamasın diye toplam puan bu katsayıyla çarpılır.', '0.00'),
 ('Puanlama', 'P_ILAN_GUN', 'İlan “eski” sayılma eşiği (gün)', 60, 'gün', 'Son güncelleme bu süreden eskiyse risk işareti.', NUM),
 ('Pazarlık', 'P_PAZARLIK', 'Standart pazarlık payı', 0.05, '%', 'Selanik merkezde ilan–satış farkı için varsayım.', PCT1),
 ('Pazarlık', 'P_PAZARLIK_MAX', 'Azami önerilen pazarlık', 0.12, '%', '', PCT1),
 ('Pazarlık', 'P_HEDEF', 'Hedef net getiri (azami fiyat hesabı için)', 0.04, '%', '“Hedef net getiri için azami alım fiyatı” sütununu belirler.', PCT1),
]
PROW = {}
r0 = 5
for i, (cat, nm, lab, val, unit, note, fmt) in enumerate(PARAMS):
    rr = r0 + i
    put(P, rr, 1, cat, font=fN); put(P, rr, 2, lab, font=fN, align=WRAP)
    c = put(P, rr, 3, val, fmt=fmt)
    if not (isinstance(val, str) and val.startswith('=')):
        c.fill = fillIn; c.font = fIn
    put(P, rr, 4, unit, font=fN); put(P, rr, 5, note, font=fSm, align=WRAP)
    dname(nm, f"{q('par')}$C${rr}")
    PROW[nm] = rr
dv = DataValidation(type='list', formula1='"1,2,3"', allow_blank=False); P.add_data_validation(dv); dv.add(f'C{PROW["P_SENARYO"]}')
dv2 = DataValidation(type='list', formula1='"1,2"'); P.add_data_validation(dv2); dv2.add(f'C{PROW["P_MUKIM"]}'); dv2.add(f'C{PROW["P_VERGI_YONTEM"]}')
dv3 = DataValidation(type='list', formula1='"0,1"'); P.add_data_validation(dv3); dv3.add(f'C{PROW["P_ESYA_UYGULA"]}')
rr = r0 + len(PARAMS) + 1
put(P, rr, 2, 'Ağırlık toplamı (100% olmalı)', font=fB); put(P, rr, 3, '=W_GETIRI+W_KONUM+W_FIYAT+W_KALITE+W_RISK+W_LIKIT', fmt=PCT1)
put(P, rr + 1, 2, 'Satın alma masrafları toplamı (% fiyat, sabitler hariç)', font=fB); put(P, rr + 1, 3, '=P_DEVIR+P_NOTER+P_TAPU+P_AVUKAT+P_KOMISYON', fmt=PCT)
dname('P_MASRAF_PCT', f"{q('par')}$C${rr + 1}")
P.freeze_panes = 'A5'

# =====================================================================================
# SENARYO PARAMETRELERİ (SENARYO sayfası üst kısım)
# =====================================================================================
S = ws['scen']
title(S, 'SENARYO ANALİZİ – Kötümser / Baz / İyimser',
      'Sarı hücreler değiştirilebilir. Aktif senaryo PARAMETRELER!P_SENARYO ile seçilir; tüm sayfalar otomatik güncellenir.', 30)
hdr(S, 4, 1, 'Senaryo varsayımı', 46); hdr(S, 4, 2, 'Birim', 10)
hdr(S, 4, 3, 'Kötümser', 12); hdr(S, 4, 4, 'Baz', 12); hdr(S, 4, 5, 'İyimser', 12); hdr(S, 4, 6, 'AKTİF', 12); hdr(S, 4, 7, 'Açıklama', 70)
SCN = [
 ('KIRA', 'Kira seviyesi (emsal ortalamasına çarpan)', 'x', (0.92, 1.00, 1.06), '0.00', 'Kötümser: talep zayıflaması; İyimser: hızlı ve uygun fiyatla kiralama.'),
 ('BOS', 'Yıllık boş kalma (ay)', 'ay', (2.0, 1.0, 0.5), '0.0', 'Kiracı değişimi + tahsilat kaybı. Kiranın 12 ay kesin tahsil edildiği varsayılmaz.'),
 ('IND', 'Alım fiyatında pazarlık indirimi', '%', (0.0, 0.03, 0.06), PCT1, 'İlan fiyatına göre gerçekleşen indirim.'),
 ('MAS', 'İşletme gideri çarpanı (bakım + sigorta)', 'x', (1.15, 1.00, 0.90), '0.00', ''),
 ('KARTIS', 'Yıllık kira artışı (5/10 yıl)', '%', (0.010, 0.025, 0.035), PCT1, 'Yunanistan’da kira artışı sözleşmede serbestçe belirlenir; TÜFE civarı varsayım.'),
 ('GARTIS', 'Yıllık gider artışı (5/10 yıl)', '%', (0.035, 0.025, 0.020), PCT1, ''),
 ('DEGER', 'Yıllık mülk değer değişimi (5/10 yıl)', '%', (0.0, 0.025, 0.040), PCT1, 'Nominal. Geçmiş yüksek artışların sürmesi varsayılmamıştır.'),
 ('CIKIS', 'Çıkışta satış maliyeti (komisyon + avukat, KDV dahil)', '%', (0.035, 0.030, 0.025), PCT1, ''),
 ('DKV', 'Değer artış kazancı vergisi (satışta)', '%', (0.15, 0.0, 0.0), PCT1, 'Yunanistan’da %15 değer artış kazancı vergisi 31.12.2026’ya kadar askıda; sonrası belirsiz → kötümserde uygulanır.'),
]
SC = {}
for i, (k, lab, unit, vals, fmt, note) in enumerate(SCN):
    rr = 5 + i
    put(S, rr, 1, lab, font=fN); put(S, rr, 2, unit, font=fN)
    for j, v in enumerate(vals):
        c = put(S, rr, 3 + j, v, fmt=fmt); c.fill = fillIn
    put(S, rr, 6, f'=INDEX(C{rr}:E{rr},1,P_SENARYO)', fmt=fmt, fill=fillGr)
    put(S, rr, 7, note, font=fSm, align=WRAP)
    for j, s in enumerate(['K', 'B', 'I', 'A']):
        dname(f'S{s}_{k}', f"{q('scen')}${'CDEF'[j]}${rr}")
    SC[k] = rr
put(S, 15, 1, 'Aktif senaryo adı', font=fB); put(S, 15, 6, '=CHOOSE(P_SENARYO,"Kötümser","Baz","İyimser")', font=Font(name=F, size=10, bold=True, color=GREEN_D))

# =====================================================================================
# MAIN LIST
# =====================================================================================
M = ws['main']
N = len(R); FR = 5; LR = FR + N - 1
COLS = [
 # key, header, width, group
 ('no', 'Daire No', 8, 'A. GENEL BİLGİLER'), ('title', 'İlan başlığı (özet, TR)', 34, ''), ('url', 'İlan URL (tıklanabilir)', 16, ''),
 ('plat', 'İlan platformu', 18, ''), ('adv', 'Emlak şirketi', 20, ''), ('contact', 'Emlakçı iletişim / ilan ref.', 26, ''),
 ('upd', 'Son güncelleme (kaynak)', 11, ''), ('first', 'İlk yayın (portal)', 11, ''), ('ver', 'Doğrulama (portalda son görülme)', 11, ''),
 ('status', 'İlan durumu', 18, ''), ('photo', 'Fotoğraf bağlantısı', 12, ''),
 ('mah', 'Mahalle', 20, 'B. KONUM'), ('grp', 'Öncelikli bölge grubu', 22, ''), ('zone', 'Bölge sınıfı', 13, ''), ('street', 'Cadde/sokak (ilan metninde geçen)', 22, ''),
 ('addr', 'Tam adres', 12, ''), ('maps', 'Google Maps', 11, ''), ('dcenter', 'Merkeze mesafe (m, Aristotelous)', 10, ''),
 ('road', 'En yakın ana cadde', 16, ''), ('droad', 'Ana caddeye mesafe (m, yaklaşık)', 10, ''), ('metro', 'En yakın metro', 14, ''),
 ('dmetro', 'Metroya mesafe (m)', 9, ''), ('duni', 'Üniversiteye mesafe (m)', 9, ''), ('dsea', 'Denize mesafe (m, portal)', 9, ''),
 ('adv_txt', 'Kiralama avantajları', 30, ''), ('dis_txt', 'Dezavantajlar (gürültü, erişim vb.)', 30, ''),
 ('price', 'Satış fiyatı (€)', 11, 'C. DAİRE ÖZELLİKLERİ'), ('net', 'Net m² (ilan metninden)', 8, ''), ('sqm', 'İlan m² (brüt/beyan)', 8, ''),
 ('ppsm', 'Satış fiyatı / m² (€)', 9, ''), ('beds', 'Yatak odası', 7, ''), ('two', '2+1 teyidi', 14, ''), ('living', 'Salon', 6, ''), ('baths', 'Banyo', 6, ''),
 ('yb', 'Bina yapım yılı', 8, ''), ('ry', 'Yenileme yılı', 8, ''), ('scope', 'Tadilat kapsamı', 20, ''), ('under', 'Tadilat sürüyor mu?', 9, ''),
 ('floor', 'Kat', 7, ''), ('floorn', 'Kat (sayı)', 6, ''), ('tf', 'Binadaki toplam kat', 8, ''), ('elev', 'Asansör', 9, ''), ('balc', 'Balkon', 9, ''),
 ('heat', 'Isıtma sistemi', 14, ''), ('ac', 'Klima', 9, ''), ('energy', 'Enerji sınıfı', 8, ''), ('furn', 'Mobilya durumu', 11, ''), ('park', 'Otopark', 10, ''),
 ('fees', 'Aidat (€/ay, ilan)', 8, ''), ('bldg', 'Binanın genel durumu', 18, ''), ('tenant', 'Kiracı durumu', 16, ''), ('rentnow', 'Mevcut sözleşme kirası (€/ay) – varsa 1. yıl kira olarak kullanılır', 12, ''),
 ('legal', 'Hukuki / teknik risk notları', 28, ''), ('dups', 'Aynı daire – diğer ilanlar ve fiyatlar', 28, ''),
 ('ncomp', 'Kira emsal sayısı', 7, 'D. MALİYETLER VE GELİRLER'), ('compavg', 'Emsal ort. ilan kirası (€/m²)', 9, ''), ('complink', 'Kira emsalleri', 11, ''),
 ('basem2', 'Baz kira €/m² (ilan indirimi sonrası)', 9, ''),
 ('rk', 'Aylık kira – Kötümser (€)', 9, ''), ('rb', 'Aylık kira – Baz (€)', 9, ''), ('ri', 'Aylık kira – İyimser (€)', 9, ''), ('ra', 'Aylık kira – AKTİF senaryo (€)', 9, ''),
 ('gross', 'Yıllık brüt kira (€)', 10, ''), ('vac', 'Boş kalma (ay/yıl)', 7, ''), ('coll', 'Tahsil edilen yıllık kira (€)', 10, ''),
 ('maint', 'Bakım gideri (€/yıl)', 9, ''), ('mgmt', 'Yönetim + kiracı bulma (€/yıl)', 9, ''), ('enfia', 'Emlak vergisi ENFIA (€/yıl)', 9, ''),
 ('ins', 'Sigorta (€/yıl)', 8, ''), ('feeown', 'Mal sahibi aidat payı (€/yıl)', 9, ''), ('opex', 'Toplam işletme gideri (€/yıl)', 10, ''),
 ('buy', 'Alım bedeli (senaryo pazarlığı sonrası, €)', 11, ''), ('ttax', 'Satın alma (devir) vergisi (€)', 10, ''), ('notary', 'Noter + tapu + avukat + sabit (€)', 10, ''),
 ('comm', 'Emlakçı komisyonu (€)', 9, ''), ('furncost', 'Eşya + tadilat + hazırlık (€)', 10, ''), ('tinv', 'TOPLAM YATIRIM (€)', 11, ''),
 ('noi', 'Vergi öncesi net işletme geliri (€)', 10, ''), ('rtax', 'Yıllık kira geliri vergisi (€)', 9, ''), ('ncf', 'Vergi sonrası net kira geliri (€)', 10, ''),
 ('gy', 'Brüt kira getirisi (ilan fiyatına)', 9, ''), ('ny', 'NET KİRA GETİRİSİ (toplam yatırıma)', 10, ''), ('payback', 'Geri dönüş süresi (yıl)', 8, ''),
 ('mult', 'Kira çarpanı (fiyat/yıllık brüt kira)', 8, ''),
 ('exitppsm', 'Emsal satış €/m² (aynı sokak/çevre medyanı)', 10, ''), ('exitn', 'Emsal satış n', 6, ''), ('pratio', 'Fiyat / emsal oranı', 8, ''),
 ('nego', 'Pazarlık önerilen alım fiyatı (€)', 11, ''), ('maxp', 'Hedef net getiri için azami alım fiyatı (€)', 11, ''),
 ('risk', 'Risk puanı (0–100, yüksek=riskli)', 8, ''), ('risklvl', 'Yatırım riski', 9, ''), ('mainrisk', 'En büyük risk(ler) – ağırlık sırasıyla', 30, ''),
 ('s1', 'Puan: Net getiri', 7, 'E. YATIRIM PUANI'), ('s2', 'Puan: Konum & talep', 7, ''), ('s3', 'Puan: Fiyat uygunluğu', 7, ''), ('s4', 'Puan: Kalite', 7, ''),
 ('s5', 'Puan: Boş kalma/kiracı riski', 7, ''), ('s6', 'Puan: Satış kolaylığı', 7, ''), ('dataok', 'Veri yeterliliği', 9, ''),
 ('score', 'GENEL YATIRIM PUANI (100)', 9, ''), ('helper', 'Sıralama yardımcı', 6, ''), ('rank', 'SIRA', 6, ''), ('nyrank', 'Net getiri sırası', 6, ''),
]
C = {}
for i, (k, h, w, g) in enumerate(COLS, 1):
    C[k] = GL(i)
title(M, '100 DAİRE ANA LİSTESİ – Selanik merkez, yenilenmiş 2+1 satılık daireler',
      'Mavi = ilandan alınan veri (girdi) · Siyah = formül · Mesafeler kuş uçuşu; ana cadde mesafesi yaklaşık güzergâh koordinatlarından · Kira emsalleri ilan (istenen) kiralarıdır', len(COLS))
grp_fill = PatternFill('solid', fgColor='8EA9DB')
for i, (k, h, w, g) in enumerate(COLS, 1):
    if g:
        c = M.cell(3, i, g); c.font = Font(name=F, size=9, bold=True, color=NAVY); c.fill = grp_fill
    else:
        M.cell(3, i).fill = grp_fill
    hdr(M, 4, i, h, w)
M.row_dimensions[4].height = 62

def ref(k, r): return f'{C[k]}{r}'
FLOOR_TR = {'Ground': 'Zemin', 'Mezzanine': 'Asma kat', 'Basement': 'Bodrum', 'Semi-basement': 'Yarı bodrum'}
def floor_tr(f):
    if f in FLOOR_TR: return FLOOR_TR[f]
    m = re.match(r'(\d+)', str(f or '')); return f'{m.group(1)}. kat' if m else (f or '')
HEAT_TR = {'autonomous/gas': 'Bağımsız – doğalgaz', 'autonomous': 'Bağımsız (yakıt belirtilmemiş)', 'autonomous/electricity': 'Bağımsız – elektrik',
           'autonomous/heat-pump': 'Bağımsız – ısı pompası', 'central/oil': 'Merkezi – mazot', 'central/gas': 'Merkezi – doğalgaz', 'central': 'Merkezi', 'none': 'Yok', '': 'Belirtilmemiş', 'autonomous/oil': 'Bağımsız – mazot'}
MAH_TXT = {
 'Tsimiski / Aristotelous / Tarihi Merkez': ('Ticari kalp; Venizelou/Agias Sofias metro; üniversite öğrencisi, genç profesyonel ve kamu çalışanı talebi çok güçlü; kısa kiralama talebi dolaylı olarak uzun dönem arzı daraltıyor.',
     'Trafik ve gece gürültüsü; otopark yok; binaların çoğu 1960–70 yapımı; m² fiyatı yüksek → getiri baskılanıyor.'),
 'Agia Sofia': ('Agias Sofias metro; Tsimiski–Egnatia arası; çok güçlü kiracı talebi ve yüksek ilan kiraları.', 'Pahalı alım; meydan çevresinde kalabalık ve gürültü; otopark yok.'),
 'Beyaz Kule / XANTH / Fuar': ('Deniz ve sahil, AUTH ve Makedonya Üniv. yürüme mesafesi, prestijli adres; yüksek likidite.', 'En yüksek m² fiyatları → düşük getiri; fuar ve etkinlik dönemlerinde trafik.'),
 'Kamara - Rotonda': ('AUTH kampüsüne yürüme; Sintrivani/Agias Sofias metro; öğrenci ve akademisyen talebi en güçlü bölge.', 'Gece hayatı gürültüsü; eski apartmanlar; öğrenci kiracı sirkülasyonu (yaz aylarında boşalma riski).'),
 'Agios Dimitrios': ('Agiou Dimitriou ana arteri; hastaneler, adliye ve üniversiteye yakın; dengeli fiyat.', 'Ana arterde trafik gürültüsü; metroya 300–700 m; yokuşlu sokaklar.'),
 'Vardaris': ('En düşük alım fiyatları; Dimokratias metro ve OSE garına yakın; adliye çevresi çalışan talebi; dönüşüm potansiyeli.', 'Sosyal profil karışık; yoğun trafik ve gürültü; kiracı seçimi önemli.'),
 'Dioikitirio': ('Kamu binaları (Makedonya-Trakya Bakanlığı) çevresi; Ano Poli eteği; merkeze yürüme.', 'Yokuş; otopark sıkıntısı; metroya 400–800 m.'),
 'Antigonidon': ('Egnatia üzerinde/yakını; Venizelou–Dimokratias metro; makul fiyat, güçlü merkez talebi.', 'Egnatia trafik gürültüsü; bazı sokaklarda bakımsız yapılar.'),
 'OSE Tren İstasyonu': ('Yeni Tren Garı metro terminali; en uygun fiyatlar; ulaşım düğümü.', 'Çevre algısı zayıf; kiracı profili ve boş kalma riski görece yüksek; gürültü.'),
 'Faliro': ('Ippokrateio Hastanesi (sağlık çalışanı talebi), Nea Paralia sahili, Papafi/Efkleidis metro; aileler için uygun.', 'Merkeze 1,5–2,5 km; bazı sokaklar dar ve otopark sıkıntılı.'),
}
DEMAND = {'Tsimiski / Aristotelous / Tarihi Merkez': (90, 88, 'Düşük'), 'Agia Sofia': (90, 88, 'Düşük'), 'Beyaz Kule / XANTH / Fuar': (88, 92, 'Düşük'),
          'Kamara - Rotonda': (92, 85, 'Düşük–Orta (öğrenci döngüsü)'), 'Agios Dimitrios': (82, 78, 'Orta'), 'Vardaris': (70, 65, 'Orta–Yüksek'),
          'Dioikitirio': (78, 72, 'Orta'), 'Antigonidon': (80, 75, 'Orta'), 'OSE Tren İstasyonu': (65, 60, 'Yüksek'), 'Faliro': (80, 80, 'Orta')}

def esc(s): return (s or '').replace('"', "'")
for idx, rec in enumerate(R):
    r = FR + idx
    put(M, r, 1, rec['no'], font=Font(name=F, size=9, bold=True, color=NAVY))
    t = f"Yenilenmiş 2+1 daire, {int(rec['sqm'])} m², {floor_tr(rec['floor'])} – {rec['mah']}"
    put(M, r, 2, t)
    link(M, r, 3, rec['url'], 'İlanı aç ↗')
    put(M, r, 4, 'spitimou.net (ajans ilanlarını toplayan portal)', align=WRAP)
    put(M, r, 5, rec['adv'])
    contact = (rec['advweb'] or 'web yok') + (f" | Ajans ref.: {rec['ref']}" if rec['ref'] else '') + ' | Tel/form: ilan sayfasından'
    put(M, r, 6, contact, align=WRAP)
    put(M, r, 7, datetime.date.fromisoformat(rec['upd']), fmt=DATE)
    put(M, r, 8, datetime.date.fromisoformat(rec['first']) if rec['first'] else 'Bilinmiyor', fmt=DATE)
    put(M, r, 9, datetime.date.fromisoformat(rec['lastseen'][:10]), fmt=DATE)
    put(M, r, 10, f'=IF(P_KONTROL-{ref("ver", r)}<=35,"Yayında (portal teyitli)","Teyit gerekli – son görülme "&DAY({ref("ver", r)})&"."&MONTH({ref("ver", r)})&"."&YEAR({ref("ver", r)}))')
    link(M, r, 11, rec['url'] + '?gallery=1', 'Fotoğraflar ↗')
    put(M, r, 12, rec['mah']); put(M, r, 13, rec['grp']); put(M, r, 14, 'Öncelikli merkez')
    st = rec['onstreet'] and f"{rec['onstreet']} üzerinde (ilan beyanı)" or (', '.join(rec['streets']) if rec['streets'] else 'İlanda belirtilmemiş')
    put(M, r, 15, st, align=WRAP)
    put(M, r, 16, rec['address'] or 'İlanda yok')
    link(M, r, 17, f"https://www.google.com/maps?q={rec['lat']:.6f},{rec['lng']:.6f}", 'Harita ↗')
    put(M, r, 18, rec['dcenter'], fmt=NUM)
    put(M, r, 19, rec['onstreet'] or rec['road'])
    put(M, r, 20, 0 if rec['onstreet'] else rec['droad'], fmt=NUM)
    put(M, r, 21, rec['metro']); put(M, r, 22, rec['dmetro'], fmt=NUM); put(M, r, 23, rec['duni'], fmt=NUM)
    put(M, r, 24, rec['dsea'] if rec['dsea'] is not None else 'Bilinmiyor', fmt=NUM)
    a, d = MAH_TXT[rec['grp']]
    put(M, r, 25, a, align=WRAP); put(M, r, 26, d, align=WRAP)
    put(M, r, 27, rec['price'], fmt=EUR)
    put(M, r, 28, rec['net'] if rec['net'] else 'Belirtilmemiş', fmt=NUM)
    put(M, r, 29, rec['sqm'], fmt=NUM)
    put(M, r, 30, f'={ref("price", r)}/{ref("sqm", r)}', fmt=EUR)
    put(M, r, 31, rec['beds'], fmt=NUM)
    put(M, r, 32, 'Teyitli (ilan metni)' if rec['two'] else 'Yalnız portal alanı')
    put(M, r, 33, rec['living'] if rec['living'] is not None else 'Bilinmiyor', fmt=NUM)
    put(M, r, 34, rec['baths'] if rec['baths'] is not None else 'Bilinmiyor', fmt=NUM)
    put(M, r, 35, rec['yb'] if rec['yb'] else 'Bilinmiyor', fmt='0')
    put(M, r, 36, rec['ry'] if rec['ry'] else 'Bilinmiyor', fmt='0')
    put(M, r, 37, rec['scope'] + ('' if rec['ry_src'] in ('portal', 'portal + ilan metni') else f" – yenileme yılı {rec['ry_src']}"), align=WRAP)
    put(M, r, 38, 'Evet' if rec['under'] else 'Hayır')
    put(M, r, 39, floor_tr(rec['floor'])); put(M, r, 40, rec['floorn'] if rec['floorn'] is not None else 0, fmt=NUM1)
    put(M, r, 41, rec['tf'] if rec['tf'] else 'Bilinmiyor')
    put(M, r, 42, rec['elev']); put(M, r, 43, rec['balcony'])
    put(M, r, 44, HEAT_TR.get(rec['heat'], rec['heat'])); put(M, r, 45, rec['ac'])
    put(M, r, 46, {'exempt': 'Muaf'}.get(rec['energy'], rec['energy'] or 'Bilinmiyor'))
    put(M, r, 47, rec['furn']); put(M, r, 48, rec['parking'])
    put(M, r, 49, rec['fees'] if rec['fees'] is not None else '', fmt=EUR)
    bl = 'İyi bakımlı (ilan beyanı)' if re.search(r'καλοσυντηρημέν|προσεγμέν|άριστη κατάσταση', rec['desc']) else ('Cephe yenilenmiş (ilan)' if 'πρόσοψη' in rec['desc'] and 'ανακαιν' in rec['desc'] else 'İlanda bilgi yok – yerinde kontrol')
    put(M, r, 50, bl, align=WRAP)
    put(M, r, 51, rec['tenant'], align=WRAP)
    put(M, r, 52, rec['rent_lock'] if rec['rent_lock'] else '', fmt=EUR)
    lg = []
    if rec['yb'] and rec['yb'] < 1985: lg.append(f"{rec['yb']} yapımı: 1985 öncesi deprem yönetmeliği → statik kontrol")
    if not rec['yb']: lg.append('Yapım yılı yok')
    lg.append('Tapu/kadastro, imar (αυθαίρετα) ve enerji sertifikası avukat+mühendisçe kontrol edilmeli')
    if rec['under']: lg.append('Tadilat devam ediyor – teslim şartı sözleşmeye')
    if rec['tenant'].startswith('Kiracılı'): lg.append('Kiracılı – kira sözleşmesi, depozito ve bitiş tarihi incelenmeli')
    if rec['split']: lg.append('İlan: iki ayrı üniteye bölünmüş – imar ve kat mülkiyeti uygunluğu kontrol edilmeli')
    put(M, r, 53, '; '.join(lg), align=WRAP)
    du = '; '.join(f"{x['adv']}: {x['price']:,} €".replace(',', '.') for x in rec['dups']) if rec['dups'] else 'Tekrarı bulunamadı'
    put(M, r, 54, du, align=WRAP)
    # D – calculations
    CR = f"{q('comp')}$A$6:$A$2000"; CV = f"{q('comp')}$J$6:$J$2000"
    put(M, r, 55, f'=COUNTIF({CR},{ref("no", r)})', fmt=NUM)
    put(M, r, 56, f'=IFERROR(AVERAGEIF({CR},{ref("no", r)},{CV}),INDEX({q("mah")}$K$6:$K$40,MATCH({ref("grp", r)},{q("mah")}$A$6:$A$40,0)))', fmt=EUR2)
    c = M.cell(r, 57, 'Emsaller ↗'); c.hyperlink = f"#'{SN['comp']}'!A{6 + [i for i, x in enumerate(COMPS) if x['no'] == rec['no']][0]}" if rec['ncomp'] else None; c.font = fLink; c.border = BOX
    put(M, r, 58, f'={ref("compavg", r)}*(1-P_ILAN_INDIRIM)', fmt=EUR2)
    LK = ref('rentnow', r)
    put(M, r, 59, f'=IF(ISNUMBER({LK}),{LK},ROUND({ref("basem2", r)}*{ref("sqm", r)}*SK_KIRA,0))', fmt=EUR)
    put(M, r, 60, f'=IF(ISNUMBER({LK}),{LK},ROUND({ref("basem2", r)}*{ref("sqm", r)}*SB_KIRA,0))', fmt=EUR)
    put(M, r, 61, f'=IF(ISNUMBER({LK}),{LK},ROUND({ref("basem2", r)}*{ref("sqm", r)}*SI_KIRA,0))', fmt=EUR)
    put(M, r, 62, f'=CHOOSE(P_SENARYO,{ref("rk", r)},{ref("rb", r)},{ref("ri", r)})', fmt=EUR, fill=fillGr)
    put(M, r, 63, f'={ref("ra", r)}*12', fmt=EUR)
    put(M, r, 64, '=SA_BOS', fmt=NUM1)
    put(M, r, 65, f'={ref("ra", r)}*(12-{ref("vac", r)})', fmt=EUR)
    put(M, r, 66, f'={ref("gross", r)}*P_BAKIM*SA_MAS', fmt=EUR)
    put(M, r, 67, f'={ref("coll", r)}*P_YONETIM+{ref("ra", r)}*P_KIRACI_AY/P_SOZLESME', fmt=EUR)
    put(M, r, 68, f'={ref("sqm", r)}*P_ENFIA', fmt=EUR)
    put(M, r, 69, f'={ref("sqm", r)}*P_SIGORTA*SA_MAS', fmt=EUR)
    put(M, r, 70, f'=IF({ref("fees", r)}="",P_AIDAT_VARS,{ref("fees", r)})*({ref("vac", r)}+(12-{ref("vac", r)})*P_AIDAT_PAY)', fmt=EUR)
    put(M, r, 71, f'=SUM({ref("maint", r)}:{ref("feeown", r)})', fmt=EUR)
    put(M, r, 72, f'={ref("price", r)}*(1-SA_IND)', fmt=EUR)
    put(M, r, 73, f'={ref("buy", r)}*P_DEVIR', fmt=EUR)
    put(M, r, 74, f'={ref("buy", r)}*(P_NOTER+P_TAPU+P_AVUKAT)+P_SABIT', fmt=EUR)
    put(M, r, 75, f'={ref("buy", r)}*P_KOMISYON', fmt=EUR)
    put(M, r, 76, f'=IF(AND(P_ESYA_UYGULA=1,{ref("furn", r)}<>"Eşyalı"),P_ESYA,0)+P_HAZIRLIK+IF({ref("under", r)}="Evet",P_TADILAT_DEVAM,0)', fmt=EUR)
    put(M, r, 77, f'=SUM({ref("buy", r)}:{ref("furncost", r)})', fmt=EUR, font=Font(name=F, size=9, bold=True))
    put(M, r, 78, f'={ref("coll", r)}-{ref("opex", r)}', fmt=EUR)
    T = f'{ref("coll", r)}*(1-P_GOTURU)'
    put(M, r, 79, f'=IF(P_VERGI_YONTEM=1,MIN({T},P_V1_LIM)*P_V1+MAX(0,MIN({T},P_V2_LIM)-P_V1_LIM)*P_V2+MAX(0,MIN({T},P_V3_LIM)-P_V2_LIM)*P_V3+MAX(0,{T}-P_V3_LIM)*P_V4,{T}*P_VERGI_SABIT)', fmt=EUR)
    put(M, r, 80, f'={ref("noi", r)}-{ref("rtax", r)}', fmt=EUR, font=Font(name=F, size=9, bold=True))
    put(M, r, 81, f'={ref("gross", r)}/{ref("price", r)}', fmt=PCT)
    put(M, r, 82, f'={ref("ncf", r)}/{ref("tinv", r)}', fmt=PCT, font=Font(name=F, size=9, bold=True))
    put(M, r, 83, f'=IF({ref("ncf", r)}>0,{ref("tinv", r)}/{ref("ncf", r)},"Negatif")', fmt=NUM1)
    put(M, r, 84, f'={ref("price", r)}/{ref("gross", r)}', fmt=NUM1)
    put(M, r, 85, rec['exit_ppsm'] if rec['exit_ppsm'] else '', fmt=EUR)
    put(M, r, 86, rec['exit_n'] if rec['exit_n'] else '', fmt=NUM)
    put(M, r, 87, f'=IF({ref("exitppsm", r)}="",{ref("ppsm", r)}/INDEX({q("mah")}$I$6:$I$40,MATCH({ref("grp", r)},{q("mah")}$A$6:$A$40,0)),{ref("ppsm", r)}/{ref("exitppsm", r)})', fmt='0.00')
    put(M, r, 88, f'=ROUND(MAX({ref("price", r)}*(1-P_PAZARLIK_MAX),MIN({ref("price", r)}*(1-P_PAZARLIK),IF({ref("pratio", r)}>1,{ref("price", r)}/{ref("pratio", r)},{ref("price", r)}*(1-P_PAZARLIK)))),-3)', fmt=EUR)
    put(M, r, 89, f'=IF({ref("ncf", r)}>0,ROUND(({ref("ncf", r)}/P_HEDEF-P_SABIT-{ref("furncost", r)})/(1+P_MASRAF_PCT),-3),0)', fmt=EUR)
    RK = f"{q('risk')}"
    put(M, r, 90, f"={RK}R{6 + idx}", fmt=NUM, font=fX)
    put(M, r, 91, f'=IF({ref("risk", r)}>=40,"Yüksek",IF({ref("risk", r)}>=22,"Orta","Düşük"))')
    put(M, r, 92, f"={RK}S{6 + idx}", font=fX, align=WRAP)
    # scores
    put(M, r, 93, f'=MAX(0,MIN(100,({ref("ny", r)}-P_NY_MIN)/(P_NY_MAX-P_NY_MIN)*100))', fmt=NUM)
    MI = f"MATCH({ref('grp', r)},{q('mah')}$A$6:$A$40,0)"
    put(M, r, 94, f"=0.5*INDEX({q('mah')}$O$6:$O$40,{MI})+0.3*MAX(0,100-{ref('dmetro', r)}/8)+0.2*MAX(0,100-{ref('droad', r)}/3)", fmt=NUM)
    put(M, r, 95, f'=MAX(0,MIN(100,(1.25-{ref("pratio", r)})/0.5*100))', fmt=NUM)
    E = ref('energy', r)
    put(M, r, 96, f'=MIN(100,IF(N({ref("ry", r)})>=2024,40,IF(N({ref("ry", r)})>=2019,30,15))+IF({ref("elev", r)}="Evet",20,0)+IF(OR({E}="A+",{E}="A",{E}="B+",{E}="B"),20,IF({E}="C",15,IF({E}="D",10,IF({E}="E",5,0))))+IF(N({ref("yb", r)})>=1985,10,IF(N({ref("yb", r)})>=1960,5,0))+IF({ref("balc", r)}="Evet",5,0)+IF({ref("ac", r)}="Evet",5,0))', fmt=NUM)
    put(M, r, 97, f'=100-{ref("risk", r)}', fmt=NUM)
    put(M, r, 98, f"=MIN(100,0.5*INDEX({q('mah')}$P$6:$P$40,{MI})+IF(AND({ref('sqm', r)}>=50,{ref('sqm', r)}<=85),30,15)+IF(AND({ref('floorn', r)}>=2,{ref('elev', r)}=\"Evet\"),20,IF({ref('floorn', r)}>=1,10,0)))", fmt=NUM)
    put(M, r, 99, f'=IF(OR({ref("ncomp", r)}<3,{ref("two", r)}<>"Teyitli (ilan metni)"),"Yetersiz","Yeterli")')
    put(M, r, 100, f'=ROUND((W_GETIRI*{ref("s1", r)}+W_KONUM*{ref("s2", r)}+W_FIYAT*{ref("s3", r)}+W_KALITE*{ref("s4", r)}+W_RISK*{ref("s5", r)}+W_LIKIT*{ref("s6", r)})*IF({ref("dataok", r)}="Yetersiz",P_VERI_CEZA,1),1)', fmt=NUM1, font=Font(name=F, size=9, bold=True))
    put(M, r, 101, f'={ref("score", r)}+ROW()/10000000', fmt='0.0000')
    put(M, r, 102, f'=RANK({ref("helper", r)},${C["helper"]}${FR}:${C["helper"]}${LR})', fmt=NUM, font=Font(name=F, size=9, bold=True, color=NAVY))
    put(M, r, 103, f'=RANK({ref("ny", r)},${C["ny"]}${FR}:${C["ny"]}${LR})+COUNTIF(${C["ny"]}${FR}:{C["ny"]}{r},{ref("ny", r)})-1', fmt=NUM)
    if idx % 2 == 1:
        for cc in range(1, len(COLS) + 1):
            cell = M.cell(r, cc)
            if cell.fill is None or cell.fill.fgColor.rgb in (None, '00000000'):
                cell.fill = fillG
    M.row_dimensions[r].height = 30
assert C['nyrank'] == GL(103), C['nyrank']
M.freeze_panes = 'C5'
M.auto_filter.ref = f'A4:{GL(len(COLS))}{LR}'
rng = lambda k: f'{C[k]}{FR}:{C[k]}{LR}'
M.conditional_formatting.add(rng('ny'), ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
M.conditional_formatting.add(rng('gy'), ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
M.conditional_formatting.add(rng('risk'), ColorScaleRule(start_type='min', start_color='63BE7B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='F8696B'))
M.conditional_formatting.add(rng('score'), DataBarRule(start_type='num', start_value=0, end_type='num', end_value=100, color='2E7D32'))
M.conditional_formatting.add(rng('pratio'), CellIsRule(operator='greaterThan', formula=['1.15'], fill=PatternFill('solid', fgColor='F4CCCC')))
M.conditional_formatting.add(rng('pratio'), CellIsRule(operator='lessThan', formula=['0.9'], fill=PatternFill('solid', fgColor='C6EFCE')))
M.conditional_formatting.add(rng('risklvl'), CellIsRule(operator='equal', formula=['"Yüksek"'], fill=PatternFill('solid', fgColor='F4CCCC'), font=Font(name=F, size=9, bold=True, color='9C0006')))
M.conditional_formatting.add(rng('rank'), CellIsRule(operator='lessThanOrEqual', formula=['10'], fill=PatternFill('solid', fgColor='C6EFCE'), font=Font(name=F, size=9, bold=True, color=GREEN_D)))
M.conditional_formatting.add(rng('status'), FormulaRule(formula=[f'LEFT({C["status"]}{FR},5)="Teyit"'], fill=PatternFill('solid', fgColor='FCE4D6')))

# =====================================================================================
# KİRA EMSALLERİ
# =====================================================================================
K = ws['comp']
title(K, 'KİRA EMSALLERİ – Her satılık daire için aynı çevrede benzer kiralık daireler',
      'UYARI: Tüm emsal kiralar ilanlarda İSTENEN aylık kiradır; gerçekleşmiş sözleşme kirası değildir. Seçim: ≤900 m, ±%30 m², 2 yatak odası, 06/2026 sonrası güncellenmiş, aykırı değerler (%65–150 bandı dışı) çıkarıldı.', 15)
put(K, 3, 1, f'Kiralık ilan havuzu: spitimou.net, Selanik Belediyesi, 2+ yatak odası, 40–125 m² → {DATA["rentpool"]} tekil, makul fiyatlı ilan (2.050 ham ilan). Mesafe kuş uçuşudur.', font=fSm, border=False)
KH = [('Daire No', 9), ('Emsal #', 7), ('Kiralık ilan URL', 14), ('Emsal mahallesi (portal)', 24), ('m²', 7), ('Kat', 8), ('Durum', 13), ('Eşya (ilan metni)', 12),
      ('Aylık ilan kirası (€)', 10), ('Kira €/m²', 9), ('Satılık daireye mesafe (m)', 10), ('İlan güncelleme', 11), ('Satılık daire m²', 9), ('Satılık daire mahallesi', 22), ('Satılık daire grubu', 24)]
for i, (h, w) in enumerate(KH, 1): hdr(K, 5, i, h, w)
COND_TR = {'RENOVATED': 'Yenilenmiş', 'UNVERIFIED': 'Belirtilmemiş', 'NEW': 'Yeni bina', 'TO BE RENOVATED': 'Tadilat gerekli', 'FULLY RENOVATED': 'Tam yenilenmiş', 'UNFINISHED': 'Bitmemiş'}
MR = f"{q('main')}$A${FR}:$A${LR}"
for i, cpx in enumerate(COMPS):
    r = 6 + i
    put(K, r, 1, cpx['no'], font=Font(name=F, size=9, bold=True, color=NAVY)); put(K, r, 2, cpx['k'])
    link(K, r, 3, cpx['url'], 'Kiralık ilan ↗')
    put(K, r, 4, cpx['mah']); put(K, r, 5, cpx['sqm'], fmt=NUM); put(K, r, 6, floor_tr(cpx['floor'])); put(K, r, 7, COND_TR.get(cpx['cond'], cpx['cond']))
    put(K, r, 8, cpx['furn']); put(K, r, 9, cpx['rent'], fmt=EUR); put(K, r, 10, f'=I{r}/E{r}', fmt=EUR2)
    put(K, r, 11, cpx['dist'], fmt=NUM); put(K, r, 12, datetime.date.fromisoformat(cpx['upd']), fmt=DATE)
    put(K, r, 13, f"=INDEX({q('main')}${C['sqm']}${FR}:${C['sqm']}${LR},MATCH(A{r},{MR},0))", fmt=NUM, font=fX)
    put(K, r, 14, f"=INDEX({q('main')}${C['mah']}${FR}:${C['mah']}${LR},MATCH(A{r},{MR},0))", font=fX)
    put(K, r, 15, f"=INDEX({q('main')}${C['grp']}${FR}:${C['grp']}${LR},MATCH(A{r},{MR},0))", font=fX)
    if cpx['k'] == 1:
        for cc in range(1, 16): K.cell(r, cc).border = Border(left=thin, right=thin, bottom=thin, top=Side(style='medium', color=NAVY))
KL = 6 + len(COMPS) - 1
K.freeze_panes = 'A6'; K.auto_filter.ref = f'A5:O{KL}'
K.conditional_formatting.add(f'J6:J{KL}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))

# =====================================================================================
# MAHALLE ANALİZİ
# =====================================================================================
H = ws['mah']
title(H, 'MAHALLE ANALİZİ – Satış ve kira €/m², brüt getiri, talep ve boş kalma riski',
      'Sol blok: 100 dairelik örneklemden formülle. Sağ blok: tüm piyasa ilanlarından (spitimou, 17.11.2026) medyanlar. Talep/likidite puanları analist değerlendirmesidir (sarı – değiştirilebilir).', 20)
MH = [('Bölge grubu', 30), ('Örneklemdeki daire', 9), ('Ort. satış fiyatı (€)', 11), ('Ort. satış €/m² (örneklem)', 10), ('Ort. emsal ilan kirası €/m²', 10),
      ('Brüt kira getirisi (örneklem)', 10), ('Kira çarpanı (yıl)', 9), ('Ort. NET getiri (aktif senaryo)', 10),
      ('Piyasa medyan satış €/m² (yenilenmiş 2-yatak ilanları)', 12), ('Satış ilan sayısı (n)', 8), ('Piyasa medyan kira €/m² (2-yatak kiralık ilanları)', 12),
      ('Kiralık ilan sayısı (n)', 8), ('Piyasa brüt getirisi', 9), ('Piyasa kira çarpanı', 9), ('Talep puanı (0–100)', 9), ('Likidite puanı (0–100)', 9),
      ('Boş kalma riski', 16), ('Yatırım avantajları', 48), ('Dezavantajlar', 44)]
for i, (h, w) in enumerate(MH, 1): hdr(H, 5, i, h, w)
GRPS = list(DEMAND.keys())
GMAH = {}
for rec in R: GMAH.setdefault(rec['grp'], set()).add(rec['mah'])
mk_g = {}
# aggregate market medians per group: weighted by n (approximation of median using weighted mean of sub-area medians)
for g in GRPS:
    rows = [(MARKET[m]['sale_med'], MARKET[m]['sale_n'], MARKET[m]['rent_med'], MARKET[m]['rent_n']) for m in MARKET if m in {rec['mah']: 1 for rec in R if rec['grp'] == g} or (g == 'Tsimiski / Aristotelous / Tarihi Merkez' and m in ('Ladadika', 'Ladadika / Emboriou', 'Aristotelous / Tarihi Merkez', 'Roma Agorası - Venizelou', 'Merkez', 'Merkez (Tsimiski-Egnatia)')) or (g == 'Beyaz Kule / XANTH / Fuar' and m in ('Beyaz Kule', 'XANTH (Beyaz Kule)', 'Fuar (HELEXPO)')) or (g == 'Kamara - Rotonda' and m in ('Kamara - Rotonda', 'Navarinou Meydanı')) or (g == 'Agios Dimitrios' and m in ('Agios Dimitrios', 'Türk Konsolosluğu çevresi')) or (g == 'Faliro' and m.startswith('Faliro'))]
    sn = sum(x[1] for x in rows); rn = sum(x[3] for x in rows if x[2])
    mk_g[g] = (round(sum(x[0] * x[1] for x in rows) / sn) if sn else None, sn, round(sum(x[2] * x[3] for x in rows if x[2]) / rn, 2) if rn else None, rn)
MAINR = lambda k: f"{q('main')}${C[k]}${FR}:${C[k]}${LR}"
for i, g in enumerate(GRPS):
    r = 6 + i
    put(H, r, 1, g, font=Font(name=F, size=9, bold=True, color=NAVY))
    put(H, r, 2, f"=COUNTIF({MAINR('grp')},A{r})", fmt=NUM)
    put(H, r, 3, f"=IFERROR(AVERAGEIF({MAINR('grp')},A{r},{MAINR('price')}),\"-\")", fmt=EUR)
    put(H, r, 4, f"=IFERROR(AVERAGEIF({MAINR('grp')},A{r},{MAINR('ppsm')}),\"-\")", fmt=EUR)
    put(H, r, 5, f"=IFERROR(AVERAGEIF({q('comp')}$O$6:$O${KL},A{r},{q('comp')}$J$6:$J${KL}),K{r})", fmt=EUR2)
    put(H, r, 6, f'=IFERROR(E{r}*12/D{r},"-")', fmt=PCT)
    put(H, r, 7, f'=IFERROR(1/F{r},"-")', fmt=NUM1)
    put(H, r, 8, f"=IFERROR(AVERAGEIF({MAINR('grp')},A{r},{MAINR('ny')}),\"-\")", fmt=PCT)
    sm, sn, rm, rn = mk_g[g]
    put(H, r, 9, sm, fmt=EUR); put(H, r, 10, sn, fmt=NUM); put(H, r, 11, rm, fmt=EUR2); put(H, r, 12, rn, fmt=NUM)
    put(H, r, 13, f'=IFERROR(K{r}*12/I{r},"-")', fmt=PCT); put(H, r, 14, f'=IFERROR(1/M{r},"-")', fmt=NUM1)
    d1, d2, d3 = DEMAND[g]
    put(H, r, 15, d1, fill=fillIn); put(H, r, 16, d2, fill=fillIn); put(H, r, 17, d3, fill=fillIn)
    a, d = MAH_TXT[g]
    put(H, r, 18, a, align=WRAP); put(H, r, 19, d, align=WRAP)
    H.row_dimensions[r].height = 48
HL = 6 + len(GRPS) - 1
for col in ('F', 'H', 'M'):
    H.conditional_formatting.add(f'{col}6:{col}{HL}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
H.freeze_panes = 'B6'
put(H, HL + 2, 1, 'Not: Piyasa medyanları her alt bölgenin medyanının ilan sayısıyla ağırlıklı ortalamasıdır. Satış: spitimou yenilenmiş, 2 yatak odalı, 45–110 m² ilanlar (mükerrerler dahil). Kira: 2 yatak odalı, 40–125 m², 06/2026 sonrası, tekilleştirilmiş kiralık ilanlar.', font=fSm, border=False)
# alt-mahalle detail table
r = HL + 4
hdr(H, r, 1, 'Alt mahalle (portal)'); hdr(H, r, 2, 'Satış n'); hdr(H, r, 3, 'Medyan satış €/m²'); hdr(H, r, 4, 'Kira n'); hdr(H, r, 5, 'Medyan kira €/m²'); hdr(H, r, 6, 'Brüt getiri'); hdr(H, r, 7, 'Kira çarpanı')
for m, v in sorted(MARKET.items(), key=lambda x: -(x[1]['rent_med'] or 0) * 12 / (x[1]['sale_med'] or 1)):
    r += 1
    put(H, r, 1, m); put(H, r, 2, v['sale_n'], fmt=NUM); put(H, r, 3, v['sale_med'], fmt=EUR); put(H, r, 4, v['rent_n'], fmt=NUM); put(H, r, 5, v['rent_med'] if v['rent_med'] else '', fmt=EUR2)
    put(H, r, 6, f'=IF(E{r}="","-",E{r}*12/C{r})', fmt=PCT); put(H, r, 7, f'=IFERROR(1/F{r},"-")', fmt=NUM1)

# =====================================================================================
# RİSK VE EKSİK VERİLER
# =====================================================================================
RS = ws['risk']
title(RS, 'RİSK VE EKSİK VERİLER – Daire bazında risk işaretleri (1 = risk var)',
      '3. satırdaki ağırlıklar (sarı) değiştirilebilir. Risk puanı = Σ(işaret × ağırlık), en fazla 100. Ana listeye ve puanlamaya otomatik aktarılır.', 22)
FLAGS = [
 ('Asansör belirtilmemiş', 15, lambda r: f'=IF({q("main")}{C["elev"]}{r}="Evet",0,1)'),
 ('Tadilat devam ediyor', 12, lambda r: f'=IF({q("main")}{C["under"]}{r}="Evet",1,0)'),
 ('Kira emsali < 3', 12, lambda r: f'=IF({q("main")}{C["ncomp"]}{r}<3,1,0)'),
 ('Eski / teyitsiz ilan (son görülme > 35 gün veya güncelleme > eşik)', 10, lambda r: f'=IF(OR(P_KONTROL-{q("main")}{C["ver"]}{r}>35,P_KONTROL-{q("main")}{C["upd"]}{r}>P_ILAN_GUN),1,0)'),
 ('1985 öncesi bina / yapım yılı yok', 10, lambda r: f'=IF(N({q("main")}{C["yb"]}{r})<1985,1,0)'),
 ('2+1 yalnız portal alanı (metinde teyitsiz)', 10, lambda r: f'=IF({q("main")}{C["two"]}{r}="Teyitli (ilan metni)",0,1)'),
 ('Enerji sınıfı düşük (E–H) / bilinmiyor / muaf', 8, lambda r: f'=IF(OR({q("main")}{C["energy"]}{r}="A+",{q("main")}{C["energy"]}{r}="A",{q("main")}{C["energy"]}{r}="B+",{q("main")}{C["energy"]}{r}="B",{q("main")}{C["energy"]}{r}="C",{q("main")}{C["energy"]}{r}="D"),0,1)'),
 ('Kiracılı (sözleşme şartları bilinmiyor)', 8, lambda r: f'=IF(LEFT({q("main")}{C["tenant"]}{r},8)="Kiracılı",1,0)'),
 ('Fiyat emsalin %15+ üstünde', 8, lambda r: f'=IF({q("main")}{C["pratio"]}{r}>1.15,1,0)'),
 ('Zemin / asma kat / 1. kat', 5, lambda r: f'=IF({q("main")}{C["floorn"]}{r}<2,1,0)'),
 ('Tam adres yok (yalnız konum)', 2, lambda r: f'=IF({q("main")}{C["addr"]}{r}="İlanda yok",1,0)'),
 ('Aynı daire farklı fiyatlarla birden çok ilanda', 3, lambda r: f'=IF({q("main")}{C["dups"]}{r}="Tekrarı bulunamadı",0,1)'),
 ('Aidat bilgisi yok', 2, lambda r: f'=IF({q("main")}{C["fees"]}{r}="",1,0)'),
 ('Ana caddeye > 300 m (yaklaşık)', 3, lambda r: f'=IF({q("main")}{C["droad"]}{r}>300,1,0)'),
]
hdr(RS, 5, 1, 'Daire No', 9); hdr(RS, 5, 2, 'Mahalle', 20); hdr(RS, 5, 3, 'Fiyat (€)', 10)
for j, (nm, w, fn) in enumerate(FLAGS):
    hdr(RS, 5, 4 + j, nm, 11)
    c = put(RS, 3, 4 + j, w, fill=fillIn, align=CENTER)
put(RS, 3, 3, 'Ağırlık →', font=fB, border=False)
RC = 4 + len(FLAGS)  # = R (18)? compute
RISKCOL = GL(RC); MAINRISKCOL = GL(RC + 1); MISSCOL = GL(RC + 2)
assert RISKCOL == 'R' and MAINRISKCOL == 'S', (RISKCOL, MAINRISKCOL)
hdr(RS, 5, RC, 'RİSK PUANI', 9); hdr(RS, 5, RC + 1, 'Risk işaretleri (ağırlığa göre, en büyük önce)', 45); hdr(RS, 5, RC + 2, 'Eksik / teyit edilmesi gereken veriler', 50)
RS.row_dimensions[5].height = 70
for idx in range(N):
    r = 6 + idx; mr = FR + idx
    put(RS, r, 1, f"={q('main')}A{mr}", font=fX); put(RS, r, 2, f"={q('main')}{C['mah']}{mr}", font=fX); put(RS, r, 3, f"={q('main')}{C['price']}{mr}", fmt=EUR, font=fX)
    for j, (nm, w, fn) in enumerate(FLAGS):
        put(RS, r, 4 + j, fn(mr), fmt='0', align=CENTER)
    last = GL(RC - 1)
    put(RS, r, RC, f'=MIN(100,SUMPRODUCT(D{r}:{last}{r},$D$3:${last}$3))', fmt=NUM, font=Font(name=F, size=9, bold=True))
    # risk flags listed by weight (highest first)
    order = sorted(range(len(FLAGS)), key=lambda j: -FLAGS[j][1])
    parts = '&'.join(f'IF({GL(4 + j)}{r}=1,{GL(4 + j)}$5&"; ","")' for j in order)
    put(RS, r, RC + 1, f'=IF({RISKCOL}{r}=0,"Belirgin risk yok",{parts})', align=WRAP)
    mm = (f'=IF(N({q("main")}{C["yb"]}{mr})=0,"Yapım yılı; ","")&IF({q("main")}{C["energy"]}{mr}="Bilinmiyor","Enerji sınıfı; ","")'
          f'&IF({q("main")}{C["fees"]}{mr}="","Aidat; ","")&IF({q("main")}{C["net"]}{mr}="Belirtilmemiş","Net m²; ","")'
          f'&IF({q("main")}{C["addr"]}{mr}="İlanda yok","Tam adres; ","")&IF(ISNUMBER({q("main")}{C["tf"]}{mr}),"","Toplam kat; ")'
          f'&IF({q("main")}{C["park"]}{mr}="Belirtilmemiş","Otopark; ","")&IF(LEFT({q("main")}{C["bldg"]}{mr},5)="İland","Bina durumu; ","")'
          f'&IF(LEFT({q("main")}{C["status"]}{mr},5)="Teyit","İlanın güncelliği; ","")&"Tapu/imar/statik (her daire)"')
    put(RS, r, RC + 2, mm, align=WRAP)
RL = 6 + N - 1
RS.freeze_panes = 'D6'; RS.auto_filter.ref = f'A5:{GL(RC + 2)}{RL}'
RS.conditional_formatting.add(f'D6:{GL(RC - 1)}{RL}', CellIsRule(operator='equal', formula=['1'], fill=PatternFill('solid', fgColor='F4CCCC'), font=Font(name=F, size=9, bold=True, color='9C0006')))
RS.conditional_formatting.add(f'{RISKCOL}6:{RISKCOL}{RL}', ColorScaleRule(start_type='min', start_color='63BE7B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='F8696B'))
r = RL + 2
put(RS, r, 1, 'Risk özeti', font=fB, border=False)
for j, (nm, w, fn) in enumerate(FLAGS):
    put(RS, r + 1 + j, 2, nm, font=fN); put(RS, r + 1 + j, 3, f'=SUM({GL(4 + j)}6:{GL(4 + j)}{RL})', fmt=NUM)
put(RS, r + 1 + len(FLAGS), 2, 'Yüksek riskli daire (puan ≥40)', font=fB); put(RS, r + 1 + len(FLAGS), 3, f'=COUNTIF({RISKCOL}6:{RISKCOL}{RL},">=40")', fmt=NUM)

# =====================================================================================
# SENARYO – per apartment
# =====================================================================================
hdr(S, 17, 1, 'Daire bazında üç senaryo karşılaştırması (aynı formül zinciri; yalnız senaryo varsayımları değişir)', None)
S.merge_cells('A17:G17')
blocks = [('Kötümser', 'SK'), ('Baz', 'SB'), ('İyimser', 'SI')]
sub = [('Aylık kira (€)', EUR), ('Tahsil edilen yıllık kira (€)', EUR), ('İşletme gideri (€)', EUR), ('Kira vergisi (€)', EUR), ('Vergi sonrası net (€)', EUR), ('Toplam yatırım (€)', EUR), ('Net getiri', PCT)]
hdr(S, 19, 1, 'Daire No'); hdr(S, 19, 2, 'Bölge grubu'); hdr(S, 19, 3, 'İlan fiyatı (€)')
col = 4
for bn, pf in blocks:
    c = S.cell(18, col, bn); c.font = fH; c.fill = fillT; c.alignment = CENTER
    S.merge_cells(start_row=18, start_column=col, end_row=18, end_column=col + len(sub) - 1)
    for sname, fmt in sub:
        hdr(S, 19, col, sname, 10); col += 1
hdr(S, 19, col, 'Fark: İyimser – Kötümser net getiri', 12)
S.row_dimensions[19].height = 48
for idx in range(N):
    r = 20 + idx; mr = FR + idx
    m = lambda k: f"{q('main')}{C[k]}{mr}"
    put(S, r, 1, f"={m('no')}", font=fX); put(S, r, 2, f"={m('grp')}", font=fX); put(S, r, 3, f"={m('price')}", fmt=EUR, font=fX)
    col = 4
    for bn, pf in blocks:
        L = [GL(col + i) for i in range(7)]
        put(S, r, col, f"=IF(ISNUMBER({m('rentnow')}),{m('rentnow')},ROUND({m('basem2')}*{m('sqm')}*{pf}_KIRA,0))", fmt=EUR)
        put(S, r, col + 1, f'={L[0]}{r}*(12-{pf}_BOS)', fmt=EUR)
        put(S, r, col + 2, f"={L[0]}{r}*12*P_BAKIM*{pf}_MAS+{L[1]}{r}*P_YONETIM+{L[0]}{r}*P_KIRACI_AY/P_SOZLESME+{m('sqm')}*P_ENFIA+{m('sqm')}*P_SIGORTA*{pf}_MAS+IF({m('fees')}=\"\",P_AIDAT_VARS,{m('fees')})*({pf}_BOS+(12-{pf}_BOS)*P_AIDAT_PAY)", fmt=EUR)
        T = f'{L[1]}{r}*(1-P_GOTURU)'
        put(S, r, col + 3, f'=IF(P_VERGI_YONTEM=1,MIN({T},P_V1_LIM)*P_V1+MAX(0,MIN({T},P_V2_LIM)-P_V1_LIM)*P_V2+MAX(0,MIN({T},P_V3_LIM)-P_V2_LIM)*P_V3+MAX(0,{T}-P_V3_LIM)*P_V4,{T}*P_VERGI_SABIT)', fmt=EUR)
        put(S, r, col + 4, f'={L[1]}{r}-{L[2]}{r}-{L[3]}{r}', fmt=EUR)
        put(S, r, col + 5, f"={m('price')}*(1-{pf}_IND)*(1+P_MASRAF_PCT)+P_SABIT+{m('furncost')}", fmt=EUR)
        put(S, r, col + 6, f'={L[4]}{r}/{L[5]}{r}', fmt=PCT, font=Font(name=F, size=9, bold=True))
        col += 7
    put(S, r, col, f'={GL(col - 1)}{r}-{GL(col - 8)}{r}', fmt=PCT)
SL = 20 + N - 1
for c0 in (10, 17, 24):
    S.conditional_formatting.add(f'{GL(c0)}20:{GL(c0)}{SL}', ColorScaleRule(start_type='num', start_value=0, start_color='F8696B', mid_type='num', mid_value=0.03, mid_color='FFEB84', end_type='num', end_value=0.05, end_color='63BE7B'))
S.freeze_panes = 'D20'
# summary rows
put(S, 16, 1, 'Portföy ortalaması net getiri →', font=fB, border=False)
for j, c0 in enumerate((10, 17, 24)):
    put(S, 16, 3 + j, f'=AVERAGE({GL(c0)}20:{GL(c0)}{SL})', fmt=PCT, font=Font(name=F, size=10, bold=True, color=GREEN_D))
dname('SCN_AVG_K', f"{q('scen')}$C$16"); dname('SCN_AVG_B', f"{q('scen')}$D$16"); dname('SCN_AVG_I', f"{q('scen')}$E$16")

# =====================================================================================
# 5 VE 10 YILLIK NAKİT AKIŞI
# =====================================================================================
CF = ws['cf']
title(CF, '5 VE 10 YILLIK NAKİT AKIŞI – Kira artışı, gider artışı, değer değişimi, çıkış maliyeti ve IRR',
      'Üst blok: seçilen daire için yıl yıl tablo (B4’ten daire seçin). Alt blok: 100 dairenin tamamı için 5 ve 10 yıllık IRR ve toplam getiri. Varsayımlar SENARYO ANALİZİ’ndeki AKTİF sütundan gelir.', 26)
put(CF, 4, 1, 'Seçili daire →', font=fB, border=False)
c = put(CF, 4, 2, R[0]['no'], fill=fillIn)
dvn = DataValidation(type='list', formula1=f"={q('main')}$A${FR}:$A${LR}"); CF.add_data_validation(dvn); dvn.add('B4')
def mlook(k): return f"INDEX({q('main')}${C[k]}${FR}:${C[k]}${LR},MATCH($B$4,{q('main')}$A${FR}:$A${LR},0))"
info = [('Mahalle', 'mah', None), ('İlan fiyatı', 'price', EUR), ('Alım bedeli', 'buy', EUR), ('Toplam yatırım', 'tinv', EUR), ('1. yıl tahsil edilen kira', 'coll', EUR), ('1. yıl işletme gideri', 'opex', EUR)]
for i, (lab, k, fmt) in enumerate(info):
    put(CF, 5 + i, 1, lab, font=fN); put(CF, 5 + i, 2, f'={mlook(k)}', fmt=fmt, font=fX)
CF.column_dimensions['A'].width = 34; CF.column_dimensions['B'].width = 14
hdr(CF, 12, 1, 'Kalem')
for y in range(11): hdr(CF, 12, 2 + y, f'Yıl {y}', 11)
rows_cf = ['Tahsil edilen kira', 'İşletme giderleri', 'Kira geliri vergisi', 'Vergi sonrası net nakit akışı', 'Kümülatif net kira', 'Mülk değeri (yıl sonu)',
           '5. yılda çıkış: net satış geliri', 'Nakit akışı – 5 yıl tutma', '10. yılda çıkış: net satış geliri', 'Nakit akışı – 10 yıl tutma']
for i, lab in enumerate(rows_cf): put(CF, 13 + i, 1, lab, font=fB if 'Nakit akışı' in lab else fN)
for y in range(11):
    cl = GL(2 + y)
    if y == 0:
        for i in range(4): put(CF, 13 + i, 2, 0, fmt=EUR, font=fN)
        put(CF, 17, 2, 0, fmt=EUR, font=fN)
        put(CF, 18, 2, '=B7', fmt=EUR)
        put(CF, 19, 2, 0, fmt=EUR, font=fN); put(CF, 21, 2, 0, fmt=EUR, font=fN)
        put(CF, 20, 2, '=-B8', fmt=EUR); put(CF, 22, 2, '=-B8', fmt=EUR)
        continue
    put(CF, 13, 2 + y, f'=$B$9*(1+SA_KARTIS)^{y - 1}', fmt=EUR)
    put(CF, 14, 2 + y, f'=$B$10*(1+SA_GARTIS)^{y - 1}', fmt=EUR)
    T = f'{cl}13*(1-P_GOTURU)'
    put(CF, 15, 2 + y, f'=IF(P_VERGI_YONTEM=1,MIN({T},P_V1_LIM)*P_V1+MAX(0,MIN({T},P_V2_LIM)-P_V1_LIM)*P_V2+MAX(0,MIN({T},P_V3_LIM)-P_V2_LIM)*P_V3+MAX(0,{T}-P_V3_LIM)*P_V4,{T}*P_VERGI_SABIT)', fmt=EUR)
    put(CF, 16, 2 + y, f'={cl}13-{cl}14-{cl}15', fmt=EUR)
    put(CF, 17, 2 + y, f'={GL(1 + y)}17+{cl}16', fmt=EUR)
    put(CF, 18, 2 + y, f'=$B$7*(1+SA_DEGER)^{y}', fmt=EUR)
    put(CF, 19, 2 + y, f'=IF({y}=5,{cl}18*(1-SA_CIKIS)-SA_DKV*MAX(0,{cl}18-$B$7),0)', fmt=EUR)
    put(CF, 20, 2 + y, f'=IF({y}<=5,{cl}16+{cl}19,0)', fmt=EUR)
    put(CF, 21, 2 + y, f'=IF({y}=10,{cl}18*(1-SA_CIKIS)-SA_DKV*MAX(0,{cl}18-$B$7),0)', fmt=EUR)
    put(CF, 22, 2 + y, f'={cl}16+{cl}21', fmt=EUR)
for rr_ in (16, 20, 22):
    for y in range(11): CF.cell(rr_, 2 + y).fill = fillGr
put(CF, 24, 1, 'Özet', font=fB, border=False)
summ = [('5 yıl – kümülatif net kira (€)', '=F17', EUR), ('5 yıl – toplam kazanç (net kira + çıkış – yatırım) (€)', '=SUM(B20:G20)', EUR), ('5 yıl – toplam getiri (yatırıma)', '=B26/B8', PCT1),
        ('5 yıl – IRR (yıllık)', '=IRR(B20:G20)', PCT), ('10 yıl – kümülatif net kira (€)', '=K17', EUR), ('10 yıl – toplam kazanç (€)', '=SUM(B22:L22)', EUR),
        ('10 yıl – toplam getiri (yatırıma)', '=B30/B8', PCT1), ('10 yıl – IRR (yıllık)', '=IRR(B22:L22)', PCT)]
for i, (lab, f_, fmt) in enumerate(summ):
    put(CF, 25 + i, 1, lab, font=fN); put(CF, 25 + i, 2, f_, fmt=fmt, font=Font(name=F, size=9, bold=True))
# all-apartments grid
G0 = 36
put(CF, G0 - 1, 1, 'Tüm daireler – 5 ve 10 yıllık tutma senaryoları (aktif senaryo)', font=fB, border=False)
heads = ['Daire No', 'Bölge grubu', 'Toplam yatırım'] + [f'5Y: Yıl {y}' for y in range(6)] + [f'10Y: Yıl {y}' for y in range(11)] + ['5Y IRR', '10Y IRR', '5Y toplam getiri', '10Y toplam getiri', '10Y kümülatif net kira']
for i, h in enumerate(heads): hdr(CF, G0, 1 + i, h, None if i < 2 else 10)
CF.row_dimensions[G0].height = 30
for idx in range(N):
    r = G0 + 1 + idx; mr = FR + idx
    m = lambda k: f"{q('main')}{C[k]}{mr}"
    put(CF, r, 1, f"={m('no')}", font=fX); put(CF, r, 2, f"={m('grp')}", font=fX); put(CF, r, 3, f"={m('tinv')}", fmt=EUR, font=fX)
    def ncf(y):
        T = f"{m('coll')}*(1+SA_KARTIS)^{y - 1}*(1-P_GOTURU)"
        tax = f'IF(P_VERGI_YONTEM=1,MIN({T},P_V1_LIM)*P_V1+MAX(0,MIN({T},P_V2_LIM)-P_V1_LIM)*P_V2+MAX(0,MIN({T},P_V3_LIM)-P_V2_LIM)*P_V3+MAX(0,{T}-P_V3_LIM)*P_V4,{T}*P_VERGI_SABIT)'
        return f"{m('coll')}*(1+SA_KARTIS)^{y - 1}-{m('opex')}*(1+SA_GARTIS)^{y - 1}-{tax}"
    def exitv(y): return f"{m('buy')}*(1+SA_DEGER)^{y}*(1-SA_CIKIS)-SA_DKV*MAX(0,{m('buy')}*((1+SA_DEGER)^{y}-1))"
    for y in range(6):
        put(CF, r, 4 + y, f'=-C{r}' if y == 0 else f'={ncf(y)}' + (f'+{exitv(5)}' if y == 5 else ''), fmt=EUR)
    for y in range(11):
        put(CF, r, 10 + y, f'=-C{r}' if y == 0 else f'={ncf(y)}' + (f'+{exitv(10)}' if y == 10 else ''), fmt=EUR)
    put(CF, r, 21, f'=IRR(D{r}:I{r})', fmt=PCT, font=Font(name=F, size=9, bold=True))
    put(CF, r, 22, f'=IRR(J{r}:T{r})', fmt=PCT, font=Font(name=F, size=9, bold=True))
    put(CF, r, 23, f'=SUM(D{r}:I{r})/C{r}', fmt=PCT1)
    put(CF, r, 24, f'=SUM(J{r}:T{r})/C{r}', fmt=PCT1)
    put(CF, r, 25, f'=SUM(K{r}:S{r})+{ncf(10)}', fmt=EUR)
GL_ = G0 + N
for cc in ('U', 'V'):
    CF.conditional_formatting.add(f'{cc}{G0 + 1}:{cc}{GL_}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
put(CF, G0 - 1, 21, 'Ortalama →', font=fB, border=False)
put(CF, 33, 1, 'Portföy ortalaması: 5Y IRR / 10Y IRR', font=fB); put(CF, 33, 2, f'=AVERAGE(U{G0 + 1}:U{GL_})', fmt=PCT); put(CF, 33, 3, f'=AVERAGE(V{G0 + 1}:V{GL_})', fmt=PCT)
CF.freeze_panes = 'B13'

# =====================================================================================
# EN İYİ 20
# =====================================================================================
TP = ws['top']
title(TP, 'EN İYİ 20 YATIRIM – Genel yatırım puanına göre (net getiri %40 + risk ve diğer kriterler)',
      'Sıralama formülle (INDEX/MATCH) çalışır; parametre veya senaryo değişince liste otomatik güncellenir. Yetersiz veri cezası uygulanmış puanlar kullanılır.', 20)
TH = [('Sıra', 6), ('Daire No', 9), ('İlan', 9), ('Bölge grubu', 24), ('Mahalle', 20), ('Fiyat (€)', 10), ('m²', 6), ('€/m²', 8), ('Aylık kira – aktif (€)', 9),
      ('Toplam yatırım (€)', 11), ('Net nakit (€/yıl)', 10), ('Brüt getiri', 8), ('NET GETİRİ', 9), ('Risk puanı', 7), ('GENEL PUAN', 8), ('10Y IRR', 8),
      ('Pazarlık önerilen fiyat (€)', 11), ('Neden seçildi (formülle üretilir)', 48), ('En büyük risk', 30)]
for i, (h, w) in enumerate(TH, 1): hdr(TP, 4, i, h, w)
TP.row_dimensions[4].height = 40
for k in range(20):
    r = 5 + k
    put(TP, r, 1, k + 1, font=Font(name=F, size=10, bold=True, color=NAVY), align=CENTER)
    MT = f"MATCH($A{r},{q('main')}${C['rank']}${FR}:${C['rank']}${LR},0)"
    put(TP, r, 2, f"=INDEX({q('main')}$A${FR}:$A${LR},{MT})", font=Font(name=F, size=9, bold=True, color='008000'))
    I = lambda kk: f"INDEX({q('main')}${C[kk]}${FR}:${C[kk]}${LR},{MT})"
    # Hyperlink via formula to listing: rebuild from id map
    put(TP, r, 3, f'=HYPERLINK(INDEX({q("src")}$C$7:$C${6 + N},MATCH(B{r},{q("src")}$A$7:$A${6 + N},0)),"İlan ↗")', font=fLink)
    for j, (kk, fmt) in enumerate([('grp', None), ('mah', None), ('price', EUR), ('sqm', NUM), ('ppsm', EUR), ('ra', EUR), ('tinv', EUR), ('ncf', EUR), ('gy', PCT), ('ny', PCT), ('risk', NUM), ('score', NUM1)]):
        put(TP, r, 4 + j, f'={I(kk)}', fmt=fmt, font=fX)
    put(TP, r, 16, f"=INDEX({q('cf')}$V${G0 + 1}:$V${GL_},{MT})", fmt=PCT, font=fX)
    put(TP, r, 17, f'={I("nego")}', fmt=EUR, font=fX)
    why = (f'=IF({I("ny")}>=AVERAGE({q("main")}{C["ny"]}{FR}:{C["ny"]}{LR}),"Net getiri ortalamanın üstünde (%"&ROUND({I("ny")}*100,2)&"); ","")'
           f'&IF({I("pratio")}<0.95,"Fiyat emsal €/m² altında ("&ROUND({I("pratio")},2)&"x); ","")'
           f'&IF({I("dmetro")}<=400,"Metroya "&{I("dmetro")}&" m; ","")'
           f'&IF({I("droad")}<=150,"Ana caddeye yakın ("&{I("road")}&"); ","")'
           f'&IF(N({I("ry")})>=2024,"Yeni yenilenmiş ("&{I("ry")}&"); ","")'
           f'&IF({I("risk")}<22,"Düşük risk profili; ","")&"Kira "&{I("ncomp")}&" emsalle destekli"')
    put(TP, r, 18, why, align=WRAP)
    put(TP, r, 19, f'={I("mainrisk")}', font=fX, align=WRAP)
    TP.row_dimensions[r].height = 42
TP.conditional_formatting.add('M5:M24', ColorScaleRule(start_type='min', start_color='FFEB84', end_type='max', end_color='63BE7B'))
TP.conditional_formatting.add('O5:O24', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=100, color='2E7D32'))
TP.conditional_formatting.add('N5:N24', ColorScaleRule(start_type='min', start_color='63BE7B', end_type='max', end_color='F8696B'))
TP.freeze_panes = 'C5'
# net getiri top 20 (pure)
put(TP, 27, 1, 'Ek tablo: Yalnızca NET KİRA GETİRİSİNE göre en yüksek 20 daire (risk dikkate alınmadan)', font=fB, border=False)
for i, h in enumerate(['Net getiri sırası', 'Daire No', 'Bölge grubu', 'Fiyat (€)', 'Net getiri', 'Genel puan sırası', 'Risk puanı', 'Veri yeterliliği'], 1): hdr(TP, 28, i, h)
for k in range(20):
    r = 29 + k
    put(TP, r, 1, k + 1, align=CENTER)
    MT = f"MATCH($A{r},{q('main')}${C['nyrank']}${FR}:${C['nyrank']}${LR},0)"
    I = lambda kk: f"INDEX({q('main')}${C[kk]}${FR}:${C[kk]}${LR},{MT})"
    for j, (kk, fmt) in enumerate([('no', None), ('grp', None), ('price', EUR), ('ny', PCT), ('rank', NUM), ('risk', NUM), ('dataok', None)]):
        put(TP, r, 2 + j, f'={I(kk)}', fmt=fmt, font=fX)

# =====================================================================================
# KAYNAKLAR
# =====================================================================================
SR = ws['src']
title(SR, 'KAYNAKLAR – İlan bağlantıları, kira emsalleri, veri kaynakları ve kontrol tarihleri', 'Tüm bağlantılar tıklanabilir. Kontrol tarihi: 17.11.2026.', 9)
put(SR, 4, 1, 'A) 100 satılık ilan', font=fB, border=False)
for i, h in enumerate(['Daire No', 'spitimou ilan ID', 'İlan URL', 'Emlak ofisi', 'Ofis web sitesi', 'Ajans ilan ref.', 'Son güncelleme', 'Portalda son görülme', 'Kontrol tarihi'], 1):
    hdr(SR, 6, i, h, [9, 11, 60, 26, 30, 12, 12, 12, 12][i - 1])
for idx, rec in enumerate(R):
    r = 7 + idx
    put(SR, r, 1, rec['no']); put(SR, r, 2, rec['id'])
    c = put(SR, r, 3, rec['url']); c.hyperlink = rec['url']; c.font = fLink
    put(SR, r, 4, rec['adv'])
    if rec['advweb']:
        c = put(SR, r, 5, rec['advweb']); c.hyperlink = rec['advweb']; c.font = fLink
    else: put(SR, r, 5, '-')
    put(SR, r, 6, str(rec['ref'] or '-')); put(SR, r, 7, datetime.date.fromisoformat(rec['upd']), fmt=DATE)
    put(SR, r, 8, datetime.date.fromisoformat(rec['lastseen'][:10]), fmt=DATE); put(SR, r, 9, CHECK, fmt=DATE)
r = 7 + N + 1
put(SR, r, 1, 'B) Aynı dairenin diğer ofislerdeki ilanları (mükerrer – fiyat karşılaştırması)', font=fB, border=False); r += 1
for i, h in enumerate(['Daire No', 'spitimou ilan ID', 'İlan URL', 'Emlak ofisi', 'İlan fiyatı (€)', 'Ajans ref.'], 1): hdr(SR, r, i, h)
for rec in R:
    for x in rec['dups']:
        r += 1
        put(SR, r, 1, rec['no']); put(SR, r, 2, x['id']); u = 'https://spitimou.net/en' + x['path']
        c = put(SR, r, 3, u); c.hyperlink = u; c.font = fLink
        put(SR, r, 4, x['adv'] or '-'); put(SR, r, 5, x['price'], fmt=EUR); put(SR, r, 6, str(x['ref'] or '-'))
r += 2
put(SR, r, 1, 'C) Kira emsali ilanları: KİRA EMSALLERİ sayfasında tıklanabilir olarak listelenmiştir (' + str(len(COMPS)) + ' emsal satırı).', font=fB, border=False)
r += 2
put(SR, r, 1, 'D) Veri ve resmî/yarı resmî kaynaklar', font=fB, border=False); r += 1
SRCS = [
 ('İlan verisi', 'spitimou.net – Yunan emlak ofislerinin ilanlarını toplayan portal (JSON API: /api/listings, /api/listings/{id}).', 'https://spitimou.net/en/sale/thessaloniki-kentro', '17.11.2026'),
 ('İlan verisi', 'Satış: Selanik Belediyesi (thessaloniki-kentro), Daire, 2+ yatak odası, 45–110 m², yenilenmiş → 2.657 ilan; detayı çekilen 984 ilan; 498 tekil daire.', 'https://spitimou.net/en/sale/thessaloniki-kentro?types=Apartment', '17.11.2026'),
 ('Kira verisi', 'Kiralık: aynı bölge, 2+ yatak odası, 40–125 m² → 2.050 ham ilan, 1.275 tekil/makul ilan.', 'https://spitimou.net/en/rent/thessaloniki-kentro', '17.11.2026'),
 ('Vergi', 'Kira geliri vergisi 2026 tarifesi (%15/%25/%35/%45) – Kanun 5246/2025 (basın özeti)', 'https://greekcitytimes.com/2026/03/13/greece-landlords-tax-exemption-rental-income-e2-codes-2026/', '17.11.2026'),
 ('Vergi', 'KPMG Yunanistan vergi güncellemesi (3. dilim eşiği 36.000 €)', 'https://kpmg.com/gr/en/insights/2025/11/tax-updates-20112025.html', '17.11.2026'),
 ('Vergi', 'Devir vergisi %3,09, noter/tapu/avukat maliyetleri', 'https://www.globalpropertyguide.com/europe/greece/taxes-and-costs', '17.11.2026'),
 ('Vergi', 'AB/AEA dışı vergi mukimleri için %15 devir vergisi yasası (Kanun 5355/2026, 01.07.2027’den itibaren)', 'https://greekcitytimes.com/2026/11/17/greece-15-percent-property-transfer-tax-law-non-eu', '17.11.2026'),
 ('Vergi', 'IMI Daily – Parlamento %15 devir vergisini kabul etti', 'https://www.imidaily.com/program-updates/greek-parliament-passes-15-transfer-tax-on-homes-bought-by-non-eu-tax-residents', '17.11.2026'),
 ('Vergi', 'ENFIA hesaplama mantığı (bölge değeri bantları)', 'https://hellascalculator.pages.dev/blog/enfia-property-tax-explained', '17.11.2026'),
 ('Vergi', 'Kira ödemelerinde nakit yasağı (01.01.2026)', 'https://www.tovima.com/society/renting-in-greece-cash-payments-end-on-jan-1-2026/', '17.11.2026'),
 ('Resmî', 'AADE – Yunan Vergi İdaresi (ENFIA, E2 kira beyanı, AFM)', 'https://www.aade.gr', 'Teyit için'),
 ('Konum', 'Metro istasyon koordinatları – Wikipedia (Venizelou, Agias Sofias, Sintrivani, Panepistimio, Dimokratias, Papafi, Efkleidis, Fleming, Analipseos, Yeni Gar)', 'https://en.wikipedia.org/wiki/List_of_Thessaloniki_Metro_stations', '17.11.2026'),
 ('Konum', 'Ana cadde mesafesi: Egnatia, Tsimiski, Mitropoleos, Nikis, Agiou Dimitriou, Ethnikis Aminis, Vas. Olgas, Lagkada, Monastiriou güzergâhları elle yaklaşık koordinatlarla tanımlandı (±100–150 m hata payı).', '', '17.11.2026'),
 ('Erişilemeyen', 'spitogatos.gr (bot koruması), xe.gr / indomio.gr / plot.gr / properstar.com (HTTP 403), remax.gr (HTTP 500), century21global.com (bağlantı yok). spitimou bu portallardaki ofis ilanlarının çoğunu topladığı için ana kaynak olarak kullanıldı.', '', '17.11.2026'),
]
for i, h in enumerate(['Tür', 'Açıklama', 'Bağlantı', 'Kontrol'], 1): hdr(SR, r, i, h)
for t_, d_, u_, k_ in SRCS:
    r += 1
    put(SR, r, 1, t_); put(SR, r, 2, d_, align=WRAP)
    if u_:
        c = put(SR, r, 3, u_); c.hyperlink = u_; c.font = fLink
    else: put(SR, r, 3, '-')
    put(SR, r, 4, k_)
    SR.row_dimensions[r].height = 30
SR.column_dimensions['B'].width = 60

# =====================================================================================
# YEDEK İLANLAR
# =====================================================================================
X = ws['ext']
title(X, 'YEDEK İLANLAR – Ana listeye alınmayan tekil merkez ilanları (kriter dışı veya kota nedeniyle)',
      'Bu ilanlar modele DAHİL DEĞİLDİR. Eksik kriter sütunu nedenini gösterir. İhtiyaç halinde ana listeye taşınabilir.', 10)
for i, (h, w) in enumerate([('#', 5), ('spitimou ID', 10), ('İlan URL', 12), ('Mahalle', 22), ('Fiyat (€)', 10), ('m²', 6), ('€/m²', 8), ('Kat', 9), ('Yenileme yılı', 8), ('Emlak ofisi', 24), ('Son güncelleme', 11), ('Neden ana listede değil', 52)], 1):
    hdr(X, 4, i, h, w)
for i, e in enumerate(EXTRAS):
    r = 5 + i
    put(X, r, 1, i + 1); put(X, r, 2, e['id']); link(X, r, 3, e['url'], 'İlan ↗'); put(X, r, 4, e['mah']); put(X, r, 5, e['price'], fmt=EUR); put(X, r, 6, e['sqm'], fmt=NUM)
    put(X, r, 7, f'=E{r}/F{r}', fmt=EUR); put(X, r, 8, floor_tr(e['floor'])); put(X, r, 9, e['ry'] or 'Bilinmiyor'); put(X, r, 10, e['adv'] or '-')
    put(X, r, 11, datetime.date.fromisoformat(e['upd']) if e['upd'] else '-', fmt=DATE); put(X, r, 12, e['why'], align=WRAP)
X.freeze_panes = 'A5'; X.auto_filter.ref = f'A4:L{4 + len(EXTRAS)}'

# =====================================================================================
# VERGİ-HUKUK NOTLARI
# =====================================================================================
V = ws['tax']
title(V, 'YABANCI YATIRIMCI – SATIN ALMA, VERGİ VE HUKUKİ KONTROL NOTLARI', 'Genel bilgidir; hukuki/vergisel tavsiye değildir. Vergi mukimliğinize göre değişen konular ayrıca işaretlenmiştir. Kontrol: 17.11.2026', 3)
V.column_dimensions['A'].width = 34; V.column_dimensions['B'].width = 120
TXT = [
 ('1. SATIN ALMA SÜRECİ', ''),
 ('Vergi numarası (AFM)', 'Yabancı alıcı için Yunan vergi numarası zorunlu; AB/AEA dışında yaşayanlar için uygulamada Yunanistan’da vergi temsilcisi atanması istenebilir. Banka hesabı (ödeme izi için) önerilir.'),
 ('Avukat ve noter', 'Alıcı avukatı tapu (Ktimatologio) ve ipotek/haciz kayıtlarını inceler; satış noter huzurunda yapılır; tescil Ktimatologio’da.'),
 ('Sınır bölgesi kısıtı', 'AB/AEA dışı vatandaşlar için bazı sınır bölgelerinde izin gerekir (Kanun 1892/1990). Selanik şehir merkezi genelde bu kapsamda değildir – avukatınızla teyit edin.'),
 ('Devir vergisi', '%3 + belediye payı = %3,09; satış bedeli ile objektif (vergi) değerden yüksek olana uygulanır. Yeni yapılarda %24 KDV rejimi (askıda/istisna) söz konusu olabilir.'),
 ('YENİ: %15 devir vergisi', 'Raporlara göre Kanun 5355/2026 (FEK A 188, 13.11.2026): 01.07.2027 ve sonrası sözleşmelerde AB/AEA DIŞI VERGİ MUKİMLERİ için oran %15 (+belediye payı ≈ %15,45). Belirleyici olan uyruk değil vergi mukimliği. Türkiye vergi mukimi bir alıcı 01.07.2027’den sonra alım yaparsa 100.000 €’luk dairede ek ≈12.360 € maliyet. PARAMETRELER’de P_MUKIM=2 seçerek etkisini görün.'),
 ('Golden Visa', 'Selanik’te tek mülkle oturum izni eşiği 800.000 € (≥120 m²). Bu listedeki 2+1 daireler tek başına uygun değildir; ilanlardaki “Golden Visa uyumlu” ifadelerine dikkat edin.'),
 ('2. KİRA GELİRİNİN VERGİLENDİRİLMESİ', ''),
 ('Kaynak ülke vergisi', 'Yunanistan’daki gayrimenkulden elde edilen kira geliri, sahibinin mukimliğinden bağımsız olarak Yunanistan’da vergilendirilir; yıllık beyan (E1 + E2) gerekir.'),
 ('2026 kira tarifesi', '0–12.000 €: %15 · 12.000–24.000 €: %25 (yeni) · 24.000–36.000 €: %35 · üstü: %45. Matrah = brüt kira – %5 götürü indirim. Birden fazla daire alınırsa toplam kira geliri dilimleri yükseltir (P_VERGI_YONTEM=2 ile marjinal oranı girin).'),
 ('Mukimlik ülkesi vergisi (DEĞİŞKEN)', 'Türkiye vergi mukimi iseniz yurt dışı kira geliriniz Türkiye’de de beyan kapsamına girebilir; Türkiye–Yunanistan Çifte Vergilendirmeyi Önleme Anlaşması’na göre Yunanistan’da ödenen vergi mahsup edilebilir. Başka bir ülkede mukimseniz o ülkenin kuralları geçerlidir. Bu modelde YALNIZ Yunan vergisi hesaplanmıştır.'),
 ('Diğer yükümlülükler', 'Kira sözleşmesi AADE (Taxisnet) üzerinden elektronik bildirilmeli; 2026’dan itibaren kira ödemeleri nakit yapılamaz (banka). Konut kira sözleşmesi asgari 3 yıldır.'),
 ('Yıllık emlak vergileri', 'ENFIA (bölge değeri, m², kat, yaş katsayılarına göre) + belediye TAP (elektrik faturasına yansır). Modeldeki ENFIA €/m² tahminidir; satın almadan önce daire bazında hesaplatın.'),
 ('Satışta değer artış vergisi', 'Yunanistan’da %15 değer artış kazancı vergisi 31.12.2026’ya kadar askıda; sonrası belirsiz – kötümser senaryoda uygulanmıştır.'),
 ('3. AVUKATIN KONTROL ETMESİ GEREKENLER', ''),
 ('Mülkiyet', 'Ktimatologio kaydı ve tapu zinciri (20 yıl), satıcının tam yetkisi, miras/ortaklık durumu, vekâletnameler.'),
 ('Takyidatlar', 'İpotek, haciz, şerh, dava; belediye ve vergi borcu yokluk belgeleri (ENFIA ödeme belgesi).'),
 ('İmar', 'Kaçak/izin dışı alanlar (αυθαίρετα) için mühendis beyanı; kapalı balkon, çatı katı ilaveleri, depo/oda dönüşümleri.'),
 ('Kiracı', 'Kiracılı satılıyorsa mevcut sözleşme, bitiş tarihi, depozito ve kira artış maddeleri; tahliye kısıtları.'),
 ('Apartman yönetmeliği', 'Kat irtifakı/yönetim planı (κανονισμός πολυκατοικίας), aidat borcu yokluğu, planlanan büyük onarım (asansör, çatı, cephe) kararları.'),
 ('Ödeme ve kaynak', 'Kara para mevzuatı kapsamında fonların kaynağı belgesi; ödeme banka yoluyla.'),
 ('4. MÜHENDİSİN KONTROL ETMESİ GEREKENLER', ''),
 ('Statik', 'Özellikle 1985 öncesi binalarda betonarme durumu, kolon/kiriş hasarı, nem, 1978 depremi sonrası güçlendirme.'),
 ('Tesisat', 'Elektrik (yeni tesisat/ΥΔΕ belgesi), sıhhi tesisat, doğalgaz bağlantısı, ısı pompası/kombi.'),
 ('Enerji sertifikası', 'ΠΕΑ (enerji performans sertifikası) kira ilanı için zorunlu; sınıf iyileştirme maliyeti.'),
 ('Ortak alanlar', 'Asansörün ruhsat/bakım durumu, çatı yalıtımı, cephe, merdiven, kapı güvenliği.'),
 ('Ölçüm', 'Net ve brüt m² ölçümü; ilan m²’si genellikle brüttür. Ölçüm farkı fiyat pazarlığında kullanılır.'),
]
r = 4
for a_, b_ in TXT:
    if not b_:
        r += 1; c = V.cell(r, 1, a_); c.font = Font(name=F, size=11, bold=True, color='FFFFFF'); c.fill = fillH; V.cell(r, 2).fill = fillH
    else:
        put(V, r, 1, a_, font=Font(name=F, size=9, bold=True)); put(V, r, 2, b_, font=fN, align=WRAP)
        V.row_dimensions[r].height = max(18, 14 * (len(b_) // 120 + 1))
    r += 1

# =====================================================================================
# YÖNETİCİ PANELİ
# =====================================================================================
D = ws['panel']
title(D, 'SELANİK 100 DAİRE – YATIRIM ÖZETİ (Uzun dönem kiralama)', 'Merkez ve merkeze yakın, yenilenmiş 2+1 daireler · Veri kontrol tarihi 17.11.2026 · Tüm göstergeler formülle hesaplanır ve senaryo/parametre değişikliğiyle güncellenir', 14)
for cc, w in zip('ABCDEFGHIJK', [26, 15, 26, 15, 26, 15, 26, 15, 11, 11, 4]):
    D.column_dimensions[cc].width = w
put(D, 3, 1, 'Aktif senaryo:', font=fB, border=False); put(D, 3, 2, "='" + SN['scen'] + "'!F15", font=Font(name=F, size=11, bold=True, color=GREEN_D), border=False)
put(D, 3, 4, 'Alıcı profili:', font=fB, border=False); put(D, 3, 5, '=IF(P_MUKIM=2,"AB dışı mukim (%15,45)","AB/AEA mukim (%3,09)")', font=fN, border=False)
MRN = lambda k: f"{q('main')}${C[k]}${FR}:${C[k]}${LR}"
KPIS = [
 ('Araştırılan daire sayısı', f"=COUNTA({MRN('no')})", NUM), ('Ortalama satış fiyatı', f"=AVERAGE({MRN('price')})", EUR), ('Ortalama m² fiyatı', f"=AVERAGE({MRN('ppsm')})", EUR), ('Ortalama aylık kira (aktif)', f"=AVERAGE({MRN('ra')})", EUR),
 ('Ortalama brüt getiri', f"=AVERAGE({MRN('gy')})", PCT), ('Ortalama NET getiri', f"=AVERAGE({MRN('ny')})", PCT), ('En yüksek net getiri', f"=MAX({MRN('ny')})", PCT),
 ('En yüksek net getirili daire', f"=INDEX({MRN('no')},MATCH(MAX({MRN('ny')}),{MRN('ny')},0))", None),
 ('Medyan toplam yatırım', f"=MEDIAN({MRN('tinv')})", EUR), ('Ortalama 10Y IRR', f"='{SN['cf']}'!C33", PCT), ('Kira emsal satırı', f"=COUNTA({q('comp')}A6:A{KL})", NUM),
 ('Yüksek riskli daire', f"=COUNTIF({MRN('risklvl')},\"Yüksek\")", NUM),
]
kfill = PatternFill('solid', fgColor='DDEBF7')
for i, (lab, f_, fmt) in enumerate(KPIS):
    row = 5 + (i // 4) * 3; colx = 1 + (i % 4) * 2
    c = D.cell(row, colx, lab); c.font = Font(name=F, size=8, bold=True, color='595959'); c.fill = kfill
    D.cell(row, colx + 1).fill = kfill
    c = D.cell(row + 1, colx, f_); c.font = Font(name=F, size=15, bold=True, color=NAVY); c.fill = kfill
    if fmt: c.number_format = fmt
    D.cell(row + 1, colx + 1).fill = kfill
    D.merge_cells(start_row=row + 1, start_column=colx, end_row=row + 1, end_column=colx + 1)
# Top 10 table
r0 = 15
put(D, r0, 1, 'EN İYİ 10 FIRSAT (genel puana göre)', font=Font(name=F, size=11, bold=True, color=NAVY), border=False)
for i, h in enumerate(['Sıra', 'Daire No', 'Bölge', 'Fiyat (€)', 'Aylık kira (€)', 'Toplam yatırım (€)', 'Net getiri', 'Risk', 'Puan', '10Y IRR'], 1):
    hdr(D, r0 + 1, i, h)
for k in range(10):
    r = r0 + 2 + k; tr = 5 + k
    put(D, r, 1, k + 1, align=CENTER, font=fN)
    for j, (col_, fmt) in enumerate([('B', None), ('D', None), ('F', EUR), ('I', EUR), ('J', EUR), ('M', PCT), ('N', NUM), ('O', NUM1), ('P', PCT)]):
        put(D, r, 2 + j, f"={q('top')}{col_}{tr}", fmt=fmt, font=fX)
D.conditional_formatting.add(f'G{r0 + 2}:G{r0 + 11}', ColorScaleRule(start_type='min', start_color='FFEB84', end_type='max', end_color='63BE7B'))
# regional table (links to MAHALLE)
r1 = r0 + 14
put(D, r1, 1, 'BÖLGESEL KARŞILAŞTIRMA', font=Font(name=F, size=11, bold=True, color=NAVY), border=False)
for i, h in enumerate(['Bölge grubu', 'Daire', 'Ort. €/m² (satış)', 'Ort. kira €/m²', 'Brüt getiri', 'Kira çarpanı', 'Ort. NET getiri', 'Piyasa brüt getirisi', 'Talep'], 1):
    hdr(D, r1 + 1, i, h)
for k in range(len(GRPS)):
    r = r1 + 2 + k; hr = 6 + k
    for j, (col_, fmt) in enumerate([('A', None), ('B', NUM), ('D', EUR), ('E', EUR2), ('F', PCT), ('G', NUM1), ('H', PCT), ('M', PCT), ('O', NUM)]):
        put(D, r, 1 + j, f"={q('mah')}{col_}{hr}", fmt=fmt, font=fX)
RE = r1 + 1 + len(GRPS)
D.conditional_formatting.add(f'G{r1 + 2}:G{RE}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
# scenario summary
r2 = RE + 2
put(D, r2, 1, 'SENARYO ÖZETİ – portföy ortalama net getiri', font=Font(name=F, size=11, bold=True, color=NAVY), border=False)
for j, (lab, nm) in enumerate([('Kötümser', 'SCN_AVG_K'), ('Baz', 'SCN_AVG_B'), ('İyimser', 'SCN_AVG_I')]):
    hdr(D, r2 + 1, 1 + j, lab); put(D, r2 + 2, 1 + j, f'={nm}', fmt=PCT, font=Font(name=F, size=11, bold=True, color=NAVY))
# charts
ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Bölgelere göre ortalama NET getiri'; ch.roundedCorners = False
ch.add_data(Reference(D, min_col=7, min_row=r1 + 1, max_row=RE), titles_from_data=True)
ch.set_categories(Reference(D, min_col=1, min_row=r1 + 2, max_row=RE))
ch.y_axis.number_format = '0.0%'; ch.legend = None; ch.height = 9; ch.width = 16
ch.x_axis.delete = False; ch.y_axis.delete = False; ch.title.overlay = False
ch.series[0].graphicalProperties.solidFill = '2E7D32'
D.add_chart(ch, 'L4')
ch2 = BarChart(); ch2.type = 'col'; ch2.title = 'Bölgelere göre ortalama satış €/m²'; ch2.roundedCorners = False
ch2.add_data(Reference(D, min_col=3, min_row=r1 + 1, max_row=RE), titles_from_data=True)
ch2.set_categories(Reference(D, min_col=1, min_row=r1 + 2, max_row=RE))
ch2.legend = None; ch2.height = 9; ch2.width = 16; ch2.x_axis.delete = False; ch2.y_axis.delete = False; ch2.title.overlay = False
ch2.series[0].graphicalProperties.solidFill = NAVY
D.add_chart(ch2, 'L23')
sc = ScatterChart(); sc.title = 'Satış fiyatı (€) – Net getiri (100 daire)'; sc.style = 13; sc.roundedCorners = False
xs = Reference(M, min_col=27, min_row=FR, max_row=LR); ys = Reference(M, min_col=82, min_row=FR, max_row=LR)
se = Series(ys, xs, title='Daireler'); se.marker.symbol = 'circle'; se.marker.size = 6; se.graphicalProperties.line.noFill = True
se.marker.graphicalProperties.solidFill = NAVY2; se.marker.graphicalProperties.line.solidFill = NAVY2
sc.series.append(se); sc.x_axis.title = 'Satış fiyatı (€)'; sc.y_axis.title = 'Net getiri'; sc.y_axis.number_format = '0.0%'; sc.x_axis.number_format = '#,##0'
sc.legend = None; sc.height = 9; sc.width = 16; sc.x_axis.delete = False; sc.y_axis.delete = False; sc.title.overlay = False
sc.x_axis.title.overlay = False; sc.y_axis.title.overlay = False
D.add_chart(sc, 'L42')
ch3 = BarChart(); ch3.type = 'bar'; ch3.title = 'En iyi 10 – genel yatırım puanı'; ch3.roundedCorners = False
ch3.add_data(Reference(D, min_col=9, min_row=r0 + 1, max_row=r0 + 11), titles_from_data=True)
ch3.set_categories(Reference(D, min_col=2, min_row=r0 + 2, max_row=r0 + 11))
ch3.legend = None; ch3.height = 9; ch3.width = 16; ch3.x_axis.delete = False; ch3.y_axis.delete = False; ch3.title.overlay = False
ch3.series[0].graphicalProperties.solidFill = '548235'
D.add_chart(ch3, 'L61')
put(D, r2 + 4, 1, 'Not: Kira tahminleri yakın çevredeki kiralık İLAN fiyatlarına (istenen kira) dayanır; %5 sözleşme indirimi ve boş kalma varsayımı uygulanmıştır. Ayrıntı: KİRA EMSALLERİ ve PARAMETRELER.', font=fSm, border=False)

# =====================================================================================
# ANALİZ SORULARI
# =====================================================================================
A = ws['qa']
title(A, 'ANALİZ SORULARI – Talimatnamedeki 10 soru', 'Tablolar formülle güncellenir. Yazılı yorumlar BAZ senaryo ve 17.11.2026 verisiyle hazırlanmıştır; parametre değişirse metinler güncellenmez.', 8)
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
    c = A.cell(r, 1, t); c.font = fN; c.alignment = WRAP
    A.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    A.row_dimensions[r].height = max(15, 13 * (len(t) // 135 + 1))
    r += 1
notes = NOTES or {}
qhead('1. En yüksek net kira getirisi sağlayan 20 daire hangileri?')
qtext(notes.get('q1', 'EN İYİ 20 YATIRIM sayfasının alt tablosunda (salt net getiriye göre) formülle listelenir.'))
qhead('2. Aynı kira gelirini daha düşük satın alma maliyetiyle sağlayan daireler var mı?')
qtext(notes.get('q2', ''))
qhead('3. Hangi mahallelerde kira çarpanı daha avantajlı?')
hdrs = ['Bölge grubu', 'Kira çarpanı (örneklem)', 'Brüt getiri', 'Ort. net getiri', 'Piyasa kira çarpanı']
for i, h in enumerate(hdrs, 1): hdr(A, r, i, h)
r += 1
for k in range(len(GRPS)):
    hr = 6 + k
    for j, (col_, fmt) in enumerate([('A', None), ('G', NUM1), ('F', PCT), ('H', PCT), ('N', NUM1)]):
        put(A, r, 1 + j, f"={q('mah')}{col_}{hr}", fmt=fmt, font=fX)
    r += 1
qtext(notes.get('q3', ''))
qhead('4. Ana cadde üzerinde olmanın satın alma primini kira geliri karşılıyor mu?')
for i, h in enumerate(['Grup', 'Daire sayısı', 'Ort. satış €/m²', 'Ort. emsal kira €/m²', 'Ort. brüt getiri', 'Ort. net getiri'], 1): hdr(A, r, i, h)
r += 1
for lab, cond in [('Ana caddeye ≤100 m', '"<=100"'), ('Ana caddeye >100 m', '">100"')]:
    put(A, r, 1, lab, font=fN)
    put(A, r, 2, f"=COUNTIF({MRN('droad')},{cond})", fmt=NUM)
    for j, (k_, fmt) in enumerate([('ppsm', EUR), ('compavg', EUR2), ('gy', PCT), ('ny', PCT)]):
        put(A, r, 3 + j, f"=AVERAGEIF({MRN('droad')},{cond},{MRN(k_)})", fmt=fmt)
    r += 1
put(A, r, 1, 'Fiyat primi / kira primi', font=fB); put(A, r, 3, f'=C{r - 2}/C{r - 1}-1', fmt=PCT1); put(A, r, 4, f'=D{r - 2}/D{r - 1}-1', fmt=PCT1); r += 1
qtext(notes.get('q4', ''))
qhead('5. Hangi daireler gereğinden pahalı?')
qtext(notes.get('q5', 'Ana listede “Fiyat / emsal oranı” > 1,15 olan hücreler kırmızı ile işaretlidir.'))
qhead('6. Hangi dairelerde pazarlık potansiyeli daha yüksek?')
qtext(notes.get('q6', ''))
qhead('7. Hangi binalarda teknik ve hukuki riskler daha yüksek?')
qtext(notes.get('q7', ''))
qhead('8. Düzenli kira geliri için hangi 10 daire tercih edilmeli?')
qtext(notes.get('q8', ''))
qhead('9. Aynı sermayeyle bir büyük daire mi, birkaç küçük 2+1 mi?')
for i, h in enumerate(['m² bandı', 'Daire sayısı', 'Ort. fiyat (€)', 'Ort. toplam yatırım (€)', 'Ort. aylık kira (€)', 'Ort. net getiri', 'Ort. risk puanı'], 1): hdr(A, r, i, h)
r += 1
for lab, lo, hi in [('45–60 m²', 45, 60), ('60–75 m²', 60.01, 75), ('75–90 m²', 75.01, 90), ('90–110 m²', 90.01, 110)]:
    put(A, r, 1, lab, font=fN)
    cnd = f'{MRN("sqm")},">={lo}",{MRN("sqm")},"<={hi}"'
    put(A, r, 2, f'=COUNTIFS({cnd})', fmt=NUM)
    for j, (k_, fmt) in enumerate([('price', EUR), ('tinv', EUR), ('ra', EUR), ('ny', PCT), ('risk', NUM)]):
        put(A, r, 3 + j, f'=IFERROR(AVERAGEIFS({MRN(k_)},{cnd}),"-")', fmt=fmt)
    r += 1
qtext(notes.get('q9', ''))
qhead('10. Beş ve on yıllık elde tutma senaryolarında toplam getiri nasıl değişiyor?')
for i, h in enumerate(['Gösterge', 'Değer'], 1): hdr(A, r, i, h)
r += 1
for lab, f_, fmt in [('Ortalama 5 yıllık IRR', f"=AVERAGE({q('cf')}U{G0 + 1}:U{GL_})", PCT), ('Ortalama 10 yıllık IRR', f"=AVERAGE({q('cf')}V{G0 + 1}:V{GL_})", PCT),
                     ('Ortalama 5 yıllık toplam getiri', f"=AVERAGE({q('cf')}W{G0 + 1}:W{GL_})", PCT1), ('Ortalama 10 yıllık toplam getiri', f"=AVERAGE({q('cf')}X{G0 + 1}:X{GL_})", PCT1)]:
    put(A, r, 1, lab, font=fN); put(A, r, 2, f_, fmt=fmt); r += 1
qtext(notes.get('q10', ''))

# print setup
for sh in wb.worksheets:
    sh.sheet_properties.pageSetUpPr.fitToPage = True
    sh.page_setup.orientation = 'landscape'; sh.page_setup.fitToWidth = 1; sh.page_setup.fitToHeight = 0
for k, color in [('panel', NAVY), ('main', NAVY2), ('top', '548235'), ('par', 'BF8F00'), ('scen', 'BF8F00'), ('risk', 'C00000')]:
    ws[k].sheet_properties.tabColor = color
wb.save(OUT)
print('saved', OUT)
