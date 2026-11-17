import json, re, math, collections, statistics, datetime
SALE = json.load(open('sale_det.json')); RENT = json.load(open('rent_det.json'))
SL = {**json.load(open('sale_list.json')), **json.load(open('sale_res.json'))}
RL = {**json.load(open('rent_list.json')), **json.load(open('rent_kadikoy.json')), **json.load(open('rent_res.json'))}
MST = json.load(open('mahalle_stats.json'))
CHECK = datetime.date(2026, 11, 17)
LUX = {'Caferağa': 'caferaga', 'Kozyatağı': 'kozyatagi', 'Caddebostan': 'caddebostan', 'Suadiye': 'suadiye', 'Fenerbahçe': 'fenerbahce',
       'Feneryolu': 'feneryolu', 'Göztepe': 'goztepe', 'Erenköy': 'erenkoy', 'Zühtüpaşa': 'zuhtupasa', 'Koşuyolu': 'kosuyolu'}
def hav(a, b, c, d):
    R = 6371000; p1, p2 = math.radians(a), math.radians(c); dp = p2 - p1; dl = math.radians(d - b)
    return 2 * R * math.asin(math.sqrt(math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2))
def segd(p, a, b):
    ky = 111320; kx = 111320 * math.cos(math.radians(p[0]))
    ax, ay = (a[1] - p[1]) * kx, (a[0] - p[0]) * ky; bx, by = (b[1] - p[1]) * kx, (b[0] - p[0]) * ky
    dx, dy = bx - ax, by - ay; t = 0 if dx == dy == 0 else max(0, min(1, -(ax * dx + ay * dy) / (dx * dx + dy * dy)))
    return math.hypot(ax + t * dx, ay + t * dy)
def pline(p, pts): return min(segd(p, a, b) for a, b in zip(pts, pts[1:]))
RAIL = {'Marmaray Söğütlüçeşme': (40.9906, 29.0377), 'Marmaray Feneryolu': (40.9787, 29.0488), 'Marmaray Göztepe': (40.9790, 29.0666),
        'Marmaray Erenköy (yaklaşık)': (40.9700, 29.0750), 'Marmaray Suadiye': (40.9607, 29.0833), 'Marmaray Bostancı': (40.95389, 29.09522),
        'M4 Kadıköy (yaklaşık)': (40.9910, 29.0245), 'M4 Ayrılık Çeşmesi': (40.9999, 29.02985), 'M4 Acıbadem': (41.00155, 29.04559), 'M4 Ünalan': (40.99725, 29.05952),
        'M4 Göztepe': (40.9936, 29.0703), 'M4 Yenisahra': (40.9845, 29.0902), 'M4 Kozyatağı': (40.9745, 29.1000), 'M4 Bostancı': (40.9646, 29.1048)}
BAGDAT = [(40.9840, 29.0410), (40.9765, 29.0505), (40.9712, 29.0615), (40.9650, 29.0720), (40.9585, 29.0830), (40.9545, 29.0950)]
SAHIL = [(40.9850, 29.0255), (40.9790, 29.0300), (40.9700, 29.0330), (40.9640, 29.0370), (40.9690, 29.0450), (40.9660, 29.0600), (40.9600, 29.0700), (40.9545, 29.0820), (40.9505, 29.0950)]
CENTER = (40.9925, 29.0235)  # Kadıköy iskele / çarşı
def num(s):
    if s is None: return None
    m = re.search(r'([\d.]+)', str(s).replace('.', '')) if False else re.search(r'(\d[\d.]*)', str(s))
    return float(m.group(1).replace('.', '')) if m else None
def age_info(s):
    s = s or ''
    if 'Yapım' in s: return ('construction', 0, 0)
    if 'Oturuma' in s: return ('ok', 0, 0)
    m = re.match(r'(\d+)-(\d+)', s)
    if m: return ('ok', int(m.group(1)), int(m.group(2)))
    m = re.match(r'(\d+)', s)
    if m and 'Üzeri' in s: return ('old', int(m.group(1)), 99)
    if m: return ('ok', int(m.group(1)), int(m.group(1)))
    return ('unknown', None, None)
def floornum(s):
    s = s or ''
    for k, v in {'Bahçe': 0, 'Giriş': 0, 'Yüksek Giriş': 0.5, 'Zemin': 0, 'Kot': -1, 'Bodrum': -1, 'Çatı': 9, 'Villa': 0, 'Ara Kat': 1}.items():
        if k in s: return v
    m = re.search(r'(-?\d+)', s); return int(m.group(1)) if m else None
def mahalle(v):
    m = v.get('_mahalle') or ''
    for k in LUX:
        if k in m: return k
    return None
def mk(s):
    return s
def parse_date(s):
    if not s: return None
    return datetime.date.fromisoformat(s[:10])
DESC_TEN = re.compile(r'kiracılı|kiracı oturuyor|kiracı mevcut|kira getirili|kirada', re.I)
records = []
for u, v in SALE.items():
    if not v or '_error' in v: continue
    mah = mahalle(v)
    if not mah: continue
    room = v.get('Oda Sayısı')
    if room not in ('1+1', '2+1'): continue
    st, a0, a1 = age_info(v.get('Bina Yaşı'))
    brut = num(v.get('Brüt m2')); net = num(v.get('Net m2'))
    price = v.get('_price') or (SL.get(u) or {}).get('price')
    if not price or not v.get('_lat'): continue
    records.append(dict(url=u, v=v, mah=mah, room=room, age_status=st, age_lo=a0, age_hi=a1, brut=brut, net=net, price=float(price),
                        lat=float(v['_lat']), lng=float(v['_lng'])))
print('lux 1+1/2+1 with details', len(records), collections.Counter(r['age_status'] for r in records))
# dedupe
records.sort(key=lambda r: r['price'])
groups = []
for r in records:
    for g in groups:
        m = g[0]
        if m['room'] == r['room'] and abs((m['brut'] or 0) - (r['brut'] or 0)) <= 3 and m['v'].get('Bulunduğu Kat') == r['v'].get('Bulunduğu Kat') and hav(m['lat'], m['lng'], r['lat'], r['lng']) < 200 and abs(m['price'] - r['price']) / m['price'] <= 0.10:
            g.append(r); break
    else:
        groups.append([r])
print('groups', len(groups))
main, other = [], []
for g in groups:
    r = g[0]; r['dups'] = [dict(url=x['url'], price=x['price'], adv=x['v'].get('_agency')) for x in g[1:]]
    reasons = []
    if r['age_status'] == 'construction': reasons.append('Yapım aşamasında (kiraya hazır değil)')
    if r['age_status'] in ('old', 'unknown') or (r['age_hi'] or 0) > 16: reasons.append('Bina 2010 öncesi / yaş bilinmiyor')
    if not r['brut'] or r['brut'] < 35 or r['brut'] > 160: reasons.append('m² aralık dışı')
    if re.search(r'Dubai|Kıbrıs|yurt ?dışı|Antalya', r['v'].get('_desc') or '', re.I): reasons.append('İlan metni başka bir proje/ülkeyi tanıtıyor (güvenilmez ilan)')
    upd = parse_date(r['v'].get('_updated'))
    r['why_not'] = reasons
    (other if reasons else main).append(r)
print('main', len(main), collections.Counter(r['mah'] for r in main), collections.Counter(r['room'] for r in main))
print('other', len(other), collections.Counter(tuple(r['why_not']) for r in other))
# rent pool
rp = []
for u, v in RENT.items():
    if not v or '_error' in v or not v.get('_lat'): continue
    room = v.get('Oda Sayısı'); brut = num(v.get('Brüt m2')); price = v.get('_price') or (RL.get(u) or {}).get('price')
    if not price or not brut: continue
    pp = float(price) / brut
    if not (250 <= pp <= 2500): continue
    st, a0, a1 = age_info(v.get('Bina Yaşı'))
    x = dict(url=u, room=room, brut=brut, net=num(v.get('Net m2')), rent=float(price), lat=float(v['_lat']), lng=float(v['_lng']), mah=(v.get('_mahalle') or '').replace('Kadıköy ', '').replace(' Mahallesi', ''),
             floor=v.get('Bulunduğu Kat'), age=v.get('Bina Yaşı'), age_hi=a1, esya=v.get('Eşya Durumu'), site=v.get('Site İçerisinde'), upd=(v.get('_updated') or '')[:10], adv=v.get('_agency'))
    if any(abs(y['brut'] - x['brut']) <= 2 and y['rent'] == x['rent'] and hav(y['lat'], y['lng'], x['lat'], x['lng']) < 200 for y in rp): continue
    rp.append(x)
print('rent pool', len(rp), collections.Counter(x['room'] for x in rp), collections.Counter(x['mah'] for x in rp).most_common(12))
FEAT_KEYS = ['Marmaray', 'Metro', 'Dolmuş', 'Minibüs', 'Otobüs Durağı', 'Deniz Manzaralı', 'Şehir Manzaralı', 'Denize Sıfır', 'Sahil', 'Alışveriş Merkezi', 'Hastane', 'Okula Yakın', 'Üniversite', 'Ana Yol', 'Cadde', 'E-5', 'Belediye', 'Kapalı Otopark', 'Güvenlik', 'Jeneratör', 'Isı Yalıtımı', 'Ses Yalıtımı', 'Klima', 'Isıcam', 'Ebeveyn Banyosu']
out = []
for n, r in enumerate(sorted(main, key=lambda r: (list(LUX).index(r['mah']), r['price'])), 1):
    v = r['v']; p = (r['lat'], r['lng'])
    rail = min((hav(*p, *c), k) for k, c in RAIL.items())
    desc = v.get('_desc') or ''
    bag_txt = None
    m = re.search(r'Bağdat\s*Cad\w*[^.]{0,40}?(\d{2,4})\s*(m|mt|metre)', desc, re.I)
    if m: bag_txt = int(m.group(1))
    elif re.search(r'Bağdat\s*Cad\w*\s*(üzeri|üzerinde|cepheli)', desc, re.I): bag_txt = 0
    feats = [k for k in FEAT_KEYS if k in v]
    # comps
    cs = []
    for x in rp:
        d = hav(*p, x['lat'], x['lng'])
        if d > 2500 or abs(x['brut'] - r['brut']) / r['brut'] > 0.35: continue
        same_room = x['room'] == r['room'] or (r['room'] == '2+1' and x['room'] in ('2.5+1',)) or (r['room'] == '1+1' and x['room'] in ('1.5+1',))
        if not same_room: continue
        newb = (x['age_hi'] is not None and x['age_hi'] <= 16)
        cs.append((d / 500 + abs(x['brut'] - r['brut']) / r['brut'] * 3 + (0 if newb else 0.8), d, x))
    cs.sort(key=lambda t: t[0])
    pick = cs[:8]
    if len(pick) >= 4:
        med = statistics.median(t[2]['rent'] / t[2]['brut'] for t in pick)
        pick = [t for t in pick if 0.65 * med <= t[2]['rent'] / t[2]['brut'] <= 1.5 * med]
    pick = pick[:6]
    ten = v.get('Kullanım Durumu')
    rent_now = None
    if ten == 'Kiracı Oturuyor':
        m2 = re.search(r'(?:kira(?:sı|cı)?[^.]{0,30}?|getiri[^.]{0,20}?)(\d{2,3}[.,]?\d{3})\s*(?:TL|₺)', desc, re.I)
        if m2:
            val = float(m2.group(1).replace('.', '').replace(',', ''))
            if 10000 <= val <= 400000: rent_now = val
    rc = None
    for m3 in re.finditer(r'kira', desc, re.I):
        y3 = re.search(r'(\d{2,3}[.,]\d{3}|\d{2,3}\s*bin)\s*(TL|₺)?', desc[m3.start():m3.start()+80], re.I)
        if y3:
            t3 = y3.group(1).lower().replace(' ', '')
            val = float(t3.replace('bin', '')) * 1000 if 'bin' in t3 else float(t3.replace('.', '').replace(',', ''))
            if 15000 <= val <= 300000: rc = val; break
    rec = dict(no=f'K-{n:03d}', url=r['url'], title=(v.get('_header') or [''])[0], adv=v.get('_agency') or 'Bilinmiyor', agent=next((h for h in (v.get('_header') or []) if '·' in h), ''),
               ilan_no=v.get('İlan Numarası'), created=(v.get('_created') or '')[:10], updated=(v.get('_updated') or '')[:10], mah=r['mah'], lat=r['lat'], lng=r['lng'],
               dcenter=round(hav(*p, *CENTER)), rail=rail[1], drail=round(rail[0]), dbagdat=round(pline(p, BAGDAT)), bag_txt=bag_txt, dsahil=round(pline(p, SAHIL)),
               price=r['price'], brut=r['brut'], net=r['net'], room=r['room'], baths=v.get('Banyo Sayısı'), age=v.get('Bina Yaşı'), age_hi=r['age_hi'],
               floor=v.get('Bulunduğu Kat'), floorn=floornum(v.get('Bulunduğu Kat')), tf=v.get('Kat Sayısı'), heat=v.get('Isıtma Tipi'), park=v.get('Otopark'), usage=ten,
               kredi=v.get('Krediye Uygunluk'), site=v.get('Site İçerisinde'), tapu=v.get('Tapu Durumu'), elev=v.get('Asansör'), esya=v.get('Eşya Durumu'), balkon=v.get('Balkon'),
               aidat=num(v.get('Aidat')) if v.get('Aidat') else None, feats=feats, desc=desc[:1500], images=v.get('_images') or [], dups=r['dups'], rent_now=rent_now,
               ncomp=len(pick), rent_claim=rc)
    out.append(rec)
    for k, (sc, d, x) in enumerate(pick, 1):
        rec.setdefault('comps', []).append(dict(k=k, url=x['url'], mah=x['mah'], room=x['room'], brut=x['brut'], net=x['net'], floor=x['floor'], age=x['age'], esya=x['esya'], site=x['site'], rent=x['rent'], dist=round(d), upd=x['upd'], adv=x['adv']))
# emsal satış ₺/m²: aynı mahalle + aynı oda tipi, ≤15 yaş (ana liste + yeni yapım aşamasındaki yedekler hariç), kendisi hariç medyan
pool=[(x['mah'],x['room'],x['price']/x['brut'],x['url']) for x in main if x['brut']]
for rec in out:
    vals=[pp for (m_,ro,pp,u) in pool if m_==rec['mah'] and ro==rec['room'] and u!=rec['url']]
    if len(vals)<3: vals=[pp for (m_,ro,pp,u) in pool if m_==rec['mah'] and u!=rec['url']]
    rec['exit_ppsm']=round(statistics.median(vals)) if len(vals)>=3 else None; rec['exit_n']=len(vals) if len(vals)>=3 else None
print('final', len(out), collections.Counter(r['ncomp'] for r in out))
print('comps dist median', statistics.median(c['dist'] for r in out for c in r.get('comps', [])))
extras = []
for r in other:
    v = r['v']
    extras.append(dict(url=r['url'], mah=r['mah'], room=r['room'], price=r['price'], brut=r['brut'], age=v.get('Bina Yaşı'), floor=v.get('Bulunduğu Kat'), adv=v.get('_agency'), upd=(v.get('_updated') or '')[:10], why='; '.join(r['why_not'])))
market = {}
for k, slug in LUX.items():
    s = MST.get(f'{slug}|satilik-daire', {}); rr = MST.get(f'{slug}|kiralik-daire', {})
    f = lambda x: float(x.replace('.', '')) if x else None
    market[k] = dict(sale_m2=f(s.get('m2')), amort=f(s.get('amort')), avg_rent=f(s.get('kira')), avg_price=f(s.get('ort_fiyat')), rent_m2=f(rr.get('m2')), exp=f(s.get('beklenti')), n_sale=f(s.get('n')), n_rent=f(rr.get('n')))
json.dump(dict(recs=out, extras=extras, market=market, rentpool=len(rp)), open('kdk_model.json', 'w'), ensure_ascii=False)
