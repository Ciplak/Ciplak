import re, html, json, sys, time, os, urllib.request
UA = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
def toks_of(h):
    t = re.sub(r'<script.*?</script>|<style.*?</style>', '', h, flags=re.S)
    return [x for x in (html.unescape(y).strip() for y in re.split(r'<[^>]+>', t)) if x]
def after(toks, label, n=1):
    for i, x in enumerate(toks):
        if x == label or x == label + ':':
            return toks[i + n] if i + n < len(toks) else None
    return None
def num(s):
    if s is None: return None
    m = re.search(r'(\d[\d.]*)', str(s)); return float(m.group(1).replace('.', '')) if m else None
def parse_c21(h, url):
    T = toks_of(h)
    loc = []
    if 'Konum' in T:
        i = T.index('Konum'); loc = [x.strip(' ,') for x in T[i + 1:i + 6]]
    title = next((x for x in T if x.isupper() and len(x) > 25), '')
    m = re.search(r'kadikoy-([a-z0-9-]+)-(satilik|kiralik)/', url)
    desc = ''
    if 'Devamını Oku' in T:
        j = T.index('Devamını Oku'); desc = ' '.join(T[max(0, j - 40):j])
    return dict(source='Century 21', url=url, deal=m.group(2) if m else None, mah_slug=m.group(1) if m else None,
                title=title, price=num(after(T, 'Fiyat')), brut=num(after(T, 'Metre Kare (Brüt)')), net=num(after(T, 'Metre Kare (Net)')),
                room=after(T, 'Oda Sayısı'), age=after(T, 'Bina Yaşı'), floor=after(T, 'Bulunduğu Kat'), tf=after(T, 'Kat Sayısı'),
                tapu=after(T, 'Tapu Durumu'), esya=after(T, 'Eşyalı'), usage=after(T, 'Kullanım Durumu'), site=after(T, 'Site İçerisinde'),
                heat=after(T, 'Isıtma'), kredi=None, aidat=num(after(T, 'Aidat')), updated=None, lat=None, lng=None,
                elev='Var' if 'Asansör' in T else 'Belirtilmemiş', park='Kapalı Otopark' if 'Kapalı Otopark' in T else ('Açık Otopark' if 'Açık Otopark' in T else 'Belirtilmemiş'),
                office=None, desc=desc[:1500])
def parse_ty(h, url):
    T = toks_of(h)
    loc = next((x for x in T if x.startswith('İstanbul - Kadıköy')), '')
    upd = next((x for x in T if x.startswith('Son Güncelleme')), '')
    m = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{4})', upd)
    lat = re.search(r'(4[01]\.\d{5,})', h); lng = re.search(r'(2[89]\.\d{5,})', h)
    j = T.index('Portföy Açıklaması') if 'Portföy Açıklaması' in T else None
    desc = ' '.join(T[j + 1:j + 40]) if j else ''
    title = T[T.index('WhatsApp') - 1] if False else next((x for x in T if 'TURYAP' in x.upper() and len(x) > 20), '')
    return dict(source='Turyap', url=url, deal='satilik' if after(T, 'Durumu') == 'Satılık' else 'kiralik', mah_slug=None, mah=loc.split(' - ')[-1] if loc else None,
                title=title, price=num(after(T, 'Fiyat')), brut=num(after(T, 'Brüt Alan')), net=num(after(T, 'Net Alan')), room=after(T, 'Oda Sayısı'),
                age=after(T, 'Bina Yaşı'), floor=after(T, 'Bulunduğu Kat'), tf=after(T, 'Bina Kat Sayısı'), tapu=after(T, 'Tapu'), esya=None,
                usage=None, site=None, heat=after(T, 'Isıtma'), kredi=after(T, 'Uygunluk'), aidat=num(after(T, 'Aidat')),
                updated=f'{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}' if m else None,
                lat=float(lat.group(1)) if lat else None, lng=float(lng.group(1)) if lng else None,
                elev='Var' if 'Asansör' in T else 'Belirtilmemiş', park='Kapalı Otopark' if any('Kapalı Otopark' in x for x in T) else 'Belirtilmemiş',
                office=next((x for x in T if x.upper().startswith('TURYAP ') and len(x) < 40), None), desc=desc[:1500])
if __name__ == '__main__':
    urls = json.load(open(sys.argv[1])); out = sys.argv[2]
    have = json.load(open(out)) if os.path.exists(out) else {}
    for u in urls:
        if u in have: continue
        try:
            h = urllib.request.urlopen(urllib.request.Request(u, headers={'User-Agent': UA}), timeout=40).read().decode('utf-8', 'ignore')
            have[u] = parse_c21(h, u) if 'century21' in u else parse_ty(h, u)
        except Exception as e:
            print('err', u, e, flush=True)
        time.sleep(1.0)
    json.dump(have, open(out, 'w'), ensure_ascii=False); print('done', len(have))
