import json,re,math,collections,statistics
D=json.load(open('alldet.json')); SEL=json.load(open('selection.json'))
RENT=json.load(open('rent.json')); SALE=json.load(open('sale_ren.json'))
def hav(a,b,c,d):
    R=6371000;p1,p2=math.radians(a),math.radians(c);dp=p2-p1;dl=math.radians(d-b)
    return 2*R*math.asin(math.sqrt(math.sin(dp/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2))
def segd(p,a,b):
    # local equirectangular
    ky=111320;kx=111320*math.cos(math.radians(p[0]))
    ax,ay=(a[1]-p[1])*kx,(a[0]-p[0])*ky; bx,by=(b[1]-p[1])*kx,(b[0]-p[0])*ky
    dx,dy=bx-ax,by-ay; t=0 if dx==dy==0 else max(0,min(1,-(ax*dx+ay*dy)/(dx*dx+dy*dy)))
    return math.hypot(ax+t*dx,ay+t*dy)
METRO={'Neos Sidirodromikos Stathmos':(40.64361,22.92917),'Dimokratias':(40.64111,22.93417),'Venizelou':(40.63694,22.94194),'Agias Sofias':(40.63444,22.94639),'Sintrivani':(40.63056,22.95417),'Panepistimio':(40.62611,22.96000),'Papafi':(40.6197,22.9625),'Efkleidis':(40.61611,22.96028),'Fleming':(40.60917,22.95722),'Analipseos':(40.60556,22.95778)}
ROADS={'Egnatia':[(40.6411,22.9342),(40.63694,22.94194),(40.63444,22.94639),(40.6323,22.9516),(40.63056,22.95417)],
'Tsimiski':[(40.6373,22.9355),(40.6341,22.9415),(40.6316,22.9465),(40.6286,22.9512)],
'Mitropoleos':[(40.6362,22.9352),(40.6332,22.9408),(40.6303,22.9460),(40.6278,22.9500)],
'Nikis (sahil)':[(40.6355,22.9345),(40.6326,22.9403),(40.6290,22.9460),(40.6265,22.9484)],
'Agiou Dimitriou':[(40.6425,22.9395),(40.6390,22.9475),(40.6365,22.9530),(40.6345,22.9580)],
'Ethnikis Aminis':[(40.6268,22.9484),(40.63056,22.95417)],
'Vas. Olgas':[(40.6262,22.9520),(40.6180,22.9545),(40.6090,22.9550),(40.6000,22.9560)],
'Lagkada':[(40.6430,22.9320),(40.6500,22.9290)],
'Monastiriou / 26is Oktovriou':[(40.6420,22.9310),(40.6435,22.9230)],
'Papafi (metro ekseni)':[(40.62611,22.96000),(40.6197,22.9625),(40.61611,22.96028),(40.60917,22.95722),(40.60556,22.95778)]}
UNIS={'AUTH (Aristoteles Üniv.)':[(40.6300,22.9575),(40.6330,22.9600),(40.6285,22.9555)],'Makedonya Üniv.':[(40.6253,22.9601)]}
CENTER=(40.6326,22.9406)
TR={'Αριστοτέλους - Ιστορικό Κέντρο':'Aristotelous / Tarihi Merkez','Ρωμαϊκή Αγορά - Βενιζέλου - Ιστορικό Κέντρο':'Roma Agorası - Venizelou','Κέντρο Θεσσαλονίκης':'Merkez (Tsimiski-Egnatia)','Κέντρο':'Merkez','Λαδάδικα':'Ladadika','Πλατεία Εμπορίου - Ανω Λαδάδικα - Ιστορικό Κέντρο':'Ladadika / Emboriou','Αγία Σοφία':'Agia Sofia','Λευκός Πύργος':'Beyaz Kule','Καμάρα - Ροτόντα':'Kamara - Rotonda','Άγιος Δημήτριος':'Agios Dimitrios','Βαρδάρης - Δικαστήρια':'Vardaris - Adliye','Διοικητήριο':'Dioikitirio','Αντιγονιδών':'Antigonidon','Σταθμός ΟΣΕ':'OSE Tren İstasyonu','Πλατεία Ναυαρίνου':'Navarinou Meydanı','ΔΕΘ':'Fuar (HELEXPO)','ΧΑΝΘ':'XANTH (Beyaz Kule)','Τουρκικό Προξενείο':'Türk Konsolosluğu çevresi','Φάληρο - Ιπποκράτειο':'Faliro','Ιπποκράτειο':'Faliro - Ippokrateio','Αγία Τριάδα':'Faliro - Agia Triada'}
GROUPS={'Aristotelous / Tarihi Merkez':'Tsimiski / Aristotelous / Tarihi Merkez','Roma Agorası - Venizelou':'Tsimiski / Aristotelous / Tarihi Merkez','Merkez (Tsimiski-Egnatia)':'Tsimiski / Aristotelous / Tarihi Merkez','Merkez':'Tsimiski / Aristotelous / Tarihi Merkez','Ladadika':'Tsimiski / Aristotelous / Tarihi Merkez','Ladadika / Emboriou':'Tsimiski / Aristotelous / Tarihi Merkez','Agia Sofia':'Agia Sofia','Beyaz Kule':'Beyaz Kule / XANTH / Fuar','XANTH (Beyaz Kule)':'Beyaz Kule / XANTH / Fuar','Fuar (HELEXPO)':'Beyaz Kule / XANTH / Fuar','Kamara - Rotonda':'Kamara - Rotonda','Navarinou Meydanı':'Kamara - Rotonda','Agios Dimitrios':'Agios Dimitrios','Türk Konsolosluğu çevresi':'Agios Dimitrios','Vardaris - Adliye':'Vardaris','Dioikitirio':'Dioikitirio','Antigonidon':'Antigonidon','OSE Tren İstasyonu':'OSE Tren İstasyonu','Faliro':'Faliro','Faliro - Ippokrateio':'Faliro','Faliro - Agia Triada':'Faliro'}
def meta(l,feat):
    pass
FLOOR={'Basement':-1,'Semi-basement':-0.5,'Ground':0,'Mezzanine':0.5}
def floornum(f):
    if f is None: return None
    if f in FLOOR: return FLOOR[f]
    m=re.match(r'(\d+)',str(f)); return int(m.group(1)) if m else None
# rent pool dedupe
rp=[]
for r in RENT:
    if not r.get('lat') or not r['price'] or not r['sqm']: continue
    pp=r['price']/r['sqm']
    if not (350<=r['price']<=1800 and 5<=pp<=22): continue
    if (r['sourceModified'] or '')<'2026-06': continue
    if any(abs(x['sqm']-r['sqm'])<=2 and x['price']==r['price'] and hav(x['lat'],x['lng'],r['lat'],r['lng'])<200 for x in rp): continue
    rp.append(r)
print('rent pool',len(rp))
recs=[];comps=[]
for n,s in enumerate(SEL['sel'],1):
    v=D[str(s['id'])];l=v['listing'];f={x['name']:(x['present'],x['detail']) for x in v['features']}
    desc=re.sub(r'\s+',' ',l['description'] or '')
    p=(l['lat'],l['lng'])
    md=min(((hav(*p,*c),k) for k,c in METRO.items()))
    rd=min(((min(segd(p,a,b) for a,b in zip(pts,pts[1:])) if len(pts)>1 else hav(*p,*pts[0]),k) for k,pts in ROADS.items()))
    ud=min(((min(hav(*p,*q) for q in pts),k) for k,pts in UNIS.items()))
    net=None
    m=re.search(r'(\d{2,3}(?:[.,]\d+)?)\s*τ\.?\s*μ\.?\s*\(?\s*(?:καθαρ|ωφέλιμ)',desc) or re.search(r'καθαρ\w*\s*(?:επιφάνειας|εμβαδού)?\s*(\d{2,3}(?:[.,]\d+)?)',desc)
    if m:
        try:
            net=float(m.group(1).replace(',','.'))
            if not (25<net<=l['sqm']): net=None
        except: net=None
    scope='Tam kapsamlı yenileme' if re.search(r'πλήρως|ριζικά|ολικώς|ολική ανακαίνιση|εκ βάθρων|fully',desc,re.I) else ('Kısmi yenileme' if re.search(r'μερικώς|partially',desc,re.I) else 'Yenilenmiş (kapsam belirtilmemiş)')
    if s['under']: scope+=' – ilan: tadilat devam ediyor'
    ex=l.get('exit') or {}
    ry=l['renovationYear']; ry_src='portal'
    mr=None
    for mm_ in re.finditer(r'ανακα[ιί]ν',desc,re.I):
        yy_=re.search(r'(20[0-2]\d)',desc[mm_.start():mm_.start()+70])
        if yy_: mr=yy_;break
    if mr and 2000<=int(mr.group(1))<=2026:
        if int(mr.group(1))!=ry: ry_src='ilan metni (portal: %s)'%ry
        else: ry_src='portal + ilan metni'
        ry=int(mr.group(1))
    adv=l['advertiser'] or {}
    tenant_txt='Belirtilmemiş'; rent_now=None; lease_end=None; rent_lock=None
    if s['tenant']:
        m2=re.search(r'(?:μισθωμένο|μίσθωμα|ενοίκιο)[^.€]{0,25}?(\d{3,4})\s*(?:€|ευρώ)',desc,re.I)
        if m2: rent_now=int(m2.group(1))
        m3=re.search(r'(?:έως|εως|μέχρι|διαθέσιμο από)\s*(?:και\s*)?(?:τις\s*|τον\s*)?(\d{1,2}/\d{1,2}/\d{4}|[^\s]+\s+(?:του\s+)?\d{4})',desc,re.I)
        if m3: lease_end=m3.group(1).replace('Αύγουστο του','Ağustos').replace('Αυγούστου','Ağustos')
        ended=False
        if lease_end:
            y=re.search(r'(\d{4})',lease_end); mo=re.match(r'\d{1,2}/(\d{1,2})/',lease_end)
            if y:
                yy=int(y.group(1)); mm=int(mo.group(1)) if mo else 12
                ended=(yy,mm)<(2026,11)
        tenant_txt='Kiracılı (ilan)'+(f' – {rent_now} €/ay' if rent_now else ' – kira tutarı yok')+(f'; bitiş: {lease_end}'+(' (geçmiş tarih → boş olabilir, teyit)' if ended else '') if lease_end else '; bitiş belirtilmemiş')
        if rent_now and not ended: rent_lock=rent_now
    split=bool(re.search(r'δύο (ανεξάρτητα )?διαμερίσματα|2 (ανεξάρτητα )?διαμερίσματα',desc,re.I))
    onstreet=None
    for g,e in [('Τσιμισκή','Tsimiski'),('Εγνατία','Egnatia'),('Μητροπόλεως','Mitropoleos'),('Αγίου Δημητρίου','Agiou Dimitriou'),('Νίκης','Nikis'),('Βασ. Όλγας','Vas. Olgas'),('Βασιλίσσης Όλγας','Vas. Olgas'),('Εθνικής Αμύνης','Ethnikis Aminis'),('Κασσάνδρου','Kassandrou'),('Αγίας Σοφίας','Agias Sofias'),('Βενιζέλου','Venizelou'),('Αριστοτέλους','Aristotelous'),('Παπαναστασίου','Papanastasiou'),('Λαγκαδά','Lagkada')]:
        if re.search(r'(επί της|επί τη[ςσ]|στην οδό|οδός|on)\s+(οδού\s+)?'+g[:-2],desc): onstreet=e;break
    T=TR[s['place']]
    rec=dict(no=f'D-{n:03d}',id=l['id'],url='https://spitimou.net/en'+l['path'],path=l['path'],adv=adv.get('name') or 'Bilinmiyor',advweb=adv.get('website') or '',ref=l.get('agentRef'),
        upd=l['sourceModified'],first=l.get('firstPublished'),lastseen=l.get('lastSeen'),gone=l.get('sourceGone'),
        mah=T,grp=GROUPS[T],mah_gr=s['place'],lat=l['lat'],lng=l['lng'],address=l.get('addressLatin') or l.get('address'),streets=s['streets'],onstreet=onstreet,
        dcenter=round(hav(*p,*CENTER)),road=rd[1],droad=round(rd[0]),metro=md[1],dmetro=round(md[0]),uni=ud[1],duni=round(ud[0]),dsea=l.get('seaDistanceM'),
        price=l['price'],sqm=l['sqm'],net=net,beds=l['bedrooms'],living=l['livingRooms'],baths=l['bathrooms'],wc=l.get('wc'),yb=l['yearBuilt'],ry=ry,ry_src=ry_src,scope=scope,
        floor=l['floor'],floorn=floornum(l['floor']),tf=l.get('totalFloors'),elev='Evet' if f.get('elevator',(None,))[0] else 'Belirtilmemiş',
        balcony='Evet' if f.get('veranda',(None,))[0] else 'Belirtilmemiş',heat=(l.get('heating') or '')+('/'+l['heatingMedium'] if l.get('heatingMedium') else ''),
        ac='Evet' if f.get('ac',(None,))[0] else 'Belirtilmemiş',energy=l.get('energyClass'),furn='Eşyalı' if f.get('furnished',(None,))[0] else 'Belirtilmemiş',
        parking='Evet' if f.get('parking',(None,))[0] else ('Hayır' if f.get('parking',(None,None))[0] is False else 'Belirtilmemiş'),fees=l.get('commonFees'),
        tenant=tenant_txt,rent_now=rent_now,rent_lock=rent_lock,split=split,two=s['two'],under=s['under'],crit=s['crit'],
        exit_ppsm=ex.get('ppsm'),exit_n=ex.get('n'),exit_basis=ex.get('basis'),area_med=l.get('areaMedianPpsm'),cond_med=l.get('condMedianPpsm'),
        desc=desc[:1500],dups=s['dups'],penthouse=l.get('isPenthouse'),student='student' in f,lux='luxury' in f)
    # comps
    cs=[]
    for r in rp:
        d=hav(*p,r['lat'],r['lng'])
        if d>900 or r['bedrooms']!=2 or abs(r['sqm']-l['sqm'])/l['sqm']>0.30: continue
        cs.append((d/300+abs(r['sqm']-l['sqm'])/l['sqm']*3+(0.4 if r['condition'] not in ('RENOVATED','NEW') else 0),d,r))
    cs.sort(key=lambda x:x[0])
    pick=cs[:8]
    if len(pick)>=4:
        med=statistics.median(x[2]['price']/x[2]['sqm'] for x in pick)
        pick=[x for x in pick if 0.65*med<=x[2]['price']/x[2]['sqm']<=1.5*med]
    pick=pick[:6]
    for k,(sc,d,r) in enumerate(pick,1):
        ex_=re.sub(r'\s+',' ',r.get('excerpt') or '')
        comps.append(dict(no=rec['no'],k=k,id=r['id'],url='https://spitimou.net/en'+r['path'],mah=r['place'],sqm=r['sqm'],floor=r['floor'],cond=r['condition'],
            furn='Eşyalı' if re.search(r'επιπλωμέν|furnished',ex_,re.I) and not re.search(r'μη επιπλωμ|χωρίς έπιπλα|unfurnished',ex_,re.I) else ('Eşyasız' if re.search(r'μη επιπλωμ|χωρίς έπιπλα|unfurnished',ex_,re.I) else 'Belirtilmemiş'),
            rent=r['price'],dist=round(d),upd=r['sourceModified']))
    rec['ncomp']=len(pick)
    recs.append(rec)
print(collections.Counter(r['ncomp'] for r in recs))
print('metro',statistics.median(r['dmetro'] for r in recs),'road',statistics.median(r['droad'] for r in recs),'onstreet',sum(1 for r in recs if r['onstreet']))
print('net parsed',sum(1 for r in recs if r['net']))
# market medians by mahalle (sale ppsm all renovated 2bd listing, rent ppsm pool)
mk={}
for r in SALE:
    if r['bedrooms']==2 and r['place'] in TR: mk.setdefault(TR[r['place']],{'s':[],'r':[]})['s'].append(r['price']/r['sqm'])
for r in rp:
    if r['place'] in TR and r['bedrooms']==2: mk.setdefault(TR[r['place']],{'s':[],'r':[]})['r'].append(r['price']/r['sqm'])
market={k:dict(sale_med=round(statistics.median(v['s'])) if v['s'] else None,sale_n=len(v['s']),rent_med=round(statistics.median(v['r']),2) if v['r'] else None,rent_n=len(v['r'])) for k,v in mk.items()}
print(json.dumps(market,ensure_ascii=False,indent=0)[:2000])
# extras (not selected) for reserve sheet
sel_ids={s['id'] for s in SEL['sel']}
extras=[]
for x in SEL['rest']:
    l=D[str(x['id'])]['listing']
    extras.append(dict(id=l['id'],url='https://spitimou.net/en'+l['path'],mah=TR.get(x['place'],x['place']),price=l['price'],sqm=l['sqm'],floor=l['floor'],ry=l['renovationYear'],adv=(l['advertiser'] or {}).get('name'),upd=l['sourceModified'],
       why=', '.join(k for k,vv in {'2+1 açıklamada teyitsiz':not x['crit']['beds'],'yenileme yılı <2019/bilinmiyor':not x['crit']['ren'],'asansör belirtilmemiş':not x['crit']['elev'],'tadilat devam ediyor':not x['crit']['ready']}.items() if vv) or 'Kriterleri karşılıyor – mahalle kotası nedeniyle yedek'))
json.dump(dict(recs=recs,comps=comps,market=market,extras=extras,rentpool=len(rp)),open('model_data.json','w'),ensure_ascii=False,default=str)
print('extras',len(extras),'comps',len(comps))
