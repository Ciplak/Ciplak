import json,re,math,collections
D=json.load(open('alldet.json'))
def hav(a,b,c,d):
    R=6371000;p1,p2=math.radians(a),math.radians(c);dp=p2-p1;dl=math.radians(d-b)
    return 2*R*math.asin(math.sqrt(math.sin(dp/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2))
TWO=re.compile(r'(2|δύο|δυο)\s*(άνετα |μεγάλα |ευρύχωρα )?(υπνοδωμάτι|υ/δ|υπνοδ|bedroom)',re.I)
BAD=re.compile(r'(\b(1|3|4)\b|ένα|τρία|τρια|τέσσερα)\s*(άνετα |μεγάλα |ευρύχωρα )?(υπνοδωμάτι|υ/δ|bedroom)',re.I)
ROOMS3=re.compile(r'(3|τρία)\s*(μεγάλα )?δωμάτια',re.I)
UNDER=re.compile(r'υπό ανακαίνιση|υπο ανακαινιση|υπό κατασκευή',re.I)
TEN=re.compile(r'μισθωμέν|μισθωμεν|ενοικιασμέν|ενοικιασμεν|με ενοικιαστ|με μισθωτ|έχει ενοικιαστ|rented|tenant',re.I)
STREETS={'Τσιμισκή':'Tsimiski','Εγνατία':'Egnatia','Μητροπόλεως':'Mitropoleos','Αγίου Δημητρίου':'Agiou Dimitriou','Αγ. Δημητρίου':'Agiou Dimitriou','Αριστοτέλους':'Aristotelous','Αγίας Σοφίας':'Agias Sofias','Νίκης':'Nikis','Βασ. Όλγας':'Vas. Olgas','Βασιλίσσης Όλγας':'Vas. Olgas','Παπαναστασίου':'Papanastasiou','Δελφών':'Delfon','Λαγκαδά':'Lagkada','Μοναστηρίου':'Monastiriou','26ης Οκτωβρίου':'26is Oktovriou','Εθνικής Αμύνης':'Ethnikis Aminis','Βενιζέλου':'Venizelou','Κασσάνδρου':'Kassandrou','Ολυμπιάδος':'Olympiados','Φιλίππου':'Filippou','Ερμού':'Ermou','Καρόλου Ντηλ':'Karolou Diehl','Παύλου Μελά':'Pavlou Mela','Γούναρη':'Gounari','Κωνσταντινουπόλεως':'Konstantinoupoleos','Μαρτίου':'Martiou','Ιασωνίδου':'Iasonidou','Βασ. Σοφίας':'Vas. Sofias','Αγγελάκη':'Aggelaki','Εγνατίας':'Egnatia','Τσιμισκη':'Tsimiski','Καστριτσίου':'Kastritsiou','Φράγκων':'Fragkon','Λέοντος Σοφού':'Leontos Sofou','Ίωνος Δραγούμη':'Ionos Dragoumi','Πολυτεχνείου':'Polytechniou','Κ. Καραμανλή':'K. Karamanli','Ανδρέου Παπανδρέου':'Andreou Papandreou'}
MAIN={'Tsimiski','Egnatia','Mitropoleos','Agiou Dimitriou','Nikis','Vas. Olgas','Papanastasiou','Delfon','Lagkada','Monastiriou','26is Oktovriou','Ethnikis Aminis','Venizelou','Aristotelous','Agias Sofias','Kassandrou','Vas. Sofias','K. Karamanli','Polytechniou','Andreou Papandreou'}
CORE={'Αριστοτέλους - Ιστορικό Κέντρο':'Aristotelous / Tarihi Merkez','Ρωμαϊκή Αγορά - Βενιζέλου - Ιστορικό Κέντρο':'Roma Agorası - Venizelou','Κέντρο Θεσσαλονίκης':'Merkez (Tsimiski-Egnatia)','Κέντρο':'Merkez','Λαδάδικα':'Ladadika','Πλατεία Εμπορίου - Ανω Λαδάδικα - Ιστορικό Κέντρο':'Ladadika / Emboriou Meyd.','Αγία Σοφία':'Agia Sofia','Λευκός Πύργος':'Beyaz Kule','Καμάρα - Ροτόντα':'Kamara - Rotonda','Άγιος Δημήτριος':'Agios Dimitrios','Βαρδάρης - Δικαστήρια':'Vardaris - Adliye','Διοικητήριο':'Dioikitirio','Αντιγονιδών':'Antigonidon','Σταθμός ΟΣΕ':'OSE Tren İstasyonu','Πλατεία Ναυαρίνου':'Navarinou Meydanı','ΔΕΘ':'Fuar (HELEXPO)','ΧΑΝΘ':'XANTH (Beyaz Kule)','Τουρκικό Προξενείο':'Türk Konsolosluğu (Ag. Dimitriou)','Φάληρο - Ιπποκράτειο':'Faliro','Ιπποκράτειο':'Faliro - Ippokrateio','Αγία Τριάδα':'Faliro - Agia Triada'}
EXT={'Ευζώνων - Πεδίον Άρεως':'Evzonon - Pedion Areos','Παπάφη - Άγιος Φανούριος':'Papafi','Παναγία Φανερωμένη':'Panagia Faneromeni','Νέα Παραλία':'Nea Paralia','Ανάληψη - Μπότσαρη - Νέα Παραλία':'Analipsi - Nea Paralia','Ανάληψη - Μπότσαρη':'Analipsi - Botsari'}
TIER={**{k:1 for k in CORE},**{k:2 for k in EXT}}
rows=[]
for k,v in D.items():
    l=v['listing'];desc=l['description'] or ''
    feats={x['name']:(x['present'],x['detail']) for x in v['features']}
    place=None
    for pl in [ (v.get('place') or {}).get('name'), l.get('neighbourhood'), l.get('municipality')]:
        if pl in TIER: place=pl;break
    if not place:
        # find from list data
        continue
    two=bool(TWO.search(desc)); bad=bool(BAD.search(desc)) and not two; r3=bool(ROOMS3.search(desc))
    st=[e for g,e in STREETS.items() if g in desc]
    rows.append(dict(id=l['id'],l=l,feats=feats,place=place,tier=TIER[place],two=two,bad=bad or (r3 and not two),under=bool(UNDER.search(desc)),tenant=bool(TEN.search(desc)),streets=list(dict.fromkeys(st))))
print('rows',len(rows))
# dedupe
rows.sort(key=lambda r:r['l']['price'])
groups=[]
for r in rows:
    l=r['l']
    for g in groups:
        m=g[0]['l']
        if abs(m['sqm']-l['sqm'])<=3 and (m['floor']==l['floor']) and hav(m['lat'],m['lng'],l['lat'],l['lng'])<250 and abs(m['price']-l['price'])/m['price']<=0.10:
            g.append(r);break
    else: groups.append([r])
print('groups',len(groups),'dups',sum(len(g)-1 for g in groups))
def ren(r):
    y=r['l']['renovationYear'] or 0
    return y
cands=[]
for g in groups:
    r=g[0]; r['dups']=[x for x in g[1:]]
    l=r['l']
    el=r['feats'].get('elevator',(None,))[0]
    ok_ren = ren(r)>=2019
    crit = dict(center=r['tier']==1, beds=r['two'], ren=ok_ren, size=45<=l['sqm']<=110, elev=el is True, ready=not r['under'])
    r['crit']=crit
    if r['bad']: continue
    cands.append(r)
print('cands',len(cands))
def pre(r):
    c=r['crit'];l=r['l']
    s=(3 if c['center'] else 0)+2*c['beds']+2*c['ren']+2*c['elev']+1*c['ready']+ (1 if (l['sourceModified'] or '')>='2026-10' else 0)+(1 if r['streets'] and set(r['streets'])&MAIN else 0)+(2 if (l.get('lastSeen') or '')>='2026-10-15' else 0)
    return s
cands.sort(key=lambda r:(-all(r['crit'].values()),-((r['l'].get('lastSeen') or '')>='2026-10-01'),-pre(r),-(int((r['l']['sourceModified'] or '0').replace('-','')))))
cap=collections.Counter();sel=[];rest=[]
for r in cands:
    if len(sel)<100 and cap[r['place']]<8 and pre(r)>=9:
        sel.append(r);cap[r['place']]+=1
    else: rest.append(r)
for r in list(rest):
    if len(sel)>=100: break
    if cap[r['place']]<11 and pre(r)>=9: sel.append(r);rest.remove(r);cap[r['place']]+=1
print('selected',len(sel),cap.most_common())
print('allcrit',sum(all(r['crit'].values()) for r in sel), collections.Counter(pre(r) for r in sel))
print('tier',collections.Counter(r['tier'] for r in sel))
json.dump({'sel':[dict(id=r['id'],place=r['place'],tier=r['tier'],two=r['two'],under=r['under'],tenant=r['tenant'],streets=r['streets'],crit=r['crit'],dups=[dict(id=x['id'],price=x['l']['price'],adv=(x['l']['advertiser'] or {}).get('name'),path=x['l']['path'],ref=x['l'].get('agentRef')) for x in r['dups']]) for r in sel],
           'rest':[dict(id=r['id'],place=r['place'],crit=r['crit'],pre=pre(r)) for r in rest]},open('selection.json','w'),ensure_ascii=False)
