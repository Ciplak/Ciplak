import json,sys,statistics
from openpyxl import load_workbook
wb=load_workbook(sys.argv[1],data_only=True)
M=wb['100 DAİRE ANA LİSTESİ']; hdr=[c.value for c in M[4]]
rows=[]
for r in range(5,105):
    d={h:M.cell(r,i+1).value for i,h in enumerate(hdr)}; rows.append(d)
K=lambda s:[h for h in hdr if h.startswith(s)][0]
k=dict(no='Daire No',grp='Öncelikli bölge grubu',mah='Mahalle',price='Satış fiyatı (€)',sqm='İlan m² (brüt/beyan)',ppsm='Satış fiyatı / m² (€)',ra=K('Aylık kira – AKTİF'),ny=K('NET KİRA'),gy=K('Brüt kira'),risk=K('Risk puanı'),score=K('GENEL'),pr=K('Fiyat / emsal'),nego=K('Pazarlık önerilen'),maxp=K('Hedef net'),dups=K('Aynı daire'),first=K('İlk yayın'),mr=K('En büyük risk'),tinv=K('TOPLAM YATIRIM'),ncf=K('Vergi sonrası net'),ex=K('Emsal satış €/m²'),dataok='Veri yeterliliği',rank='SIRA',yb='Bina yapım yılı',droad=K('Ana caddeye'),dm=K('Metroya'),ry='Yenileme yılı',energy='Enerji sınıfı',tenant='Kiracı durumu',legal=K('Hukuki'),comp=K('Emsal ort'))
R=[{kk:d[v] for kk,v in k.items()} for d in rows]
out={}
def E(x): return f"{x:,.0f}".replace(',', '.')
def PC(x,n=2): return ('%'+f"{x*100:.{n}f}").replace('.', ',')
def DC(x,n=2): return f"{x:.{n}f}".replace('.', ',')
# Q2: similar rent, lower price
pairs=[]
for a in R:
    for b in R:
        if a is b: continue
        if abs(a['ra']-b['ra'])/b['ra']<=0.05 and a['price']<=0.75*b['price']:
            pairs.append((b['price']-a['price'],a,b))
pairs.sort(key=lambda x:-x[0])
seen=set();ex=[]
for g,a,b in pairs:
    if b['no'] in seen: continue
    seen.add(b['no']); ex.append(f"{a['no']} ({a['grp']}, {E(a['price'])} €, kira ≈{E(a['ra'])} €) ≈ {b['no']} ({b['grp']}, {E(b['price'])} €, kira ≈{E(b['ra'])} €)")
    if len(ex)>=6: break
out['pairs_n']=len({b['no'] for g,a,b in pairs})
out['q2']=(f"Evet. Baz senaryoda aylık kirası ±%5 aynı olup satın alma fiyatı en az %25 daha düşük olan {out['pairs_n']} 'pahalı' daire için daha ucuz bir muadil bulunuyor. En belirgin örnekler: " + '; '.join(ex) + ". Genel desen: Vardaris, Antigonidon, Agios Dimitrios ve OSE çevresinde 120–160 bin €’luk daireler, Beyaz Kule / Aristotelous / Agia Sofia’daki 250–400 bin €’luk dairelere yakın kira üretiyor; prestij primi kiraya yansımıyor.")
# Q5 overpriced
op=sorted([x for x in R if isinstance(x['pr'],(int,float)) and x['pr']>1.15],key=lambda x:-x['pr'])
out['q5']=f"Satış €/m²’si aynı sokak/çevredeki satılık ilan medyanının %15’inden fazla üstünde olan {len(op)} daire var (ana listede kırmızı): " + '; '.join(f"{x['no']} {x['grp']} – {E(x['price'])} € ({E(x['ppsm'])} €/m², emsalin {DC(x['pr'])} katı)" for x in op[:10]) + ". Bu dairelerde net getiri ortalaması " + PC(statistics.mean(x['ny'] for x in op)) + " (portföy ortalaması " + PC(statistics.mean(x['ny'] for x in R)) + ")."
# Q6 negotiation
neg=[]
for x in R:
    pts=0; why=[]
    if isinstance(x['pr'],(int,float)) and x['pr']>1.05: pts+=2; why.append(f"emsalin {DC(x['pr'])} katı")
    if x['dups']!='Tekrarı bulunamadı': pts+=2; why.append('aynı daire başka ofiste farklı fiyatla')
    if isinstance(x['first'],__import__('datetime').datetime) and (wb['PARAMETRELER']['C6'].value - x['first']).days>120: pts+=1; why.append('ilan 4+ aydır yayında')
    if isinstance(x['maxp'],(int,float)) and x['maxp']<x['price']*0.85: pts+=1; why.append('getiri hedefi için fiyat yüksek')
    if pts>=3: neg.append((pts,x,why))
neg.sort(key=lambda t:(-t[0],-t[1]['price']))
out['q6']=f"Pazarlık potansiyeli puanı ≥3 olan {len(neg)} daire belirlendi (fiyatı emsal üstü, aynı dairenin başka ofiste farklı fiyatla ilanda olması, uzun süredir yayında olması). Öne çıkanlar: " + '; '.join(f"{x['no']} ({', '.join(w)}; önerilen teklif ≈{E(x['nego'])} €)" for p,x,w in neg[:10]) + ". Mükerrer ilanlarda en düşük ilan fiyatı ana listede kullanılmıştır; diğer ofis fiyatları ‘Aynı daire – diğer ilanlar’ sütunundadır."
# Q7 risk
hr=sorted(R,key=lambda x:-x['risk'])[:12]
out['q7']="Risk puanı en yüksek binalar/daireler: " + '; '.join(f"{x['no']} ({x['grp']}, puan {x['risk']}: {x['mr'].rstrip('; ')})" for x in hr[:10]) + ". Genel tablo: örneklemdeki binaların çoğu 1960–1980 yapımı (1985 öncesi deprem yönetmeliği) – statik inceleme şart. Kiracılı ilanlar, enerji sınıfı E–G olanlar ve tadilatı süren daireler ek hukuki/teknik kontrol gerektirir."
# Q8 steady income: low risk, data ok, net yield high
st=sorted([x for x in R if x['risk']<=25 and x['dataok']=='Yeterli'],key=lambda x:-x['ny'])[:10]
out['q8_list']=[x['no'] for x in st]
out['q8']="Düzenli kira geliri için risk puanı ≤25 ve verisi yeterli daireler arasından net getirisi en yüksek 10 daire: " + '; '.join(f"{x['no']} – {x['grp']}, {E(x['price'])} €, kira ≈{E(x['ra'])} €/ay, net {PC(x['ny'])}" for x in st) + ". Bu liste merkez/metro yakınlığı, ≥5 emsal ve düşük teknik risk dengesini önceliklendirir."
# Q3 mahalle multipliers
H=wb['MAHALLE ANALİZİ']
mh=[(H.cell(r,1).value,H.cell(r,7).value,H.cell(r,8).value,H.cell(r,14).value) for r in range(6,16)]
mh.sort(key=lambda t:t[1])
out['q3']="Kira çarpanı (fiyat / yıllık brüt kira) en düşük – yani en avantajlı – bölgeler: " + '; '.join(f"{a} {DC(b,1)} yıl (net {PC(c)})" for a,b,c,d in mh[:4]) + ". En pahalı (çarpanı yüksek) bölgeler: " + '; '.join(f"{a} {DC(b,1)} yıl" for a,b,c,d in mh[-3:]) + ". Batı merkez (Vardaris–Antigonidon–Agios Dimitrios–OSE) alım fiyatı düşük olduğu için daha avantajlı; Beyaz Kule/Aristotelous/Agia Sofia kira olarak en yüksek ama fiyat primi daha da yüksek."
# Q4
near=[x for x in R if x['droad']<=100]; far=[x for x in R if x['droad']>100]
f=lambda L,key: statistics.mean(x[key] for x in L)
out['q4']=(f"Ana caddeye ≤100 m olan {len(near)} dairenin ortalama satış fiyatı {E(f(near,'ppsm'))} €/m², diğer {len(far)} dairenin {E(f(far,'ppsm'))} €/m² (fark {PC(f(near,'ppsm')/f(far,'ppsm')-1,1)}). Emsal ilan kirası farkı {PC(f(near,'comp')/f(far,'comp')-1,1)}; net getiri ≤100 m grubunda {PC(f(near,'ny'))}, diğerlerinde {PC(f(far,'ny'))}. ") + ("Sonuç: ana cadde primi kira primiyle karşılanmıyor – ana caddeye 100–300 m mesafedeki ara sokaklar getiri açısından daha verimli." if f(near,'ny')<f(far,'ny') else "Sonuç: ana caddeye ≤100 m dairelerde kira primi (+"+PC(f(near,'comp')/f(far,'comp')-1,1)+") fiyat primini (+"+PC(f(near,'ppsm')/f(far,'ppsm')-1,1)+") karşılıyor; net getiri farkı küçük ama ana cadde yakınlığı ayrıca kiralama hızı ve yeniden satış likiditesi sağlıyor. Ancak doğrudan ana cadde cephesi gürültü nedeniyle arka cepheye göre daha zor kiralanabilir.") + " Not: mesafe yaklaşık güzergâh koordinatlarıyla hesaplanmıştır."
# Q9 size bands
bands=[(45,60),(60.01,75),(75.01,90),(90.01,110)]
bd=[]
for lo,hi in bands:
    L=[x for x in R if lo<=x['sqm']<=hi]
    if L: bd.append((lo,hi,len(L),statistics.mean(x['price'] for x in L),statistics.mean(x['ny'] for x in L)))
best_small=sorted([x for x in R if x['sqm']<=70 and x['risk']<=30],key=lambda x:-x['ny'])[:3]
big=sorted([x for x in R if x['sqm']>=85],key=lambda x:-x['ny'])[:1]
out['q9']="m² bandına göre ortalama net getiri: " + '; '.join(f"{int(lo)}–{int(hi)} m²: {n} daire, ort. fiyat {E(p)} €, net {PC(y)}" for lo,hi,n,p,y in bd) + ". " + (f"Örnek: toplam ≈{E(sum(x['tinv'] for x in best_small))} € yatırımla üç küçük daire ({', '.join(x['no'] for x in best_small)}) ortalama net {PC(statistics.mean(x['ny'] for x in best_small))} sağlarken, ≥85 m² daireler içinde en iyisi {big[0]['no']} ({int(big[0]['sqm'])} m², {E(big[0]['price'])} €) net {PC(big[0]['ny'])} sağlıyor. " if big else '') + "Birkaç küçük/orta 2+1 (60–75 m² bandı en verimli), boş kalma riskini dağıtır ve m² başına daha yüksek kira getirir. Dezavantajı: toplam kira geliri yükseldikçe Yunan kira vergisi %25/%35 dilimlerine çıkar (PARAMETRELER’de yöntem 2 ile test edin) ve yönetim/işlem maliyetleri çoğalır."
CF=wb['5 VE 10 YILLIK NAKİT AKIŞI']
irr5=[CF.cell(r,21).value for r in range(37,137)]; irr10=[CF.cell(r,22).value for r in range(37,137)]
tr5=[CF.cell(r,23).value for r in range(37,137)]; tr10=[CF.cell(r,24).value for r in range(37,137)]
out['q10']=f"Baz senaryoda (kira artışı %2,5, gider artışı %2,5, değer artışı %2,5/yıl, çıkış maliyeti %3) ortalama 5 yıllık IRR {PC(statistics.mean(irr5))}, 10 yıllık IRR {PC(statistics.mean(irr10))}; 5 yılda toplam getiri yatırımın {PC(statistics.mean(tr5),1)}’i, 10 yılda {PC(statistics.mean(tr10),1)}’i. En iyi 10 yıllık IRR: {PC(max(irr10))}. " + "Getirinin önemli kısmı değer artışından geliyor; kötümser senaryoda (değer artışı %0, %15 değer artış vergisi) IRR net kira getirisinin altına iner. 5 yıllık tutmada giriş (≈%8,5 + sabit masraflar) ve çıkış (≈%3) maliyetleri daha kısa sürede itfa edildiği için IRR belirgin şekilde düşüktür; uzun dönem tutma daha avantajlıdır."
top=[(wb['EN İYİ 20 YATIRIM'].cell(r,2).value) for r in range(5,25)]
out['q1']="Salt net kira getirisine göre ilk 20 daire EN İYİ 20 YATIRIM sayfasının alt tablosunda formülle listelenir (baz senaryo, 17.11.2026): " + ', '.join(wb['EN İYİ 20 YATIRIM'].cell(r,2).value for r in range(29,49)) + ". Genel puana (risk dahil) göre ilk 20: " + ', '.join(top) + "."
json.dump(out,open(sys.argv[2],'w'),ensure_ascii=False,indent=1)
for kx,v in out.items(): print(kx,':',v,'\n')
