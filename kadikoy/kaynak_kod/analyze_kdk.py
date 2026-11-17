import json,sys,statistics,datetime
from openpyxl import load_workbook
wb=load_workbook(sys.argv[1],data_only=True)
M=wb['DAİRE ANA LİSTESİ']; hdr=[c.value for c in M[4]]
rows=[]
r=5
while M.cell(r,1).value: rows.append({h:M.cell(r,i+1).value for i,h in enumerate(hdr)}); r+=1
K=lambda s:[h for h in hdr if h and h.startswith(s)][0]
k=dict(no='Daire No',mah='Mahalle',room='Oda tipi',price='Satış fiyatı (₺)',eur='Satış fiyatı (€)',sqm='Brüt m²',ppsm='₺/brüt m²',ra=K('Aylık kira – AKTİF'),ny=K('NET KİRA'),gy=K('Brüt kira'),risk=K('Risk puanı'),score=K('GENEL'),pr=K('Fiyat / emsal'),nego=K('Pazarlık önerilen'),maxp=K('Hedef net'),dups=K('Aynı daire'),first='İlk yayın',mr=K('Risk işaretleri'),tinv='TOPLAM YATIRIM (₺)',dataok='Veri yeterliliği',bag=K('Bağdat'),comp=K('Emsal ort'),tapu='Tapu durumu',age='Bina yaşı (ilan)')
R=[{kk:d[v] for kk,v in k.items()} for d in rows]
def E(x): return f"{x:,.0f}".replace(',', '.')
def PC(x,n=2): return ('%'+f"{x*100:.{n}f}").replace('.', ',')
def DC(x,n=2): return f"{x:.{n}f}".replace('.', ',')
def MN(x): return DC(x/1e6,2)+' mn ₺'
out={}
pairs=[]
for a in R:
    for b in R:
        if a is not b and abs(a['ra']-b['ra'])/b['ra']<=0.05 and a['price']<=0.75*b['price']: pairs.append((b['price']-a['price'],a,b))
pairs.sort(key=lambda x:-x[0]); seen=set(); ex=[]
for g,a,b in pairs:
    if b['no'] in seen: continue
    seen.add(b['no']); ex.append(f"{a['no']} ({a['mah']} {a['room']}, {MN(a['price'])}, kira ≈{E(a['ra'])} ₺) ≈ {b['no']} ({b['mah']} {b['room']}, {MN(b['price'])}, kira ≈{E(b['ra'])} ₺)")
    if len(ex)>=6: break
out['q2']=f"Evet. Aylık kirası ±%5 aynı olup fiyatı en az %25 düşük muadili bulunan {len(seen)} daire var. Örnekler: "+'; '.join(ex)+". Genel desen: Kozyatağı, Göztepe ve Erenköy’deki daireler, Fenerbahçe/Caddebostan/Suadiye’nin Bağdat Caddesi’ne yakın pahalı dairelerine benzer kira üretiyor."
op=sorted([x for x in R if isinstance(x['pr'],(int,float)) and x['pr']>1.15],key=lambda x:-x['pr'])
out['q5']=f"₺/m²’si aynı mahalle ve oda tipindeki ≤15 yaş ilanların medyanından %15’ten fazla yüksek {len(op)} daire var: "+'; '.join(f"{x['no']} {x['mah']} {x['room']} – {MN(x['price'])} ({E(x['ppsm'])} ₺/m², emsalin {DC(x['pr'])} katı)" for x in op[:10])+(f". Bu dairelerin ortalama net getirisi {PC(statistics.mean(x['ny'] for x in op))}, portföy ortalaması {PC(statistics.mean(x['ny'] for x in R))}." if op else '.')
CHECK=datetime.datetime(2026,11,17)
neg=[]
for x in R:
    pts=0; why=[]
    if isinstance(x['pr'],(int,float)) and x['pr']>1.05: pts+=2; why.append(f"emsalin {DC(x['pr'])} katı")
    if x['dups']!='Tekrarı bulunamadı': pts+=2; why.append('aynı daire başka ofiste farklı fiyatla')
    if isinstance(x['first'],datetime.datetime) and (CHECK-x['first']).days>90: pts+=1; why.append('3+ aydır yayında')
    if isinstance(x['maxp'],(int,float)) and x['maxp']<x['price']*0.85: pts+=1; why.append('getiri hedefi için fiyat yüksek')
    if pts>=3: neg.append((pts,x,why))
neg.sort(key=lambda t:(-t[0],-t[1]['price']))
out['q6']=f"Pazarlık potansiyeli puanı ≥3 olan {len(neg)} daire var. Öne çıkanlar: "+'; '.join(f"{x['no']} ({', '.join(w)}; önerilen teklif ≈{MN(x['nego'])})" for p,x,w in neg[:10])+". Not: Emlakjet’te ilk yayın tarihi çoğu ilanda yeniden yayınla sıfırlanabilir."
hr=sorted(R,key=lambda x:-x['risk'])[:10]
out['q7']="Risk puanı en yüksek daireler: "+'; '.join(f"{x['no']} ({x['mah']}, puan {x['risk']}: {x['mr'].rstrip('; ')})" for x in hr)+f". Örneklemde tapusu kat mülkiyeti olmayan {sum(1 for x in R if x['tapu']!='Kat Mülkiyeti')} daire var (kat irtifakı, tapu kaydı yok, intifa hakkı vb.); bunlarda iskân ve kat mülkiyetine geçiş teyit edilmeden alım yapılmamalı. 2010–2015 yapımı binalar 2007 deprem yönetmeliğine, 2019 sonrası ruhsatlılar TBDY 2018’e tabidir; ruhsat tarihi ve yapı denetim raporları kontrol edilmeli."
st=sorted([x for x in R if x['risk']<=20 and x['dataok']=='Yeterli' and x['tapu']=='Kat Mülkiyeti'],key=lambda x:-x['ny'])[:10]
out['q8']="Risk puanı ≤20, tapusu kat mülkiyetli ve verisi yeterli daireler arasından net getirisi en yüksek 10 daire: "+'; '.join(f"{x['no']} – {x['mah']} {x['room']}, {MN(x['price'])}, kira ≈{E(x['ra'])} ₺/ay, net {PC(x['ny'])}" for x in st)+"."
H=wb['MAHALLE ANALİZİ']
mh=[(H.cell(r,1).value,H.cell(r,7).value,H.cell(r,8).value,H.cell(r,2).value) for r in range(6,16) if isinstance(H.cell(r,7).value,(int,float))]
mh.sort(key=lambda t:t[1])
out['q3']="Kira çarpanı (fiyat / yıllık brüt kira) en avantajlı mahalleler: "+'; '.join(f"{a} {DC(b,1)} yıl (net {PC(c)}, n={n})" for a,b,c,n in mh[:4])+". En yüksek çarpan: "+'; '.join(f"{a} {DC(b,1)} yıl" for a,b,c,n in mh[-3:])+". Emlakjet’in tüm daire ilanlarına dayanan mahalle amortismanı 15–23 yıl; en uzun Fenerbahçe (23) ve Caddebostan/Zühtüpaşa (21). Caddebostan, Fenerbahçe, Caferağa, Zühtüpaşa ve Koşuyolu’nda 2010 sonrası 1+1/2+1 örneklemi küçük (n≤4) – bu mahallelerin sonuçları dikkatle yorumlanmalı."
near=[x for x in R if x['bag']<=300]; far=[x for x in R if x['bag']>300]
f=lambda L,key: statistics.mean(x[key] for x in L)
if near and far:
    out['q4']=f"Bağdat Caddesi’ne ≤300 m {len(near)} dairenin ortalama fiyatı {E(f(near,'ppsm'))} ₺/m², diğer {len(far)} dairenin {E(f(far,'ppsm'))} ₺/m² (prim {PC(f(near,'ppsm')/f(far,'ppsm')-1,1)}). Emsal kira farkı {PC(f(near,'comp')/f(far,'comp')-1,1)}; net getiri ≤300 m grubunda {PC(f(near,'ny'))}, diğerlerinde {PC(f(far,'ny'))}. " + ("Sonuç: Bağdat Caddesi primi kira ile karşılanmıyor; caddeye 300–800 m mesafede, raylı sisteme yakın sokaklar getiri açısından daha verimli." if f(near,'ny')<f(far,'ny') else "Sonuç: Bağdat Caddesi yakınlığı primi kira ile büyük ölçüde karşılanıyor.") + " Not: mesafe yaklaşık güzergâhla hesaplandı."
bd=[]
for lab,cond in [('1+1',lambda x:x['room']=='1+1'),('2+1',lambda x:x['room']=='2+1')]:
    L=[x for x in R if cond(x)]
    if L: bd.append((lab,len(L),statistics.mean(x['price'] for x in L),statistics.mean(x['ny'] for x in L),statistics.mean(x['ra'] for x in L)))
best_small=sorted([x for x in R if x['room']=='1+1' and x['risk']<=25],key=lambda x:-x['ny'])[:2]
big=sorted([x for x in R if x['room']=='2+1' and x['sqm']>=100],key=lambda x:-x['ny'])[:1]
out['q9']="Oda tipine göre: "+'; '.join(f"{lab}: {n} daire, ort. fiyat {MN(p)}, ort. kira {E(rr)} ₺/ay, net {PC(y)}" for lab,n,p,y,rr in bd)+". "+(f"Örnek: iki 1+1 ({', '.join(x['no'] for x in best_small)}; toplam ≈{MN(sum(x['tinv'] for x in best_small))}) ortalama net {PC(statistics.mean(x['ny'] for x in best_small))}; büyük 2+1’lerin en iyisi {big[0]['no']} ({int(big[0]['sqm'])} m², {MN(big[0]['price'])}) net {PC(big[0]['ny'])}. " if best_small and big else '')+"Kadıköy’de 1+1 arzı az (örneklemde yalnızca "+str(sum(1 for x in R if x['room']=='1+1'))+" adet) ama m² başına kira daha yüksek. Birkaç küçük daire boş kalma riskini dağıtır; ancak konut kira istisnası kişi başına bir kez uygulandığından toplam kira arttıkça vergi dilimi yükselir (yöntem 2 ile test edin)."
CF=wb['5 VE 10 YILLIK NAKİT AKIŞI']
G0=46; n=len(R)
e5=[CF.cell(r,21).value for r in range(G0+1,G0+1+n)]; e10=[CF.cell(r,22).value for r in range(G0+1,G0+1+n)]
t5=[CF.cell(r,23).value for r in range(G0+1,G0+1+n)]; t10=[CF.cell(r,24).value for r in range(G0+1,G0+1+n)]
out['q10']=f"Baz senaryoda (TÜFE %25’ten %12’ye iniyor, kira TÜFE kadar artıyor, reel değer artışı %1/yıl, TL reel olarak sabit) ortalama IRR: 5 yıl {PC(statistics.mean(e5))} (€) / {PC(statistics.mean(t5))} (TL), 10 yıl {PC(statistics.mean(e10))} (€) / {PC(statistics.mean(t10))} (TL). TL getirisi yüksek görünür ama enflasyonun büyük kısmını telafi eder; € bazlı IRR gerçek karşılaştırma ölçüsüdür. Giriş maliyeti (%4,4 + sabit) ve çıkış maliyeti (%4,4) 5 yılda daha ağır basar; 10 yıllık tutma daha avantajlı. Kötümser senaryoda (TL’nin yılda %5 reel değer kaybı, reel fiyat düşüşü) € getirisi belirgin düşer."
T=wb['EN İYİ 20 YATIRIM']
out['q1']="Salt net kira getirisine göre ilk 20 (baz, 17.11.2026): "+', '.join(str(T.cell(r,2).value) for r in range(29,49) if T.cell(r,2).value)+". Genel puana göre ilk 20: "+', '.join(str(T.cell(r,2).value) for r in range(5,25) if T.cell(r,2).value)+"."
json.dump(out,open(sys.argv[2],'w'),ensure_ascii=False,indent=1)
for kx,v in out.items(): print(kx,':',v[:400],'\n')
