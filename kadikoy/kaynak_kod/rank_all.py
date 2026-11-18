import json,statistics,collections,sys
d=json.load(open(sys.argv[1])); R=d['recs']; MK=d['market']
LUX={'Caferağa','Kozyatağı','Caddebostan','Suadiye','Fenerbahçe','Feneryolu','Göztepe','Erenköy','Zühtüpaşa','Koşuyolu'}
def tax(T):
    M=max(0,T-58000)*0.85
    return min(M,190000)*.15+max(0,min(M,400000)-190000)*.2+max(0,min(M,1e6)-400000)*.27+max(0,min(M,5.3e6)-1e6)*.35+max(0,M-5.3e6)*.4
out=[]
for r in R:
    cs=r.get('comps',[])
    if not cs: continue
    pm=statistics.mean(c['rent']/(c['net'] or c['brut']*0.8) for c in cs)
    netm=r['net'] or r['brut']*0.8
    ground=(r['floorn'] is not None and r['floorn']<1)
    rent=r['rent_now'] or round(pm*0.95*netm*(0.9 if ground else 1),-2)
    coll=rent*11.25
    aid=r['aidat'] or 2500
    opex=rent*12*0.03+rent*1.2/3+r['price']*0.0007+r['brut']*60+aid*(0.75+11.25*0.1)
    tinv=r['price']*0.95*1.044+2534+25000+20000
    ncf=coll-opex-tax(coll)
    out.append(dict(no=r['no'],mah=r['mah'],lux=r['mah'] in LUX,room=r['room'],brut=r['brut'],net=r['net'],floor=r['floor'],age=r['age'],price=r['price'],rent=rent,
        carpan=r['price']/(rent*12),gy=rent*12/r['price'],ny=ncf/tinv,tapu=r['tapu'],usage=r['usage'],kredi=r['kredi'],rail=r['rail'],drail=r['drail'],
        ncomp=len(cs),cdist=statistics.median(c['dist'] for c in cs),url=r['url'],src=r['source'],ground=ground,adv=r['adv'],dups=len(r['dups']),upd=r['updated'],claim=r.get('rent_claim')))
json.dump(out,open(sys.argv[2],'w'),ensure_ascii=False)
for grp,L in (('lüks',[x for x in out if x['lux']]),('diğer',[x for x in out if not x['lux']])):
    print(grp,len(L),'çarpan medyan',round(statistics.median(x['carpan'] for x in L),1),'net medyan',round(statistics.median(x['ny'] for x in L)*100,2))
bym=collections.defaultdict(list)
for x in out: bym[x['mah']].append(x)
for m,L in sorted(bym.items(),key=lambda t:statistics.median(x['carpan'] for x in t[1])):
    print(f"{m:14s} n={len(L):3d} çarpan={statistics.median(x['carpan'] for x in L):5.1f} net={statistics.median(x['ny'] for x in L)*100:5.2f}% fiyat={statistics.median(x['price'] for x in L)/1e6:5.2f}mn emlakjet_amort={MK.get(m,{}).get('amort')}")
