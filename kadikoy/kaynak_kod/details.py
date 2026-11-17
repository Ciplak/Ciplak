import json,re,sys,time,os,html,urllib.request
UA='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
LABELS=['Oda Sayısı','Bulunduğu Kat','Banyo Sayısı','Kat Sayısı','Bina Yaşı','İlan Numarası','İlan Güncelleme Tarihi','Kategori','Isıtma Tipi','Otopark','Kullanım Durumu','Krediye Uygunluk','Site İçerisinde','Tapu Durumu','Takas','Asansör','Mutfak','Eşya Durumu','Balkon','Aidat','Depozito','Yapı Durumu','Cephe','Yapı Tipi','Kimden','Brüt / Net']
def parse(h):
    t=re.sub(r'<script.*?</script>|<style.*?</style>','',h,flags=re.S)
    toks=[html.unescape(x).strip() for x in re.split(r'<[^>]+>',t)]; toks=[x for x in toks if x]
    d={}
    try:
        i0=toks.index('Oda Sayısı'); iD=toks.index('İlan Açıklaması',i0)
    except ValueError: return None
    blk=toks[i0:iD]
    i=0
    while i<len(blk):
        if blk[i]=='Brüt / Net':
            # e.g. '105 m²','Brüt','79 m²','Net'
            seg=blk[i+1:i+5]
            for j in range(0,len(seg)-1,2):
                if seg[j+1] in ('Brüt','Net'): d[seg[j+1]+' m2']=seg[j]
            i+=5; continue
        if blk[i] in LABELS or (i+1<len(blk) and blk[i] not in d and i%1==0 and blk[i].istitle() and len(blk[i])<30):
            if i+1<len(blk): d[blk[i]]=blk[i+1]; i+=2; continue
        i+=1
    # header: agent and price, location
    try:
        k=toks.index(d.get('_t','')) if False else None
    except: pass
    j=i0
    d['_header']=toks[max(0,i0-6):i0]
    desc=[]
    for x in toks[iD+1:iD+200]:
        if 'Nasıl Bir Semt' in x: break
        desc.append(x)
    d['_desc']=' '.join(desc)[:3000]
    for k,pat in [('lat',r'\\?"(?:lat|latitude)\\?":\s*\\?"?([0-9.]+)'),('lng',r'\\?"(?:lon|lng|longitude)\\?":\s*\\?"?([0-9.]+)'),('agency',r'\\?"(?:storeName|officeName|companyName|ownerName|agencyName|firmName)\\?":\\?"([^"\\]+)'),('created',r'\\?"createdAt\\?":\\?"([0-9T:\-Z]+)'),('updated',r'\\?"updatedAt\\?":\\?"([0-9T:\-Z]+)')]:
        m=re.search(pat,h); d['_'+k]=m.group(1) if m else None
    m=re.search(r'<script[^>]*type="application/ld\+json"[^>]*>(\{"@context":"https://schema.org","@graph":\[\{"@type":"Product".*?)</script>',h,re.S)
    if m:
        try:
            g=json.loads(m.group(1))['@graph'][0]; d['_images']=g.get('image',[])[:3]
            o=g.get('offers') or {}; d['_price']=o.get('price')
        except Exception: pass
    bc=re.findall(r'"name":"(Kadıköy [^"]+ Mahallesi)',h); d['_mahalle']=bc[0] if bc else None
    return d
if __name__=='__main__':
    urls=json.load(open(sys.argv[1])); out=sys.argv[2]
    have=json.load(open(out)) if os.path.exists(out) else {}
    for n,u in enumerate(urls):
        if u in have: continue
        for a in range(4):
            try:
                h=urllib.request.urlopen(urllib.request.Request(u,headers={'User-Agent':UA}),timeout=40).read().decode('utf-8','ignore')
                have[u]=parse(h) or {'_error':'parse'}; break
            except Exception as e:
                print('retry',u,e,flush=True); time.sleep(6*(a+1))
        if n%20==0: json.dump(have,open(out,'w'),ensure_ascii=False); print(n,len(have),flush=True)
        time.sleep(1.0)
    json.dump(have,open(out,'w'),ensure_ascii=False); print('done',len(have))
