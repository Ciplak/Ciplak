import json,re,sys,time,urllib.request,html
UA='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
MAH=['caferaga','kozyatagi','caddebostan','suadiye','fenerbahce','feneryolu','goztepe','erenkoy','zuhtupasa','kosuyolu']
out=sys.argv[1]; deals=sys.argv[2].split(',')
res={}
def get(u):
    for a in range(4):
        try:
            return urllib.request.urlopen(urllib.request.Request(u,headers={'User-Agent':UA,'Accept-Language':'tr-TR'}),timeout=40).read().decode('utf-8','ignore')
        except Exception as e:
            print('retry',u,e,flush=True); time.sleep(6*(a+1))
    return ''
for deal in deals:
    for m in MAH:
        seen=0
        for p in range(1,60):
            u=f'https://www.emlakjet.com/{deal}/istanbul-kadikoy-{m}-mahallesi'+(f'?sayfa={p}' if p>1 else '')
            h=get(u); new=0
            for s in re.findall(r'<script id="listing-realestate-schema"[^>]*>(.*?)</script>',h,re.S):
                try: g=json.loads(s)['@graph']
                except Exception: continue
                for it in g:
                    url=it.get('url'); 
                    if not url or url in res: continue
                    props={x['name']:x['value'] for x in it.get('additionalProperty',[])}
                    res[url]=dict(deal=deal,mah=m,name=it.get('name'),url=url,price=(it.get('offers') or {}).get('price'),cur=(it.get('offers') or {}).get('priceCurrency'),posted=it.get('datePosted'),props=props,img=it.get('image'))
                    new+=1
            print(deal,m,p,new,len(res),flush=True)
            if new==0: break
            time.sleep(1.2)
        json.dump(res,open(out,'w'),ensure_ascii=False)
json.dump(res,open(out,'w'),ensure_ascii=False)
print('done',len(res))
