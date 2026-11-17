import json,re,sys,time,urllib.request
UA='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
out=sys.argv[1]; base=sys.argv[2]
res={}
for p in range(1,80):
    u=base+(f'?sayfa={p}' if p>1 else '')
    h=''
    for a in range(4):
        try: h=urllib.request.urlopen(urllib.request.Request(u,headers={'User-Agent':UA}),timeout=40).read().decode('utf-8','ignore');break
        except Exception as e: print('retry',e,flush=True);time.sleep(6*(a+1))
    new=0
    for s in re.findall(r'<script id="listing-realestate-schema"[^>]*>(.*?)</script>',h,re.S):
        for it in json.loads(s)['@graph']:
            url=it.get('url')
            if url and url not in res:
                props={x['name']:x['value'] for x in it.get('additionalProperty',[])}
                res[url]=dict(name=it.get('name'),url=url,price=(it.get('offers') or {}).get('price'),cur=(it.get('offers') or {}).get('priceCurrency'),posted=it.get('datePosted'),props=props);new+=1
    print(p,new,len(res),flush=True)
    if new==0: break
    time.sleep(1.2)
json.dump(res,open(out,'w'),ensure_ascii=False); print('done',len(res))
