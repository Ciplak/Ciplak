import re,sys,time,json,urllib.request
UA='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
deal=sys.argv[1]; out=sys.argv[2]; urls=set()
for p in range(1,200):
    u=f'https://www.century21.com.tr/{deal}/konut'+(f'?pager_p={p}' if p>1 else '')
    try: h=urllib.request.urlopen(urllib.request.Request(u,headers={'User-Agent':UA}),timeout=40).read().decode('utf-8','ignore')
    except Exception as e: print('err',p,e,flush=True); time.sleep(5); continue
    found=set(re.findall(r'href="(/[a-z0-9-]+-(?:satilik|kiralik)/[a-z0-9-]+/\d+)"',h))
    new=found-urls; urls|=found; print(p,len(new),len(urls),flush=True)
    if not new: break
    time.sleep(1.0)
json.dump(sorted(urls),open(out,'w')); print('done',len(urls))
