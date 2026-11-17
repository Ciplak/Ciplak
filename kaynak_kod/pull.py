import json,sys,time,urllib.request,urllib.parse
base,out,maxp=sys.argv[1],sys.argv[2],int(sys.argv[3])
items=[];p=1;d={'total':0}
while p<=maxp:
    u=base+f"&pageSize=60&page={p}"
    for a in range(4):
        try:
            r=urllib.request.urlopen(urllib.request.Request(u,headers={'User-Agent':'Mozilla/5.0','Accept':'application/json'}),timeout=40)
            d=json.load(r);break
        except Exception as e:
            print('retry',p,e,file=sys.stderr);time.sleep(5*(a+1))
    else: break
    items+=d['items']
    if p*d['pageSize']>=d['total'] or not d['items']: break
    p+=1;time.sleep(1.2)
json.dump(items,open(out,'w'),ensure_ascii=False)
print(out,len(items),d['total'])
