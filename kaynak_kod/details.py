import json,sys,time,os,urllib.request
ids=json.load(open(sys.argv[1])); out=sys.argv[2]
have=json.load(open(out)) if os.path.exists(out) else {}
for n,i in enumerate(ids):
    if str(i) in have: continue
    for a in range(4):
        try:
            r=urllib.request.urlopen(urllib.request.Request(f'https://spitimou.net/api/listings/{i}',headers={'User-Agent':'Mozilla/5.0','Accept':'application/json'}),timeout=40)
            have[str(i)]=json.load(r);break
        except Exception as e:
            print('retry',i,e,flush=True);time.sleep(8*(a+1))
    if n%25==0:
        json.dump(have,open(out,'w'),ensure_ascii=False);print(n,len(have),flush=True)
    time.sleep(0.9)
json.dump(have,open(out,'w'),ensure_ascii=False);print('done',len(have))
