import re,html,urllib.request,urllib.parse,http.cookiejar,sys,json,time
UA='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
cj=http.cookiejar.CookieJar(); op=urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
URL='https://www.turyap.com.tr/Portfoyler.aspx'
def get(u,data=None):
    req=urllib.request.Request(u,data=urllib.parse.urlencode(data,doseq=True).encode() if data else None,headers={'User-Agent':UA})
    return op.open(req,timeout=60).read().decode('utf-8','ignore')
def form(h):
    f={}
    for m in re.finditer(r'<input[^>]*type="hidden"[^>]*>',h):
        n=re.search(r'name="([^"]*)"',m.group(0)); v=re.search(r'value="([^"]*)"',m.group(0))
        if n: f[n.group(1)]=html.unescape(v.group(1)) if v else ''
    for m in re.finditer(r'<select[^>]*name="([^"]*)"[^>]*>(.*?)</select>',h,re.S):
        sel=re.search(r'<option[^>]*selected="selected"[^>]*value="([^"]*)"|<option[^>]*value="([^"]*)"[^>]*selected',m.group(2))
        if sel: f[m.group(1)]=sel.group(1) or sel.group(2)
    return f
def selects(h):
    return {m.group(1):[(a,html.unescape(b).strip()) for a,b in re.findall(r'<option[^>]*value="([^"]*)"[^>]*>([^<]*)',m.group(2))] for m in re.finditer(r'<select[^>]*name="([^"]*)"[^>]*>(.*?)</select>',h,re.S)}
P='ctl00$ContentPlaceHolder1$'
h=get(URL)
def post(h,target,**vals):
    f=form(h); f.update({P+k:v for k,v in vals.items()}); f['__EVENTTARGET']=target; f['__EVENTARGUMENT']=''
    return get(URL,f)
h=post(h,P+'cmbKategori',cmbKategori=sys.argv[1]); time.sleep(1)
h=post(h,P+'cmbCinsi',cmbKategori=sys.argv[1],cmbCinsi='1'); time.sleep(1)
h=post(h,P+'cmbIL',cmbKategori=sys.argv[1],cmbCinsi='1',cmbIL='34'); time.sleep(1)
sl=selects(h)
print([k for k in sl]); 
for k,v in sl.items():
    if any('Kadıköy' in b or 'KADIKÖY' in b for a,b in v): print(k,[x for x in v if 'ADIK' in x[1].upper()])
open('ty_step.html','w').write(h)
h=post(h,P+'cmbIlce',cmbKategori=sys.argv[1],cmbCinsi='1',cmbIL='34',cmbIlce='572'); time.sleep(1)
sl=selects(h)
print('after ilce',[k for k in sl if k not in (P+'ODA_SAYISI',)])
for k in sl:
    if 'Mah' in k or 'mah' in k or 'Semt' in k: print(k,sl[k][:40])
print('tipi',sl.get(P+'cmbTipi'))
open('ty_step2.html','w').write(h)
def search(h,semt):
    f=form(h)
    f.update({P+'cmbKategori':sys.argv[1],P+'cmbCinsi':'1',P+'cmbIL':'34',P+'cmbIlce':'572',P+'cmbSemt':semt,P+'cmbTipi':'1'})
    if len(sys.argv)>3: f[P+'ODA_SAYISI']=['2','3']; f[P+'BINANIN_YASI']=[str(i) for i in range(1,17)]
    f['__EVENTTARGET']=''; f['__EVENTARGUMENT']=''; f[P+'Button1']='İlan Ara'
    return get(URL,f)
if len(sys.argv)>2:
    r=search(h,sys.argv[2]); open('ty_res.html','w').write(r)
    ids=sorted(set(re.findall(r'Portfoy_Bilgileri.aspx\?ProductID=(\d+)',r))); print('ids',len(ids),ids[:10])
    print(re.findall(r"__doPostBack\(&#39;([^&]*)&#39;,&#39;(Page\$[^&]*)&#39;\)",r)[:10])
    print(re.findall(r'(\d+)\s*(?:adet\s*)?(?:portföy|ilan|sonuç)',r,re.I)[:5])
