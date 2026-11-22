import re, requests, json
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
def nd(u):
    r=requests.get(u,headers=H,timeout=30); print("==",u,r.status_code,len(r.text))
    m=re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>',r.text,re.S)
    return json.loads(m.group(1)) if m else None
d=nd("https://www.sportinglife.com/racing/tips-centre")
pp=d["props"]["pageProps"]; print(list(pp.keys()))
s=json.dumps(pp)[:3000]; print(s)
d=nd("https://www.sportinglife.com/tips/tipping-records")
pp=d["props"]["pageProps"] if d else {}; print(list(pp.keys())); print(json.dumps(pp)[:2500])
