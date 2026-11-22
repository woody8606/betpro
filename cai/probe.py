import re, requests, json
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
def nd(u):
    r=requests.get(u,headers=H,timeout=30); print("==",u,r.status_code,len(r.text))
    m=re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>',r.text,re.S)
    return json.loads(m.group(1)) if m else None
d=nd("https://www.sportinglife.com/racing/news/x/235577")
pp=d["props"]["pageProps"]; print(list(pp.keys()))
a=pp.get("article") or pp
s=json.dumps(a)
for w in re.findall(r'\{"keys": \[[^\]]*\], "widget_type": "Racerunner"\}', s)[:5]: print("RR",w[:1500])
i=s.find("Racerunner"); print(s[max(0,i-2500):i+300])
# results page structure for one race with winner
d=nd("https://www.sportinglife.com/racing/results/2026-11-21")
m=d["props"]["pageProps"]["meetings"][0]["races"][0]; print(json.dumps(m)[:2500])
