import re, requests, sys
H={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
for u in ["https://betting.betfair.com/horse-racing/","https://betting.betfair.com/horse-racing/tips/timeform/",
          "https://www.sportinglife.com/racing/results/2026-11-21","https://www.sportinglife.com/racing/fast-results"]:
    try:
        r=requests.get(u,headers=H,timeout=30); print("==",u,r.status_code,len(r.text))
        links=sorted(set(re.findall(r'href="([^"]*(?:timeform|tips)[^"]*)"',r.text)))[:40]
        for l in links: print("  L",l)
        if "sportinglife" in u:
            m=re.search(r'__NEXT_DATA__[^>]*>(.{0,1500})',r.text,re.S); print("  NEXT",bool(m), (m.group(1)[:800] if m else r.text[:500]))
    except Exception as ex: print("ERR",u,ex)
