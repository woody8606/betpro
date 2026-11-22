import requests, json
r=requests.get("https://data.ny.gov/resource/6nbc-h7bj.json",params={"$order":"draw_date DESC","$limit":5},timeout=30)
print(r.status_code); print(r.text[:1500])
