import requests
import os

url = "https://github.com/StellarCN/scp_zh/raw/master/fonts/SimHei.ttf"
dest = "SimHei.ttf"

print(f"Downloading {url}...")
r = requests.get(url, allow_redirects=True)
if r.status_code == 200:
    with open(dest, 'wb') as f:
        f.write(r.content)
    print(f"Downloaded {len(r.content)} bytes.")
    if len(r.content) < 100000:
        print("Warning: File too small, likely HTML.")
else:
    print(f"Failed: {r.status_code}")
