import urllib.request
import json
import time

TOKEN = 'YKtVKOIprzQoLKqB7foUkyxIwvzGaWxY/lnBmm4GaoJVNVDgbEUOTs8MOZRWBtEfzX8X6k0pX+pJSyave60Ka//baM6waKsQE/Ho43TkMod6YcyLcreDpjVC85MCXv7NxSj47Bh6bI2a2Xuls5hnkAdB04t89/1O/w1cDnyilFU='
WEBHOOK_URL = 'https://webhook.site/token/d43cd402-b87b-4c7f-a8a8-8e58b0cc27ba/requests'

def check_group_id():
    req = urllib.request.Request(WEBHOOK_URL, headers={'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json'})
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode())
            for item in data.get('data', []):
                raw = item.get('content')
                if not raw:
                    continue
                try:
                    payload = json.loads(raw)
                    events = payload.get('events', [])
                    for ev in events:
                        source = ev.get('source', {})
                        if source.get('type') == 'group':
                            gid = source.get('groupId')
                            print(f"FOUND_GROUP_ID:{gid}")
                            return gid
                except Exception:
                    pass
    except Exception as e:
        print("Error fetching requests:", e)
    return None

if __name__ == '__main__':
    gid = check_group_id()
    if not gid:
        print("WAITING_FOR_MESSAGE")
