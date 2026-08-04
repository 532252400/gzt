import urllib.request, json
data = urllib.request.urlopen('http://127.0.0.1:8932/workshop_data', timeout=10).read().decode('utf-8')
items = json.loads(data)
statuses = {}
for i in items:
    s = i.get('status', '?')
    statuses[s] = statuses.get(s, 0) + 1
print(f'Total items: {len(items)}')
for s, c in sorted(statuses.items()):
    print(f'  {s}: {c}')
