import urllib.request, re
url = 'http://127.0.0.1:8932/workshop_board'
data = urllib.request.urlopen(url, timeout=10).read().decode('utf-8')
m = re.search(r'<script>(.*?)</script>', data, re.DOTALL)
if m:
    js = m.group(1)
    lines = js.split('\n')
    for i, line in enumerate(lines):
        opens = line.count('{')
        closes = line.count('}')
        if opens != closes and len(line.strip()) > 0:
            print(f'Line {i+1}: Unbalanced ({{={opens}, }}={closes}): {line.strip()[:80]}')
    print(f'Total: {len(lines)} lines, {js.count("{")} open, {js.count("}")} close')
    if js.count('{') != js.count('}'):
        print('ERROR: Unbalanced braces!')
    else:
        print('OK: Balanced braces')
else:
    print('No script found')
