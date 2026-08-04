path = r'D:\桌面\工厂工作台.py'
with open(path, 'r', encoding='utf-8') as f:
    c = f.read()
idx = c.find('openComplete(this)')
if idx > 0:
    print(repr(c[idx-5:idx+80]))
