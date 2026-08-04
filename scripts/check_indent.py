with open('C:/Users/pc/ZCodeProject/test/工厂工作台.py','r',encoding='utf-8') as f:
    lines = f.readlines()
# Check for tab/space issues around line 1580-1590
for i in range(1575, 1590):
    line = lines[i-1]
    if '\t' in line:
        print(f'Line {i}: has TAB')
    print(f'Line {i}: {repr(line[:80])}')
