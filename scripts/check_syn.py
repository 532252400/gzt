import sys
filepath = 'C:/Users/pc/ZCodeProject/test/工厂工作台.py'
with open(filepath, 'r', encoding='utf-8') as f:
    source = f.read()
try:
    compile(source, 'test', 'exec')
    print('OK')
except SyntaxError as e:
    print(f'Line {e.lineno}: {e.msg}')
    lines = source.split('\n')
    if 0 < e.lineno <= len(lines):
        print(f'Content: {repr(lines[e.lineno-1][:100])}')
