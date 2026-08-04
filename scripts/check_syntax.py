try:
    compile(open('C:/Users/pc/ZCodeProject/server/工厂工作台.py','r',encoding='utf-8').read(), 'test.py', 'exec')
    print('OK')
except SyntaxError as e:
    print(f'Line {e.lineno}: {e.msg}')
    if e.text:
        print(f'Text: {repr(e.text.strip()[:100])}')
