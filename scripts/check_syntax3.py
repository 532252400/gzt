import py_compile, sys
try:
    py_compile.compile('C:/Users/pc/ZCodeProject/test/工厂工作台.py', doraise=True)
    print('OK')
except py_compile.PyCompileError as e:
    print(f'LINE {e.lineno}: {e.msg}')
except Exception as e:
    print(f'Error: {e}')
