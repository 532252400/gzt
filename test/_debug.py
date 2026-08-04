path = r"C:\Users\pc\ZCodeProject\test\工厂工作台.py"
with open(path, "r", encoding="utf-8") as f:
    content = f.read()

# Add debug printing at the start of do_POST
old = "def do_POST(self):\n        try:\n            ct = self.headers.get('Content-Type','')"
new = """def do_POST(self):
        try:
            ct = self.headers.get('Content-Type','')
            print(f'[POST] CT={ct[:80]}', flush=True)
            print(f'[POST] CL={self.headers.get(\"Content-Length\",\"?\")}', flush=True)"""

content = content.replace(old, new)

# Also add debug after reading body
old2 = "b = self.rfile.read(int(self.headers['Content-Length']))"
new2 = "b = self.rfile.read(int(self.headers['Content-Length'])); print(f'[POST] read {len(b)} bytes, first 100: {b[:100]}', flush=True)"
content = content.replace(old2, new2)

# Debug after boundary detection
old3 = "if not boundary: return self._json({'status':'error','message':'No boundary'})"
new3 = "if not boundary: print('[POST] NO boundary found in CT', flush=True); return self._json({'status':'error','message':'No boundary'})"
content = content.replace(old3, new3)

# Debug parts count
old4 = "parts = b.split(('--'+bnd).encode())"
new4 = "parts = b.split(('--'+bnd).encode()); print(f'[POST] boundary={bnd}, parts={len(parts)}', flush=True)"
content = content.replace(old4, new4)

# Debug action found
old5 = "if key == 'action': action = val"
new5 = "if key == 'action': action = val; print(f'[POST] action={action}', flush=True)"
content = content.replace(old5, new5)

# Debug final dispatch
old6 = "func = {'lbl100':run_lbl100"
new6 = "print(f'[POST] dispatching action={action}', flush=True)\n            func = {'lbl100':run_lbl100"
content = content.replace(old6, new6)

with open(path, "w", encoding="utf-8") as f:
    f.write(content)
print("Debug added, size:", len(content))
