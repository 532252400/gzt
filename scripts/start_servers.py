import subprocess, sys, os, time, socket

SERVERS = [
    (r"C:\Users\pc\AppData\Local\Python\bin\python.exe",
     r"C:\Users\pc\ZCodeProject\server\工厂工作台.py",
     r"C:\Users\pc\ZCodeProject\server", 8932),
    (r"C:\Program Files\AutoClaw\resources\python\python.exe",
     r"C:\Users\pc\ZCodeProject\test\工厂工作台.py",
     r"C:\Users\pc\ZCodeProject\test", 8933),
]

def is_port_open(port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.connect(("127.0.0.1", port))
        s.close()
        return True
    except:
        return False

procs = []
for py, script, cwd, port in SERVERS:
    if is_port_open(port):
        print(f"Port {port} already in use, skipping")
        continue
    p = subprocess.Popen(
        [py, script],
        cwd=cwd,
        creationflags=subprocess.CREATE_NO_WINDOW,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    procs.append((port, p))
    print(f"Started server on port {port} (PID {p.pid})")
    time.sleep(2)

if not procs:
    print("All servers already running")
    sys.exit(0)

print("Servers running. Press Ctrl+C to stop.")
try:
    for _, p in procs:
        p.wait()
except KeyboardInterrupt:
    for _, p in procs:
        p.terminate()
    print("Stopped")
