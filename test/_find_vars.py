with open(r"C:\Users\pc\ZCodeProject\test\工厂工作台.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "WORKSHOP_BOARD" in line and "=" in line:
        print(f"Line {i+1}: {line.rstrip()[:120]}")
    if "WORKSHOP_ADMIN" in line and "=" in line:
        print(f"Line {i+1}: {line.rstrip()[:120]}")
