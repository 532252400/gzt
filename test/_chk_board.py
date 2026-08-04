with open(r"C:\Users\pc\ZCodeProject\test\工厂工作台.py", "r", encoding="utf-8") as f:
    c = f.read()

# Find WORKSHOP_BOARD section
wsb_start = c.find("WORKSHOP_BOARD = '''") + 20
wsb_end = c.find("WORKSHOP_ADMIN = '''")
wsb = c[wsb_start:wsb_end]

print("BOARD section:", len(wsb), "chars")
print("peopleInput in BOARD:", "peopleInput" in wsb)
print("startJob in BOARD:", "startJob" in wsb)
print("setPriority in BOARD:", "setPriority" in wsb)
print("start_job in BOARD:", "start_job" in wsb)

# Check if the JS functions exist
for func in ["startJob", "setPriority", "loadBoard", "fetch"]:
    print(f"  {func}:", func in wsb)
