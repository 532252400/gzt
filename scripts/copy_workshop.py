"""Copy WORKSHOP_PAGE from stable version to test version"""
import re

stable_file = 'C:/Users/pc/ZCodeProject/server/工厂工作台.py'
test_file = 'C:/Users/pc/ZCodeProject/test/工厂工作台.py'

# Read stable version
with open(stable_file, 'r', encoding='utf-8') as f:
    stable_content = f.read()

# Extract WORKSHOP_PAGE from stable
start = stable_content.find("WORKSHOP_PAGE = '''")
end = stable_content.find("\nWORKSHOP_ADMIN = '''", start)
if start < 0 or end < 0:
    print("Could not find WORKSHOP_PAGE in stable version")
    exit(1)

# Include the newline before WORKSHOP_ADMIN
stable_ws = stable_content[start:end+1]  # Include the newline

print(f"Stable WORKSHOP_PAGE length: {len(stable_ws)} chars")
print(f"First 100 chars: {repr(stable_ws[:100])}")

# Read test version
with open(test_file, 'r', encoding='utf-8') as f:
    test_content = f.read()

# Find WORKSHOP_PAGE in test version
test_start = test_content.find("WORKSHOP_PAGE = '''")
test_end = test_content.find("\nWORKSHOP_ADMIN = '''", test_start)
if test_start < 0 or test_end < 0:
    print("Could not find WORKSHOP_PAGE in test version")
    exit(1)

# Replace
new_content = test_content[:test_start] + stable_ws + test_content[test_end+1:]

with open(test_file, 'w', encoding='utf-8') as f:
    f.write(new_content)

print("WORKSHOP_PAGE copied from stable to test successfully!")
