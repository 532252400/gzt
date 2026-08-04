#!/bin/bash
# Fix corrupted UTF-8 bytes in the Python file

file="D:\\桌面\\工厂工作台.py"

# Create a Python script to fix the file
python_script=$(cat << 'PYEOF'
import sys
file = r'D:\桌面\工厂工作台.py'
with open(file, 'rb') as f:
    data = f.read()

# Fix corrupted bytes:
# E2 9C 3F -> E2 9C 94 (✔, checkmark) - in button contexts  
# E2 9C 3F -> E2 9C 96 (✖, cancel X) - near cancel buttons
# E9 A1 3F -> E9 A1 B9 (项)

# Count corrupted sequences
count_3f = data.count(b'\xe2\x9c\x3f')
count_a1 = data.count(b'\xe9\xa1\x3f')
print(f"Found {count_3f} corrupted E2 9C 3F sequences")
print(f"Found {count_a1} corrupted E9 A1 3F sequences")

# Fix all E2 9C 3F -> E2 9C 94 (they're all checkmarks/crosses in this context)
data = data.replace(b'\xe2\x9c\x3f', b'\xe2\x9c\x94')

# Fix all E9 A1 3F -> E9 A1 B9 (项)
data = data.replace(b'\xe9\xa1\x3f', b'\xe9\xa1\xb9')

with open(file, 'wb') as f:
    f.write(data)
print("Fixed!")
PYEOF
)

echo "$python_script" > "D:\\fix_utf8_final.py"
python3 "D:\\fix_utf8_final.py" 2>/dev/null || python "D:\\fix_utf8_final.py" 2>/dev/null || echo "Python not available"
