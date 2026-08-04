# Read the test Python file and extract WORKSHOP_PAGE boundaries
with open('C:/Users/pc/ZCodeProject/test/工厂工作台.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find WORKSHOP_PAGE start and end
start_idx = None
end_idx = None
for i, line in enumerate(lines):
    if line.startswith('WORKSHOP_PAGE'):
        start_idx = i
    if start_idx and line.strip().endswith("'''") and i > start_idx and 'WORKSHOP_ADMIN' not in line:
        # Check if next line starts with WORKSHOP_ADMIN
        if i+1 < len(lines) and lines[i+1].startswith('WORKSHOP_ADMIN'):
            end_idx = i
            break

if start_idx and end_idx:
    print(f"WORKSHOP_PAGE: lines {start_idx+1} to {end_idx+1}")
    print(f"First line: {repr(lines[start_idx][:100])}")
    print(f"Last line: {repr(lines[end_idx][:100])}")
else:
    print("Could not find WORKSHOP_PAGE boundaries")
