"""Fix all sync connections in test version"""
filepath = 'C:/Users/pc/ZCodeProject/test/工厂工作台.py'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix first boardSync postMessage in savePeople to use syncChannel
count = 0
pos = 0
first_pos = None
while True:
    idx = content.find('try { boardSync.postMessage', pos)
    if idx < 0:
        break
    count += 1
    if count == 1:
        first_pos = idx
    pos = idx + 1

if first_pos:
    # Check context - if before the kanban board section, use syncChannel
    # Find the board section start
    board_start = content.find("WORKSHOP_BOARD = '''")
    if first_pos < board_start:
        # Before board section - workshop page
        content = content[:first_pos] + 'try { syncChannel.postMessage' + content[first_pos + len('try { boardSync.postMessage'):]
        print(f"Fixed workshop savePeople (position {first_pos}, before board)")
    else:
        print(f"First occurrence is in board section at {first_pos}, not changing")

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
print("Done!")
