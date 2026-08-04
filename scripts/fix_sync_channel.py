"""Fix workshop page savePeople to use syncChannel instead of boardSync"""
filepath = 'C:/Users/pc/ZCodeProject/server/工厂工作台.py'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Find the FIRST occurrence of boardSync in savePeople (workshop page)
# It should be after loadJobs and before cancelJob
# Count occurrences to find the first one
count = 0
pos = 0
while True:
    idx = content.find('try { boardSync.postMessage', pos)
    if idx < 0:
        break
    count += 1
    if count == 1:  # First occurrence - workshop page
        # Replace with syncChannel
        content = content[:idx] + 'try { syncChannel.postMessage' + content[idx + len('try { boardSync.postMessage'):]
        print(f"Fixed occurrence #{count} at position {idx}")
        break
    pos = idx + 1

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
print("Done!")
