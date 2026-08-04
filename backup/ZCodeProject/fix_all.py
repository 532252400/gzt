# -*- coding: utf-8 -*-
import re

path = r'D:\桌面\工厂工作台.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: Remove leftover junk lines (WORKSHOP_PAGE)
old_junk1 = "fetch('/job_batches')" + "\u2716" + " '+d.message);}\n\t    catch(e){alert('" + "\u2716" + " '+e.message);}\n\t}\n\tfetch('/job_batches')"
new_clean1 = "\tfetch('/job_batches')"

# The junk was introduced by the edit - it's after cancelJob and before the real fetch
# Let me use a simpler approach - find the exact pattern

# Fix 2: Repair corrupted characters
content = content.replace('鉂?', '\u2716')
content = content.replace('路 ', '路')  # This might be the bullet
# Actually let me check what the original was

print("Looking for corrupted patterns...")

# Find the problematic area
idx = content.find("fetch('/job_batches')" + "\u2716")
if idx >= 0:
    print(f"Found corrupted pattern at {idx}")
    # Find the end of the junk
    end = content.find("fetch('/job_batches').then", idx)
    if end > idx:
        # Replace the junk between cancelJob's closing brace and the real fetch
        junk_end = content.find("}", idx-5)  # Find closing brace of cancelJob
        if junk_end > 0 and junk_end < idx:
            print(f"Junk between {junk_end} and {end}")
        else:
            print(f"No clean boundary found")
else:
    print("No corrupted pattern found")

# Let me look at the raw bytes around the problematic area
for i in range(len(content)-10):
    if content[i:i+5] == "fetch":
        # Check if this is the corrupted one
        after = content[i+5:i+30]
        if "\u2716" in after or '鉂' in after:
            print(f"Corrupted fetch at {i}: {repr(content[i:i+60])}")

print("Done scanning")
