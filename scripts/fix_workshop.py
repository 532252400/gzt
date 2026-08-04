# -*- coding: utf-8 -*-
import re

path = r'D:\桌面\工厂工作台.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: WORKSHOP_PAGE - openComplete button with data attributes
old1 = "else if(i.status==='processing'){btns='<button class=\"btn-c\" onclick=\"openComplete('+i.id+',\\''+i.sku+'\\')\">\u2714 \u5b8c\u6210\u52a0\u5de5</button>';}"
new1 = "else if(i.status==='processing'){btns='<button class=\"btn-c\" data-id=\"'+i.id+'\" data-sku=\"'+i.sku+'\" onclick=\"openComplete(this)\">\u2714 \u5b8c\u6210\u52a0\u5de5</button>';}"

count1 = content.count(old1)
print(f"Pattern 1 found: {count1} times")

if count1 > 0:
    content = content.replace(old1, new1)
    print("Fixed pattern 1")
else:
    # Debug: show surrounding context
    idx = content.find('openComplete(')
    if idx >= 0:
        print(f"Found 'openComplete(' at position {idx}")
        print("Context:", repr(content[idx-50:idx+100]))
    else:
        print("'openComplete(' not found!")

# Fix 2: openComplete function signature
old2 = "function openComplete(id,sku){curItemId=id;document.getElementById('mSku').textContent=sku;document.getElementById('mQty').value='';document.getElementById('modal').style.display='flex';}"
new2 = "function openComplete(btn){curItemId=btn.dataset.id;document.getElementById('mSku').textContent=btn.dataset.sku;document.getElementById('mQty').value='';document.getElementById('modal').style.display='flex';}"

count2 = content.count(old2)
print(f"Pattern 2 found: {count2} times")

if count2 > 0:
    content = content.replace(old2, new2)
    print("Fixed pattern 2")
else:
    idx = content.find('function openComplete(')
    if idx >= 0:
        print(f"Found 'function openComplete' at position {idx}")
        print("Context:", repr(content[idx:idx+200]))

# Write back
with open(path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Done")
