# -*- coding: utf-8 -*-
path = r'D:\u684c\u9762\u5de5\u5382\u5de5\u4f5c\u53f0.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

idx = content.find('openComplete')
print('Found at:', idx)
if idx > 0:
    print(content[idx-30:idx+100])
