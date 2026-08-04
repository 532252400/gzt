"""Fix run_rc to auto-detect worksheet name"""
filepath = 'C:/Users/pc/ZCodeProject/test/工厂工作台.py'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Find the run_rc function and replace the worksheet selection
old = """    wb=openpyxl.load_workbook(fp,data_only=True)
    # 自动选择工作表\uff1a\u4f18\u5148Sheet2\u3001sheet1\u3001\u6216\u7b2c\u4e00\u4e2a\u8868
    ws = wb['Sheet2'] if 'Sheet2' in wb.sheetnames else (wb['sheet1'] if 'sheet1' in wb.sheetnames else wb[wb.sheetnames[0]])
    left={}"""

new = """    wb=openpyxl.load_workbook(fp,data_only=True)
    # Auto-select worksheet
    if 'Sheet2' in wb.sheetnames:
        ws = wb['Sheet2']
    elif 'sheet1' in wb.sheetnames:
        ws = wb['sheet1']
    else:
        ws = wb[wb.sheetnames[0]]
    left={}"""

if old in content:
    content = content.replace(old, new)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Fixed!")
else:
    print("Could not find the old string")
    # Debug: find what's actually in the file
    idx = content.find("wb=openpyxl.load_workbook(fp,data_only=True)")
    if idx >= 0:
        print(f"Found at position {idx}")
        print(f"Context: {repr(content[idx:idx+300])}")
