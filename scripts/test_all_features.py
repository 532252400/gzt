import urllib.request, json, random, os, openpyxl

# Create test Excel
wb = openpyxl.Workbook()
ws1 = wb.active
ws1.title = "Data"
ws1.cell(1,1,"货件单号"); ws1.cell(1,2,"物流中心编码"); ws1.cell(1,3,"物流商")
ws1.cell(1,4,"物流渠道"); ws1.cell(1,5,"物流商单号"); ws1.cell(1,7,"国家"); ws1.cell(1,8,"货件单号")
ws1.cell(2,1,"SHIP001"); ws1.cell(2,2,"东北部中心1"); ws1.cell(2,3,"FedEx")
ws1.cell(2,4,"Ground"); ws1.cell(2,5,"FDX123"); ws1.cell(2,7,"US"); ws1.cell(2,8,"EXTRA")

for i in range(2,6):
    wb.create_sheet(f"Extra{i}")

ws5 = wb["Extra5"]
ws5.cell(1,1,"货件单号"); ws5.cell(1,2,"箱数"); ws5.cell(1,3,"重量"); ws5.cell(1,4,"体积"); ws5.cell(1,5,"类型")
ws5.cell(2,1,"SHIP001"); ws5.cell(2,2,10); ws5.cell(2,3,100.5); ws5.cell(2,4,2.5); ws5.cell(2,5,"A")

test_file = "C:/Users/pc/ZCodeProject/test/uploads/test_features.xlsx"
wb.save(test_file)

def test_action(action):
    boundary = '----' + str(random.randint(100000,999999))
    body = bytearray()
    with open(test_file, 'rb') as f:
        file_data = f.read()
    body.extend(f'--{boundary}\r\n'.encode())
    body.extend(f'Content-Disposition: form-data; name="file"; filename="test.xlsx"\r\n'.encode())
    body.extend(b'Content-Type: application/octet-stream\r\n\r\n')
    body.extend(file_data); body.extend(b'\r\n')
    body.extend(f'--{boundary}\r\n'.encode())
    body.extend(f'Content-Disposition: form-data; name="action"\r\n\r\n'.encode())
    body.extend(f'{action}\r\n'.encode())
    body.extend(f'--{boundary}--\r\n'.encode())
    
    req = urllib.request.Request('http://localhost:8933/run', data=bytes(body))
    req.add_header('Content-Type', f'multipart/form-data; boundary={boundary}')
    try:
        resp = urllib.request.urlopen(req, timeout=15)
        result = json.loads(resp.read())
        s = result.get("status", "?")
        m = str(result.get("message", ""))[:80]
        print(f'{action:10s}: status={s}, msg={m}')
    except Exception as e:
        print(f'{action:10s}: ERROR - {e}')

# Test all actions
test_action('us')
test_action('ca')
test_action('rc')
test_action('lbl100')
test_action('lbl30')

os.remove(test_file)
print('Done')
