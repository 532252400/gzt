import requests

# Simulate browser-style FormData POST without file
url = "http://127.0.0.1:8933/run"
data = {"action": "start_job", "batch_name": "86", "worker": "test"}
r = requests.post(url, data=data, timeout=5)
print("Form POST:", r.status_code, r.text[:100])

# Simulate FormData with file (like browser)
files = {"file": ("test.txt", b"hello", "text/plain")}
data2 = {"action": "ca"}
r2 = requests.post(url, files=files, data=data2, timeout=5)
print("File POST:", r2.status_code, r2.text[:100])
