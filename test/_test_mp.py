import requests

# Test with multipart but no file (like browser's FormData without file)
url = "http://127.0.0.1:8933/run"
# Use files param with empty bytes to force multipart
r = requests.post(url, files={"action": (None, "start_job"), "batch_name": (None, "86"), "worker": (None, "test")}, timeout=5)
print("Multipart no-file:", r.status_code, r.text[:200])
