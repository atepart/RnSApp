import requests

url = "https://github.com/atepart/RnSApp/releases/download/new107b3/RnSApp_macOS_arm64_new107b3.zip"
resp = requests.get(url, stream=True, timeout=10)
print("status_code:", resp.status_code)
print("content-length:", resp.headers.get("content-length", 0))
print("headers:", resp.headers)
