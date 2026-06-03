import requests

url = "https://github.com/atepart/RnSApp/releases/download/new107b3/RnSApp_macOS_arm64_new107b3.zip"
resp = requests.get(url, stream=True, timeout=10)
print(resp.headers)
