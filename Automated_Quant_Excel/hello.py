import urllib.request, json

KEY = "AIzaSyCK4AO3TDOIsqZImB36nKH41PLxKmnKX7k"
url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={KEY}"
body = json.dumps({"contents": [{"parts": [{"text": "Say hello in one sentence."}]}]}).encode()
req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"}, method="POST")

with urllib.request.urlopen(req) as r:
    print(json.loads(r.read())["candidates"][0]["content"]["parts"][0]["text"])