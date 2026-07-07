import json
import re

txt_path = r"C:\Users\user\.gemini\antigravity\brain\b120cd73-e77b-45ef-a520-5f577420490e\.system_generated\steps\266\output.txt"

with open(txt_path, "r", encoding="utf-8") as f:
    content = f.read()

parts = content.split("### Ran Playwright code")
match = re.search(r"(\[.*\]|\{.*\})", parts[0], re.DOTALL)
data = json.loads(match.group(1))

first_row = data["aaData"][0]
print(f"Row length: {len(first_row)}")
for idx, val in enumerate(first_row):
    print(f"Index {idx}: {val}")
