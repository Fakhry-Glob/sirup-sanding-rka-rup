import json
import re

txt_path = r"C:\Users\user\.gemini\antigravity\brain\b120cd73-e77b-45ef-a520-5f577420490e\.system_generated\steps\204\output.txt"

with open(txt_path, "r", encoding="utf-8") as f:
    content = f.read()

match = re.search(r"(\[.*\]|\{.*\})", content.split("### Ran Playwright code")[0], re.DOTALL)
if not match:
    match = re.search(r"(\[.*\]|\{.*\})", content, re.DOTALL)

data = json.loads(match.group(1))

# Print first item's details for each packet
for pkt_id, items in list(data.items())[:5]:
    print(f"Packet {pkt_id}:")
    for item in items[:2]:
        print(f"  - MAK: {item['mak']} | Pagu: {item['pagu']}")
