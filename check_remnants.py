with open("sirup_exporter.user.js", "r", encoding="utf-8") as f:
    content = f.read()

import re

# Find elif
elifs = re.findall(r"\belif\b", content)
print(f"Found 'elif': {len(elifs)}")

# Find def
defs = re.findall(r"\bdef\b", content)
print(f"Found 'def': {len(defs)}")

# Find append
appends = re.findall(r"\bappend\b", content)
print(f"Found 'append': {len(appends)}")
