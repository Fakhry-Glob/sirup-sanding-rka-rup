with open("sirup_exporter.user.js", "r", encoding="utf-8") as f:
    lines = f.readlines()

for idx, line in enumerate(lines):
    if "merge_cells" in line:
        print(f"Line {idx+1}: {line.strip()}")
