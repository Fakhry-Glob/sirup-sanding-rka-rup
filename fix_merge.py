with open("sirup_exporter.user.js", "r", encoding="utf-8") as f:
    content = f.read()

# Replace merge_cells with mergeCells
content = content.replace("merge_cells", "mergeCells")

with open("sirup_exporter.user.js", "w", encoding="utf-8") as f:
    f.write(content)

print("Fixed merge_cells to mergeCells successfully!")
