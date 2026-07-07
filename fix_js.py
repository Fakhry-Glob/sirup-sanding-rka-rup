with open("sirup_exporter.user.js", "r", encoding="utf-8") as f:
    content = f.read()

# Fix border_cell to border_thin
content = content.replace("border_cell", "border_thin")

# Fix curr_row = 6 to let curr_row = 6
content = content.replace("curr_row = 6;", "let curr_row = 6;")

# Save back
with open("sirup_exporter.user.js", "w", encoding="utf-8") as f:
    f.write(content)

print("Fixed userscript successfully!")
