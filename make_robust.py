with open("sirup_exporter.user.js", "r", encoding="utf-8") as f:
    content = f.read()

content = content.replace('aktif: row[6] === "true",', 'aktif: row[6] === "true" || row[6] === true,')
content = content.replace('fd: row[7] === "true",', 'fd: row[7] === "true" || row[7] === true,')
content = content.replace('umumkan: row[8] === "true",', 'umumkan: row[8] === "true" || row[8] === true,')

with open("sirup_exporter.user.js", "w", encoding="utf-8") as f:
    f.write(content)

print("Made script boolean checks robust!")
