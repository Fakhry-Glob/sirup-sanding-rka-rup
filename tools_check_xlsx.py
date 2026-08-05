"""Periksa XML mentah di dalam .xlsx — yang tidak tertangkap saat dibaca ulang ExcelJS."""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

path = sys.argv[1] if len(sys.argv) > 1 else 'out.xlsx'
problems = []

with zipfile.ZipFile(path) as z:
    names = z.namelist()
    sheets = sorted(n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n))

    for name in sheets:
        raw = z.read(name).decode('utf-8')

        # 1. XML harus well-formed
        try:
            ET.fromstring(raw)
        except ET.ParseError as e:
            problems.append(f'{name}: XML rusak — {e}')
            continue

        # 2. <pane> beku wajib punya minimal satu split > 0
        for pane in re.findall(r'<pane\b[^>]*/?>', raw):
            xs = re.search(r'xSplit="([^"]+)"', pane)
            ys = re.search(r'ySplit="([^"]+)"', pane)
            xv = float(xs.group(1)) if xs else 0
            yv = float(ys.group(1)) if ys else 0
            state = re.search(r'state="([^"]+)"', pane)
            if xv == 0 and yv == 0:
                problems.append(f'{name}: <pane> tanpa split ({pane})')
            elif state and state.group(1) == 'frozen' and not (xs or ys):
                problems.append(f'{name}: pane frozen tanpa atribut split ({pane})')

        # 3. Warna wajib ARGB 8 digit
        for col in re.findall(r'rgb="([0-9A-Fa-f]+)"', raw):
            if len(col) != 8:
                problems.append(f'{name}: rgb bukan 8 digit: {col}')

        # 4. autoFilter ref harus masuk akal
        for ref in re.findall(r'<autoFilter ref="([^"]+)"', raw):
            if ':' not in ref:
                problems.append(f'{name}: autoFilter ref aneh: {ref}')

    # 5. definedName kosong (dari printTitlesRow undefined)
    if 'xl/workbook.xml' in names:
        wbx = z.read('xl/workbook.xml').decode('utf-8')
        for dn in re.findall(r'<definedName[^>]*>([^<]*)</definedName>', wbx):
            if not dn.strip() or 'undefined' in dn:
                problems.append(f'workbook.xml: definedName kosong/undefined: "{dn}"')
        for stx in sorted(n for n in names if 'styles' in n):
            ET.fromstring(z.read(stx))

print(f'{path}: {len(sheets)} worksheet diperiksa')
if problems:
    for p in problems:
        print('  !!', p)
    sys.exit(1)
print('  OK: XML sah, pane valid, warna 8 digit, definedName bersih')
