import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\Administrator\Desktop\数据报表\李芮计算器\赶集双模式\伯小乐计算器\【双模式】伯小乐版本计算器V1.1.xlsx', data_only=True)

def num(v):
    if v is None: return 0
    if isinstance(v, (int, float)): return int(v)
    try: return int(float(str(v)))
    except: return 0

def s(v):
    return str(v).strip() if v else ''

sku58 = [
    ('58_00001','58_00002','58_00003','58_00007','58_00005','58_00006'),
    ('58_00009','58_00010','58_00011','58_00015','58_00013','58_00014'),
    ('58_00009','58_00017','58_00011','58_00015','58_00013','58_00014'),
    ('58_00009','58_00017','58_00011','58_00015','58_00013','58_00014'),
    ('58_00009','58_00017','','58_00015','58_00013','58_00014'),
    ('58_00001','58_00002','58_00003','58_00007','58_00013','58_00014'),
    ('58_00009','58_00002','58_00003','58_00015','58_00013','58_00014'),
    ('58_00009','58_00018','58_00011','58_00015','58_00013','58_00014'),
    ('58_00009','58_00019','58_00011','58_00015','58_00013','58_00014'),
    ('58_00001','58_00020','58_00003','58_00007','58_00013','58_00014'),
]

sku_gj = [
    ('','GJ_00001','GJ_00002','','',''),
    ('58_00009','GJ_00004','GJ_00005','58_00015','58_00013','58_00014'),
    ('58_00009','GJ_00009','GJ_00005','58_00015','58_00013','58_00014'),
    ('58_00009','GJ_00009','GJ_00005','58_00015','58_00013','58_00014'),
    ('58_00009','GJ_00009','GJ_00005','58_00015','58_00013','58_00014'),
    ('58_00009','GJ_00009','GJ_00005','58_00015','58_00013','58_00014'),
]

sections_58 = [
    (2, 3, 3, True, '套餐季度（92天）'),
    (6, 7, 9, True, '套餐年度（365天-限类限城）'),
    (12, 13, 20, True, '套餐年度（365天-全国）'),
    (23, 24, 27, False, '定制/框架年度'),
    (30, 31, 32, False, '定制/框架年度（精简版）'),
    (35, 36, 38, True, '定制/框架季度（92天）'),
    (41, 42, 43, True, '定制/框架-美团特供2.0'),
    (46, 47, 48, False, '定制/框架年度（容器）'),
    (51, 52, 56, False, '定制/框架年度（容器-行业限类）'),
    (59, 60, 61, True, '定制/框架年度（容器-行业限类-92天）'),
]

sections_gj = [
    (2, 3, 9, True, '个人超值月度（30天）'),
    (11, 12, 13, True, '套餐年度（365天-限类限城）'),
    (16, 17, 21, True, '套餐年度（365天-全国）'),
    (24, 25, 28, True, '套餐年度（365天-简历宝）'),
    (31, 32, 34, True, '套餐年度（365天-行业版）'),
    (38, 39, 39, True, '定制/框架'),
]

ws58 = wb['58伯小乐计算器']
wsGJ = wb['赶集伯小乐计算器']

# Generate get58Data
lines58 = ['function get58Data(){ return {\n  sections: [']
for si, (base_row, d1, d2, use_k, sec_name) in enumerate(sections_58):
    bc = [ws58.cell(row=base_row, column=c).value for c in range(1, 17)]
    sc, sd, se, sf, sg, sh = sku58[si]
    lines58.append(f'    {{name:"{sec_name}",base:{{C:{num(bc[3])},D:{num(bc[4])},E:{num(bc[5])},F:{num(bc[6])},G:{num(bc[7])},H:{num(bc[8])},skuC:\'{sc}\',skuD:\'{sd}\',skuE:\'{se}\',skuF:\'{sf}\',skuG:\'{sg}\',skuH:\'{sh}\'}},packages:[')
    for dr in range(d1, d2+1):
        cells = [ws58.cell(row=dr, column=c).value for c in range(1, 17)]
        sp = cells[10] if use_k else cells[9]
        pn = s(cells[1])
        ln = s(cells[12]) or '\u65e0'
        nn = s(cells[14])
        jv = f'J:{num(sp)}' if sp else 'J:null'
        # Escape strings for JS
        pn_esc = pn.replace('\\', '\\\\').replace("'", "\\'")
        ln_esc = ln.replace('\\', '\\\\').replace("'", "\\'")
        nn_esc = nn.replace('\\', '\\\\').replace("'", "\\'")
        name_esc = s(cells[0]).replace('\\', '\\\\').replace("'", "\\'")
        lines58.append(f"      {{name:'{name_esc}',dur:{num(cells[2])},C:{num(cells[3])},D:{num(cells[4])},E:{num(cells[5])},F:{num(cells[6])},G:{num(cells[7])},H:{num(cells[8])},{jv},L:'{ln_esc}',M:{num(cells[13])},N:'{nn_esc}',productName:'{pn_esc}'}},")
    lines58.append('    ]},')
lines58.append('  ]};')
lines58.append('}')

# Generate getGanjiData
linesGJ = ['function getGanjiData(){ return {\n  sections: [']
for si, (base_row, d1, d2, use_k, sec_name) in enumerate(sections_gj):
    bc = [wsGJ.cell(row=base_row, column=c).value for c in range(1, 17)]
    sc, sd, se, sf, sg, sh = sku_gj[si]
    linesGJ.append(f'    {{name:"{sec_name}",base:{{C:{num(bc[3])},D:{num(bc[4])},E:{num(bc[5])},F:{num(bc[6])},G:{num(bc[7])},H:{num(bc[8])},skuC:\'{sc}\',skuD:\'{sd}\',skuE:\'{se}\',skuF:\'{sf}\',skuG:\'{sg}\',skuH:\'{sh}\'}},packages:[')
    for dr in range(d1, d2+1):
        cells = [wsGJ.cell(row=dr, column=c).value for c in range(1, 17)]
        sp = cells[10] if use_k else cells[9]
        pn = s(cells[1])
        if not pn:
            continue  # skip empty rows
        ln = s(cells[12]) or '\u65e0'
        nn = s(cells[14])
        jv = f'J:{num(sp)}' if sp else 'J:null'
        pn_esc = pn.replace('\\', '\\\\').replace("'", "\\'")
        ln_esc = ln.replace('\\', '\\\\').replace("'", "\\'")
        nn_esc = nn.replace('\\', '\\\\').replace("'", "\\'")
        name_esc = s(cells[0]).replace('\\', '\\\\').replace("'", "\\'")
        linesGJ.append(f"      {{name:'{name_esc}',dur:{num(cells[2])},C:{num(cells[3])},D:{num(cells[4])},E:{num(cells[5])},F:{num(cells[6])},G:{num(cells[7])},H:{num(cells[8])},{jv},L:'{ln_esc}',M:{num(cells[13])},N:'{nn_esc}',productName:'{pn_esc}'}},")
    linesGJ.append('    ]},')
linesGJ.append('  ]};')
linesGJ.append('}')

# Write to files
with open('gen_58data.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines58))

with open('gen_gjdata.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(linesGJ))

print('Done. Generated gen_58data.txt and gen_gjdata.txt')
print(f'58: {len(lines58)} lines, GJ: {len(linesGJ)} lines')

# Verify: count packages
import re
count58 = sum(1 for l in lines58 if l.strip().startswith('{name:'))
countGJ = sum(1 for l in linesGJ if l.strip().startswith('{name:'))
print(f'58 packages: {count58}, GJ packages: {countGJ}')
