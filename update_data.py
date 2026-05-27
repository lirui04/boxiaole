# -*- coding: utf-8 -*-
"""Update HTML tool data from V1.1.xlsx.
Parse existing HTML data, match to V1.1 rows, apply updates.
"""

import json, re

# Load V1.1 extracted data
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\data58.json','r',encoding='utf-8') as f:
    data58 = json.load(f)
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\datagj.json','r',encoding='utf-8') as f:
    datagj = json.load(f)

# Load HTML
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\伯小乐组合包定价工具.html','r',encoding='utf-8') as f:
    html = f.read()

def D(s):
    return str(s).strip() if s is not None else ''

# Extract V1.1 data rows with contract_type and subgroup tracking
def extract_v11_rows(raw_rows):
    result = []
    prev_ct = ''
    subgroup = 0
    for r in raw_rows:
        pn = D(r['product_name'])
        if '商品名称' in pn:
            subgroup += 1
            continue
        if not pn: continue
        ct = D(r['contract_type'])
        if not ct: ct = prev_ct
        else: prev_ct = ct
        r['_ct'] = ct
        r['_sg'] = subgroup
        result.append(r)
    return result

v58 = extract_v11_rows(data58)
vgj = extract_v11_rows(datagj)

# Parse existing HTML package data
def parse_html_packages(html, func_name):
    """Parse get58Data() or getGanjiData() function from HTML and return list of packages with their section info."""
    pattern = f'function {func_name}\\(\\)\\{{ return \\{{(.*?)\\n  \\]\\}};\\n\\}}'
    match = re.search(pattern, html, re.DOTALL)
    if not match:
        print(f"ERROR: Could not find {func_name}")
        return []
    
    body = match.group(1)
    
    # Parse sections: {name:"...",base:{...},packages:[...]}
    sections = []
    # Find each section block
    section_pattern = r'\{name:"([^"]+)",base:\{([^}]+)\},packages:\[(.*?)\]\}'
    for sm in re.finditer(section_pattern, body, re.DOTALL):
        sname = sm.group(1)
        sbase = sm.group(2)
        spkgs_str = sm.group(3)
        
        # Parse base attributes
        base = {}
        for k in ['C','D','E','F','G','H']:
            m = re.search(rf'{k}:(\d+)', sbase)
            if m: base[k] = int(m.group(1))
        for k in ['skuC','skuD','skuE','skuF','skuG','skuH']:
            m = re.search(rf'{k}:"([^"]*)"', sbase)
            if m: base[k] = m.group(1)
        
        # Parse package rows
        packages = []
        pkg_pattern = r'\{name:"([^"]+)",dur:(\d+),C:(\d+),D:(\d+),E:(\d+),F:(\d+),G:(\d+),H:(\d+),J:(null|\d+)(?:,L:"([^"]*)")?(?:,productName:"([^"]*)")?\}'
        for pm in re.finditer(pkg_pattern, spkgs_str):
            pkg = {
                'name': pm.group(1),
                'dur': int(pm.group(2)),
                'C': int(pm.group(3)),
                'D': int(pm.group(4)),
                'E': int(pm.group(5)),
                'F': int(pm.group(6)),
                'G': int(pm.group(7)),
                'H': int(pm.group(8)),
                'J': None if pm.group(9) == 'null' else int(pm.group(9)),
                'L': pm.group(10) if pm.group(10) else '',
                'productName': pm.group(11) if pm.group(11) else '',
            }
            # Also capture M and N if they already exist (V3.6 might not have them)
            # Strip the matched text and look for additional fields
            packages.append(pkg)
        
        sections.append({'name': sname, 'base': base, 'packages': packages})
    
    return sections

print("Parsing HTML packages...")
html58 = parse_html_packages(html, 'get58Data')
htmlgj = parse_html_packages(html, 'getGanjiData')

print(f"\nOld 58: {sum(len(s['packages']) for s in html58)} packages in {len(html58)} sections")
for s in html58:
    print(f"  [{s['name']}] {len(s['packages'])} pkgs: {[p['name'] for p in s['packages']]}")

print(f"\nOld GJ: {sum(len(s['packages']) for s in htmlgj)} packages in {len(htmlgj)} sections")
for s in htmlgj:
    print(f"  [{s['name']}] {len(s['packages'])} pkgs: {[p['name'] for p in s['packages']]}")

# ==== MATCHING ====
def match_v11(pkg, vrows, section_ct=''):
    """Find best matching V1.1 row for a package."""
    dur = pkg['dur']
    C_val = pkg['C']
    D_val = pkg['D']
    F_val = pkg['F']
    pname = pkg['name']
    
    best = None
    best_score = -1
    
    for r in vrows:
        r_dur = int(D(r['duration']) or 0)
        r_C = int(D(r['pkg_count']) or 1)
        r_D = int(D(r['online_jobs']) or 0)
        r_F = int(D(r['promotion_fee']) or 0)
        rn = D(r['product_name'])
        
        score = 0
        if r_dur == dur: score += 10
        if r_C == C_val: score += 5
        if r_D == D_val: score += 5
        if r_F == F_val: score += 3
        # Name similarity
        if rn in pname or pname in rn: score += 2
        if rn == pname: score += 5
        
        if score > best_score:
            best_score = score
            best = r
    
    return best, best_score

# Update 58 packages
print("\n===== MAPPING 58 =====")
for s in html58:
    sname = s['name']
    # Determine which V1.1 subgroup to search in based on section name
    sg_map = {
        '套餐季度（92天）': ('套餐季度', 1),
        '套餐年度（365天-限类限城）': ('套餐年度', 2),
        '套餐年度（365天-全国）': ('套餐年度', 3),
        '定制/框架年度': ('合同类型', 4),
        '定制/框架年度（精简版）': ('合同类型', 5),
        '定制/框架季度（92天）': ('合同类型', 6),
        '定制/框架-美团特供2.0': ('合同类型', 7),
        '定制/框架年度（容器）': ('合同类型', 8),
        '定制/框架年度（容器-行业限类）': ('合同类型', 9),
        '定制/框架年度（容器-行业限类-92天）': ('合同类型', 10),
    }
    
    search_ct, search_sg = sg_map.get(sname, (None, None))
    if search_ct:
        candidates = [r for r in v58 if r['_ct'] == search_ct and r['_sg'] == search_sg]
    else:
        candidates = v58  # search all
    
    if not candidates:
        print(f"  [{sname}] No V1.1 candidates for ct={search_ct} sg={search_sg}")
        continue
    
    for p in s['packages']:
        match, score = match_v11(p, candidates)
        if match and score >= 10:
            old_pn = p.get('productName', '')
            new_pn = D(match['product_name'])
            new_L = D(match['has_exchange'])
            new_ex_price = D(match['exchange_price'])
            new_N = D(match['publish_limit']).replace('\n',' ')
            
            p['productName'] = new_pn
            p['L'] = new_L
            try: p['M'] = int(float(new_ex_price)) if new_ex_price else 0
            except: p['M'] = 0
            p['N'] = new_N
            
            print(f"  [{sname}] {p['name']:20s} → {new_pn:20s} L={new_L:3s} M={p.get('M','')} N={new_N[:40]}")
        else:
            print(f"  [{sname}] {p['name']:20s} → NO MATCH (score={score})")

# Update GJ packages
print("\n===== MAPPING GJ =====")
for s in htmlgj:
    sname = s['name']
    sg_map = {
        '个人超值月度（30天）': ('', 1),
        '套餐年度（365天-限类限城）': ('套餐年度', 2),
        '套餐年度（365天-全国）': ('套餐年度', 3),
        '套餐年度（365天-简历宝）': ('套餐年度', 4),
        '套餐年度（365天-行业版）': ('套餐年度', 5),
        '定制/框架': ('套餐年度', 6),
    }
    
    search_ct, search_sg = sg_map.get(sname, (None, None))
    if search_ct is not None:
        candidates = [r for r in vgj if r['_ct'] == search_ct and r['_sg'] == search_sg]
    else:
        candidates = vgj
    
    if not candidates:
        print(f"  [{sname}] No V1.1 candidates for ct='{search_ct}' sg={search_sg}")
        continue
    
    for p in s['packages']:
        match, score = match_v11(p, candidates)
        if match and score >= 8:
            new_pn = D(match['product_name'])
            new_L = D(match['has_exchange'])
            new_ex_price = D(match['exchange_price'])
            new_N = D(match['publish_limit']).replace('\n',' ')
            
            p['productName'] = new_pn
            p['L'] = new_L
            try: p['M'] = int(float(new_ex_price)) if new_ex_price else 0
            except: p['M'] = 0
            p['N'] = new_N
            
            print(f"  [{sname}] {p['name']:25s} → {new_pn:25s} L={new_L:6s} M={p.get('M','')} N={new_N[:50]}")
        else:
            print(f"  [{sname}] {p['name']:25s} → NO MATCH (score={score}), candidates: {len(candidates)}")

# ===== DEDUPLICATION =====
# After updating, deduplicate packages with same productName within each section
for s in html58:
    seen = set()
    deduped = []
    for p in s['packages']:
        key = p.get('productName', p['name'])
        if key not in seen:
            seen.add(key)
            deduped.append(p)
        else:
            print(f"  DEDUP 58 [{s['name']}]: removed duplicate '{key}'")
    s['packages'] = deduped

for s in htmlgj:
    seen = set()
    deduped = []
    for p in s['packages']:
        key = p.get('productName', p['name'])
        if key not in seen:
            seen.add(key)
            deduped.append(p)
        else:
            print(f"  DEDUP GJ [{s['name']}]: removed duplicate '{key}'")
    s['packages'] = deduped

# ===== GENERATE NEW JS CODE =====
def js_str(s):
    return '"' + s.replace('\\','\\\\').replace('"','\\"').replace('\n','\\n') + '"'

def js_val(v):
    if v is None: return 'null'
    if isinstance(v, str): return js_str(v)
    return str(v)

def generate_js58():
    lines = ['function get58Data(){ return {','  sections: [']
    for si, s in enumerate(html58):
        b = s['base']
        lines.append(f'    {{name:"{s["name"]}",base:{{C:{b["C"]},D:{b["D"]},E:{b["E"]},F:{b["F"]},G:{b["G"]},H:{b["H"]},skuC:"{b["skuC"]}",skuD:"{b["skuD"]}",skuE:"{b["skuE"]}",skuF:"{b["skuF"]}",skuG:"{b["skuG"]}",skuH:"{b["skuH"]}"}},packages:[')
        for pi, p in enumerate(s['packages']):
            fields = f'name:"{p["name"]}",dur:{p["dur"]},C:{p["C"]},D:{p["D"]},E:{p["E"]},F:{p["F"]},G:{p["G"]},H:{p["H"]},J:{js_val(p["J"])},L:"{p.get("L","否")}"'
            if p.get('M') is not None:
                fields += f',M:{p["M"]}'
            if p.get('N'):
                fields += f',N:{js_str(p["N"])}'
            if p.get('productName'):
                fields += f',productName:{js_str(p["productName"])}'
            pkg_line = f'      {{{fields}}}'
            if pi < len(s['packages']) - 1:
                pkg_line += ','
            lines.append(pkg_line)
        tail = '    ]},' if si < len(html58)-1 else '    ]}'
        lines.append(tail)
    lines.append('  ]};')
    lines.append('}')
    return '\n'.join(lines)

def generate_jsgj():
    lines = ['function getGanjiData(){ return {','  sections: [']
    for si, s in enumerate(htmlgj):
        b = s['base']
        lines.append(f'    {{name:"{s["name"]}",base:{{C:{b["C"]},D:{b["D"]},E:{b["E"]},F:{b["F"]},G:{b["G"]},H:{b["H"]},skuC:"{b["skuC"]}",skuD:"{b["skuD"]}",skuE:"{b["skuE"]}",skuF:"{b["skuF"]}",skuG:"{b["skuG"]}",skuH:"{b["skuH"]}"}},packages:[')
        for pi, p in enumerate(s['packages']):
            fields = f'name:"{p["name"]}",dur:{p["dur"]},C:{p["C"]},D:{p["D"]},E:{p["E"]},F:{p["F"]},G:{p["G"]},H:{p["H"]},J:{js_val(p["J"])},L:"{p.get("L","否")}"'
            if p.get('M') is not None:
                fields += f',M:{p["M"]}'
            if p.get('N'):
                fields += f',N:{js_str(p["N"])}'
            if p.get('productName'):
                fields += f',productName:{js_str(p["productName"])}'
            pkg_line = f'      {{{fields}}}'
            if pi < len(s['packages']) - 1:
                pkg_line += ','
            lines.append(pkg_line)
        tail = '    ]},' if si < len(htmlgj)-1 else '    ]}'
        lines.append(tail)
    lines.append('  ]};')
    lines.append('}')
    return '\n'.join(lines)

new_js58 = generate_js58()
new_jsgj = generate_jsgj()

print("\n===== GENERATED 58 JS =====")
print(new_js58[:500])
print("...")
print("\n===== GENERATED GJ JS =====")
print(new_jsgj[:500])
print("...")

# Save
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\new_58data.js','w',encoding='utf-8') as f:
    f.write(new_js58)
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\new_gjdata.js','w',encoding='utf-8') as f:
    f.write(new_jsgj)

print("\nNew JS data saved to new_58data.js and new_gjdata.js")
