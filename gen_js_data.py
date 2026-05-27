# -*- coding: utf-8 -*-
"""Generate complete JS data for get58Data() and getGanjiData() from V1.1 JSON"""
import json

# Load extracted data
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\sec58_final.json', 'r', encoding='utf-8') as f:
    sec58 = json.load(f)
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\secgj_final.json', 'r', encoding='utf-8') as f:
    secgj = json.load(f)

# ====== 58 Section name + SKU mappings ======
SEC58_NAMES = [
    "套餐季度（92天）",
    "套餐年度（365天-限类限城）",
    "套餐年度（365天-全国）",
    "定制/框架年度",
    "定制/框架年度（精简版）",
    "定制/框架季度（92天）",
    "定制/框架-美团特供2.0",
    "定制/框架年度（容器）",
    "定制/框架年度（容器-行业限类）",
    "定制/框架年度（容器-行业限类-92天）",
]

# Fix section 9 SKU (容器-行业限类-92天: D=388 -> skuD=58_00020)
sec58[9]['skus'] = {'skuC':'58_00001','skuD':'58_00020','skuE':'58_00003','skuF':'58_00007','skuG':'58_00013','skuH':'58_00014'}

# SKU for GJ sections (all use same sku mapping for most sections)
GJ_SKU_ANNUAL = {'skuC':'58_00009','skuD':'GJ_00009','skuE':'GJ_00005','skuF':'58_00015','skuG':'58_00013','skuH':'58_00014'}
GJ_SKU_LIMITED = {'skuC':'58_00009','skuD':'GJ_00004','skuE':'GJ_00005','skuF':'58_00015','skuG':'58_00013','skuH':'58_00014'}
GJ_SKU_MONTHLY = {'skuC':'','skuD':'GJ_00001','skuE':'GJ_00002','skuF':'','skuG':'','skuH':''}

def escape_js_str(s):
    return s.replace('\\','\\\\').replace("'","\\'").replace('\n','\\n')

def gen_package(r):
    """Generate JS package object string."""
    parts = [f"name:'{escape_js_str(r['name'])}'"]
    parts.append(f"dur:{r['dur']}")
    parts.append(f"C:{r['C']}")
    parts.append(f"D:{r['D']}")
    parts.append(f"E:{100 if r['dur'] <= 92 else 300}")  # Default E values
    parts.append(f"F:{r.get('F',0)}")
    parts.append("G:0")
    parts.append("H:0")
    sell = r.get('sell_price')
    if sell is not None:
        parts.append(f"J:{sell}")
    else:
        parts.append("J:null")
    parts.append(f"L:'{r.get('L','无')}'")
    parts.append(f"M:{r.get('M','null')}")
    parts.append(f"N:'{escape_js_str(r.get('N',''))}'")
    parts.append(f"productName:'{escape_js_str(r['product_name'])}'")
    return '{' + ','.join(parts) + '}'

def gen_base(b):
    return f"C:{b['C']},D:{b['D']},E:{b['E']},F:{b['F']},G:{b['G']},H:{b['H']}"

def gen_skus(s):
    return f"skuC:'{s.get('skuC','')}',skuD:'{s.get('skuD','')}',skuE:'{s.get('skuE','')}',skuF:'{s.get('skuF','')}',skuG:'{s.get('skuG','')}',skuH:'{s.get('skuH','')}'"

# ====== Generate 58 Data ======
lines58 = []
lines58.append("function get58Data(){ return {")
lines58.append("  sections: [")
for i, s in enumerate(sec58):
    name = SEC58_NAMES[i]
    base = gen_base(s['base'])
    skus = gen_skus(s.get('skus', {}))
    lines58.append(f"    {{name:\"{name}\",base:{{{base},{skus}}},packages:[")
    for r in s['rows']:
        lines58.append(f"      {gen_package(r)},")
    lines58.append("    ]},")
lines58.append("  ]};")
lines58.append("}")
f58 = '\n'.join(lines58)

# ====== Generate GJ Data ======
GJ_NAMES = [
    "套餐年度（365天-限类限城）",
    "套餐年度（365天-全国）",
    "套餐年度（365天-简历宝）",
    "套餐年度（365天-行业版）",
    "定制/框架",
]
# Add 个人超值月度 section manually
gj_monthly_base = {'C':0,'D':50,'E':50,'F':0,'G':0,'H':0}
gj_monthly_skus = GJ_SKU_MONTHLY
gj_monthly_rows = [
    {"name":"急招限时超值包（制造业企业维度45天购买1次）","product_name":"个人超值权益包","dur":30,"C":1,"D":2,"F":0,"sell_price":200,"L":"无","M":200,"N":"不支持5大类（普工/技工、生产制造、司机/物流、安保消防、家政服务）的岗位发布，需要在线单购，制造业行业范围的营业执照：企业维度45天购买1次"},
    {"name":"急招限时超值包（制造业企业维度45天购买1次）","product_name":"个人超值权益包","dur":30,"C":1,"D":3,"F":0,"sell_price":200,"L":"无","M":200,"N":"不支持5大类（普工/技工、生产制造、司机/物流、安保消防、家政服务）的岗位发布，需要在线单购，制造业行业范围的营业执照：企业维度45天购买1次"},
    {"name":"急招限时超值包（制造业企业维度45天购买1次）","product_name":"个人超值权益包","dur":30,"C":1,"D":3,"F":0,"sell_price":200,"L":"无","M":200,"N":"不支持5大类（普工/技工、生产制造、司机/物流、安保消防、家政服务）的岗位发布，需要在线单购，制造业行业范围的营业执照：企业维度45天购买1次"},
    {"name":"AI速聘包","product_name":"个人超值权益包","dur":30,"C":1,"D":2,"F":0,"sell_price":200,"L":"无","M":200,"N":"不支持5大类（普工/技工、生产制造、司机/物流、安保消防、家政服务）的岗位发布，需要在线单购"},
    {"name":"AI速聘包","product_name":"个人超值权益包","dur":30,"C":1,"D":3,"F":0,"sell_price":200,"L":"无","M":200,"N":"不支持5大类（普工/技工、生产制造、司机/物流、安保消防、家政服务）的岗位发布，需要在线单购"},
    {"name":"AI速聘包","product_name":"个人超值权益包","dur":30,"C":1,"D":3,"F":0,"sell_price":200,"L":"无","M":200,"N":"不支持5大类（普工/技工、生产制造、司机/物流、安保消防、家政服务）的岗位发布，需要在线单购"},
]

linesgj = []
linesgj.append("function getGanjiData(){ return {")
linesgj.append("  sections: [")
# Monthly section
linesgj.append(f"    {{name:\"个人超值月度（30天）\",base:{{{gen_base(gj_monthly_base)},{gen_skus(gj_monthly_skus)}}},packages:[")
for r in gj_monthly_rows:
    linesgj.append(f"      {gen_package(r)},")
linesgj.append("    ]},")
# Other sections
GJ_SKUS = [GJ_SKU_LIMITED, GJ_SKU_ANNUAL, GJ_SKU_ANNUAL, GJ_SKU_ANNUAL, GJ_SKU_ANNUAL]
for i, s in enumerate(secgj):
    name = GJ_NAMES[i]
    sku = GJ_SKUS[i]
    base = gen_base(s['base'])
    skus = gen_skus(sku)
    linesgj.append(f"    {{name:\"{name}\",base:{{{base},{skus}}},packages:[")
    for r in s['rows']:
        linesgj.append(f"      {gen_package(r)},")
    linesgj.append("    ]},")
linesgj.append("  ]};")
linesgj.append("}")
fgj = '\n'.join(linesgj)

# Write to files
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\new_get58data.js', 'w', encoding='utf-8') as f:
    f.write(f58)
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\new_getgjdata.js', 'w', encoding='utf-8') as f:
    f.write(fgj)

# Also write combined JS for reference
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\new_data_combined.js', 'w', encoding='utf-8') as f:
    f.write(f58 + '\n\n' + fgj)

print("Generated get58Data and getGanjiData JS code.")
print(f"58: {sum(len(s['rows']) for s in sec58)} packages across {len(sec58)} sections")
print(f"GJ: {len(gj_monthly_rows) + sum(len(s['rows']) for s in secgj)} packages across {1 + len(secgj)} sections")
