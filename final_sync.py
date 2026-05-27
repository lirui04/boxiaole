# -*- coding: utf-8 -*-
"""Final V1.1 sync: correct section detection, generate JS data"""
import pandas as pd, json, re

xpath = r'C:\Users\Administrator\Desktop\数据报表\李芮计算器\赶集双模式\伯小乐计算器\【双模式】伯小乐版本计算器V1.1.xlsx'
xls = pd.ExcelFile(xpath)

# SKU mapping
SKU_MAP = {
    # Quarterly base
    ('200','66','100','1300','600','350'): {
        'skuC':'58_00001','skuD':'58_00002','skuE':'58_00003','skuF':'58_00007','skuG':'58_00005','skuH':'58_00006'},
    # Annual limited base
    ('500','100','300','2400','1780','1400'): {
        'skuC':'58_00009','skuD':'58_00010','skuE':'58_00011','skuF':'58_00015','skuG':'58_00013','skuH':'58_00014'},
    # Annual national base
    ('500','125','300','2400','1780','1400'): {
        'skuC':'58_00009','skuD':'58_00017','skuE':'58_00011','skuF':'58_00015','skuG':'58_00013','skuH':'58_00014'},
    # Annual 0-E base (simplified)
    ('500','125','0','2400','1780','1400'): {
        'skuC':'58_00009','skuD':'58_00017','skuE':'','skuF':'58_00015','skuG':'58_00013','skuH':'58_00014'},
    # Quarterly container
    ('200','66','100','1300','1780','1400'): {
        'skuC':'58_00001','skuD':'58_00002','skuE':'58_00003','skuF':'58_00007','skuG':'58_00013','skuH':'58_00014'},
    # Meituan special
    ('500','66','100','2400','1780','1400'): {
        'skuC':'58_00009','skuD':'58_00002','skuE':'58_00003','skuF':'58_00015','skuG':'58_00013','skuH':'58_00014'},
    # Container annual
    ('500','560','300','2400','1780','1400'): {
        'skuC':'58_00009','skuD':'58_00018','skuE':'58_00011','skuF':'58_00015','skuG':'58_00013','skuH':'58_00014'},
    # Container industry-limit annual
    ('500','440','300','2400','1780','1400'): {
        'skuC':'58_00009','skuD':'58_00019','skuE':'58_00011','skuF':'58_00015','skuG':'58_00013','skuH':'58_00014'},
    # Container industry-limit 92d
    ('200','260','100','1300','1780','1400'): {
        'skuC':'58_00001','skuD':'58_00020','skuE':'58_00003','skuF':'58_00007','skuG':'58_00013','skuH':'58_00014'},
}

def parse_sheet(df, sheet_type='58'):
    """Parse with correct section detection. Section rows: B=NaN, D-I=numeric."""
    sections = []
    cur_name = ''
    cur_base = None
    cur_rows = []
    cur_skus = None
    
    for i, row in df.iterrows():
        a_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        b_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        
        # Check if this is a section header row
        # Pattern: B is NaN, D,E,F,G,H,I are numeric
        d_val = row.iloc[3] if pd.notna(row.iloc[3]) else None
        e_val = row.iloc[4] if pd.notna(row.iloc[4]) else None
        f_val = row.iloc[5] if pd.notna(row.iloc[5]) else None
        g_val = row.iloc[6] if pd.notna(row.iloc[6]) else None
        h_val = row.iloc[7] if pd.notna(row.iloc[7]) else None
        i_val = row.iloc[8] if pd.notna(row.iloc[8]) else None
        
        is_section_header = (
            not b_val and
            d_val is not None and isinstance(d_val, (int, float)) and d_val > 0 and
            e_val is not None and isinstance(e_val, (int, float)) and
            f_val is not None and isinstance(f_val, (int, float)) and
            g_val is not None and isinstance(g_val, (int, float)) and
            h_val is not None and isinstance(h_val, (int, float)) and
            i_val is not None and isinstance(i_val, (int, float))
        )
        
        if is_section_header:
            # Save previous section
            if cur_name and cur_base and cur_rows:
                sections.append({
                    'name': cur_name,
                    'base': cur_base,
                    'skus': cur_skus or {},
                    'rows': cur_rows
                })
            
            base_C = int(d_val)
            base_D = int(e_val)
            base_E = int(f_val)
            base_F = int(g_val)
            base_G = int(h_val)
            base_H = int(i_val)
            
            cur_name = a_val
            cur_base = {'C': base_C, 'D': base_D, 'E': base_E, 'F': base_F, 'G': base_G, 'H': base_H}
            
            # Find SKU mapping
            base_key = (str(base_C), str(base_D), str(base_E), str(base_F), str(base_G), str(base_H))
            cur_skus = SKU_MAP.get(base_key, {})
            cur_rows = []
            continue
        
        # Skip column header rows
        if '商品名称' in b_val or '职小乐' in b_val and 'A' not in b_val and '-' not in b_val:
            # But check if b_val is actually a product name like 伯小乐-A, 职小乐-A
            import re
            if re.match(r'^[伯职小乐\-+A-Z0-9]+$', b_val) or '（直招版）' in b_val:
                pass  # This is a product name
            else:
                continue
        
        # Data row
        if b_val and cur_name:
            dur = int(float(row.iloc[2])) if pd.notna(row.iloc[2]) else 365
            C = int(float(row.iloc[3])) if pd.notna(row.iloc[3]) else 1
            D = int(float(row.iloc[4])) if pd.notna(row.iloc[4]) else 0
            F_val = int(float(row.iloc[6])) if pd.notna(row.iloc[6]) and float(row.iloc[6]) != 0 else 0
            
            sell_price = int(float(row.iloc[10])) if pd.notna(row.iloc[10]) else None
            L_val = str(row.iloc[12]).strip() if pd.notna(row.iloc[12]) else '无'
            M_val = int(float(row.iloc[13])) if pd.notna(row.iloc[13]) else None
            N_val = str(row.iloc[14]).strip() if len(row) > 14 and pd.notna(row.iloc[14]) else ''
            
            cur_rows.append({
                'name': a_val,
                'product_name': b_val,
                'dur': dur,
                'C': C,
                'D': D,
                'F': F_val,
                'sell_price': sell_price,
                'L': L_val,
                'M': M_val,
                'N': N_val,
            })
    
    # Save last section
    if cur_name and cur_base and cur_rows:
        sections.append({
            'name': cur_name,
            'base': cur_base,
            'skus': cur_skus or {},
            'rows': cur_rows
        })
    
    return sections

df58 = pd.read_excel(xls, sheet_name=0, header=None)
dfgj = pd.read_excel(xls, sheet_name=1, header=None)

sec58 = parse_sheet(df58, '58')
secgj = parse_sheet(dfgj, 'gj')

print(f"58 sections: {len(sec58)}, total rows: {sum(len(s['rows']) for s in sec58)}")
print(f"GJ sections: {len(secgj)}, total rows: {sum(len(s['rows']) for s in secgj)}")

print("\n=== 58 Sections ===")
for j, s in enumerate(sec58):
    print(f"\n[{j}] {s['name']} | base={s['base']} | skus={s['skus']} | {len(s['rows'])} rows")
    for r in s['rows']:
        print(f"  {r['name']:25s} | {r['product_name']:20s} | C={r['C']:2d} D={r['D']:3d} F={r['F']:2d} | sell={r['sell_price']} | L={r['L']:5s} M={r['M']}")

print("\n=== GJ Sections ===")
for j, s in enumerate(secgj):
    print(f"\n[{j}] {s['name']} | base={s['base']} | {len(s['rows'])} rows")
    for r in s['rows']:
        print(f"  {r['name']:35s} | {r['product_name']:25s} | C={r['C']:2d} D={r['D']:3d} F={r['F']:2d} | sell={r['sell_price']} | L={r['L']:5s} M={r['M']}")

# Save for later use
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\sec58_final.json', 'w', encoding='utf-8') as f:
    json.dump(sec58, f, ensure_ascii=False, indent=2)
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\secgj_final.json', 'w', encoding='utf-8') as f:
    json.dump(secgj, f, ensure_ascii=False, indent=2)
