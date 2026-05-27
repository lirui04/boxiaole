# -*- coding: utf-8 -*-
"""
Ultimate V1.1 sync: read xlsx raw, track sections, generate complete JS data
Plus: read-only display + column-width changes integrated
"""
import pandas as pd, json, re, os

xpath = r'C:\Users\Administrator\Desktop\数据报表\李芮计算器\赶集双模式\伯小乐计算器\【双模式】伯小乐版本计算器V1.1.xlsx'
html_path = r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\伯小乐组合包定价工具.html'

# ====== STEP 1: Parse xlsx with proper section tracking ======
xls = pd.ExcelFile(xpath)

def parse_sheet(df):
    """Robust parser: track section via '基准' keyword in col A or section change patterns."""
    sections = []
    all_rows_flat = []
    
    current_section_name = ''
    current_base = {}
    current_skus = {}
    in_section = False
    mode = None  # '套餐' or '定制' or '合同'
    
    for i, row in df.iterrows():
        vals = [row.iloc[j] for j in range(min(15, len(row)))]
        a = str(vals[0]).strip() if pd.notna(vals[0]) and str(vals[0]).strip() != 'nan' else ''
        b = str(vals[1]).strip() if pd.notna(vals[1]) and str(vals[1]).strip() != 'nan' else ''
        
        # Detect section header: "基准" in col B AND numeric base values follow
        if b == '基准' or b == '\u57fa\u51c6':
            # This is a section header - extract base data
            base_C = float(vals[2]) if pd.notna(vals[2]) else 0
            base_D = float(vals[3]) if pd.notna(vals[3]) else 0
            base_E = float(vals[4]) if pd.notna(vals[4]) else 0
            base_F = float(vals[5]) if pd.notna(vals[5]) else 0
            base_G = float(vals[6]) if pd.notna(vals[6]) else 0
            base_H = float(vals[7]) if pd.notna(vals[7]) else 0
            
            # Section name from col A, or derive from context
            sn = a if a else current_section_name
            if sn == '合同类型':
                sn = current_section_name  # keep previous section name
            
            if current_section_name and current_base:
                # Save previous section's data
                if current_section_rows:
                    sections.append({
                        'name': current_section_name,
                        'base': current_base,
                        'rows': current_section_rows
                    })
            
            current_section_name = sn
            current_base = {
                'C': int(base_C), 'D': int(base_D), 'E': int(base_E),
                'F': int(base_F), 'G': int(base_G), 'H': int(base_H)
            }
            current_section_rows = []
            in_section = True
            continue
        
        # Skip header rows
        if '商品名称' in b:
            continue
        if a in ['合同类型', '定制/框架', '套餐信息', '套餐季度', '套餐年度', '个人超值月度'] and not b:
            continue
        
        # Data row: has product name in col B
        if not in_section or not b:
            continue
        
        if b in ['基准', '商品名称', '服务时长']:
            continue
        
        # Normalize section name
        if current_section_name == '合同类型':
            # Determine if this is 定制/框架年度 or 精简版 based on row content
            if b.startswith('伯小乐-B') and b != '伯小乐-B':
                current_section_name = '定制/框架年度（精简版）'
                current_base['E'] = int(0)  # E column might be 0
        
        # Collect data row
        dur = str(int(float(vals[2]))) if pd.notna(vals[2]) else '365'
        C = str(int(float(vals[3]))) if pd.notna(vals[3]) else '1'
        D = str(int(float(vals[4]))) if pd.notna(vals[4]) else '0'
        F_val = str(int(float(vals[6]))) if pd.notna(vals[6]) and float(vals[6]) != 0 else '0'
        
        sell = str(int(float(vals[10]))) if pd.notna(vals[10]) else ''
        L = str(vals[12]).strip() if pd.notna(vals[12]) and str(vals[12]).strip() != 'nan' else '无'
        M = str(int(float(vals[13]))) if pd.notna(vals[13]) else ''
        N = str(vals[14]).strip() if len(vals) > 14 and pd.notna(vals[14]) else ''
        
        current_section_rows.append({
            'product_name': b,
            'name': a if a else '',  # row name (组合包名称)
            'dur': dur,
            'C': C,
            'D': D,
            'F': F_val,
            'sell': sell,
            'L': L,
            'M': M,
            'N': N
        })
    
    # Save last section
    if current_section_name and current_base and current_section_rows:
        sections.append({
            'name': current_section_name,
            'base': current_base,
            'rows': current_section_rows
        })
    
    return sections

# Read sheets
df58 = pd.read_excel(xls, sheet_name=0, header=None)
dfgj = pd.read_excel(xls, sheet_name=1, header=None)

sec58 = parse_sheet(df58)
secgj = parse_sheet(dfgj)

# Print for verification
print("=== 58 Sections ===")
for s in sec58:
    print(f"\n[{s['name']}] base={s['base']}")
    for r in s['rows']:
        print(f"  {r['name']:20s} | {r['product_name']:15s} | dur={r['dur']:4s} | C={r['C']:2s} | D={r['D']:3s} | F={r['F']:2s} | sell={r['sell']:5s} | L={r['L']:4s} | M={r['M']:5s}")

print("\n=== GJ Sections ===")
for s in secgj:
    print(f"\n[{s['name']}] base={s['base']}")
    for r in s['rows']:
        print(f"  {r['name']:30s} | {r['product_name']:20s} | dur={r['dur']:4s} | C={r['C']:2s} | D={r['D']:3s} | F={r['F']:2s} | sell={r['sell']:5s} | L={r['L']:4s} | M={r['M']:5s}")

# Save sections to JSON
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\sec58_clean.json', 'w', encoding='utf-8') as f:
    json.dump(sec58, f, ensure_ascii=False, indent=2)
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\secgj_clean.json', 'w', encoding='utf-8') as f:
    json.dump(secgj, f, ensure_ascii=False, indent=2)
