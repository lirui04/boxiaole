# -*- coding: utf-8 -*-
"""Extract V1.1.xlsx data and compare with current HTML"""
import pandas as pd
import json, re

xls_path = r'C:\Users\Administrator\Desktop\数据报表\李芮计算器\赶集双模式\伯小乐计算器\【双模式】伯小乐版本计算器V1.1.xlsx'
xls = pd.ExcelFile(xls_path)

def extract_sheet(df):
    """Extract data rows from pricing sheets with multi-row headers."""
    results = []
    current_group = ''
    current_subtype = ''
    
    for i, row in df.iterrows():
        vals = [row.iloc[j] for j in range(len(row))]
        if all(pd.isna(v) for v in vals):
            continue
        
        a_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        b_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        
        # Skip data row headers
        if a_val and not b_val:
            current_group = a_val
            current_subtype = a_val
            continue
        
        # Only process data rows with B column values
        if not b_val:
            continue
        
        # Skip "商品名称" header rows
        if '商品名称' in b_val:
            continue
        
        row_data = {
            'section': current_group,
            'product_name': b_val,
            'duration': str(row.iloc[2]) if pd.notna(row.iloc[2]) else '',
            'pkg_count': str(row.iloc[3]) if pd.notna(row.iloc[3]) else '',
            'D': str(row.iloc[4]) if pd.notna(row.iloc[4]) else '',
            'E': '100' if '92' in current_group and a_val else ('200' if '30' in current_group else '300'),
            'F': str(row.iloc[6]) if pd.notna(row.iloc[6]) else '',
            'G': str(row.iloc[7]) if pd.notna(row.iloc[7]) else '',
            'H': str(row.iloc[8]) if pd.notna(row.iloc[8]) else '',
            'total': str(row.iloc[9]) if pd.notna(row.iloc[9]) else '',
            'sell_price': str(row.iloc[10]) if pd.notna(row.iloc[10]) else '',
            'discount': str(row.iloc[11]) if pd.notna(row.iloc[11]) else '',
            'has_exchange': str(row.iloc[12]) if len(row) > 12 and pd.notna(row.iloc[12]) else '',
            'exchange_price': str(row.iloc[13]) if len(row) > 13 and pd.notna(row.iloc[13]) else '',
            'publish_limit': str(row.iloc[14]) if len(row) > 14 and pd.notna(row.iloc[14]) else '',
        }
        results.append(row_data)
    
    return results

# Extract all 3 sheets
df58 = pd.read_excel(xls, sheet_name=0, header=None)
dfgj = pd.read_excel(xls, sheet_name=1, header=None)
dfsku = pd.read_excel(xls, sheet_name=2, header=None)

data58 = extract_sheet(df58)
datagj = extract_sheet(dfgj)

print("=== 58同城 数据 ===")
for i, r in enumerate(data58):
    print(f"[{i}] sec={r['section'][:20]} | B={r['product_name']} | dur={r['duration']} | C={r['pkg_count']} | D={r['D']} | F={r['F']} | sell={r['sell_price']} | L={r['has_exchange']} | M={r['exchange_price']} | N={r['publish_limit'][:40]}")

print(f"\n=== 赶集 数据 ===")
for i, r in enumerate(datagj):
    print(f"[{i}] sec={r['section'][:20]} | B={r['product_name']} | dur={r['duration']} | C={r['pkg_count']} | D={r['D']} | F={r['F']} | sell={r['sell_price']} | L={r['has_exchange']} | M={r['exchange_price']} | N={r['publish_limit'][:40]}")

# Save to JSON
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\v11_data58.json', 'w', encoding='utf-8') as f:
    json.dump(data58, f, ensure_ascii=False, indent=2)
with open(r'D:\硅基流\workbuddy\日常任务存储\2026-05-26-16-25-35\v11_datagj.json', 'w', encoding='utf-8') as f:
    json.dump(datagj, f, ensure_ascii=False, indent=2)

print(f"\n58数据: {len(data58)} rows, 赶集数据: {len(datagj)} rows")
