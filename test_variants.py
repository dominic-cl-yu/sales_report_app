#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Test the app with variant input files."""

from process import generate_pivot_report_from_upload

files = [
    'Order Status Report_30th Oct 10+2.xlsx',
    'Order Status Report_30th Oct 10+2 - b.xlsx',
    'Order Status Report_30th Oct 10+2 - c.xlsx'
]

for f in files:
    print(f"\n{'='*60}")
    print(f"Testing: {f}")
    print('='*60)
    
    try:
        with open(f, 'rb') as fp:
            excel_bytes = fp.read()
        
        pivot_bytes, stats = generate_pivot_report_from_upload(
            excel_bytes=excel_bytes,
            filename=f,
            report_date='Jan-29',
        )
        
        print("[OK] SUCCESS!")
        print(f"   Rows used: {stats['rows_used']}")
        print(f"   Factories: {stats['factories']}")
        print(f"   Product types: {stats['product_types']}")
        print(f"   Output size: {len(pivot_bytes):,} bytes")
        
    except Exception as e:
        print(f"[FAIL] FAILED: {e}")

print("\n" + "="*60)
print("Test complete!")
