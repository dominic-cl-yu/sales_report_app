import sys
sys.path.insert(0, '/mnt/data/sales_report_app')
from process import generate_pivot_report_from_upload

excel_bytes = open('/mnt/data/Order Status Report_30th Oct 10+2.xlsx','rb').read()

pivot_bytes, stats = generate_pivot_report_from_upload(
    excel_bytes,
    filename='Order Status Report_30th Oct 10+2.xlsx',
    report_date='Jan-15',
)

print('rows_used:', stats['rows_used'])
print('factories:', stats['factories'])
print('product_types_count:', len(stats['product_types']))
print('bytes_len:', len(pivot_bytes))

open('/mnt/data/sales_report_app/generated_pivot.xlsx','wb').write(pivot_bytes)
print('saved to /mnt/data/sales_report_app/generated_pivot.xlsx')
