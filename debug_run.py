import sys
import time

sys.path.insert(0, '/mnt/data/sales_report_app')
import process

print('start read bytes')
excel_bytes = open('/mnt/data/Order Status Report_30th Oct 10+2.xlsx', 'rb').read()
print('bytes len', len(excel_bytes))

print('call generate...')
t0 = time.time()
pivot_bytes, stats = process.generate_pivot_report_from_upload(
    excel_bytes,
    filename='Order Status Report_30th Oct 10+2.xlsx',
    report_date='Jan-15',
)
print('done generate in', round(time.time() - t0, 2), 's')
print('rows_used', stats['rows_used'])
print('factories', stats['factories'])
print('product_types_count', len(stats['product_types']))
print('pivot_bytes_len', len(pivot_bytes))

out_path = '/mnt/data/sales_report_app/generated_pivot.xlsx'
with open(out_path, 'wb') as f:
    f.write(pivot_bytes)
print('wrote output to', out_path)
