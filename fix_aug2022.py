import os
import requests

fixes = [
    (
        '2022-08-02',
        'https://npp.gov.in/public-reports/cea/daily/dgr/02-08-2022/dgr1-2022-08-02.pdf',
    ),
    (
        '2022-08-03',
        'https://npp.gov.in/public-reports/cea/daily/dgr/03-08-2022/dgr1-2022-08-03.pdf',
    ),
    (
        '2022-08-22',
        'https://npp.gov.in/public-reports/cea/daily/dgr/22-08-2022/dgr1-2022-08-22.pdf',
    ),
]

headers = {'User-Agent': 'Mozilla/5.0'}
raw_dir = os.path.join('02_RAW_REPORTS', '2022')

print('=== RESOLVING AUGUST 2022 FILES ===')
for d, url in fixes:
  # 1. Remove empty .xls stub
  stub_path = os.path.join(raw_dir, f'dgr1-{d}.xls')
  if os.path.exists(stub_path):
    os.remove(stub_path)
    print(f'[✓] Removed empty stub: {stub_path}')

  # 2. Download valid PDF
  pdf_path = os.path.join(raw_dir, f'dgr1-{d}.pdf')
  resp = requests.get(url, headers=headers, timeout=15)
  if resp.status_code == 200 and len(resp.content) > 1000:
    with open(pdf_path, 'wb') as f:
      f.write(resp.content)
    print(f'[✓] Downloaded valid PDF ({len(resp.content)} bytes): {pdf_path}')
  else:
    print(f'[!] Failed to fetch {url}: status {resp.status_code}')

print('Done. Ready to re-synchronize pipeline.')