import os
import requests
import xlrd

dates = ['2022-08-02', '2022-08-03', '2022-08-22']
headers = {'User-Agent': 'Mozilla/5.0'}

print('=== 1. XLS INTERNAL SHEET INSPECTION ===')
for d in dates:
  fpath = f'02_RAW_REPORTS/2022/dgr1-{d}.xls'
  if os.path.exists(fpath):
    try:
      wb = xlrd.open_workbook(fpath)
      print(f'{d}: Sheet Names = {wb.sheet_names()}')
      for sname in wb.sheet_names():
        sheet = wb.sheet_by_name(sname)
        print(f'   Sheet "{sname}": nrows={sheet.nrows}, ncols={sheet.ncols}')
    except Exception as e:
      print(f'{d}: Error opening with xlrd: {e}')
  else:
    print(f'{d}: File not found at {fpath}')

print('\n=== 2. CHECKING NPP PORTAL FOR CORRESPONDING PDFs ===')
for d in dates:
  dmy = f'{d[8:10]}-{d[5:7]}-{d[:4]}'
  pdf_url = f'https://npp.gov.in/public-reports/cea/daily/dgr/{dmy}/dgr1-{d}.pdf'
  try:
    r = requests.head(pdf_url, headers=headers, timeout=5)
    print(f'{d} PDF URL [{r.status_code}] -> {pdf_url}')
  except Exception as e:
    print(f'{d} PDF URL request failed: {e}')