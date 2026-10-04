import os
import datetime
import requests
from tqdm import tqdm

RAW_DIR = "02_RAW_REPORTS"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}

def download_report_for_date(target_date):
    dmy = target_date.strftime("%d-%m-%Y")
    ymd = target_date.strftime("%Y-%m-%d")
    year = str(target_date.year)

    dest_dir = os.path.join(RAW_DIR, year)
    os.makedirs(dest_dir, exist_ok=True)

    pdf_file = os.path.join(dest_dir, f"dgr1-{ymd}.pdf")
    xls_file = os.path.join(dest_dir, f"dgr1-{ymd}.xls")
    xlsx_file = os.path.join(dest_dir, f"dgr1-{ymd}.xlsx")

    # Skip if already downloaded
    if os.path.exists(pdf_file) or os.path.exists(xls_file) or os.path.exists(xlsx_file):
        return True

    candidates = [
        (f"https://npp.gov.in/public-reports/cea/daily/dgr/{dmy}/dgr1-{ymd}.pdf", pdf_file),
        (f"https://npp.gov.in/public-reports/cea/daily/dgr/{dmy}/dgr1-{ymd}.xls", xls_file),
        (f"https://npp.gov.in/public-reports/cea/daily/dgr/{dmy}/dgr1-{ymd}.xlsx", xlsx_file),
    ]

    for url, save_path in candidates:
        try:
            resp = requests.get(url, headers=HEADERS, timeout=10)
            if resp.status_code == 200 and len(resp.content) > 3000:
                with open(save_path, "wb") as f:
                    f.write(resp.content)
                return True
        except Exception:
            continue

    return False

def run_batch_download(start_date, end_date):
    print(f"\nStarting Batch Download: {start_date} to {end_date}...")
    curr = start_date
    delta = datetime.timedelta(days=1)
    days = (end_date - start_date).days + 1

    pbar = tqdm(total=days, desc="Downloading")
    downloaded = 0
    while curr <= end_date:
        if download_report_for_date(curr):
            downloaded += 1
        curr += delta
        pbar.update(1)
    pbar.close()

    print(f"[✓] Completed: {downloaded} / {days} files ready in {RAW_DIR}/.")

if __name__ == "__main__":
    # Bridge the Phase-0 Middle Gap (342 days)
    START = datetime.date(2025, 8, 1)
    END = datetime.date(2026, 7, 8)
    run_batch_download(START, END)