"""
Batch KPay Reconciliation — generates one Excel per shop per month.

Scans assets/WS_Recon/{MonthFolder}/{Shop}/ for the 3 source files,
then runs the reconciliation engine for each combo found.

Usage:
    python tools/batch_reconcile.py                          # all months, all shops
    python tools/batch_reconcile.py --month 202602           # Feb 2026 only
    python tools/batch_reconcile.py --shop KYR               # KYR only
    python tools/batch_reconcile.py --month 202602 --shop MI # single combo

Output goes to:  output/{YYYYMM}/{ShopCode}_KPay_{Mon}{Year}_Reconciliation.xlsx
"""

import argparse
import glob
import os
import re
import sys
from datetime import datetime

from openpyxl import Workbook

# ── Import reconciliation engine ─────────────────────────────────────────────
sys.path.insert(0, os.path.join(os.path.dirname(__file__)))
from reconcile_kpay import (
    read_kpay, read_pos, read_dbs,
    determine_methods, build_sheet, validate_output,
)

# Shop code → Merchant ID  (from KPay MID Code List)
SHOP_MID = {
    'GI':   '059300002',
    'KYR':  '146100005',
    'MI':   '656600001',
    'SS':   '003200004',
    'YMIX': '766500001',
    'CF':   '002700003',
    'HNP':  '405300005',
    'LK':   '288100003',
    'OP':   '729600004',
    'TIM':  '901600002',
    'TY':   '721200003',
    'MO':   '142300002',
    'TKO':  '038100002',
    'TM':   '525600001',
    'TMTP': '783200002',
    'YT':   '967900003',
    'PCS':  '812500007',
    'CMR':  '867600004',
    'HT':   '138700001',
    'MCP':  '903800002',
    'ST':   '388300003',
    'VW':   '020200002',
    'LF':   '570100002',
    'MIRA': '087400004',
    'MOS':  '620200002',
    'HQ':   '952600002',
}

MONTH_NAMES = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
]

BASE_DIR = os.path.join(os.path.dirname(__file__), '..', 'assets', 'WS_Recon')


def discover_jobs(base, month_filter=None, shop_filter=None):
    """
    Scan base/*/Shop/ for source files and return a list of job dicts.
    Each job has: month_folder, shop, year, month, pos, kpay, dbs, merchant.
    """
    jobs = []
    base = os.path.normpath(base)

    for month_dir in sorted(glob.glob(os.path.join(base, '*'))):
        if not os.path.isdir(month_dir):
            continue
        dirname = os.path.basename(month_dir)
        # Match folder names like "Jan2026", "Feb2026", etc.
        m = re.match(r'^([A-Za-z]+)(\d{4})$', dirname)
        if not m:
            continue
        mon_name, year_str = m.group(1), m.group(2)
        try:
            mon_idx = [n.lower() for n in MONTH_NAMES].index(mon_name[:3].lower()) + 1
        except ValueError:
            continue
        year = int(year_str)
        month_str = f'{year}{mon_idx:02d}'

        if month_filter and month_str != month_filter:
            continue

        for shop_dir in sorted(glob.glob(os.path.join(month_dir, '*'))):
            if not os.path.isdir(shop_dir):
                continue
            shop = os.path.basename(shop_dir)
            if shop_filter and shop != shop_filter:
                continue
            if shop not in SHOP_MID:
                continue

            # Find the 3 source files
            pos_files = glob.glob(os.path.join(shop_dir, 'POS Sales_*.xlsx'))
            kpay_files = [f for f in glob.glob(os.path.join(shop_dir, '#*_*.xlsx'))
                          if 'Reconciliation' not in f]
            dbs_files = glob.glob(os.path.join(shop_dir, 'DBS_*.xls'))

            if not pos_files or not kpay_files or not dbs_files:
                print(f'  SKIP {shop}/{month_str}: missing source files')
                continue

            jobs.append({
                'shop':     shop,
                'year':     year,
                'month':    mon_idx,
                'month_str': month_str,
                'mon_name': MONTH_NAMES[mon_idx - 1],
                'merchant': SHOP_MID[shop],
                'pos':      pos_files[0],
                'kpay':     kpay_files[0],
                'dbs':      dbs_files[0],
            })

    return jobs


def run_job(job, output_dir):
    """Run a single reconciliation job and save the output."""
    shop = job['shop']
    year = job['year']
    month = job['month']
    mon_name = job['mon_name']
    merchant = job['merchant']

    out_folder = os.path.join(output_dir, job['month_str'])
    os.makedirs(out_folder, exist_ok=True)
    out_file = os.path.join(out_folder,
                            f'{shop}_KPay_{mon_name}{year}_Reconciliation.xlsx')

    print(f'\n{"─"*60}')
    print(f'  {shop} — {mon_name} {year}')
    print(f'{"─"*60}')

    # 1. KPay
    kpay_data, settle_dates, methods_found = read_kpay(job['kpay'])
    methods, rare = determine_methods(methods_found)
    print(f'  KPay: {len(settle_dates)} settle dates, '
          f'{len(methods_found)} methods'
          f'{f" (+{len(rare)} rare)" if rare else ""}')

    # 2. POS
    pos_daily, _ = read_pos(job['pos'], shop)
    print(f'  POS:  {len(pos_daily)} trading days')

    # 3. DBS
    dbs_by_date, _ = read_dbs(job['dbs'], merchant, year, month)
    print(f'  DBS:  {len(dbs_by_date)} credit dates, '
          f'HKD {sum(dbs_by_date.values()):,.2f}')

    # 4. Build
    wb = Workbook()
    wb.remove(wb.active)
    build_sheet(wb, pos_daily, kpay_data, settle_dates,
                dbs_by_date, shop, year, month, methods, rare)
    wb.save(out_file)

    # 5. Validate
    passed, results = validate_output(out_file, year, month)
    status = '✓' if passed else '✗'
    fails = [n for n, ok, _ in results if not ok]
    print(f'  {status} → {out_file}')
    if fails:
        print(f'    FAILED: {fails}')

    return passed, out_file


def main():
    ap = argparse.ArgumentParser(description='Batch KPay Reconciliation')
    ap.add_argument('--month', help='Filter by YYYYMM (e.g. 202602)')
    ap.add_argument('--shop',  help='Filter by shop code (e.g. KYR)')
    ap.add_argument('--output', default='output', help='Output root folder')
    args = ap.parse_args()

    base = os.path.normpath(BASE_DIR)
    print(f'Scanning: {base}')
    jobs = discover_jobs(base, args.month, args.shop)

    if not jobs:
        print('No source data found.')
        sys.exit(1)

    print(f'Found {len(jobs)} job(s):')
    for j in jobs:
        print(f'  {j["shop"]:6s}  {j["month_str"]}')

    output_dir = os.path.join(os.path.dirname(__file__), '..', args.output)
    passed_all = True
    generated = []

    for job in jobs:
        ok, path = run_job(job, output_dir)
        generated.append((job['shop'], job['month_str'], ok, path))
        if not ok:
            passed_all = False

    # Summary
    print(f'\n{"="*60}')
    print(f'  BATCH COMPLETE — {len(generated)} reports')
    print(f'{"="*60}')
    for shop, mth, ok, path in generated:
        print(f'  {"✓" if ok else "✗"} {shop:6s} {mth}  {path}')

    if not passed_all:
        print('\nSome reports had validation failures.')
        sys.exit(1)
    else:
        print('\nAll reports passed validation.')


if __name__ == '__main__':
    main()
