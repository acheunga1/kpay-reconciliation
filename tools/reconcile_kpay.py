"""
KPay × DBS × POS Monthly Reconciliation Tool
=============================================
Generates an Excel reconciliation report matching the exact format of the
reference template (KPay-By Date_Reconciliation.xlsx).

Usage:
    python tools/reconcile_kpay.py \
      --pos   input/POS_Sales_202508.xlsx \
      --kpay  input/AUg_Transaction_report.xlsx \
      --dbs   input/DBS_Sunstage_002378792_HKD_202508.xls \
      --month 202508 \
      --store KYR \
      --merchant 146100005 \
      --reference reference/KPay-By_Date_Reconciliation.xlsx \
      --output output/KPay_Aug2025_Reconciliation.xlsx
"""

import argparse
import re
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta

import openpyxl
import xlrd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS  (column numbers, 1-based, matching the workflow spec exactly)
# ─────────────────────────────────────────────────────────────────────────────

KPAY_METHODS = ['Mastercard', 'PayMe', 'Visa', '中國銀聯', '支付寶', '微信', '銀聯雲閃付']

# Section start columns (1-based)
GRS_START  = 1    # A  – 加總 - 支付金額 (label=A, methods B-H, total I)
FEE_START  = 16   # P  – 加總 - 手續費   (label=P, methods Q-W, total X)
NET_START  = 31   # AE – 加總 - Net Amount (label=AE, methods AF-AL, total AM)
AR_COL     = 44   # AR – MUST BE COMPLETELY EMPTY (spacing column)
DBS_START  = 47   # AU – DBS data         (AU=date, AV=credit, AW=txn variance)
RATE_START = 53   # BA – Fee rate check   (label=BA, methods BB-BH, total BI)

# Row positions
ROW_SECTION_HDR = 3
ROW_COL_HDR     = 4
ROW_DATA_START  = 5
ROW_TOTAL       = 35   # KPay totals ("總計")
ROW_EMPTY_1     = 36
ROW_EMPTY_2     = 37
ROW_POS         = 38   # POS monthly data
ROW_DIFF        = 39   # Difference formulas (=ROW_TOTAL - ROW_POS)

# POS column indices — detected dynamically per file; this is the Aug fallback
POS_COL_FALLBACK = {
    'store': 0, 'date': 1,
    'Cash': 2, 'FPS': 3, 'Master': 4, 'Visa': 5,
    'Octopus': 6, 'PayMe': 7, 'Alipay': 8, 'WeChat': 9,
    'AE': 10, 'CashVoucher': 11, 'UnionPay': 12, 'Total': 13,
}

# Keywords used to auto-detect POS columns from header row (case-insensitive)
POS_HEADER_KEYWORDS = {
    'Master':     ['master'],
    'Visa':       ['visa'],
    'PayMe':      ['payme'],
    'Alipay':     ['alipay', '支付寶', '支付宝'],
    'WeChat':     ['wechat', '微信'],
    'UnionPay':   ['銀聯', '银联', 'boc pay', 'unionpay'],
    'Cash':       ['cash'],
    'FPS':        ['fps'],
    'Octopus':    ['octopus'],
    'AE':         ['american express'],
    'CashVoucher':['voucher', '現金券', '现金券'],
    'Total':      ['total', '總計', '总计'],
}


def detect_pos_columns(ws):
    """
    Scan the worksheet header rows to find POS column positions dynamically.
    Returns a dict of {field: 0-based column index}.
    Falls back to the August hardcoded layout for any undetected fields.
    """
    result = {'store': 0, 'date': 1}
    for hr in range(1, 5):
        for c in range(1, 25):
            v = ws.cell(hr, c).value
            if not v or not isinstance(v, str):
                continue
            vl = v.lower()
            for field, kws in POS_HEADER_KEYWORDS.items():
                if field not in result and any(kw.lower() in vl for kw in kws):
                    result[field] = c - 1   # 0-indexed
                    break
        if 'Master' in result and 'Visa' in result and 'PayMe' in result:
            break   # found the important ones
    # fill any still-missing fields from fallback
    for k, v in POS_COL_FALLBACK.items():
        result.setdefault(k, v)
    return result

# KPay column indices (0-based)
KPAY_COL = {
    'txn_id': 0, 'status': 6, 'amount': 8, 'fee': 13,
    'method': 18, 'settle_date': 24,
}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def fmt_date(d):
    """date/datetime → 'dd/mm/yyyy' string."""
    if hasattr(d, 'strftime'):
        return d.strftime('%d/%m/%Y')
    return str(d)


def parse_date_str(s):
    """'dd/mm/yyyy' or '01-Aug-2025' or 'yyyy-mm-dd' → date, or None."""
    s = str(s).strip()
    for fmt in ('%d/%m/%Y', '%d-%b-%Y', '%Y-%m-%d', '%d-%B-%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except (ValueError, TypeError):
            pass
    return None


def col_ltr(n):
    """1-based column number → Excel letter(s). E.g. 44 → 'AR'."""
    return get_column_letter(n)


# HK public holidays for T+2 business day calculation
HK_HOLIDAYS = {
    date(2025, 1, 1), date(2025, 1, 29), date(2025, 1, 30), date(2025, 1, 31),
    date(2025, 4, 4), date(2025, 4, 18), date(2025, 4, 19), date(2025, 4, 21),
    date(2025, 5, 1), date(2025, 5, 5),
    date(2025, 6, 2),
    date(2025, 7, 1),
    date(2025, 10, 1), date(2025, 10, 7), date(2025, 10, 29),
    date(2025, 12, 25), date(2025, 12, 26),
    date(2026, 1, 1),
}


def next_biz_day(d, n=2):
    """Return d + n business days, skipping weekends and HK public holidays."""
    count, cur = 0, d
    while count < n:
        cur += timedelta(days=1)
        if cur.weekday() < 5 and cur not in HK_HOLIDAYS:
            count += 1
    return cur


# ─────────────────────────────────────────────────────────────────────────────
# 1. DATA READERS
# ─────────────────────────────────────────────────────────────────────────────

def read_pos(pos_file, store_code):
    """
    Read POS daily sales for the given store from the POS Excel file.

    Returns:
        daily    : {date_str: {Master, PayMe, Visa, UnionPay, Alipay, WeChat, ...}}
        non_kpay : {tender: monthly_total}
    """
    wb = openpyxl.load_workbook(pos_file)
    if 'Sheet (2)' in wb.sheetnames:
        ws = wb['Sheet (2)']
    else:
        ws = wb.active

    PC = detect_pos_columns(ws)
    print(f'      POS columns detected: Master={PC["Master"]}, Visa={PC["Visa"]}, '
          f'PayMe={PC["PayMe"]}, Alipay={PC["Alipay"]}, '
          f'WeChat={PC["WeChat"]}, UnionPay={PC["UnionPay"]}')

    daily = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[PC['store']] != store_code:
            continue
        d = row[PC['date']]
        if d is None:
            continue
        ds = fmt_date(d)
        daily[ds] = {
            'Master':      row[PC['Master']]      or 0,
            'PayMe':       row[PC['PayMe']]       or 0,
            'Visa':        row[PC['Visa']]        or 0,
            'UnionPay':    row[PC['UnionPay']]    or 0,
            'Alipay':      row[PC['Alipay']]      or 0,
            'WeChat':      row[PC['WeChat']]      or 0,
            'Cash':        row[PC.get('Cash', 2)] or 0,
            'FPS':         row[PC.get('FPS', 3)]  or 0,
            'Octopus':     row[PC.get('Octopus', 6)] or 0,
            'AE':          row[PC.get('AE', 9)]   or 0,
            'CashVoucher': row[PC.get('CashVoucher', 10)] or 0,
        }

    non_kpay = {k: sum(v[k] for v in daily.values())
                for k in ['Cash', 'FPS', 'Octopus', 'AE', 'CashVoucher']}
    wb.close()
    return daily, non_kpay


def read_kpay(kpay_file):
    """
    Read KPay settlement transactions.

    Returns:
        by_settle : {date_str: {method: {gross, fee}}}
        settle_dates : sorted list of date_str
    """
    # Support both .xlsx and legacy .xls KPay exports
    if str(kpay_file).lower().endswith('.xls'):
        import xlrd
        xwb = xlrd.open_workbook(kpay_file)
        try:
            xws = xwb.sheet_by_name('交易結算')
        except xlrd.XLRDError:
            xws = xwb.sheet_by_index(0)
        # Convert xlrd sheet to list-of-rows for uniform processing
        rows_data = [xws.row_values(i) for i in range(xws.nrows)]
        by_settle = defaultdict(lambda: defaultdict(lambda: {'gross': 0.0, 'fee': 0.0}))
        for row in rows_data[9:]:  # start_row=10 → index 9
            if not row[KPAY_COL['txn_id']]:
                continue
            status = str(row[KPAY_COL['status']]).strip()
            if status != '處理成功':
                continue
            method = str(row[KPAY_COL['method']]).strip()
            gross  = float(row[KPAY_COL['amount']] or 0)
            fee    = float(row[KPAY_COL['fee']]    or 0)
            raw_sd = row[KPAY_COL['settle_date']]
            if isinstance(raw_sd, float):
                import datetime as _dt
                sd = xlrd.xldate_as_datetime(raw_sd, xwb.datemode).strftime('%Y/%m/%d')
            else:
                sd = str(raw_sd).strip()[:10].replace('-', '/')
            if not sd:
                continue
            by_settle[sd][method]['gross'] += gross
            by_settle[sd][method]['fee']   += fee
        settle_dates = sorted(by_settle.keys())
        return by_settle, settle_dates

    wb = openpyxl.load_workbook(kpay_file)
    if '交易結算' in wb.sheetnames:
        ws = wb['交易結算']
    else:
        ws = wb.active

    by_settle = defaultdict(lambda: defaultdict(lambda: {'gross': 0.0, 'fee': 0.0}))
    for row in ws.iter_rows(min_row=10, values_only=True):
        if row[KPAY_COL['txn_id']] is None:
            continue
        if row[KPAY_COL['status']] != '處理成功':
            continue
        sd_raw = row[KPAY_COL['settle_date']]
        method = row[KPAY_COL['method']]
        amount = row[KPAY_COL['amount']] or 0
        fee    = row[KPAY_COL['fee']]    or 0
        if sd_raw and method:
            ds = fmt_date(sd_raw) if hasattr(sd_raw, 'strftime') else str(sd_raw).strip()
            by_settle[ds][method]['gross'] += amount
            by_settle[ds][method]['fee']   += fee

    settle_dates = sorted(by_settle.keys(),
                          key=lambda s: parse_date_str(s) or date.min)
    wb.close()
    return dict(by_settle), settle_dates


def read_dbs(dbs_file, merchant_id, year, month):
    """
    Read DBS bank statement (.xls) filtered to the target month and merchant.
    Uses xlrd directly — no LibreOffice dependency.

    Returns:
        by_date : {date_str: total_credit}  (only for merchant_id, exact month)
        batches : list of {date, credit, batch_no}
    """
    start_date = date(year, month, 1)
    if month == 12:
        end_date = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(year, month + 1, 1) - timedelta(days=1)

    xls = xlrd.open_workbook(dbs_file)
    ws  = xls.sheet_by_index(0)

    by_date = defaultdict(float)
    batches = []

    for r in range(6, ws.nrows):   # row 6 (0-indexed) = first data row
        row = [ws.cell_value(r, c) for c in range(ws.ncols)]
        if not row[2] or 'KPAY' not in str(row[2]).upper():
            continue

        # Parse date (column 0)
        date_raw = row[0]
        if isinstance(date_raw, float):
            # xlrd stores dates as floats
            tup = xlrd.xldate_as_tuple(date_raw, xls.datemode)
            d = date(*tup[:3])
        else:
            d = parse_date_str(date_raw)
        if d is None:
            continue

        # STRICT month filter (Rule 4)
        if not (start_date <= d <= end_date):
            continue

        credit = row[5] or 0
        desc2  = str(row[3]) if row[3] else ''

        if merchant_id not in desc2:
            continue

        ds = fmt_date(d)
        by_date[ds] += credit

        batch_match = re.search(rf'{re.escape(merchant_id)}K(\d+)', desc2)
        batch_no = batch_match.group(1) if batch_match else ''
        batches.append({'date': ds, 'credit': credit, 'batch_no': batch_no})

    return dict(by_date), batches


# ─────────────────────────────────────────────────────────────────────────────
# 2. STYLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _hdr(cell, text=None, bg='1F4E79', fg='FFFFFF', bold=True, size=9):
    if text is not None:
        cell.value = text
    cell.font      = Font(name='Arial', bold=bold, color=fg, size=size)
    cell.fill      = PatternFill('solid', start_color=bg)
    cell.alignment = Alignment(horizontal='center', vertical='center',
                                wrap_text=True)


def _num(cell, val=None, fmt='#,##0.00', color='000000', bold=False):
    if val is not None:
        cell.value = val
    cell.number_format = fmt
    cell.font      = Font(name='Arial', color=color, size=9, bold=bold)
    cell.alignment = Alignment(horizontal='right', vertical='center')


def _total_style(cell):
    cell.font  = Font(name='Arial', bold=True, size=9)
    cell.fill  = PatternFill('solid', start_color='FFF2CC')
    cell.alignment = Alignment(horizontal='right', vertical='center')


def _lbl(cell, text, bold=False):
    cell.value = text
    cell.font  = Font(name='Arial', bold=bold, size=9)
    cell.alignment = Alignment(horizontal='center', vertical='center')


# ─────────────────────────────────────────────────────────────────────────────
# 3. EXCEL BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def build_sheet(wb, pos_daily, kpay_by_settle, settle_dates, dbs_by_date,
                store_code, year, month):
    """
    Build the 'KPay-By Date_Reconciliation' sheet.
    Column layout (1-based):
      GRS  A(1)-I(9)     Gross / 支付金額
      FEE  P(16)-X(24)   Fee / 手續費
      NET  AE(31)-AM(39) Net Amount
      AR   44             EMPTY (spacing)
      DBS  AU(47)-AW(49) DBS credits
      RATE BA(53)-BI(61) Fee rate check
    """
    ws = wb.create_sheet('KPay-By Date_Reconciliation')

    # ── Row 3: section headers ──────────────────────────────────────────────
    sections = [
        (GRS_START,  '加總 - 支付金額'),
        (FEE_START,  '加總 - 手續費'),
        (NET_START,  '加總 - Net Amount'),
        (RATE_START, 'Checking on Credit Card Charge Rate'),
    ]
    for start, label in sections:
        _hdr(ws.cell(ROW_SECTION_HDR, start), label)
        ws.merge_cells(start_row=ROW_SECTION_HDR, start_column=start,
                       end_row=ROW_SECTION_HDR,   end_column=start + 8)

    # DBS section header (row 3)
    _hdr(ws.cell(ROW_SECTION_HDR, DBS_START), 'DBS Receipt', bg='833C00')
    ws.merge_cells(start_row=ROW_SECTION_HDR, start_column=DBS_START,
                   end_row=ROW_SECTION_HDR,   end_column=DBS_START + 2)

    ws.row_dimensions[ROW_SECTION_HDR].height = 25

    # ── Row 4: column sub-headers ───────────────────────────────────────────
    for sec in [GRS_START, FEE_START, NET_START, RATE_START]:
        _hdr(ws.cell(ROW_COL_HDR, sec), '列標籤', bg='2F75B6')
        for i, m in enumerate(KPAY_METHODS):
            _hdr(ws.cell(ROW_COL_HDR, sec + 1 + i), m, bg='2F75B6')
        _hdr(ws.cell(ROW_COL_HDR, sec + 8), '總計', bg='2F75B6')

    # DBS sub-headers
    for ci, lbl in enumerate(['DBS Receipt', 'DBS Receipt2', 'Txn Charge'],
                              DBS_START):
        _hdr(ws.cell(ROW_COL_HDR, ci), lbl, bg='833C00')

    ws.row_dimensions[ROW_COL_HDR].height = 30

    # ── Rows 5-34: daily KPay data ──────────────────────────────────────────
    last_data_row = ROW_DATA_START - 1
    for ri, ds in enumerate(settle_dates):
        r = ROW_DATA_START + ri
        if r > 34:
            break   # template only has 30 data rows (rows 5-34)
        last_data_row = r
        day = kpay_by_settle.get(ds, {})
        gross_vals = [day.get(m, {}).get('gross', 0) or None for m in KPAY_METHODS]
        fee_vals   = [day.get(m, {}).get('fee',   0) or None for m in KPAY_METHODS]

        # ─ Gross section ─
        _lbl(ws.cell(r, GRS_START), ds)
        for i, v in enumerate(gross_vals):
            _num(ws.cell(r, GRS_START + 1 + i), v)
        tot = ws.cell(r, GRS_START + 8)
        tot.value = f'=SUM({col_ltr(GRS_START+1)}{r}:{col_ltr(GRS_START+7)}{r})'
        _num(tot)

        # ─ Fee section ─
        ws.cell(r, FEE_START).value = f'={col_ltr(GRS_START)}{r}'
        ws.cell(r, FEE_START).font  = Font(name='Arial', size=9)
        ws.cell(r, FEE_START).alignment = Alignment(horizontal='center')
        for i, v in enumerate(fee_vals):
            _num(ws.cell(r, FEE_START + 1 + i), v)
        tot = ws.cell(r, FEE_START + 8)
        tot.value = f'=SUM({col_ltr(FEE_START+1)}{r}:{col_ltr(FEE_START+7)}{r})'
        _num(tot)

        # ─ Net section ─
        ws.cell(r, NET_START).value = f'={col_ltr(GRS_START)}{r}'
        ws.cell(r, NET_START).font  = Font(name='Arial', size=9)
        ws.cell(r, NET_START).alignment = Alignment(horizontal='center')
        for i in range(7):
            g = col_ltr(GRS_START + 1 + i)
            f = col_ltr(FEE_START + 1 + i)
            c = ws.cell(r, NET_START + 1 + i)
            c.value = f'={g}{r}-{f}{r}'
            _num(c)
        tot = ws.cell(r, NET_START + 8)
        tot.value = f'=SUM({col_ltr(NET_START+1)}{r}:{col_ltr(NET_START+7)}{r})'
        _num(tot)

        # ─ Rate section ─
        # Use plain division with 0.00% format — Excel multiplies by 100 for display
        # e.g. 0.0145 → shows as 1.45% (matches reference template exactly)
        ws.cell(r, RATE_START).value = f'={col_ltr(GRS_START)}{r}'
        ws.cell(r, RATE_START).font  = Font(name='Arial', size=9)
        ws.cell(r, RATE_START).alignment = Alignment(horizontal='center')
        for i in range(7):
            g = col_ltr(GRS_START + 1 + i)
            f = col_ltr(FEE_START + 1 + i)
            c = ws.cell(r, RATE_START + 1 + i)
            c.value = f'=IFERROR({f}{r}/{g}{r},"-")'
            c.number_format = '0.00%'
            c.font = Font(name='Arial', size=9)
            c.alignment = Alignment(horizontal='right')
        tot_r = ws.cell(r, RATE_START + 8)
        grs_tot = col_ltr(GRS_START + 8)
        fee_tot = col_ltr(FEE_START + 8)
        tot_r.value = f'=IFERROR({fee_tot}{r}/{grs_tot}{r},"-")'
        tot_r.number_format = '0.00%'
        tot_r.font = Font(name='Arial', size=9)
        tot_r.alignment = Alignment(horizontal='right')

    # ── Build T+2 map (KPay settle date → expected DBS receipt date) ─────────
    # For each KPay settle date (at a known row), compute T+2 business days
    # to determine which DBS receipt date it corresponds to.
    t2_map = defaultdict(list)  # {dbs_date_str: [kpay_row_nums]}
    for ri, ds in enumerate(settle_dates):
        kpay_r = ROW_DATA_START + ri
        if kpay_r > 34:
            break
        sd = parse_date_str(ds)
        if sd:
            t2_date = next_biz_day(sd, 2)
            t2_ds = fmt_date(t2_date)
            t2_map[t2_ds].append(kpay_r)

    # ── DBS receipt data (AU/AV/AW columns, rows 5+) ─────────────────────────
    last_dbs_row       = ROW_DATA_START - 1
    first_match_dbs_row = None
    last_match_dbs_row  = None
    net_tot_col = col_ltr(NET_START + 8)   # AM column
    av_col      = col_ltr(DBS_START + 1)   # AV column

    sorted_dbs = sorted(dbs_by_date.items(),
                        key=lambda x: parse_date_str(x[0]) or date.min)
    for dbs_ri, (d_str, credit) in enumerate(sorted_dbs):
        r = ROW_DATA_START + dbs_ri
        if r > 34:
            break
        last_dbs_row = r
        d_obj = parse_date_str(d_str)
        dbs_date_fmt = d_obj.strftime('%d-%b-%Y') if d_obj else d_str

        # AU: DBS receipt date
        ws.cell(r, DBS_START).value     = dbs_date_fmt
        ws.cell(r, DBS_START).font      = Font(name='Arial', size=9)
        ws.cell(r, DBS_START).alignment = Alignment(horizontal='center')

        # AV: DBS credit amount
        ws.cell(r, DBS_START + 1).value = credit
        _num(ws.cell(r, DBS_START + 1))

        # AW: Txn Charge variance (KPay net total - DBS credit) for T+2-matched rows
        kpay_rows = t2_map.get(d_str, [])
        if kpay_rows:
            net_refs = '+'.join(f'{net_tot_col}{kr}' for kr in kpay_rows)
            ws.cell(r, DBS_START + 2).value = f'={net_refs}-{av_col}{r}'
            _num(ws.cell(r, DBS_START + 2))
            if first_match_dbs_row is None:
                first_match_dbs_row = r
            last_match_dbs_row = r

    # ── Row 35: KPay totals ("總計") ────────────────────────────────────────
    for sec in [GRS_START, FEE_START, NET_START]:
        c35 = ws.cell(ROW_TOTAL, sec)
        c35.value = '總計'
        _total_style(c35)
        ws.cell(ROW_TOTAL, sec).alignment = Alignment(horizontal='right',
                                                       vertical='center')
        for i in range(8):
            col = col_ltr(sec + 1 + i)
            c = ws.cell(ROW_TOTAL, sec + 1 + i)
            c.value = f'=SUM({col}{ROW_DATA_START}:{col}{last_data_row})'
            _total_style(c)
            _num(c)

    # ── Rate section row 35 totals ──────────────────────────────────────────
    ws.cell(ROW_TOTAL, RATE_START).value = '總計'
    _total_style(ws.cell(ROW_TOTAL, RATE_START))
    ws.cell(ROW_TOTAL, RATE_START).alignment = Alignment(horizontal='right',
                                                         vertical='center')
    for i in range(7):
        g35 = col_ltr(GRS_START + 1 + i)
        f35 = col_ltr(FEE_START + 1 + i)
        c = ws.cell(ROW_TOTAL, RATE_START + 1 + i)
        c.value = f'=IFERROR({f35}{ROW_TOTAL}/{g35}{ROW_TOTAL},"-")'
        c.number_format = '0.00%'
        c.font = Font(name='Arial', size=9, bold=True)
        c.alignment = Alignment(horizontal='right')
        _total_style(c)
    tr = ws.cell(ROW_TOTAL, RATE_START + 8)
    tr.value = (f'=IFERROR({col_ltr(FEE_START+8)}{ROW_TOTAL}'
                f'/{col_ltr(GRS_START+8)}{ROW_TOTAL},"-")')
    tr.number_format = '0.00%'
    tr.font = Font(name='Arial', size=9, bold=True)
    tr.alignment = Alignment(horizontal='right')
    _total_style(tr)

    # DBS total row 35
    ws.cell(ROW_TOTAL, DBS_START).value = '總計'
    _total_style(ws.cell(ROW_TOTAL, DBS_START))
    ws.cell(ROW_TOTAL, DBS_START).alignment = Alignment(horizontal='right')
    dbs_amt_col = col_ltr(DBS_START + 1)
    ws.cell(ROW_TOTAL, DBS_START + 1).value = (
        f'=SUM({dbs_amt_col}{ROW_DATA_START}:{dbs_amt_col}{last_dbs_row})')
    _total_style(ws.cell(ROW_TOTAL, DBS_START + 1))
    _num(ws.cell(ROW_TOTAL, DBS_START + 1))
    # AW35: sum of AW variance rows
    aw_col = col_ltr(DBS_START + 2)
    ws.cell(ROW_TOTAL, DBS_START + 2).value = (
        f'=SUM({aw_col}{ROW_DATA_START}:{aw_col}{last_dbs_row})')
    _total_style(ws.cell(ROW_TOTAL, DBS_START + 2))
    _num(ws.cell(ROW_TOTAL, DBS_START + 2))

    # ── Rows 36-37: EMPTY (no data written — critical) ──────────────────────
    # (nothing to do here; just leave empty)

    # ── Row 38: POS monthly totals ──────────────────────────────────────────
    pos_totals = {k: sum(v[k] for v in pos_daily.values()) for k in
                  ['Master', 'PayMe', 'Visa', 'UnionPay', 'Alipay', 'WeChat']}

    # POS payment methods map to KPay column order:
    # B=Master, C=PayMe, D=Visa, E=UnionPay, F=Alipay, G=WeChat, H=0 (雲閃付)
    pos_row = [
        pos_totals['Master'],
        pos_totals['PayMe'],
        pos_totals['Visa'],
        pos_totals['UnionPay'],   # 中國銀聯 column
        pos_totals['Alipay'],
        pos_totals['WeChat'],
        0,                        # 銀聯雲閃付 — no separate POS category
    ]

    ws.cell(ROW_POS, GRS_START).value = f'POS ({store_code})'
    _hdr(ws.cell(ROW_POS, GRS_START), bg='833C00')
    for i, v in enumerate(pos_row):
        c = ws.cell(ROW_POS, GRS_START + 1 + i)
        c.value = v   # write 0 explicitly (not None) for empty payment methods
        _num(c, color='0000FF')
    grs_lbl = col_ltr(GRS_START + 1)
    grs_end = col_ltr(GRS_START + 7)
    ws.cell(ROW_POS, GRS_START + 8).value = (
        f'=SUM({grs_lbl}{ROW_POS}:{grs_end}{ROW_POS})')
    _num(ws.cell(ROW_POS, GRS_START + 8))

    # Row 38 NET section: AM38 = SUM of KPay net column (used by AW38)
    net_tot_col = col_ltr(NET_START + 8)   # AM column
    ws.cell(ROW_POS, NET_START + 8).value = (
        f'=SUM({net_tot_col}{ROW_DATA_START}:{net_tot_col}{last_data_row})')
    _num(ws.cell(ROW_POS, NET_START + 8))

    # Row 38 DBS section: "Per DBS" label + sum of all DBS credits + AW variance
    ws.cell(ROW_POS, DBS_START).value = 'Per DBS'
    _hdr(ws.cell(ROW_POS, DBS_START), bg='833C00')

    av_col = col_ltr(DBS_START + 1)
    last_av_row = last_match_dbs_row if last_match_dbs_row else last_dbs_row
    ws.cell(ROW_POS, DBS_START + 1).value = (
        f'=SUM({av_col}{ROW_DATA_START}:{av_col}{last_av_row})')
    _num(ws.cell(ROW_POS, DBS_START + 1))

    # AW38: KPay net total (AM38) - DBS credit total (AV38)
    ws.cell(ROW_POS, DBS_START + 2).value = f'=+{net_tot_col}{ROW_POS}-{av_col}{ROW_POS}'
    _num(ws.cell(ROW_POS, DBS_START + 2))

    # ── Row 39: Difference formulas (Rule 2 & 3) ────────────────────────────
    ws.cell(ROW_DIFF, GRS_START).value = 'Difference'
    _hdr(ws.cell(ROW_DIFF, GRS_START), bg='C00000')

    for i in range(8):   # columns B-I (GRS_START+1 to GRS_START+8)
        col = col_ltr(GRS_START + 1 + i)
        c = ws.cell(ROW_DIFF, GRS_START + 1 + i)
        c.value = f'={col}{ROW_TOTAL}-{col}{ROW_POS}'
        c.number_format = '#,##0.00;(#,##0.00);"-"'
        c.font = Font(name='Arial', bold=True, size=9, color='000000')
        c.alignment = Alignment(horizontal='right')

    # DBS difference in row 39 (Per KPay - DBS Total)
    ws.cell(ROW_DIFF, DBS_START).value = 'Checking'
    _hdr(ws.cell(ROW_DIFF, DBS_START), bg='833C00')
    dbs_amt = col_ltr(DBS_START + 1)
    ws.cell(ROW_DIFF, DBS_START + 1).value = (
        f'={dbs_amt}{ROW_POS}-{dbs_amt}{ROW_TOTAL}')
    _num(ws.cell(ROW_DIFF, DBS_START + 1))
    ws.cell(ROW_DIFF, DBS_START + 1).font = Font(name='Arial', bold=True, size=9)

    # ── Column widths ────────────────────────────────────────────────────────
    ws.column_dimensions[col_ltr(GRS_START)].width = 13
    for ci in range(2, 62):
        ws.column_dimensions[col_ltr(ci)].width = 11
    ws.column_dimensions[col_ltr(AR_COL)].width = 3    # visible spacing gap

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# 4. VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

def validate_output(output_path, year, month):
    """
    Run all critical validation checks against the generated output file.
    Returns (passed: bool, results: list of (check, ok, detail))
    """
    wb = openpyxl.load_workbook(output_path)
    results = []

    def chk(name, ok, detail=''):
        results.append((name, ok, detail))

    # Check 1: sheet name
    sheet_name = 'KPay-By Date_Reconciliation'
    has_sheet = sheet_name in wb.sheetnames
    chk('Sheet name correct', has_sheet,
        f'Found: {wb.sheetnames}' if not has_sheet else sheet_name)

    if not has_sheet:
        chk('(remaining checks skipped — wrong sheet name)', False, '')
        return False, results

    ws = wb[sheet_name]

    # Check 2: Row 35 label
    r35_a = ws.cell(ROW_TOTAL, GRS_START).value
    chk('Row 35 col A == "總計"', r35_a == '總計',
        f'Got: {r35_a!r}')

    # Check 3: Rows 36-37 empty
    rows_empty = True
    for r in [ROW_EMPTY_1, ROW_EMPTY_2]:
        for c in range(1, 62):
            v = ws.cell(r, c).value
            if v is not None:
                rows_empty = False
                chk(f'Rows 36-37 empty', False,
                    f'Row {r} col {c} ({col_ltr(c)}) = {v!r}')
                break
        if not rows_empty:
            break
    if rows_empty:
        chk('Rows 36-37 empty', True)

    # Check 4: Row 38 POS label
    r38_a = ws.cell(ROW_POS, GRS_START).value
    chk('Row 38 col A has POS label', r38_a is not None and 'POS' in str(r38_a),
        f'Got: {r38_a!r}')

    # Check 5: Row 39 Difference label
    r39_a = ws.cell(ROW_DIFF, GRS_START).value
    chk('Row 39 col A == "Difference"', r39_a == 'Difference',
        f'Got: {r39_a!r}')

    # Check 6: Column AR (44) completely empty
    ar_empty = True
    ar_violations = []
    for r in range(1, 50):
        v = ws.cell(r, AR_COL).value
        if v is not None:
            ar_empty = False
            ar_violations.append(f'Row {r}: {v!r}')
    chk('Column AR (44) completely empty', ar_empty,
        '; '.join(ar_violations) if ar_violations else '')

    # Check 7: Row 39 has difference formulas in ALL columns B-I
    diff_ok = True
    missing = []
    for i in range(8):
        col_num = GRS_START + 1 + i
        cell = ws.cell(ROW_DIFF, col_num)
        v = cell.value
        if v is None or str(v).strip() == '':
            diff_ok = False
            missing.append(col_ltr(col_num))
    chk('Row 39 has formulas in all columns B-I', diff_ok,
        f'Missing: {missing}' if missing else '')

    # Check 8: DBS dates filtered to exact month
    dbs_ok = True
    bad_dates = []
    start_d = date(year, month, 1)
    if month == 12:
        end_d = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_d = date(year, month + 1, 1) - timedelta(days=1)
    for r in range(ROW_DATA_START, 35):
        v = ws.cell(r, DBS_START).value
        if v is None or v in ('總計', 'Per KPay', 'Checking'):
            continue
        d = parse_date_str(str(v))
        if d and not (start_d <= d <= end_d):
            dbs_ok = False
            bad_dates.append(f'row{r}:{v}')
    chk('DBS dates filtered to exact month', dbs_ok,
        '; '.join(bad_dates) if bad_dates else '')

    # Check 9: Rate formulas use IFERROR with plain division (no ROUND — matches reference)
    rate_ok = True
    sample = ws.cell(ROW_DATA_START, RATE_START + 1)  # first rate cell (BB5)
    v = sample.value
    if v is not None and isinstance(v, str):
        rate_ok = 'IFERROR' in v.upper() and 'ROUND' not in v.upper()
    chk('Rate formulas use IFERROR (no ROUND)', rate_ok,
        f'Sample: {v!r}')

    # Check 10: File integrity (already opened successfully)
    chk('File opens without error', True)

    passed = all(ok for _, ok, _ in results)
    wb.close()
    return passed, results


# ─────────────────────────────────────────────────────────────────────────────
# 5. MAIN
# ─────────────────────────────────────────────────────────────────────────────

def run(pos_file, kpay_file, dbs_file, month_str, store_code, merchant_id,
        output_path, reference_path=None):
    year  = int(month_str[:4])
    month = int(month_str[4:])

    print(f'[1/5] Reading POS data ({store_code}) from {pos_file}…')
    pos_daily, pos_non_kpay = read_pos(pos_file, store_code)
    print(f'      {len(pos_daily)} trading days found for {store_code}')

    print(f'[2/5] Reading KPay transactions from {kpay_file}…')
    kpay_by_settle, settle_dates = read_kpay(kpay_file)
    total_gross = sum(v['gross'] for d in kpay_by_settle.values()
                      for v in d.values())
    print(f'      {len(settle_dates)} settle dates, gross total = {total_gross:,.2f}')

    print(f'[3/5] Reading DBS statement from {dbs_file}…')
    dbs_by_date, dbs_batches = read_dbs(dbs_file, merchant_id, year, month)
    dbs_total = sum(dbs_by_date.values())
    print(f'      {len(dbs_by_date)} DBS credit dates, total = {dbs_total:,.2f}')
    print(f'      Merchant {merchant_id} batches found: {len(dbs_batches)}')

    print(f'[4/5] Building Excel workbook…')
    wb = Workbook()
    wb.remove(wb.active)   # drop default blank sheet

    build_sheet(wb, pos_daily, kpay_by_settle, settle_dates, dbs_by_date,
                store_code, year, month)

    wb.save(output_path)
    print(f'      Saved: {output_path}')

    print(f'[5/5] Validating output…')
    passed, results = validate_output(output_path, year, month)

    print()
    print('-' * 60)
    print('VALIDATION RESULTS')
    print('-' * 60)
    all_pass = True
    for name, ok, detail in results:
        icon = '[PASS]' if ok else '[FAIL]'
        line = f'{icon} {name}'
        if detail:
            line += f'  ->  {detail}'
        print(line)
        if not ok:
            all_pass = False

    print('-' * 60)
    if all_pass:
        print('[ALL CHECKS PASSED]')
    else:
        print('[SOME CHECKS FAILED] -- review output above')
    print(f'Output: {output_path}')

    return all_pass


def main():
    ap = argparse.ArgumentParser(description='KPay × DBS × POS Reconciliation')
    ap.add_argument('--pos',       required=True,  help='POS_Sales_YYYYMM.xlsx')
    ap.add_argument('--kpay',      required=True,  help='KPay Transaction report')
    ap.add_argument('--dbs',       required=True,  help='DBS_Sunstage_*.xls')
    ap.add_argument('--month',     required=True,  help='YYYYMM e.g. 202508')
    ap.add_argument('--store',     default='KYR',  help='Store code')
    ap.add_argument('--merchant',  default='146100005', help='KPay merchant ID')
    ap.add_argument('--reference', default=None,   help='Reference template (unused)')
    ap.add_argument('--output',    required=True,  help='Output .xlsx path')
    args = ap.parse_args()

    ok = run(args.pos, args.kpay, args.dbs, args.month,
             args.store, args.merchant, args.output, args.reference)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
