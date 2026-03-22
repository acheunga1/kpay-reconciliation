# Monthly KPay × DBS × POS Reconciliation

## Objective
Generate Excel reconciliation report matching the EXACT format of the August 2025 reference template.
Reconcile 3 data sources for store KYR:
- POS daily sales by payment method
- KPay settlement transactions  
- DBS bank statement credits

Output: Single Excel file named `KPay_[Month][Year]_Reconciliation.xlsx`

## Required Inputs
1. **POS file**: `POS_Sales_YYYYMM.xlsx` in input/
2. **KPay file**: `AUg_Transaction_report.xlsx` (or similar) in input/
3. **DBS file**: `DBS_Sunstage_*.xls` in input/
4. **Month/Year**: e.g., "202508" for August 2025
5. **Store code**: "KYR"
6. **Merchant ID**: "146100005"

## Tools Required
- `tools/reconcile_kpay.py`

## Execution Steps

1. **Validate inputs exist**
```bash
   # Check all 3 files are in input/
   ls input/POS_Sales_*.xlsx
   ls input/*Transaction*.xlsx
   ls input/DBS_*.xls
```

2. **Run reconciliation tool**
```bash
   python tools/reconcile_kpay.py \
     --pos input/POS_Sales_202508.xlsx \
     --kpay input/AUg_Transaction_report.xlsx \
     --dbs input/DBS_Sunstage_002378792_HKD_202508.xls \
     --month 202508 \
     --store KYR \
     --merchant 146100005 \
     --reference reference/KPay-By_Date_Reconciliation.xlsx \
     --output output/KPay_Aug2025_Reconciliation.xlsx
```

3. **Validate output**
   - Tool automatically validates against reference template
   - All critical checks must pass

4. **Report results**
   - Confirm validation passed
   - Report any discrepancies found
   - Provide path to output file

## Template Structure (EXACT FORMAT - DO NOT DEVIATE)

### Sheet Configuration
- **Sheet name**: `KPay-By Date_Reconciliation`
- **Dimensions**: 39 rows × 61 columns (for 31-day months)

### Column Layout

#### Section 1: Columns A-I (加總 - 支付金額 / POS Daily Sales)
- Row 3: Header "加總 - 支付金額"
- Row 4: Sub-headers
  - A: 列標籤
  - B: Mastercard
  - C: PayMe
  - D: Visa
  - E: 中國銀聯
  - F: 支付寶
  - G: 微信
  - H: 銀聯雲閃付
  - I: 總計
- Rows 5-34: Daily data (dates in column A, amounts in B-I)
- Row 35: 總計 (KPay totals)
- Row 36-37: EMPTY
- Row 38: POS data (GETPIVOTDATA formulas)
- Row 39: Difference (=Row 35 - Row 38)

#### Section 2: Columns P-X (加總 - 手續費 / KPay Fees)
- Row 3: Header "加總 - 手續費"
- Row 4: Sub-headers (P-X matching B-I structure)
- Same row structure as Section 1

#### Section 3: Columns AE-AM (加總 - Net Amount)
- Row 3: Header "加總 - Net Amount"
- Row 4: Sub-headers (AE-AM matching B-I structure)
- Net Amount = Gross - Fee

#### Section 4: Columns AU-AW (DBS Data)
- Row 3: Header continues from AE section
- Row 4: Sub-headers
  - AU: DBS Receipt (settlement dates)
  - AV: DBS Receipt2 (settlement amounts)
  - AW: Txn Charge (variance formulas)

#### Section 5: Columns BA-BI (Checking on Credit Card Charge Rate)
- Row 3: Header "Checking on Credit Card Charge Rate"
- Row 4: Sub-headers (BA-BI matching B-I structure)
- Contains IFERROR formulas: `=IFERROR(Q5/B5,"-")`

### CRITICAL RULES (VIOLATIONS CAUSE FAILURE)

#### Rule 1: Column AR MUST BE EMPTY
- Column AR (column 44) is a SPACING column
- NO data, NO formulas, NO headers
- Validation: Scan all rows in column AR, must be completely empty

#### Rule 2: Row 35 is KPay Total (NOT Row 36)
- Row 35 contains "總計" in column A
- Contains SUM formulas for KPay data
- Row 36 and 37 are EMPTY
- Row 38 is POS data
- Row 39 formulas reference Row 35: `=B35-B38`

#### Rule 3: Row 39 Differences MUST Calculate
- Formula pattern: `=B35-B38`, `=C35-C38`, etc.
- If differences exist, they MUST show in the cells
- Columns E and G especially important (these were failing before)

#### Rule 4: DBS Data Month Filtering
- DBS section (AU-AW) must contain ONLY current month data
- August report: Only 2025-08-01 to 2025-08-31
- NO July carry-over
- NO September entries
- Strict date range validation required

#### Rule 5: Fee Ratio Precision
- All fee ratio formulas in BA-BI section
- Pattern: `=IFERROR(ROUND(Qn/Bn,2),"-")`
- MUST use ROUND to 2 decimals to prevent floating-point errors
- Show "-" when division by zero

#### Rule 6: Payment Method Mapping
```
POS Name           → KPay Name        → Column
─────────────────────────────────────────────
Mastercard         → Mastercard       → B, Q, AF, BB
PayMe              → PayMe            → C, R, AG, BC
Visa               → Visa             → D, S, AH, BD
銀聯/BOC Pay/雲閃付 → 中國銀聯         → E, T, AI, BE
支付寶              → 支付寶           → F, U, AJ, BF
微信支付            → 微信             → G, V, AK, BG
銀聯雲閃付          → 銀聯雲閃付       → H, W, AL, BH
```

#### Rule 7: Row 38 POS Formulas
- Row 38 uses GETPIVOTDATA to pull from external POS pivot table
- Example: `=GETPIVOTDATA("加總 - Master",'[1]POS-By Date'!$A$3)`
- Maps to POS file pivot table structure

## Data Processing Rules

### POS File Processing
```
File: POS_Sales_YYYYMM.xlsx
Sheet: Try 'Sheet (2)' first, fallback to 'Sheet'

Column mappings (0-indexed):
  0: Store code
  1: Date
  4: Master (Mastercard)
  5: Visa
  7: Payme (PayMe)
  8: Alipay (支付寶)
  9: WeChat (微信支付)
  12: UnionPay (銀聯)

Filter: Only rows where column 0 contains "KYR"
Output: Daily totals by payment method
```

### KPay File Processing
```
File: AUg_Transaction_report.xlsx (or similar)
Sheet: First sheet

Column mappings:
  6 (G): 交易狀態 (Status) - filter for "處理成功" only
  8 (I): 支付金額 (Amount)
  13 (N): 手續費 (Fee)
  18 (S): 支付方式 (Payment Method)
  24 (Y): 清算日期 (Settlement Date)

Filter:
  - Status = "處理成功"
  - Settlement date within target month
  
Output: 
  - Gross amount by payment method by settlement date
  - Fee by payment method by settlement date
```

### DBS File Processing
```
File: DBS_Sunstage_*.xls
Format: May be .xls (requires conversion to .xlsx first)

Processing:
  1. If .xls: Convert using LibreOffice or xlrd
  2. Read active sheet
  3. Start from row 7 (skip headers)
  4. Filter rows where:
     - Column 2 contains "KPAY" (uppercase)
     - Column 2 contains merchant ID "146100005"
  5. Extract:
     - Date: Column 0
     - Batch: Column 2
     - Credit: Column 5
  6. Filter by MONTH: STRICT date range YYYY-MM-01 to YYYY-MM-last_day
  
Output:
  - DBS credit by date
  - Must filter to EXACT month only
```

## Expected Output

### File Characteristics
- **Filename**: `KPay_Aug2025_Reconciliation.xlsx`
- **Sheet name**: `KPay-By Date_Reconciliation`
- **Structure**: Matches reference template exactly

### Validation Checks (All Must Pass)
✅ Sheet name correct
✅ Column headers match reference (rows 3-4)
✅ Row 35 is KPay total with "總計" label
✅ Row 36-37 are empty
✅ Row 38 has POS formulas
✅ Row 39 has difference formulas (=Row35-Row38)
✅ Column AR completely empty
✅ DBS data filtered to exact month
✅ Fee ratios use ROUND(x,2)
✅ All formulas reference correct rows/columns

## Edge Cases & Troubleshooting

### Issue: Column AR contains data
**Symptom**: Validation fails, column AR has values
**Root Cause**: Tool populating spacing column
**Fix**: Review tool code, ensure AR (column 44) skipped entirely
**Prevention**: Add explicit check to never write to column AR

### Issue: Row 39 shows #VALUE! or empty
**Symptom**: Difference row not calculating
**Root Cause**: 
  - Wrong row reference (using Row 36 instead of Row 35)
  - Formula not applied to all columns
**Fix**: Ensure formulas are `=B35-B38`, `=C35-C38`, etc. for ALL columns B-I
**Prevention**: Validate all Row 39 formulas before output

### Issue: DBS data includes wrong months
**Symptom**: September or July data appears in August report
**Root Cause**: Date filter not strict enough
**Fix**: Use exact range filter:
```python
  start_date = datetime(year, month, 1).date()
  end_date = (datetime(year, month+1, 1) - timedelta(days=1)).date()
  if not (start_date <= dbs_date <= end_date):
      continue  # Skip
```
**Prevention**: Log filtered dates, verify no out-of-month entries

### Issue: Fee ratios show .9999 or long decimals
**Symptom**: Ratio columns show imprecise values
**Root Cause**: Floating-point precision without rounding
**Fix**: Apply ROUND in formula: `=IFERROR(ROUND(Q5/B5,2),"-")`
**Prevention**: Built into tool from start

### Issue: Payment method totals don't match
**Symptom**: UnionPay vs WeChat classification differs between POS and KPay
**Root Cause**: POS terminal misclassification (known issue per user memory)
**Fix**: Document the variance, verify total still matches
**Prevention**: This is a known acceptable discrepancy - total reconciles even if individual methods differ

## Success Criteria

All checks must pass:
- ✅ All 3 data sources loaded correctly
- ✅ Payment methods reconcile at total level (variance = 0 or explainable)
- ✅ Column AR completely empty
- ✅ Row 35 labeled "總計" with KPay totals
- ✅ Row 36-37 empty
- ✅ Row 38 has POS formulas
- ✅ Row 39 calculations present in ALL columns
- ✅ DBS data filtered to exact month
- ✅ Fee ratios rounded to 2 decimals
- ✅ Output structure matches reference exactly
- ✅ File can be opened and formulas work

## Lessons Learned
(Document issues encountered and solutions - UPDATE THIS SECTION)

### 2025-08 (Initial Build):
- ✅ Confirmed Row 35 is total (not Row 36 as initially thought)
- ✅ Column AR must be completely empty (spacing column)
- ✅ DBS month filtering critical - no cross-month data
- ✅ Row 39 formula pattern established: =B35-B38 for all columns

### Future Updates:
[Add new learnings as issues are discovered and resolved]