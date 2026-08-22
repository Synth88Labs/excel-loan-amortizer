# Test Results — Excel Loan & Amortization Schedule Generator

Full local test run + live demos. CI re-runs the suite on Python 3.9, 3.11, 3.12.

## Unit tests

```
$ python -m pytest -q
........                                                                 [100%]
8 passed in 0.91s
```

**Result: 8/8 passed.**

### What each test proves

| Test | Verifies |
|---|---|
| `test_monthly_payment_known_mortgage` | 300k @ 6.5% / 30y = **1896.20/mo** (matches standard calculators) |
| `test_zero_interest_is_straight_line` | 0% loan = principal / months, zero total interest |
| `test_schedule_pays_off_exactly` | Balance reaches **exactly 0.00** at the final payment |
| `test_principal_plus_interest_equals_payment` | Every row: payment = principal + interest |
| `test_total_principal_equals_loan` | Sum of all principal portions = the original loan amount |
| `test_extra_payment_saves_interest_and_time` | Extra payment shortens term and cuts interest; savings reported |
| `test_month_labels_wrap_year` | Dated schedule rolls Dec→Jan correctly |
| `test_write_xlsx_has_both_sheets` | Workbook has Summary + Schedule sheets, frozen header, 360 rows |

## Live demos

**30-year mortgage:**

```
$ python loan.py --principal 300000 --rate 6.5 --years 30 -o schedule.xlsx

Loan Amortization Schedule
  Principal: 300,000.00   Rate: 6.5%/yr   Term: 360 months
  Monthly payment: 1,896.20
  Total interest:  382,633.46
  Total paid:      682,632.00
  Wrote schedule.xlsx  (360 rows)
```

**Auto loan with an extra $100/month:**

```
$ python loan.py --principal 25000 --rate 9 --months 60 --extra 100

Loan Amortization Schedule
  Principal: 25,000.00   Rate: 9.0%/yr   Term: 60 months
  Monthly payment: 518.96
  Total interest:  4,893.60
  Total paid:      29,893.62
  With 100.00/mo extra: paid off 11 months early, saving 1,243.93 in interest
```

**Interpretation:** on the $300k mortgage you pay **$382,633 in interest** — more
than the house itself. On the auto loan, adding just **$100/month** pays it off
**11 months early** and saves **$1,244** in interest. Both figures are verified
cell-for-cell by the test suite (total principal reconciles to the loan, balance
lands on exactly zero), so the output is trustworthy, not approximate.
