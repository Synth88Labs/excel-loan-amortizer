# Excel Loan & Amortization Schedule Generator 🏦

[![CI](https://github.com/Synth88Labs/excel-loan-amortizer/actions/workflows/ci.yml/badge.svg)](https://github.com/Synth88Labs/excel-loan-amortizer/actions/workflows/ci.yml)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

Turn any loan — **mortgage, auto, personal, or EMI** — into a complete,
month-by-month **amortization schedule** in a clean Excel file. See exactly how
much of every payment goes to principal vs. interest, watch the balance fall to
zero, and find out how much an **extra monthly payment** saves you.

Built for the question behind every big purchase: *"what will this loan really
cost me — and how do I pay less interest?"*

## What you get

- 💳 **The monthly payment** (the standard amortizing / EMI formula)
- 📉 **Full schedule** — principal, interest, and running balance for every payment
- 💰 **Total interest & total paid** over the life of the loan
- ⚡ **Extra-payment savings** — pass `--extra` and see months and interest saved
- 🗓️ **Dated rows** — pass `--start-date` to label each payment with its month
- 📊 **Two-sheet Excel** — a Summary sheet and the full Schedule, formatted

Zero third-party math — the calculations are pure Python standard library
(openpyxl is only used to write the `.xlsx`).

## Installation

```bash
git clone https://github.com/Synth88Labs/excel-loan-amortizer.git
cd excel-loan-amortizer
pip install -r requirements.txt
```

Requires Python 3.9+.

## Usage

```bash
python loan.py --principal 300000 --rate 6.5 --years 30 -o schedule.xlsx
```

Output:

```
Loan Amortization Schedule
  Principal: 300,000.00   Rate: 6.5%/yr   Term: 360 months
  Monthly payment: 1,896.20
  Total interest:  382,633.46
  Total paid:      682,632.00
  Wrote schedule.xlsx  (360 rows)
```

### See what an extra payment saves

```bash
python loan.py --principal 25000 --rate 9 --months 60 --extra 100
# With 100.00/mo extra: paid off 11 months early, saving 1,243.93 in interest
```

### Options

| Option | Description |
|---|---|
| `--principal` | **Required.** Loan amount |
| `--rate` | **Required.** Annual interest rate in percent (e.g. `6.5`) |
| `--years` **or** `--months` | **Required.** Loan term (one of the two) |
| `--extra` | Extra principal paid each month (shortens the term) |
| `--start-date` | `YYYY-MM` to label each row with its calendar month |
| `-o`, `--output` | Output `.xlsx` path (default: `amortization.xlsx`) |

## The math

```
monthly rate  r = annual_rate% / 100 / 12
payment       M = P · r / (1 − (1 + r)^−n)        (straight-line P/n when r = 0)
interest[t]   = balance[t−1] · r
principal[t]  = M − interest[t]
balance[t]    = balance[t−1] − principal[t]
```

## Test results

See [TEST_RESULTS.md](TEST_RESULTS.md), or run them yourself:

```bash
pip install pytest
python -m pytest
```

## 📚 Learn More — Free Excel Tutorials

Practical Excel, finance & planning guides at
**[ExcelGuru.io](https://excelguru.io/category/tutorials/)**.

## License

MIT — see [LICENSE](LICENSE).
