"""
Excel Loan & Amortization Schedule Generator
===========================================

Turn a loan (principal, rate, term) into a full month-by-month amortization
schedule in a formatted Excel file — mortgage, auto, personal loan, or EMI.
Shows principal vs. interest for every payment, the running balance, total
interest paid, and how much an optional extra monthly payment saves you.

Runtime dependency: openpyxl (Excel output). The math is pure standard library.

Usage
-----
    python loan.py --principal 300000 --rate 6.5 --years 30 -o schedule.xlsx
    python loan.py --principal 25000 --rate 9 --months 60 --extra 100
    python loan.py --principal 300000 --rate 6.5 --years 30 --start-date 2026-09
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass


def monthly_payment(principal: float, annual_rate_pct: float, months: int) -> float:
    """
    Standard fixed-rate amortizing payment (the "EMI" formula).

        M = P * r / (1 - (1 + r)^-n)

    where r is the monthly rate. Handles the 0% case (straight-line).
    """
    if months <= 0:
        raise ValueError("months must be positive")
    if principal <= 0:
        raise ValueError("principal must be positive")
    r = annual_rate_pct / 100.0 / 12.0
    if r == 0:
        return principal / months
    return principal * r / (1 - (1 + r) ** (-months))


@dataclass
class ScheduleRow:
    period: int
    payment: float
    principal: float
    interest: float
    balance: float


def build_schedule(
    principal: float,
    annual_rate_pct: float,
    months: int,
    extra: float = 0.0,
) -> list[ScheduleRow]:
    """
    Build the amortization schedule. An optional `extra` amount is applied to
    principal every month, which shortens the term (the loan pays off early).
    """
    r = annual_rate_pct / 100.0 / 12.0
    base = monthly_payment(principal, annual_rate_pct, months)
    rows: list[ScheduleRow] = []
    balance = principal
    period = 0
    # cap iterations well beyond the nominal term as a safety net
    max_periods = months + 1
    while balance > 0.005 and period < max_periods:
        period += 1
        interest = balance * r
        pay = base + extra
        principal_paid = pay - interest
        # final payment: don't overpay past the remaining balance
        if principal_paid >= balance:
            principal_paid = balance
            pay = principal_paid + interest
        balance -= principal_paid
        rows.append(
            ScheduleRow(
                period=period,
                payment=round(pay, 2),
                principal=round(principal_paid, 2),
                interest=round(interest, 2),
                balance=round(max(balance, 0.0), 2),
            )
        )
    return rows


def summarize(principal: float, annual_rate_pct: float, months: int, extra: float = 0.0) -> dict:
    """Compute headline figures, including interest saved by the extra payment."""
    sched = build_schedule(principal, annual_rate_pct, months, extra)
    total_interest = round(sum(row.interest for row in sched), 2)
    total_paid = round(sum(row.payment for row in sched), 2)
    result = {
        "monthly_payment": round(monthly_payment(principal, annual_rate_pct, months), 2),
        "n_payments": len(sched),
        "nominal_term": months,
        "total_interest": total_interest,
        "total_paid": total_paid,
        "schedule": sched,
    }
    if extra > 0:
        base_only = summarize(principal, annual_rate_pct, months, extra=0.0)
        result["interest_saved"] = round(base_only["total_interest"] - total_interest, 2)
        result["months_saved"] = base_only["n_payments"] - len(sched)
    return result


# --- month labels (optional dated schedule) ----------------------------------


def month_labels(start: str | None, count: int) -> list[str]:
    """Given a 'YYYY-MM' start, return that many 'YYYY-MM' labels; else empty."""
    if not start:
        return []
    year, month = (int(x) for x in start.split("-")[:2])
    out = []
    for _ in range(count):
        out.append(f"{year:04d}-{month:02d}")
        month += 1
        if month > 12:
            month = 1
            year += 1
    return out


# --- Excel output ------------------------------------------------------------


def write_xlsx(summary: dict, out_path: str, start_date: str | None = None) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    sched: list[ScheduleRow] = summary["schedule"]
    labels = month_labels(start_date, len(sched))

    wb = Workbook()

    # --- Summary sheet ---
    ws = wb.active
    ws.title = "Summary"
    green = Font(bold=True, color="217346")
    ws["A1"] = "Loan Amortization Summary"
    ws["A1"].font = Font(bold=True, size=14, color="217346")
    rows = [
        ("Monthly payment", summary["monthly_payment"]),
        ("Number of payments", summary["n_payments"]),
        ("Total interest", summary["total_interest"]),
        ("Total paid", summary["total_paid"]),
    ]
    if "interest_saved" in summary:
        rows.append(("Interest saved (extra payment)", summary["interest_saved"]))
        rows.append(("Months saved", summary["months_saved"]))
    for i, (label, value) in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=label).font = green
        ws.cell(row=i, column=2, value=value)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 18

    # --- Schedule sheet ---
    sh = wb.create_sheet("Schedule")
    header_fill = PatternFill("solid", fgColor="217346")
    header_font = Font(bold=True, color="FFFFFF")
    headers = (["Month"] if labels else []) + ["Period", "Payment", "Principal", "Interest", "Balance"]
    for c, h in enumerate(headers, start=1):
        cell = sh.cell(row=1, column=c, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    for i, row in enumerate(sched):
        base_c = 1
        r = i + 2
        if labels:
            sh.cell(row=r, column=1, value=labels[i])
            base_c = 2
        sh.cell(row=r, column=base_c, value=row.period)
        sh.cell(row=r, column=base_c + 1, value=row.payment).number_format = "#,##0.00"
        sh.cell(row=r, column=base_c + 2, value=row.principal).number_format = "#,##0.00"
        sh.cell(row=r, column=base_c + 3, value=row.interest).number_format = "#,##0.00"
        sh.cell(row=r, column=base_c + 4, value=row.balance).number_format = "#,##0.00"
    sh.freeze_panes = "A2"
    for c in range(1, len(headers) + 1):
        sh.column_dimensions[sh.cell(row=1, column=c).column_letter].width = 13

    wb.save(out_path)


# --- CLI ---------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Generate a loan amortization schedule in Excel (mortgage, auto, EMI...)."
    )
    p.add_argument("--principal", type=float, required=True, help="Loan amount")
    p.add_argument("--rate", type=float, required=True, help="Annual interest rate, percent (e.g. 6.5)")
    term = p.add_mutually_exclusive_group(required=True)
    term.add_argument("--years", type=float, help="Loan term in years")
    term.add_argument("--months", type=int, help="Loan term in months")
    p.add_argument("--extra", type=float, default=0.0, help="Optional extra principal paid each month")
    p.add_argument("--start-date", help="Optional 'YYYY-MM' to label each row with a month")
    p.add_argument("-o", "--output", help="Output .xlsx path (default: amortization.xlsx)")
    args = p.parse_args(argv)

    months = args.months if args.months is not None else int(round(args.years * 12))
    try:
        summary = summarize(args.principal, args.rate, months, extra=args.extra)
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    print("Loan Amortization Schedule")
    print(f"  Principal: {args.principal:,.2f}   Rate: {args.rate}%/yr   Term: {months} months")
    print(f"  Monthly payment: {summary['monthly_payment']:,.2f}")
    print(f"  Total interest:  {summary['total_interest']:,.2f}")
    print(f"  Total paid:      {summary['total_paid']:,.2f}")
    if "interest_saved" in summary:
        print(
            f"  With {args.extra:,.2f}/mo extra: paid off {summary['months_saved']} months early, "
            f"saving {summary['interest_saved']:,.2f} in interest"
        )

    out = args.output or "amortization.xlsx"
    write_xlsx(summary, out, start_date=args.start_date)
    print(f"  Wrote {out}  ({summary['n_payments']} rows)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
