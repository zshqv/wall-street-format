"""
Generate a sample Wall Street-formatted Excel financial model for Titan Industries.
Sheets: Income Statement, Balance Sheet, Cash Flow Statement, DCF Valuation, LBO Model.
All data is simulated. Formatting follows institutional IB/PE standards.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
FONT_NAME = "Arial"
FONT_SIZE = 10

BLUE = "0000FF"
BLACK = "000000"
GREEN = "008000"
MAROON = "800000"
WHITE = "FFFFFF"
NAVY = "002060"
GRAY = "D3D3D3"

HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
INPUT_FILL = PatternFill(start_color="EBF1FF", end_color="EBF1FF", fill_type="solid")

THIN_GRAY = Side(style="thin", color=GRAY)
THIN_DARK = Side(style="thin", color=BLACK)
DOUBLE_BOT = Side(style="double", color=BLACK)
NO_SIDE = Side(style=None)

GRID = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
TOTAL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_DARK, bottom=DOUBLE_BOT)
SUBTOTAL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_DARK, bottom=THIN_DARK)
HEADER_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=NO_SIDE, bottom=THIN_DARK)

FMT_USD = '$#,##0_);($#,##0)'
FMT_USD1 = '$#,##0.0_);($#,##0.0)'
FMT_PCT = '0.0%'
FMT_MULT = '0.0"x"'
FMT_DASH = '#,##0_);(#,##0);"-"'
FMT_NUM = '#,##0'

YEARS = ["FY 2022A", "FY 2023A", "FY 2024E", "FY 2025E", "FY 2026E"]


def _font(color=BLACK, bold=False):
    return Font(name=FONT_NAME, size=FONT_SIZE, color=color, bold=bold)


def _write_header(ws, row, labels):
    for col, val in enumerate(labels, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.font = _font(WHITE, bold=True)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = HEADER_BORDER


def _write_row(ws, row, values, fmt=FMT_DASH, color=BLACK, bold=False, is_total=False,
               is_subtotal=False, is_input=False, indent=0):
    for col, val in enumerate(values, 1):
        c = ws.cell(row=row, column=col, value=val)
        if col == 1:
            c.font = _font(BLACK, bold=bold)
            c.alignment = Alignment(horizontal="left", vertical="center", indent=indent)
        else:
            cell_color = BLUE if is_input else color
            c.font = _font(cell_color, bold=bold)
            c.alignment = Alignment(horizontal="right", vertical="center")
            if isinstance(val, (int, float)):
                c.number_format = fmt
            if is_input:
                c.fill = INPUT_FILL

        if is_total:
            c.border = TOTAL_BORDER
        elif is_subtotal:
            c.border = SUBTOTAL_BORDER
        else:
            c.border = GRID


def _auto_width(ws, max_col, max_row):
    for col in range(1, max_col + 1):
        width = 12
        for row in range(1, min(max_row + 1, 50)):
            val = ws.cell(row=row, column=col).value
            if val:
                width = max(width, min(len(str(val)) + 3, 22))
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.column_dimensions["A"].width = 32
    ws.sheet_view.showGridLines = False


# ===================================================================
# INCOME STATEMENT
# ===================================================================
def build_income_statement(wb):
    ws = wb.create_sheet("Income Statement")
    _write_header(ws, 1, ["Titan Industries - Income Statement ($M)"] + YEARS)
    r = 2

    data = [
        ("Revenue",                [1850, 2120, 2438, 2804, 3164], FMT_USD, False, True, False),
        ("  Cost of Goods Sold",   [925, 1039, 1170, 1346, 1519], FMT_USD, False, True, True, 1),
        ("Gross Profit",           [925, 1081, 1268, 1458, 1645], FMT_USD, True, False, False),
        ("  Gross Margin",         [0.500, 0.510, 0.520, 0.520, 0.520], FMT_PCT, False, False, False, 1),
        ("",                       [None]*5, FMT_DASH, False, False, False),
        ("  Selling, General & Administrative", [370, 403, 439, 477, 506], FMT_USD, False, True, True, 1),
        ("  Research & Development",[148, 170, 195, 224, 253], FMT_USD, False, True, True, 1),
        ("  Depreciation & Amortization", [111, 127, 146, 168, 190], FMT_USD, False, False, False, 1),
        ("Total Operating Expenses",[629, 700, 780, 869, 949], FMT_USD, True, False, False),
        ("",                       [None]*5, FMT_DASH, False, False, False),
        ("EBITDA",                 [407, 508, 634, 757, 886], FMT_USD, True, False, False),
        ("  EBITDA Margin",        [0.220, 0.240, 0.260, 0.270, 0.280], FMT_PCT, False, False, False, 1),
        ("EBIT",                   [296, 381, 488, 589, 696], FMT_USD, True, False, False),
        ("",                       [None]*5, FMT_DASH, False, False, False),
        ("  Interest Expense",     [74, 68, 61, 53, 44], FMT_USD, False, True, True, 1),
        ("  Other Income / (Expense)", [5, 8, 10, 12, 14], FMT_USD, False, True, True, 1),
        ("Earnings Before Tax",    [227, 321, 437, 548, 666], FMT_USD, True, False, False),
        ("  Income Tax (25%)",     [57, 80, 109, 137, 167], FMT_USD, False, False, False, 1),
        ("Net Income",             [170, 241, 328, 411, 500], FMT_USD, True, False, False),
        ("  Net Margin",           [0.092, 0.114, 0.134, 0.147, 0.158], FMT_PCT, False, False, False, 1),
    ]

    for item in data:
        label, vals, fmt, is_sub, *rest = item
        is_input = rest[0] if len(rest) > 0 else False
        _ = rest[1] if len(rest) > 1 else False
        indent = rest[2] if len(rest) > 2 else 0
        is_total = label in ("Net Income",)
        _write_row(ws, r, [label] + vals, fmt=fmt, bold=is_sub or is_total,
                   is_subtotal=is_sub and not is_total, is_total=is_total,
                   is_input=is_input, indent=indent)
        r += 1

    _auto_width(ws, 6, r)


# ===================================================================
# BALANCE SHEET
# ===================================================================
def build_balance_sheet(wb):
    ws = wb.create_sheet("Balance Sheet")
    _write_header(ws, 1, ["Titan Industries - Balance Sheet ($M)"] + YEARS)
    r = 2

    sections = [
        ("ASSETS", None),
        ("  Cash & Equivalents",    [320, 415, 560, 740, 980], True, 1),
        ("  Accounts Receivable",   [231, 265, 305, 350, 395], True, 1),
        ("  Inventory",             [185, 212, 244, 280, 316], True, 1),
        ("  Prepaid Expenses",      [37, 42, 49, 56, 63], True, 1),
        ("Total Current Assets",    [773, 934, 1158, 1426, 1754], False, 0),
        ("",                        [None]*5, False, 0),
        ("  Property, Plant & Equipment", [740, 810, 890, 970, 1050], True, 1),
        ("  Goodwill & Intangibles",[480, 465, 450, 435, 420], False, 1),
        ("Total Assets",            [1993, 2209, 2498, 2831, 3224], False, 0),
        ("",                        [None]*5, False, 0),
        ("LIABILITIES & EQUITY", None),
        ("  Accounts Payable",      [148, 170, 195, 224, 253], True, 1),
        ("  Accrued Liabilities",   [92, 106, 122, 140, 158], True, 1),
        ("  Current Portion of Debt",[50, 50, 50, 50, 50], True, 1),
        ("Total Current Liabilities",[290, 326, 367, 414, 461], False, 0),
        ("",                        [None]*5, False, 0),
        ("  Long-Term Debt",        [600, 530, 455, 375, 290], True, 1),
        ("  Deferred Tax Liabilities",[85, 92, 100, 108, 117], False, 1),
        ("Total Liabilities",       [975, 948, 922, 897, 868], False, 0),
        ("",                        [None]*5, False, 0),
        ("  Common Stock & APIC",   [400, 400, 400, 400, 400], False, 1),
        ("  Retained Earnings",     [618, 861, 1176, 1534, 1956], False, 1),
        ("Total Shareholders' Equity",[1018, 1261, 1576, 1934, 2356], False, 0),
        ("",                        [None]*5, False, 0),
        ("Total Liabilities & Equity",[1993, 2209, 2498, 2831, 3224], False, 0),
    ]

    for item in sections:
        if item[1] is None:
            _write_row(ws, r, [item[0]] + [None]*5, bold=True)
            r += 1
            continue
        label, vals, is_input, indent = item
        is_total = label.startswith("Total")
        is_sub = is_total and label != "Total Assets" and label != "Total Liabilities & Equity"
        _write_row(ws, r, [label] + vals, fmt=FMT_USD, bold=is_total,
                   is_subtotal=is_sub, is_total=(label in ("Total Assets", "Total Liabilities & Equity")),
                   is_input=is_input, indent=indent)
        r += 1

    _auto_width(ws, 6, r)


# ===================================================================
# CASH FLOW STATEMENT
# ===================================================================
def build_cash_flow(wb):
    ws = wb.create_sheet("Cash Flow Statement")
    _write_header(ws, 1, ["Titan Industries - Cash Flow Statement ($M)"] + YEARS)
    r = 2

    rows = [
        ("OPERATING ACTIVITIES", None),
        ("  Net Income",                [170, 241, 328, 411, 500], False, 1),
        ("  Depreciation & Amortization",[111, 127, 146, 168, 190], False, 1),
        ("  Changes in Working Capital", [-25, -18, -22, -28, -30], False, 1),
        ("  Stock-Based Compensation",   [15, 18, 22, 26, 30], False, 1),
        ("Cash from Operations",         [271, 368, 474, 577, 690], False, 0),
        ("",                             [None]*5, False, 0),
        ("INVESTING ACTIVITIES", None),
        ("  Capital Expenditures",       [-148, -170, -195, -224, -253], True, 1),
        ("  Acquisitions",               [0, -50, 0, 0, 0], True, 1),
        ("Cash from Investing",          [-148, -220, -195, -224, -253], False, 0),
        ("",                             [None]*5, False, 0),
        ("FINANCING ACTIVITIES", None),
        ("  Debt Repayment",             [-50, -70, -75, -80, -85], True, 1),
        ("  Dividends Paid",             [-30, -33, -37, -41, -46], True, 1),
        ("  Share Repurchases",          [0, -20, -22, -32, -66], True, 1),
        ("Cash from Financing",          [-80, -123, -134, -153, -197], False, 0),
        ("",                             [None]*5, False, 0),
        ("Net Change in Cash",           [43, 25, 145, 200, 240], False, 0),
        ("Beginning Cash Balance",       [277, 320, 415, 560, 740], False, 0),
        ("Ending Cash Balance",          [320, 415, 560, 740, 980], False, 0),
        ("",                             [None]*5, False, 0),
        ("Free Cash Flow",               [123, 198, 279, 353, 437], False, 0),
        ("  FCF Margin",                 [0.066, 0.093, 0.114, 0.126, 0.138], False, 0),
    ]

    for item in rows:
        if item[1] is None:
            _write_row(ws, r, [item[0]] + [None]*5, bold=True)
            r += 1
            continue
        label, vals, is_input, indent = item
        is_total = label in ("Ending Cash Balance", "Free Cash Flow")
        is_sub = label.startswith("Cash from") or label == "Net Change in Cash"
        fmt = FMT_PCT if "Margin" in label else FMT_USD
        _write_row(ws, r, [label] + vals, fmt=fmt, bold=is_sub or is_total,
                   is_subtotal=is_sub, is_total=is_total,
                   is_input=is_input, indent=indent)
        r += 1

    _auto_width(ws, 6, r)


# ===================================================================
# DCF VALUATION
# ===================================================================
def build_dcf(wb):
    ws = wb.create_sheet("DCF Valuation")
    _write_header(ws, 1, ["Titan Industries - DCF Valuation ($M)"] + YEARS)
    r = 2

    rows = [
        ("KEY ASSUMPTIONS", None),
        ("  Revenue Growth",           [None, 0.146, 0.150, 0.150, 0.128], FMT_PCT, True, 1),
        ("  EBITDA Margin",            [0.220, 0.240, 0.260, 0.270, 0.280], FMT_PCT, True, 1),
        ("  Tax Rate",                 [0.250, 0.250, 0.250, 0.250, 0.250], FMT_PCT, True, 1),
        ("  D&A % of Revenue",         [0.060, 0.060, 0.060, 0.060, 0.060], FMT_PCT, True, 1),
        ("  CapEx % of Revenue",       [0.080, 0.080, 0.080, 0.080, 0.080], FMT_PCT, True, 1),
        ("  Change in NWC % of Rev",   [0.014, 0.008, 0.009, 0.010, 0.009], FMT_PCT, True, 1),
        ("", [None]*5, FMT_DASH, False, 0),
        ("UNLEVERED FREE CASH FLOW", None),
        ("  EBITDA",                   [407, 508, 634, 757, 886], FMT_USD, False, 1),
        ("  Less: D&A",               [-111, -127, -146, -168, -190], FMT_USD, False, 1),
        ("  EBIT",                     [296, 381, 488, 589, 696], FMT_USD, False, 1),
        ("  Less: Taxes on EBIT",      [-74, -95, -122, -147, -174], FMT_USD, False, 1),
        ("  NOPAT",                    [222, 286, 366, 442, 522], FMT_USD, False, 1),
        ("  Plus: D&A",               [111, 127, 146, 168, 190], FMT_USD, False, 1),
        ("  Less: CapEx",             [-148, -170, -195, -224, -253], FMT_USD, False, 1),
        ("  Less: Change in NWC",      [-25, -18, -22, -28, -30], FMT_USD, False, 1),
        ("Unlevered Free Cash Flow",   [160, 225, 295, 358, 429], FMT_USD, False, 0),
        ("", [None]*5, FMT_DASH, False, 0),
        ("WACC CALCULATION", None),
        ("  Risk-Free Rate",           [0.042]*5, FMT_PCT, True, 1),
        ("  Equity Risk Premium",      [0.060]*5, FMT_PCT, True, 1),
        ("  Beta (Levered)",           [1.15]*5, FMT_MULT, True, 1),
        ("  Cost of Equity",           [0.111]*5, FMT_PCT, False, 1),
        ("  Pre-Tax Cost of Debt",     [0.055]*5, FMT_PCT, True, 1),
        ("  Debt / Total Cap",         [0.300]*5, FMT_PCT, True, 1),
        ("  WACC",                     [0.090]*5, FMT_PCT, False, 1),
        ("", [None]*5, FMT_DASH, False, 0),
        ("TERMINAL VALUE & VALUATION", None),
        ("  Terminal Growth Rate",     [None]*4 + [0.025], FMT_PCT, True, 1),
        ("  Terminal Value",           [None]*4 + [6764], FMT_USD, False, 1),
        ("  PV of Terminal Value",     [None]*4 + [4397], FMT_USD, False, 1),
        ("  PV of FCFs (Sum)",         [None]*4 + [1086], FMT_USD, False, 1),
        ("Enterprise Value",           [None]*4 + [5483], FMT_USD, False, 0),
        ("  Less: Net Debt",           [None]*4 + [-290], FMT_USD, False, 1),
        ("  Plus: Cash",               [None]*4 + [980], FMT_USD, False, 1),
        ("Equity Value",               [None]*4 + [6173], FMT_USD, False, 0),
        ("  Shares Outstanding (M)",   [None]*4 + [250], FMT_NUM, True, 1),
        ("Implied Share Price",        [None]*4 + [24.69], '$#,##0.00', False, 0),
    ]

    for item in rows:
        if item[1] is None:
            _write_row(ws, r, [item[0]] + [None]*5, bold=True)
            r += 1
            continue
        label, vals, fmt, is_input, indent = item
        is_total = label in ("Enterprise Value", "Equity Value", "Implied Share Price")
        is_sub = label == "Unlevered Free Cash Flow" or label == "WACC"
        _write_row(ws, r, [label] + vals, fmt=fmt, bold=is_sub or is_total,
                   is_subtotal=is_sub, is_total=is_total,
                   is_input=is_input, indent=indent)
        r += 1

    _auto_width(ws, 6, r)


# ===================================================================
# LBO MODEL
# ===================================================================
def build_lbo(wb):
    ws = wb.create_sheet("LBO Model")
    _write_header(ws, 1, ["Titan Industries - LBO Model ($M)"] + YEARS)
    r = 2

    rows = [
        ("TRANSACTION ASSUMPTIONS", None),
        ("  Entry EV / EBITDA",        [None]*2 + [10.0]*3, FMT_MULT, True, 1),
        ("  Entry Enterprise Value",   [None]*2 + [6340]*3, FMT_USD, False, 1),
        ("  Sponsor Equity",           [None]*2 + [2536]*3, FMT_USD, False, 1),
        ("  Total Debt at Entry",      [None]*2 + [3804]*3, FMT_USD, True, 1),
        ("  Leverage at Entry",        [None]*2 + [6.0]*3, FMT_MULT, False, 1),
        ("", [None]*5, FMT_DASH, False, 0),
        ("OPERATING PROJECTIONS", None),
        ("  Revenue",                  [1850, 2120, 2438, 2804, 3164], FMT_USD, False, 1),
        ("  Revenue Growth",           [None, 0.146, 0.150, 0.150, 0.128], FMT_PCT, False, 1),
        ("  EBITDA",                   [407, 508, 634, 757, 886], FMT_USD, False, 1),
        ("  EBITDA Margin",            [0.220, 0.240, 0.260, 0.270, 0.280], FMT_PCT, False, 1),
        ("  CapEx",                    [-148, -170, -195, -224, -253], FMT_USD, True, 1),
        ("  Free Cash Flow",           [123, 198, 279, 353, 437], FMT_USD, False, 1),
        ("", [None]*5, FMT_DASH, False, 0),
        ("DEBT SCHEDULE", None),
        ("  Beginning Debt Balance",   [None, None, 3804, 3525, 3172], FMT_USD, False, 1),
        ("  Mandatory Amortization",   [None, None, -50, -50, -50], FMT_USD, True, 1),
        ("  Cash Sweep Repayment",     [None, None, -229, -303, -387], FMT_USD, False, 1),
        ("  Ending Debt Balance",      [None, None, 3525, 3172, 2735], FMT_USD, False, 1),
        ("  Leverage (Debt/EBITDA)",   [None, None, 5.6, 4.2, 3.1], FMT_MULT, False, 1),
        ("  Interest Expense (5.5%)",  [None, None, 202, 184, 162], FMT_USD, False, 1),
        ("", [None]*5, FMT_DASH, False, 0),
        ("EXIT & RETURNS", None),
        ("  Exit Year EBITDA",         [None]*4 + [886], FMT_USD, False, 1),
        ("  Exit EV / EBITDA",         [None]*4 + [10.0], FMT_MULT, True, 1),
        ("  Exit Enterprise Value",    [None]*4 + [8860], FMT_USD, False, 1),
        ("  Less: Net Debt at Exit",   [None]*4 + [-2735], FMT_USD, False, 1),
        ("  Exit Equity Value",        [None]*4 + [6125], FMT_USD, False, 1),
        ("", [None]*5, FMT_DASH, False, 0),
        ("  MOIC",                     [None]*4 + [2.4], FMT_MULT, False, 1),
        ("  IRR",                      [None]*4 + [0.338], FMT_PCT, False, 1),
    ]

    for item in rows:
        if item[1] is None:
            _write_row(ws, r, [item[0]] + [None]*5, bold=True)
            r += 1
            continue
        label, vals, fmt, is_input, indent = item
        is_total = label in ("MOIC", "IRR", "Exit Equity Value")
        is_sub = label in ("Ending Debt Balance", "Exit Enterprise Value")
        _write_row(ws, r, [label] + vals, fmt=fmt, bold=is_sub or is_total,
                   is_subtotal=is_sub, is_total=is_total,
                   is_input=is_input, indent=indent)
        r += 1

    _auto_width(ws, 6, r)


# ===================================================================
# MAIN
# ===================================================================
def main():
    wb = Workbook()
    wb.remove(wb.active)

    build_income_statement(wb)
    build_balance_sheet(wb)
    build_cash_flow(wb)
    build_dcf(wb)
    build_lbo(wb)

    out = Path(__file__).parent.parent / "assets" / "Titan_Industries_Financial_Model.xlsx"
    wb.save(str(out))
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
