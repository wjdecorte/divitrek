#!/usr/bin/env python3
import datetime as dt
from pathlib import Path
import xlsxwriter

today = dt.date.today()
out_path = Path("/Users/jason/Documents/Dividend_Tracker_Skeleton.xlsx")
wb = xlsxwriter.Workbook(out_path.as_posix())

# Formats
fmt_header = wb.add_format({"bold": True, "bg_color": "#EFEFEF", "border": 1})
fmt_date = wb.add_format({"num_format": "yyyy-mm-dd"})
fmt_money = wb.add_format({"num_format": "$#,##0.00"})
fmt_pct = wb.add_format({"num_format": "0.00%"})
fmt_text = wb.add_format({"text_wrap": True, "valign": "top"})
fmt_num = wb.add_format({"num_format": "#,##0.####"})

# ---------------- 1) Transactions (Tx) ----------------
ws_tx = wb.add_worksheet("Transactions")
tx_headers = [
    "Run_Date",
    "Action",
    "Symbol",
    "Description",
    "Quantity",
    "Price",
    "Amount",
    "Settle_Date",
    "Action_Type",
    "IsDivIncome",
    "IsReinvestment",
    "BuyQty",
    "SellQty",
    "BuyCost",
    "CashDividend",
    "DRIPDividend",
    "NetContribution",
    "IsETF",
]
ws_tx.write_row(0, 0, tx_headers, fmt_header)
# one blank row so table formulas persist when you append new rows
for c in range(len(tx_headers)):
    ws_tx.write_blank(1, c, None)

ws_tx.add_table(
    0,
    0,
    1,
    len(tx_headers) - 1,
    {
        "name": "Tx",
        "style": "Table Style Medium 2",
        "columns": [
            {"header": "Run_Date"},
            {"header": "Action"},
            {"header": "Symbol"},
            {"header": "Description"},
            {"header": "Quantity"},
            {"header": "Price"},
            {"header": "Amount"},
            {"header": "Settle_Date"},
            {
                "header": "Action_Type",
                "formula": '=LET(a,[@Action],IF(ISNUMBER(SEARCH("DIVIDEND RECEIVED",a)),"DIVIDEND RECEIVED",IF(ISNUMBER(SEARCH("REINVESTMENT",a)),"REINVESTMENT",IF(OR(ISNUMBER(SEARCH("BOUGHT",a)),ISNUMBER(SEARCH("BUY",a))),"BOUGHT",IF(OR(ISNUMBER(SEARCH("SOLD",a)),ISNUMBER(SEARCH("SELL",a))),"SOLD","OTHER")))))',
            },
            {
                "header": "IsDivIncome",
                "formula": '=--([@Action_Type]="DIVIDEND RECEIVED")',
            },
            {
                "header": "IsReinvestment",
                "formula": '=--([@Action_Type]="REINVESTMENT")',
            },
            {
                "header": "BuyQty",
                "formula": '=IF(OR([@Action_Type]="BOUGHT",[@Action_Type]="REINVESTMENT"),[@Quantity],0)',
            },
            {
                "header": "SellQty",
                "formula": '=IF([@Action_Type]="SOLD",[@Quantity],0)',
            },
            {
                "header": "BuyCost",
                "formula": '=IF(OR([@Action_Type]="BOUGHT",[@Action_Type]="REINVESTMENT"),ABS([@Amount]),0)',
            },
            {
                "header": "CashDividend",
                "formula": "=IF([@IsDivIncome]=1,ABS([@Amount]),0)",
            },
            {
                "header": "DRIPDividend",
                "formula": "=IF([@IsReinvestment]=1,ABS([@Amount]),0)",
            },
            {
                "header": "NetContribution",
                "formula": '=IF([@Action_Type]="BOUGHT",-ABS([@Amount]),IF([@Action_Type]="SOLD",ABS([@Amount]),0))',
            },
            {"header": "IsETF", "formula": '=--ISNUMBER(SEARCH("ETF",[@Description]))'},
        ],
    },
)
ws_tx.set_column(0, 0, 13, fmt_date)
ws_tx.set_column(1, 1, 50)
ws_tx.set_column(2, 2, 10)
ws_tx.set_column(3, 3, 40)
ws_tx.set_column(4, 6, 12, fmt_num)
ws_tx.set_column(7, 7, 13, fmt_date)
ws_tx.set_column(8, 8, 16)
ws_tx.set_column(9, 16, 14)
ws_tx.set_column(17, 17, 8)

# ---------------- 2) Prices ----------------
ws_px = wb.add_worksheet("Prices")
ws_px.write("A1", "AsOfDate", fmt_header)
ws_px.write_datetime("B1", dt.datetime(today.year, today.month, today.day), fmt_date)
wb.define_name("AsOfDate", "=Prices!$B$1")

px_headers = [
    "Symbol",
    "Date_Pulled",
    "Price",
    "Date_5D",
    "Price_5D",
    "Chg_5D",
    "ChgPct_5D",
    "Date_1M",
    "Price_1M",
    "Chg_1M",
    "ChgPct_1M",
    "Date_6M",
    "Price_6M",
    "Chg_6M",
    "ChgPct_6M",
    "Date_1Y",
    "Price_1Y",
    "Chg_1Y",
    "ChgPct_1Y",
]
ws_px.write_row(1, 0, px_headers, fmt_header)
ws_px.write_formula(2, 0, '=SORT(UNIQUE(FILTER(Tx[Symbol],Tx[Symbol]<>"")))')
ws_px.write_formula(2, 1, '=IF(A3="","",AsOfDate)')
# change formulas (row level; copy down with Fill Handle if you want, or let updater write values)
ws_px.write_formula(2, 5, '=IF(C3="","",C3-E3)')
ws_px.write_formula(2, 6, '=IF(E3>0,(F3/E3),"")')
ws_px.write_formula(2, 9, '=IF(C3="","",C3-I3)')
ws_px.write_formula(2, 10, '=IF(I3>0,(J3/I3),"")')
ws_px.write_formula(2, 13, '=IF(C3="","",C3-M3)')
ws_px.write_formula(2, 14, '=IF(M3>0,(N3/M3),"")')
ws_px.write_formula(2, 17, '=IF(C3="","",C3-Q3)')
ws_px.write_formula(2, 18, '=IF(Q3>0,(R3/Q3),"")')
ws_px.set_column(0, 0, 10)
ws_px.set_column(1, 1, 12, fmt_date)
ws_px.set_column(2, 2, 12, fmt_money)
for col in [3, 7, 11, 15]:
    ws_px.set_column(col, col, 12, fmt_date)
for col in [4, 8, 12, 16, 5, 9, 13, 17]:
    ws_px.set_column(col, col, 12, fmt_money)
for col in [6, 10, 14, 18]:
    ws_px.set_column(col, col, 10, fmt_pct)
wb.define_name("PxSymbol", "=Prices!$A$3#")
wb.define_name("PxPrice", "=Prices!$C$3#")

# ---------------- 3) ETF ----------------
ws_etf = wb.add_worksheet("ETF")
etf_headers = ["Symbol", "Type", "NAV", "NAV_Date", "AUM_USD", "AUM_Date", "Source"]
ws_etf.write_row(0, 0, etf_headers, fmt_header)
ws_etf.add_table(
    0,
    0,
    1,
    len(etf_headers) - 1,
    {
        "name": "Etf",
        "style": "Table Style Medium 2",
        "columns": [{"header": h} for h in etf_headers],
    },
)
ws_etf.set_column(0, 0, 10)
ws_etf.set_column(1, 1, 10)
ws_etf.set_column(2, 2, 12, fmt_money)
ws_etf.set_column(3, 3, 12, fmt_date)
ws_etf.set_column(4, 4, 14, fmt_money)
ws_etf.set_column(5, 5, 12, fmt_date)
ws_etf.set_column(6, 6, 30)

# ---------------- 4) DivCal ----------------
ws_dc = wb.add_worksheet("DivCal")
dc_headers = [
    "Symbol",
    "Frequency",
    "Months_Between",
    "Last_ExDiv",
    "Last_Pay",
    "Declared_Next_ExDiv",
    "Declared_Next_Pay",
    "Declared_Amount",
    "Inferred_Next_Pay",
]
ws_dc.write_row(0, 0, dc_headers, fmt_header)
ws_dc.add_table(
    0,
    0,
    1,
    len(dc_headers) - 1,
    {
        "name": "DivCal",
        "style": "Table Style Medium 2",
        "columns": [{"header": h} for h in dc_headers],
    },
)
ws_dc.set_column(0, 0, 10)
ws_dc.set_column(1, 1, 12)
ws_dc.set_column(2, 2, 16)
ws_dc.set_column(3, 8, 14, fmt_date)

# ---------------- 5) Holdings & Summary ----------------
ws_sum = wb.add_worksheet("Holdings & Summary")
sum_headers = [
    "Symbol",
    "Type",
    "Shares",
    "Cost Basis",
    "Average Cost / Share",
    "Last Dividend Date",
    "Next Dividend Date",
    "TTM Dividends",
    "TTM Yield on Cost",
    "Calculated Yield",
]
ws_sum.write_row(1, 0, sum_headers, fmt_header)

# A3: symbol spill
ws_sum.write_formula(2, 0, '=SORT(UNIQUE(FILTER(Tx[Symbol],Tx[Symbol]<>"")))')

# B..J: dynamic array formulas that spill to match A3#
ws_sum.write_formula(
    2,
    1,
    '=BYROW($A$3#,LAMBDA(r,LET(sym,INDEX(r,1),IFERROR(XLOOKUP(sym,Etf[Symbol],Etf[Type],IF(MAX(IF(Tx[Symbol]=sym,Tx[IsETF]))=1,"ETF","STOCK")),""))))',
)
ws_sum.write_formula(
    2,
    2,
    "=BYROW($A$3#,LAMBDA(r,LET(sym,INDEX(r,1),SUMIFS(Tx[BuyQty],Tx[Symbol],sym)-SUMIFS(Tx[SellQty],Tx[Symbol],sym))))",
)
ws_sum.write_formula(
    2,
    4,
    '=BYROW($A$3#,LAMBDA(r,LET(sym,INDEX(r,1),bsh,SUMIFS(Tx[BuyQty],Tx[Symbol],sym),bcost,SUMIFS(Tx[BuyCost],Tx[Symbol],sym),IF(bsh>0,bcost/bsh,""))))',
)
ws_sum.write_formula(2, 3, "=$C$3#*$E$3#")
ws_sum.write_formula(
    2,
    5,
    '=BYROW($A$3#,LAMBDA(r,LET(sym,INDEX(r,1),IFERROR(MAX(FILTER(Tx[Run_Date],(Tx[Symbol]=sym)*(Tx[IsDivIncome]=1))),""))))',
)
ws_sum.write_formula(
    2,
    6,
    '=BYROW($A$3#,LAMBDA(r,LET(sym,INDEX(r,1),IFERROR(XLOOKUP(sym,DivCal[Symbol],DivCal[Declared_Next_Pay],EDATE(XLOOKUP(sym,DivCal[Symbol],DivCal[Last_Pay],""),XLOOKUP(sym,DivCal[Symbol],DivCal[Months_Between],1))),""))))',
)
ws_sum.write_formula(
    2,
    7,
    '=BYROW($A$3#,LAMBDA(r,LET(sym,INDEX(r,1),SUMIFS(Tx[CashDividend],Tx[Symbol],sym,Tx[Run_Date],">="&EDATE(AsOfDate,-12)+1,Tx[Run_Date],"<="&AsOfDate))))',
)
ws_sum.write_formula(2, 8, "=$H$3#/$D$3#")
ws_sum.write_formula(
    2,
    9,
    '=BYROW($A$3#,LAMBDA(r,LET(sym,INDEX(r,1),sh,SUMIFS(Tx[BuyQty],Tx[Symbol],sym)-SUMIFS(Tx[SellQty],Tx[Symbol],sym),ttm,SUMIFS(Tx[CashDividend],Tx[Symbol],sym,Tx[Run_Date],">="&EDATE(AsOfDate,-12)+1,Tx[Run_Date],"<="&AsOfDate),px,IFERROR(XLOOKUP(sym,PxSymbol,PxPrice),NA()),IFERROR(ttm/(sh*px),""))))',
)
ws_sum.set_column(0, 0, 10)
ws_sum.set_column(1, 1, 10)
ws_sum.set_column(2, 2, 10)
ws_sum.set_column(3, 3, 14, fmt_money)
ws_sum.set_column(4, 4, 14, fmt_money)
ws_sum.set_column(5, 6, 14, fmt_date)
ws_sum.set_column(7, 7, 14, fmt_money)
ws_sum.set_column(8, 9, 14, fmt_pct)

# ---------------- 6) Monthly Dividend History ----------------
ws_m = wb.add_worksheet("Monthly Dividend History")
ws_m.write(0, 0, "", fmt_header)
for j in range(12):
    ws_m.write_formula(
        0, 1 + j, f'=UPPER(TEXT(EDATE(AsOfDate, {j-11}),"MMM YY"))', fmt_header
    )
ws_m.write(1, 0, "Symbol", fmt_header)
ws_m.write_formula(2, 0, "='Holdings & Summary'!$A$3#")
for j in range(12):
    ws_m.write_formula(
        2,
        1 + j,
        f'=BYROW($A$3#,LAMBDA(r,LET(sym,INDEX(r,1),mStart,EOMONTH(EDATE(AsOfDate,{j-11}),-1)+1,mEnd,EOMONTH(EDATE(AsOfDate,{j-11}),0),SUMIFS(Tx[CashDividend],Tx[Symbol],sym,Tx[Run_Date],">="&mStart,Tx[Run_Date],"<="&mEnd))))',
    )
ws_m.set_column(0, 0, 10)
ws_m.set_column(1, 12, 12, fmt_money)

# ---------------- 7) Dividend Forecast ----------------
ws_f = wb.add_worksheet("Dividend Forecast")
ws_f.write_row(0, 0, [""] + [None] * 12, fmt_header)
for j in range(12):
    ws_f.write_formula(
        0, 1 + j, f'=UPPER(TEXT(EDATE(AsOfDate, {j-11}),"MMM YY"))', fmt_header
    )
ws_f.write(1, 0, "Symbol", fmt_header)
ws_f.write_formula(2, 0, "='Holdings & Summary'!$A$3#")
for j in range(12):
    ws_f.write_formula(
        2,
        1 + j,
        f'=BYROW($A$3#,LAMBDA(r,LET(sym,INDEX(r,1),mStart,EOMONTH(EDATE(AsOfDate,{j-11}),-1)+1,mEnd,EOMONTH(EDATE(AsOfDate,{j-11}),0),dPay,XLOOKUP(sym,DivCal[Symbol],DivCal[Declared_Next_Pay],""),dAmt,XLOOKUP(sym,DivCal[Symbol],DivCal[Declared_Amount],""),dVal,IF(AND(dAmt<>"", dPay<>"" , dPay>=mStart, dPay<=mEnd), dAmt, 0),iPay,XLOOKUP(sym,DivCal[Symbol],DivCal[Inferred_Next_Pay],""),freq,XLOOKUP(sym,DivCal[Symbol],DivCal[Months_Between],1),iVal,IF(AND(dVal=0, iPay<>"", iPay>=mStart, iPay<=mEnd), MEDIAN(TAKE(FILTER(Tx[CashDividend],Tx[Symbol]=sym),-3)), 0),dVal + iVal)))',
    )
ws_f.set_column(0, 0, 10)
ws_f.set_column(1, 12, 12, fmt_money)

# ---------------- 8) Dashboard (simple totals row) ----------------
ws_d = wb.add_worksheet("Dashboard")
ws_d.write(0, 0, "TTM Monthly Income", fmt_header)
for j in range(12):
    from xlsxwriter.utility import xl_col_to_name

    ws_d.write_formula(
        0, 1 + j, f"='Monthly Dividend History'!{xl_col_to_name(1+j)}1", fmt_header
    )
ws_d.write(1, 0, "Income (TTM)", fmt_header)
for j in range(12):
    from xlsxwriter.utility import xl_col_to_name

    col = xl_col_to_name(1 + j)
    ws_d.write_formula(
        1,
        1 + j,
        f"=LET(rws,ROWS('Monthly Dividend History'!$A$3#),rng,'Monthly Dividend History'!{col}3:INDEX('Monthly Dividend History'!{col}:{col},2+rws),SUM(rng))",
        fmt_money,
    )
ws_d.set_column(0, 0, 26)

# ---------------- 9) README ----------------
ws_readme = wb.add_worksheet("README")
ws_readme.set_column(0, 0, 110)
ws_readme.write(
    0,
    0,
    "Paste your Fidelity export into the Transactions table and run the two Python updaters. AsOfDate (Prices!B1) controls the rolling windows.",
    fmt_text,
)

wb.close()
print("Wrote", out_path.resolve())
