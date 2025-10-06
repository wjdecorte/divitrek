import pandas as pd
import datetime as dt
from pathlib import Path

# Use repository data directory for input files
project_root = Path(__file__).resolve().parents[1]
data_dir = project_root / "data"
positions_file = data_dir / "current_positions.xlsx"
history_file = data_dir / "Fidelity_Full_History_20240701_20251001.csv"

# Load data
positions = pd.read_excel(positions_file)
history = pd.read_csv(history_file)
history["Run Date"] = pd.to_datetime(history["Run Date"], errors="coerce")

# Extract different transaction types
div_received = history[
    history["Action"].str.contains("DIVIDEND RECEIVED", na=False)
].copy()
purchases = history[history["Action"].str.contains("BOUGHT", na=False)].copy()
reinvestments = history[history["Action"].str.contains("REINVESTMENT", na=False)].copy()


# Calculate dividends per payment date with shares owned
def calculate_dividends_by_date(divs, purchases):
    purchases = purchases.sort_values(by="Run Date")
    divs = divs.sort_values(by="Run Date")
    results = []
    for _, div in divs.iterrows():
        symbol = div["Symbol"]
        pay_date = div["Run Date"]
        amount_received = div["Amount ($)"]
        # Shares owned before dividend date (assumed day before payout)
        shares_bought = purchases[
            (purchases["Symbol"] == symbol) & (purchases["Run Date"] < pay_date)
        ]
        total_shares = shares_bought["Quantity"].sum()
        results.append(
            {
                "Symbol": symbol,
                "Dividend Date": pay_date,
                "Shares Owned": total_shares,
                "Amount Received": amount_received,
            }
        )
    return pd.DataFrame(results)


dividend_details = calculate_dividends_by_date(div_received, purchases)

# Calculate trailing 12 months dividend totals
one_year_ago = dt.datetime.now() - pd.DateOffset(years=1)
recent_dividends = dividend_details[dividend_details["Dividend Date"] >= one_year_ago]
total_dividends = (
    recent_dividends.groupby("Symbol")["Amount Received"].sum().reset_index()
)

# Merge trailing 12M dividends to positions
positions = positions.merge(
    total_dividends.rename(
        columns={"Amount Received": "Trailing 12M Dividend Received"}
    ),
    how="left",
    left_on="Symbol",
    right_on="Symbol",
)
positions["Trailing 12M Dividend Received"] = positions[
    "Trailing 12M Dividend Received"
].fillna(0)

# Add last dividend amount per ticker
last_dividend = (
    div_received.sort_values("Run Date")
    .groupby("Symbol")
    .tail(1)[["Symbol", "Amount ($)"]]
)
positions = positions.merge(
    last_dividend.rename(columns={"Amount ($)": "Last Dividend Paid"}),
    how="left",
    on="Symbol",
)
positions["Last Dividend Paid"] = positions["Last Dividend Paid"].fillna(0)

positions.fillna(0, inplace=True)

# Sample price history data for price changes - replace with real data source if available
price_data_examples = {
    "BITO": pd.Series([19, 19.5, 20, 20.3, 20.1, 19.8, 19.67]),
    "BTCI": pd.Series([60, 61, 62, 61, 60.5, 60.75, 60.62]),
    "HOOW": pd.Series([63, 63.5, 64, 64.2, 64.5, 64.4, 64.29]),
    "HPE": pd.Series([17, 17.5, 18, 18.2, 18.1, 17.9, 18]),
    "SCHD": pd.Series([27, 27.2, 27.5, 27.6, 27.7, 27.6, 27.52]),
    "ULTY": pd.Series([5.5, 5.52, 5.53, 5.48, 5.5, 5.52, 5.47]),
}


def calculate_price_changes(price_history, days):
    if len(price_history) < days + 1:
        return None, None
    current_close = price_history.iloc[-1]
    past_close = price_history.iloc[-1 - days]
    price_change = current_close - past_close
    pct_change = (price_change / past_close) * 100 if past_close != 0 else None
    return price_change, pct_change


price_change_data = []
for ticker, prices in price_data_examples.items():
    change_5d, pct_5d = calculate_price_changes(prices, 5)
    change_1m, pct_1m = calculate_price_changes(prices, 22)
    change_3m, pct_3m = calculate_price_changes(prices, 66)
    change_6m, pct_6m = calculate_price_changes(prices, 132)
    price_change_data.append(
        {
            "Symbol": ticker,
            "Price Change 5D": change_5d or 0,
            "Price Change % 5D": pct_5d or 0,
            "Price Change 1M": change_1m or 0,
            "Price Change % 1M": pct_1m or 0,
            "Price Change 3M": change_3m or 0,
            "Price Change % 3M": pct_3m or 0,
            "Price Change 6M": change_6m or 0,
            "Price Change % 6M": pct_6m or 0,
        }
    )

price_change_df = pd.DataFrame(price_change_data)

# Add NAV and AUM data (update as you get real-time data)
additional_data = {
    "Symbol": ["BITO", "BTCI", "HOOW", "HOOY", "HPE", "MSTY", "QQQI", "SCHD", "ULTY"],
    "NAV": [19.67, 60.62, 64.29, 0, 0, 0, 0, 27.52, 5.47],
    "AUM": [2.75e9, 8.34e8, 1.0048e8, 0, 0, 0, 0, 7.111e10, 3.38e9],
}
additional_df = pd.DataFrame(additional_data)

# Merge price changes, NAV, AUM with positions
positions = positions.merge(price_change_df, how="left", on="Symbol")
positions = positions.merge(additional_df, how="left", on="Symbol")

positions.fillna(0, inplace=True)

# Prepare Monthly Dividend History tab
dividend_details["Month"] = (
    dividend_details["Dividend Date"].dt.to_period("M").astype(str)
)
monthly_dividends = (
    dividend_details.groupby(["Symbol", "Month"])["Amount Received"].sum().reset_index()
)

# Prepare Dividend Forecast (assuming last dividend repeated monthly - adjust as needed)
forecast_months = pd.date_range(
    start=pd.Timestamp(dt.datetime.today()).to_period("M").to_timestamp(),
    periods=12,
    freq="M",
)


def generate_forecast(symbol, last_dividend):
    if last_dividend == 0:
        return pd.DataFrame()
    return pd.DataFrame(
        [
            {
                "Symbol": symbol,
                "Month": month.strftime("%Y-%m"),
                "Projected Dividend": last_dividend,
            }
            for month in forecast_months
        ]
    )


forecast_list = []
for _, row in positions.iterrows():
    forecast_list.append(generate_forecast(row["Symbol"], row["Last Dividend Paid"]))

forecast_df = pd.concat(forecast_list) if forecast_list else pd.DataFrame()

# Export all data to an Excel file with multiple tabs
output_file = data_dir / "Dividend_Tracker_Complete.xlsx"
with pd.ExcelWriter(output_file) as writer:
    positions.to_excel(writer, index=False, sheet_name="Holdings & Summary")
    monthly_dividends.to_excel(
        writer, index=False, sheet_name="Monthly Dividend History"
    )
    forecast_df.to_excel(writer, index=False, sheet_name="Dividend Forecast")
    dividend_details.to_excel(writer, index=False, sheet_name="Dividend Details")

print(f"Wrote {output_file.resolve()}")
