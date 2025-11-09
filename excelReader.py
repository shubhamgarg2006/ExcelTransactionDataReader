import pandas as pd
import os

# --- Automatically use Excel file in the same folder as the script ---
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "portfolio.xlsx")

# --- Load Excel ---
start_positions = pd.read_excel(file_path, sheet_name="Positions")
transactions = pd.read_excel(file_path, sheet_name="Transactions")

# --- Initialize portfolio ---
portfolio = {}
for idx, row in start_positions.iterrows():
    ticker = row['Ticker']
    qty = row.get('Quantity', 0) if not pd.isna(row.get('Quantity', 0)) else 0
    cash = row.get('Cash', 0) if not pd.isna(row.get('Cash', 0)) else 0
    if ticker.lower() == 'cash':
        portfolio['Cash'] = cash
    else:
        portfolio[ticker] = qty

if 'Cash' not in portfolio:
    portfolio['Cash'] = 0

# --- Sort and group transactions by date ---
transactions['Date'] = pd.to_datetime(transactions['Date'])
transactions = transactions.sort_values('Date')
grouped = transactions.groupby('Date')

# --- Process each date’s transactions cumulatively ---
positions_over_time = []

for date, day_trades in grouped:
    for idx, row in day_trades.iterrows():
        action = row['Action'].lower().strip()
        ticker = row['Ticker']
        qty = row.get('Quantity', 0) if not pd.isna(row.get('Quantity', 0)) else 0
        amount = row.get('Amount', 0) if not pd.isna(row.get('Amount', 0)) else 0
        trans_price = row.get('TransactionPrice', 0) if 'TransactionPrice' in row and not pd.isna(row['TransactionPrice']) else 0

        # Ensure ticker exists if stock transaction
        if ticker not in portfolio and action in ['buy', 'sell']:
            portfolio[ticker] = 0

        # --- Apply transaction rules ---
        if action == 'buy':
            portfolio[ticker] = portfolio.get(ticker, 0) + qty
            portfolio['Cash'] += (amount)   # pay cash for stock + transaction fee
        elif action == 'sell':
            portfolio[ticker] = portfolio.get(ticker, 0) + qty  # qty negative for sells
            portfolio['Cash'] += (amount - trans_price)  # receive sale proceeds minus transaction fee
        elif action in ['div', 'int', 'other']:
            portfolio['Cash'] += amount

    # --- Remove tickers with zero quantity ---
    tickers_to_remove = [t for t, q in portfolio.items() if t != 'Cash' and q == 0]
    for t in tickers_to_remove:
        del portfolio[t]

    # --- Snapshot after all trades that day ---
    snapshot = {'Date': date}
    snapshot.update(portfolio)
    positions_over_time.append(snapshot)

# --- Write formatted text file ---
text_output = os.path.join(script_dir, "portfolio_positions.txt")
with open(text_output, "w") as f:
    for snapshot in positions_over_time:
        f.write(f"Date: {snapshot['Date'].strftime('%Y-%m-%d')}\n")
        f.write(f"Cash: {snapshot['Cash']}\n")
        for ticker, qty in snapshot.items():
            if ticker not in ['Date', 'Cash']:
                f.write(f"{ticker}: {qty}\n")
        f.write("\n")

# --- Write Excel file (same readable format) ---
excel_rows = []
for snapshot in positions_over_time:
    date_str = snapshot['Date'].strftime('%Y-%m-%d')
    excel_rows.append({'Date': date_str, 'Field': 'Cash', 'Value': snapshot['Cash']})
    for ticker, qty in snapshot.items():
        if ticker not in ['Date', 'Cash']:
            excel_rows.append({'Date': date_str, 'Field': ticker, 'Value': qty})

excel_output = os.path.join(script_dir, "portfolio_positions_readable.xlsx")
excel_df = pd.DataFrame(excel_rows)
excel_df.to_excel(excel_output, index=False)

print("✅ Portfolio positions saved to:")
print(f"   - {text_output}")
print(f"   - {excel_output}")