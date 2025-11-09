import datetime

def main():
    print("=== Portfolio Transaction Tracker ===")
    cash = float(input("Enter starting cash: "))
    holdings = {}
    history = {}

    while True:
        date = input("\nEnter date (YYYY-MM-DD) or 'save' to finish: ").strip()
        if date.lower() == "save":
            break

        if not date:
            continue

        print(f"Adding transactions for {date}...")
        daily_changes = {}
        while True:
            ticker = input("Enter ticker (or 'done' to finish this date): ").strip().upper()
            if ticker.lower() == "done":
                break
            qty = float(input("Enter quantity (+ for buy, - for sell): "))
            amount = float(input("Enter amount ($): "))
            fee = float(input("Enter transaction fee ($): "))

            # Adjust cash and holdings
            cash_change = -amount - fee if qty > 0 else amount - fee
            cash += cash_change

            holdings[ticker] = holdings.get(ticker, 0) + qty

            # Track daily changes
            daily_changes[ticker] = daily_changes.get(ticker, 0) + qty

            print("Transaction added.")

        # Save snapshot of this day's holdings
        history[date] = {
            "cash": round(cash, 2),
            "holdings": holdings.copy()
        }
        print(f"Data for {date} saved.")

    # Save results to file
    with open("account_history.txt", "w") as f:
        for date, data in sorted(history.items()):
            f.write(f"Date: {date}\n")
            f.write(f"Cash: {data['cash']}\n")
            for ticker, qty in data["holdings"].items():
                f.write(f"{ticker}: {qty}\n")
            f.write("\n")

    print("\n✅ All data saved to account_history.txt")

if __name__ == "__main__":
    main()
