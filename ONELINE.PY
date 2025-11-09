def main():
    print("=== Portfolio Tracker (One-Line Input Version) ===")

    start_date = input("Enter starting date (YYYY-MM-DD): ").strip()
    cash = float(input("Enter starting cash amount: "))

    holdings = {}
    print("Enter starting holdings (enter 'done' when finished):")
    while True:
        ticker = input("Ticker: ").strip().upper()
        if ticker.lower() == "done":
            break
        qty = float(input(f"Starting quantity for {ticker}: "))
        holdings[ticker] = qty

    history = {start_date: {"cash": cash, "holdings": holdings.copy()}}

    while True:
        date = input("\nEnter date for transactions (YYYY-MM-DD) or 'save' to finish: ").strip()
        if date.lower() == "save":
            break

        print(f"\nAdding transactions for {date}...")
        print("Format examples:")
        print("  BUY:  AAPL, buy, 10, 1700")
        print("  SELL: MSFT, sell, 5, 900, 5")
        print("  DIVIDEND: dividend, 20")
        print("  INTEREST: interest, 15")
        print("Type 'undo' to revert last transaction, or 'done' to finish this date.\n")

        daily_transactions = []
        temp_cash = cash
        temp_holdings = holdings.copy()

        while True:
            line = input("> ").strip()
            if not line:
                continue

            if line.lower() == "done":
                confirm = input("Save this date’s transactions? (y/n): ").strip().lower()
                if confirm == "y":
                    cash = temp_cash
                    holdings = temp_holdings
                    history[date] = {"cash": round(cash, 2), "holdings": holdings.copy()}
                    print(f"✅ Data for {date} saved.")
                    break
                else:
                    print("❌ Discarding changes for this date.")
                    break

            if line.lower() == "undo":
                if daily_transactions:
                    last = daily_transactions.pop()
                    temp_cash = last["prev_cash"]
                    temp_holdings = last["prev_holdings"]
                    print(f"↩️ Undid last transaction.")
                else:
                    print("No transactions to undo.")
                continue

            prev_cash = temp_cash
            prev_holdings = temp_holdings.copy()

            parts = [p.strip() for p in line.split(",")]

            if not parts:
                continue

            # Identify transaction type
            trans_type = parts[1].lower() if len(parts) > 1 else parts[0].lower()

            if trans_type == "buy":
                ticker, _, qty, amount = parts[:4]
                qty, amount = float(qty), float(amount)
                temp_cash -= amount
                temp_holdings[ticker.upper()] = temp_holdings.get(ticker.upper(), 0) + qty

            elif trans_type == "sell":
                ticker, _, qty, amount, fee = parts[:5]
                qty, amount, fee = float(qty), float(amount), float(fee)
                temp_cash += (amount - fee)
                temp_holdings[ticker.upper()] = temp_holdings.get(ticker.upper(), 0) - qty

            elif trans_type in ["dividend", "interest"]:
                _, amount = parts[:2]
                amount = float(amount)
                temp_cash += amount

            else:
                print("❌ Invalid transaction format. Try again.")
                continue

            daily_transactions.append({
                "prev_cash": prev_cash,
                "prev_holdings": prev_holdings.copy(),
                "input": line
            })
            print("✅ Transaction recorded.")

    # Save results to text file
    with open("account_history.txt", "w") as f:
        for date, data in sorted(history.items()):
            f.write(f"Date: {date}\n")
            f.write(f"Cash: {data['cash']}\n")
            for ticker, qty in data["holdings"].items():
                f.write(f"{ticker}: {qty}\n")
            f.write("\n")

    print("\n✅ All data saved to 'account_history.txt'")

if __name__ == "__main__":
    main()
