# ExcelTransactionDataReader

Hello! Thank you for viewing this project. The goal of this project was to create a transaction analysis script that could figure out the holdings of a portfolio at a given date. This is useful for our case, if we are using the Bloomberg terminal to track our portfolios historical positions. The final positions can then be used to conduct attribution analysis (learn about it here: https://www.investopedia.com/terms/a/attribution-analysis.asp).

This project was effectively vibe coded with the help of ChatGPT. I only take credit for developing the lying logic for this script.

The current best file to use is: excelReader.py

## How to use excelReader.py

I used vscode to develop and run this program. This program employed the use of python libraries pandas(data manipulation and analysis) and os (operating system interations). You will have to downlad them into your virtual engine. Use some pip command?

You'll need your input excel to be in the same folder as your program titled: "portfolio.xlsx". This indicates that to the program that that excel file is the one that it needs to read. Follow the prompts below to understand more about the inputs required to make this program work.

### Inputs:

### Excel Input Setup:
This is how I set up my excel sheet. You'll have to do it the same way if you want the program to work!

#### Sheet 1

#### Sheet 2

### Flaws of this system:
These are things I plan to fix at one point!
#### 1:
I don't show you the amount you bought the equity for. This can make it hard to rely on just one spreadsheet, for inputting. Its important for buying to know the purchase price (it can different than the closing price).
#### 2:
No summary of the transactions for that date of input. This could help you figure out if you have only a cash adjustment or euity position adjustment.
#### 3: 
You'll have to create your own input excel file. I think there are ways to just optimize for this, I'll expirerment on it and just back to y'all.

## How did excelReader.py come to be?
I realized very quickly after "ONELINE.PY" that just using the excel file and a script might be the play! The underlying logic still applied, but now it read an excel input and output and excel file. This made working with Bloomberg, a lot more easier!

## How did ONELINE.PY come to be?
This was an improved version of "hello.py", allowing users to list out transaction data in one line itself for a specfic date.

Of course, it fell prey to same logic as "hello.py", typing in the inputs just did not make sense. It was still a time sink.

## How did hello.py come to be?
This was the first attempt to make this project work. At this point, I was entering each transaction into the terminal. This was not effective nor efficient. If I made a mistake inputting data, I would have to remake the entire program... That is not ideal.

In term of output, it would generate a .txt file. The value of this interation was that it was a proof of concept. It just needed to be optimized.

## What started the whole project?
Basically, we needed to find the date by date position history of our club's investment account to help us run attribution analysis (basically shows us how our investment descions fared against the idustry standard).

With 140 lines of transaction data, it would not be worth our time to scale through all that data and would take way more time than needed. Thinking this, the python script idea was born, to optimize this process. 

There were five types of transactions we needed to account for: buy, sell, interest, dividend, cash withdrawals, and other. Interest, Dividends, Cash Withdrawals and other were basically a cash account and would only affect our cash balance. The method for them was just just add them to to cash. NOTE: Cash Withdrawals were already a negative input (I'll explain more in the inputs section).

Now with buy and sell I had to be careful. The cash value of what we were buying would already include the transaction price, and given that it had a negative value, we could add it to the cash balance at that date. For selling the stock, we would have to subtract the transaction price from the cash value.

# Final Takeaways
### Can't you just take the very last transaction date's position history?
If we only need to see the very last date, yes. But, we want to see how we did over time.
