# Excel Finance Organizer

This project is a financial data organization utility built on top of Mint's finance tracker. The user may export the transaction data as "transaction.csv", and then run (in sequence) `excel_finance.py` and `excel_finance_venmo.py`. The utility provides the following improvements on Mint's exported data.

- Truncate transactions to the last transaction marked \[END\] in the original description
- Negates all credit transactions
- Replaces categories with custom categories in `label_rep.txt` 
- Replaces categories based on description text in `desc_rep.txt`
- Adds an aggregation worksheet to sum dollar amount attributed to each unique category
- Presents the last 50 venmo transactions from / to the user
- Adds venmo notes / sender / reciever information to all matched venmo transactions
- Replaces categories based on venmo transaction note in `venmo_rep.txt`

Required files:
- `transactions.csv`, obtained from Mint as input
- `.env` with field `venmo_at` (venmo access token)
- `label_rep.txt` a tab separated list of the format `old_category` / `new_category`
- `desc_rep.txt` a tab separated list of the format `description_regex` / `new_category`
- `venmo_rep.txt` a tab separated list of the format `description_regex` / `new_category`
