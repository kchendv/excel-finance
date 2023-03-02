def excel_finance_gsheets_base(override = False):
    # TODO  Excel finance mint net worth integration
    # TODO Excel finance google sheets auto fill in
    # TODO Excel finance charles schwab stock integration
    # TODO Excel finance programm organization
    import pygsheets, openpyxl
    from dotenv import dotenv_values
    from datetime import datetime

    config = dotenv_values(".env")
    INIT_BALANCE = "INIT BALANCE"
    END_BALANCE = "END BALANCE"
    WB_NAME = config["WB_NAME"]
    GSHEET_NAME = config["GSHEET_NAME"]
    DIV = config["DIV"]

    # Open google sheets
    sheets_client = pygsheets.authorize(service_file='gsheet_creds.json')
    finance_sheet = sheets_client.open(GSHEET_NAME)
    income_st = finance_sheet[0]
    balance_st = finance_sheet[1]
    print(f"Connected to google sheets \n{DIV}\n")

    # Map category columns to index
    spending_cols = income_st.get_values("A1", "A50", include_tailing_empty  = False)
    income_map = dict()
    for i in range(len(spending_cols)):
        income_map[spending_cols[i][0]] = i + 1
    print(f"Mapped {len(income_map)} income statement columns to index \n{DIV}\n")

    # Fill out next month date
    today_date = datetime.now().strftime("%m/%d")
    next_month = len(income_st.get_values("A1", "L1", include_tailing_empty  = False)[0]) + 1
    # Override last month when error was made
    if override:
        next_month -= 1
    income_st.update_value((1, next_month), today_date)
    print(f"Filled current date: {today_date} \n{DIV}\n")

    # Fill init balance formula
    if next_month > 2:
        init_bal_cell = income_st.cell((income_map[INIT_BALANCE], next_month))
        end_bal_cell = income_st.cell((income_map[END_BALANCE], next_month - 1))
        income_st.update_value(init_bal_cell.label, f"={end_bal_cell.label}")

    print(f"Filled init balance \n{DIV}\n")

    # Build spending update targets
    wb = openpyxl.load_workbook(WB_NAME)
    final_sheet = wb["Final"]
    spendings = final_sheet['A2:B50']
    update_targets = dict()

    for row in range(len(spendings)):
        spend_category = spendings[row][0].value
        spend_amount = spendings[row][1].value
        if spend_amount:
            spend_amount = abs(float(spend_amount))
            if spend_category in update_targets:
                update_targets[spend_category] += spend_amount
            else:
                update_targets[spend_category] = spend_amount

    print(f"Built {len(update_targets)} update targets \n{DIV}\n")

    # Attempt to match spending categories
    count = 0
    for cat, amt in update_targets.items():
        if cat in income_map:
            income_st.update_value((income_map[cat], next_month), f"={amt}")
            count += 1
        else:
            print(f"Cannot update target [{cat}] ${amt}")

    # Pad income st with 0s
    for cat, ind in income_map.items():
        if cat not in update_targets and cat.upper() != cat:
            income_st.update_value((ind, next_month), f"=0")

    print(f"\nFilled {count} spending / income categories \n{DIV}\n")

    # ----
    # Map account columns to index
    account_cols = balance_st.get_values("A1", "A50", include_tailing_empty  = False)
    account_map = dict()
    for i in range(len(account_cols)):
        account_map[account_cols[i][0]] = i + 1


    print(account_map)

    # Get latest month
    today_date = datetime.now().strftime("%m/%d/%Y")
    next_month = len(balance_st.get_values("A1", "Y1", include_tailing_empty  = False)[0]) + 1
    balance_st.update_value((1, next_month), today_date)
    print(f"Filled current date: {today_date} \n{DIV}\n")

    # Build account update targets
    wb = openpyxl.load_workbook(WB_NAME)
    final_sheet = wb["Final"]
    accounts = final_sheet["C2:D20"]
    acc_targets = dict()

    for row in range(len(accounts)):
        account_name = accounts[row][0].value
        account_balance = accounts[row][1].value
        if account_balance:
            account_balance = float(account_balance)
            acc_targets[account_name] = account_balance

    print(f"Built {len(acc_targets)} account update targets \n{DIV}\n")


    # Fill in relevant columns
    count = 0
    for cat, amt in acc_targets.items():
        if cat in balance_st:
            income_st.update_value((balance_st[cat], next_month), f"={amt}")
            count += 1
        else:
            print(f"Cannot update account target [{cat}] ${amt}")

    # Pad income st with same value as last month
    for cat, ind in account_map.items():
        if cat not in acc_targets and cat.upper() != cat:
            last_month_acc = balance_st.cell((ind, next_month - 1))
            income_st.update_value((ind, next_month), f"={last_month_acc.label}")

    print(f"Updated info of {count} accounts\n{DIV}\n")
