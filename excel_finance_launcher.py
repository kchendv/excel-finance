from scripts.excel_finance import excel_finance_base
from scripts.excel_finance_venmo import excel_finance_venmo_base
from scripts.excel_finance_gsheets import excel_finance_gsheets_base
try: 
    print("Welcome to the Excel finance manager.")
    ef_switch = input("Launch excel_finance (y/n):").lower()
    if ef_switch not in ('n','y'):
        print("Unknown command, please try again.")
        ef_switch = input("Launch excel_finance (y/n):").lower()
    elif ef_switch == "y":
        excel_finance_base()

    efv_switch = input("Launch excel_finance_venmo (y/n):").lower()
    if efv_switch not in ('n','y'):
        print("Unknown command, please try again.")
        efv_switch = input("Launch excel_finance_venmo (y/n):").lower()
    elif efv_switch == "y":
        excel_finance_venmo_base()

    efg_switch = input("Launch excel_finance_gsheets (y/o/n):").lower()
    if efg_switch not in ('n','y','o'):
        print("Unknown command, please try again.")
        efg_switch = input("Launch excel_finance_gsheets (y/o/n):").lower()
    elif efg_switch == 'y':
        excel_finance_gsheets_base(False)
    elif efg_switch == 'o':
        excel_finance_gsheets_base(True)
except Exception as e:
    print(e)
input("Press any key to continue...")
