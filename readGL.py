import pandas as pd
import matplotlib.pyplot as plt
from tkinter.filedialog import askopenfilename

consumableForScreens = ['vwr','rigaku','qiagen','ttp labtech','molecular dimensions','e&k','anatrace','hampton research']
managerSalaryBenefits = ['loren']

# checks if strings of lst1 is a substring of any string in lst2 and returns array of bool
def substringInListOfStrings(x,lst):
    for s in lst:
        if (x.find(s) > -1):
            return True
    return False

class UserException(Exception):
    pass

def main()

    # Read from GL (general ledger) file
    while True:
        try:
            filename = askopenfilename(initialdir = 'C:/Users/ljiang/Google Drive/UCSF_XTAL/MSG_recharge_automation/GL')
            title = "Select file (.csv or .tsv only)"
            path = filename #change to file name

            if not(filename): # empty file name occurs when user tries to close input window
                raise SystemExit()
            elif filename.endswith('.tsv'):
                df = pd.read_csv(path, delimiter="\t")
                break
            elif filename.endswith('.csv'):
                df = pd.read_csv(path, delimiter=",")
                break
            else:
                raise UserException
        except UserException as e:    
            print('\nWrong file type selected (ONLY .csv or .tsv files)')

    # MAIN SCRIPT BELOW
    print("\nRunning script on:\n" + filename +"\n")

    df.index = pd.to_datetime(df.pop(' Ldgr Post Date'))


    df['Recharge Category'] = df['  Description'].apply(lambda s: 'consumableScreen' 
        if substringInListOfStrings(str(s).lower(), consumableForScreens) else ('managerSalaryBenefits' 
            if substringInListOfStrings(str(s).lower(), managerSalaryBenefits) else 'distributedExpense'))

    a=df[df['Recharge Category'] == 'distributedExpense']

    monthlyTotals = a.groupby([pd.Grouper(freq="M")]).sum()

    return monthlyTotals
