import pandas as pd
import matplotlib.pyplot as plt
from tkinter.filedialog import askopenfilename

class UserException(Exception):
    pass


# Read from Quartzy's usage log file
while True:
    try:
        filename = askopenfilename(initialdir = 'C:/Users/ljiang/Google Drive/UCSF_XTAL',
            title = "Select screenOrders file (.csv or .tsv or .txt only)")
        path = filename #change to file name

        if not(filename): # empty file name occurs when user tries to close input window
            raise SystemExit()
        elif filename.endswith('.tsv'):
            df = pd.read_csv(path, delimiter="\t",encoding = "utf-8")
            break
        elif filename.endswith('.txt'):
            df = pd.read_csv(path, delimiter="\t",encoding = "utf-8")
            break
        elif filename.endswith('.csv'):
            df = pd.read_csv(path, delimiter=",",encoding = "utf-8")
            break
        else:
            raise UserException
    except UserException as e:    
        print('\nWrong file type selected (ONLY .csv or .tsv files)')

def removeBadChar(x):
    if isinstance(x, str):
        x=x.replace('"','').replace('=','')
    return x


# MAIN SCRIPT BELOW
print("\nRunning script on:\n" + filename +"\n")

# make 'Date Requested' the index
df = df.applymap(removeBadChar) #removes unwanted " and = characters
df.rename(removeBadChar,axis='columns',inplace=True)
df.rename(index=str, columns={"Spend Tracking Code": "Group"},inplace=True)
df.index = pd.to_datetime(df.pop('Date Requested')) # Change index to datetime index
df.sort_values(['Group'],inplace=True)
usageByMonthByPI = df.groupby([pd.Grouper(freq="M"),'Group'])
usageByMonthByPI = usageByMonthByPI.apply(lambda x: x.head(
    len(x.index)))

df.loc[df['Group'] == 'Shokat']
