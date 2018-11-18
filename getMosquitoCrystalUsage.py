import pandas as pd
import matplotlib.pyplot as plt
from tkinter.filedialog import askopenfilename

class UserException(Exception):
    pass

def main():
    # Read from Mosquito's usage log file
    while True:
        try:
            filename = askopenfilename(initialdir = 'C:/Users/ljiang/Google Drive/UCSF_XTAL/MSG_recharge_automation/mosquitoLogs',
                title = "Select mosquitoCrysalUsage file (.csv or .tsv or .txt only)")
            path = filename #change to file name

            if not(filename): # empty file name occurs when user tries to close input window
                raise SystemExit()
            elif filename.endswith('.tsv'):
                df = pd.read_csv(path, delimiter="\t")
                break
            elif filename.endswith('.txt'):
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

    # Rename columns to make easier to read
    df.rename(columns={'How many tips did you use in total?':'Tips',
    	'How many plates did you set up?':'Plates',
        'What is your name?':'Name',
    	'Which lab are you from?':'Group',
        'Did you use your own plates or those from the common supply?':'Source of plates',
        'What day did you use the Mosquito?':'Self-report use date',
        'Which protocol did you use?':'Protocol Used',
    	'Was this a training session?':'Training sesh?',
    	'Did you park the Mosquito in position #3 and put everything back into the drawer after you were finished?':'Cleaned-up?',
    	'How long did you use the Mosquito? (1 hr and 15 min → 1.25)':'Dur (hr)'}, inplace=True)
    # Look at 'Time' column and convert to datetime type
    # df.index = pd.to_datetime(df.pop('What day did you use the Mosquito?')) # Change index to datetime index
    df.index = pd.to_datetime(df.pop('Timestamp')) # Change index to datetime index
    df['Tips'] = pd.to_numeric(df['Tips'], errors='coerce') # convert objects to numeric
    df['Plates'] = pd.to_numeric(df['Plates'], errors='coerce') # convert objects to numeric
    df.dropna(subset=['Group', 'Plates', 'Tips'], inplace=True) # Removes rows in the labels subset that are NaN

    # Group by PIs and aggregate total duration
    groupedByPI = df.groupby('Group')
    usagePlatesByPI = groupedByPI.sum()['Plates'] # num of plates
    usageTipsByPI = groupedByPI.sum()['Tips'] # num of tips
    usagePlatesSum = usagePlatesByPI.sum()
    usageTipsSum = usageTipsByPI.sum()

    # print(100*usagePlatesByPI/usagePlatesSum) #percent usage 
    # print(100*usageTipsByPI/usageTipsSum)

    # Create pie chart of usage by PI group -- too crowded, need to reformat...
    # labels = usageByPI.index
    # usageByPI.plot.pie(figsize=(6, 6), labels=labels, autopct='%1.0f%%')
    # plt.show()

    usageByMonthByPI = df.groupby([pd.Grouper(freq="M"),'Group']).sum()
    usageByYearByPI = df.groupby([df.index.year,'Group']).sum()

    # usageByMonth.xs('Minor', level = 1)
    return df