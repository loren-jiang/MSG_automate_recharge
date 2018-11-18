import sys
import pandas as pd 
import pygsheets as pyg 
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

def clean_DF_MosquitoCrystal(df):
    # Rename columns to make easier to read, NEED TO CHANGE IF CHANGES MADE IN GOOGLE FORM
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

    return df

def clean_DF_Dragonfly(df):
    return df

# obtain PI and their use_types from Google drive spreadsheet
def getPITypes():
    wks_groups = gc.open_by_key('1Yom-H6j04TJ5_W1Ic2dlqKsVWrHQbdInQyq_Q7ZVnZA').sheet1
    row_data = wks_groups.get_all_values(include_empty=False)
    coreUsers = []
    associateUsers = []
    regUsers = []

    k = 1 #skip first row
    while (k < len(row_data)):
        pi, user_type = row_data[k]
        if (user_type=='regular'):
            regUsers.append(pi)
        elif (user_type=='associate'):
            associateUsers.append(pi)
        else:
            coreUsers.append(pi)
        k +=1
    return [coreUsers,associateUsers,regUsers]

# obtain usage log data from Google drive forms/spreadsheets (mosquito, mosquitoLCP, dragonfly)
def getGDriveLogData():
    df_mosquitoLCPLogRAW = gc.open_by_key('1MpwGvh6xlOb4mrn8BtJgs7Fux7hmlZRRdmhBUkhqRAY').sheet1.get_as_df()
    df_mosquitoLogRAW = gc.open_by_key('1demabrSE50t_euIpP3AhM8V64I3BQuaK1VRcNJCSJmA').sheet1.get_as_df()
    df_dragonflyLogRAW = gc.open_by_key('1JciEUj4dg1AZedcmi42InLIQs5XINQt5aaok4-vnUwg').sheet1.get_as_df()

    # df_mosquitoLCPLog = clean_DF_MosquitoCrystal(df_mosquitoLCPLogRAW)
    df_mosquitoLog = clean_DF_MosquitoCrystal(df_mosquitoLogRAW)
    df_dragonflyLog = clean_DF_Dragonfly(df_dragonflyLogRAW)
    return [df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog]

if __name__ == '__main__':
    gc = pyg.authorize(service_file='msg-Recharge-24378e029f2d.json')
    coreUsers, associateUsers, regUsers = getPITypes()
    df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog = getGDriveLogData()