import sys
import pandas as pd 
import pygsheets as pyg 
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from tkinter.filedialog import askopenfilename


class UserException(Exception):
    pass

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
    
#TO BE IMPLEMENTED LATER
def clean_DF_MosquitoLCP(df):
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
def getGDriveLogUsage():
    df_mosquitoLCPLogRAW = gc.open_by_key('1MpwGvh6xlOb4mrn8BtJgs7Fux7hmlZRRdmhBUkhqRAY').sheet1.get_as_df()
    df_mosquitoLogRAW = gc.open_by_key('1demabrSE50t_euIpP3AhM8V64I3BQuaK1VRcNJCSJmA').sheet1.get_as_df()
    df_dragonflyLogRAW = gc.open_by_key('1JciEUj4dg1AZedcmi42InLIQs5XINQt5aaok4-vnUwg').sheet1.get_as_df()

    df_mosquitoLCPLog = clean_DF_MosquitoLCP(df_mosquitoLCPLogRAW)
    df_mosquitoLog = clean_DF_MosquitoCrystal(df_mosquitoLogRAW)
    df_dragonflyLog = clean_DF_Dragonfly(df_dragonflyLogRAW)

    return [df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog]

def getRockImagerUsage():
    # Read from RockImagers usage log file
    while True:
        try:
            filename = askopenfilename(initialdir = 'C:/Users/ljiang/Google Drive/UCSF_XTAL/MSG_recharge_automation/rockImagerLogs',
                title = "Select rockImagerUsage file (.csv or .tsv only)")
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
    # Remove "groups" not in recharge system (administrators, etc.)
    df = df[df['Group'] != 'Administrators']
    # Look at 'Duration' column and extract numerical data inplace
    durStr = 'Dur (min)'
    df.rename(columns={'Duration':durStr}, inplace=True)
    df[durStr] = df[durStr].str.replace(r'[^0-9.]','', regex=True).astype('float')

    # Look at 'Time' column and convert to datetime type
    df.index = pd.to_datetime(df.pop('Time')) # Change index to datetime index
    return df

def getGL(start_date,end_date):
    consumableForScreens = ['vwr','rigaku','qiagen','ttp labtech','molecular dimensions','e&k','anatrace','hampton research']
    managerSalaryBenefits = ['loren']
    # checks if strings of lst1 is a substring of any string in lst2 and returns array of bool
    def substringInListOfStrings(x,lst):
        for s in lst:
            if (x.find(s) > -1):
                return True
        return False

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

    return monthlyTotals.loc[start_date:end_date]

# Make the output files more Excel friendly
# assumes index in TimeStamp format categorized by month
def makeExcelFriendly(df):
    df.index.set_names(['Month/Year'],[0],inplace=True)
    j = 0
    while (j < len(df.index.levels)):
        if df.index.levels[j].dtype == 'datetime64[ns]':
            df.index.set_levels(df.index.levels[j]
                .strftime('%m/%d/%Y'),level=j,inplace=True)
        j+=1

#takes a list of df and list of their output paths w/ specified file extension
def exportToFile(dfLst, pathLst, fileExt='.xlsx'):
    if not(len(dfLst)==len(pathLst)):
        print('Lists need to be same length')
    elif not(fileExt == '.xlsx' or fileExt == '.csv'):
        print("fileExt needs to be '.csv' or '.xlsx'")
    else:
        k = 0
        if (fileExt == '.csv'):
            while (k < len(dfLst)):
                dfLst[k].to_csv(pathLst[k]+fileExt)
                k+=1
        else: 
            while (k < len(dfLst)):
                dfLst[k].to_excel(pathLst[k]+fileExt)
                k+=1   

def calculateRecharge(start_date,end_date):
    coreMultl = 1.0
    assocMult = 1.5
    regMult = 2.0
    # Get data from facilitySuppliesPricing (https://docs.google.com/spreadsheets/d/1d6GVWGwwrlh_lTKxVRI08xZSiE__Zieu3WWtwbmOMlE/edit#gid=0)

    costHDScreen = 25.00
    costSDScreen = 25.00
    costPerTip = 0.07

    costMosquitoTime = 0.00 #per hr
    costDragonflyTime = 0.00 #per hr
    costRockImagerTime = 5.00 / 60.0 #per min

    costHDSConsume = 7.50 #cost of consumables for mosquito hanging drop plate setup (optical seal, new plates, micro reservoirs)
    costSDSConsume = 3.89 #cost of consumables for mosquito sitting drop plate setup (optical seal, new plates, micro reservoirs)

    costDragonflyPlate = 0 #cost of consumables for one dragonfly plate setup (new reservoir, new pipettes, micro reservoirs)

    # Extract relevant data from df_RockImager in the given date range
    wantedRockCol = ['User Name','Project','Experiment','Plate Type', 'Dur (min)']

    rockImagerUsageMonthByPI = df_RockImager.groupby([pd.Grouper(freq="M"),'Group'])
    rockImagerUsageYearByPI = df_RockImager.groupby([df_RockImager.index.year,'Group'])
    itemizedRockMonthByPI = rockImagerUsageMonthByPI.apply(lambda x: x.head(
        len(x.index)))[wantedRockCol].loc[start_date:end_date]

    itemizedRockMonthByPI = itemizedRockMonthByPI[(itemizedRockMonthByPI['Dur (min)'] > 0)] #remove rows if 'Dur (min)' <= 0
    itemizedRockMonthByPI['Cost/min'] = costRockImagerTime
    itemizedRockMonthByPI['RockImager Cost'] = costRockImagerTime * itemizedRockMonthByPI['Dur (min)']

    itemizedRockMonthByPI=itemizedRockMonthByPI.iloc[::-1] #reverse index (descending date)

    # Extract relevant data from df_mosquitoLog in the given date range
    wantedMosqCol = ['Name','Plates','Protocol Used','Tips','Dur (hr)']

    mosqCrystalUsageMonthByPI = df_mosquitoLog.groupby([pd.Grouper(freq="M"),'Group'])
    mosqCrystalUsageYearByPI = df_mosquitoLog.groupby([df_mosquitoLog.index.year,'Group'])
    itemizedMosqCryMonthByPI = mosqCrystalUsageMonthByPI.apply(lambda x: x.head(
        len(x.index)))[wantedMosqCol].loc[start_date:end_date]

    #finds the type of protocol used (sitting or hanging) and assigns approp. cost/plate
    itemizedMosqCryMonthByPI['Cost/plate'] = itemizedMosqCryMonthByPI['Protocol Used'].apply(lambda x: 
        costHDSConsume if x.find('sitting')>0 else costSDSConsume)

    itemizedMosqCryMonthByPI['Cost/tip'] = costPerTip

    itemizedMosqCryMonthByPI['Mosquito Cost'] = costPerTip*itemizedMosqCryMonthByPI['Tips'] \
    + itemizedMosqCryMonthByPI['Cost/plate']*itemizedMosqCryMonthByPI['Plates']

    itemizedMosqCryMonthByPI=itemizedMosqCryMonthByPI.iloc[::-1] #reverse index (descending date)

    a = itemizedMosqCryMonthByPI.sum(level=[0,1],numeric_only=True)[['Tips','Mosquito Cost']]
    b = itemizedRockMonthByPI.sum(level=[0,1],numeric_only=True)[['Dur (min)','RockImager Cost']]

    monthlyRechargeTotal = pd.concat([a,b], axis=1).fillna(0)
    monthlyRechargeTotal['Raw Usage'] = monthlyRechargeTotal['Tips'] + monthlyRechargeTotal['Dur (min)']

    a1 = monthlyRechargeTotal.groupby(level=[0,1]).sum().groupby(level=0)

    rawUsagePercent = a1.apply(lambda x: x['Raw Usage'] / x['Raw Usage'].sum())
    monthlyRechargeTotal['Usage prop'] = pd.Series.ravel(rawUsagePercent)

    lst = []
    for index, row in monthlyRechargeTotal.iterrows():
        lst.append(row['Usage prop'] * df_GL.loc[index[0]][' Actual'])

    monthlyRechargeTotal['Dist Cost'] = lst  #distributed costs include everything except pay-per-use consumables and base salary/benefits
    monthlyRechargeTotal['Use Multiplier'] = regMult

    # Set Use Multiplier column for Core and Assoc users
    monthlyRechargeTotal.loc[((monthlyRechargeTotal.index.levels[0],
        coreUsers),'Use Multiplier')] = coreMultl
    monthlyRechargeTotal.loc[((monthlyRechargeTotal.index.levels[0],
        associateUsers),'Use Multiplier')] = assocMult

    monthlyRechargeTotal['Total Charge'] = (monthlyRechargeTotal['Dist Cost'] + monthlyRechargeTotal['Mosquito Cost'] + monthlyRechargeTotal['RockImager Cost'])

    outSummary = monthlyRechargeTotal[['Mosquito Cost','RockImager Cost','Usage prop','Dist Cost','Use Multiplier','Total Charge']]

    # itemizedMosqCryMonthByPI.index.set_names(['Month/Year'],[0],inplace=True)
    # itemizedRockMonthByPI.index.set_names(['Month/Year'],[0],inplace=True)
    outSummary.index.set_names(['Month/Year', 'Group'],[0,1],inplace=True)

    # itemizedMosqCryMonthByPI.index.set_levels(itemizedMosqCryMonthByPI.
    #   index.levels[0].strftime('%m/%Y'),level=0,inplace=True)
    # itemizedMosqCryMonthByPI.index.set_levels(itemizedMosqCryMonthByPI.
    #   index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)

    # itemizedRockMonthByPI.index.set_levels(itemizedRockMonthByPI.
    #   index.levels[0].strftime('%m/%Y'),level=0,inplace=True)
    # itemizedRockMonthByPI.index.set_levels(itemizedRockMonthByPI.
    #   index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)

    outSummary.index.set_levels(monthlyRechargeTotal.
        index.levels[0].strftime('%m/%Y'),level=0,inplace=True)


    fileOut = ['monthlyRecharges/recharge'+str(start_date)[0:10]+'_TO_'+str(end_date)[0:10],
    'monthlyRecharges/mosquitoUsage'+str(start_date)[0:10]+'_TO_'+str(end_date)[0:10],
    'monthlyRecharges/rockImagerUsage'+str(start_date)[0:10] +'_TO_'+str(end_date)[0:10]]

    dfOut = [outSummary, itemizedMosqCryMonthByPI, itemizedRockMonthByPI]

    return fileOut, dfOut

if __name__ == '__main__':

    # Get start_date and end_date from user input
    while 1:
        try:    
            start_date = datetime.strptime(input('Enter Start date in the format m/d/yyyy: '), '%m/%d/%Y')
            end_date = datetime.strptime(input('Enter End date in the format m/d/yyyy: '), '%m/%d/%Y')
            break
        except ValueError as e:
            if (e.args[0][11] == '0'):
                sys.exit(1)
            else:
                print('\nWrong format. Please try again or enter 0 to exit.')

    gc = pyg.authorize(service_file='msg-Recharge-24378e029f2d.json')
    coreUsers, associateUsers, regUsers = getPITypes()
    df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog = getGDriveLogUsage()
    df_RockImager = getRockImagerUsage()
    df_GL = getGL(start_date,end_date)
    fileOut, dfOut = calculateRecharge(start_date,end_date)