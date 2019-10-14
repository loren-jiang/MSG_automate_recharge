import sys
import os
import pandas as pd 
import pygsheets as pyg 
import numpy as np
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from tkinter.filedialog import askopenfilename
import shutil

class UserException(Exception):
    pass    

def clean_DF_MosquitoCrystal(df,dates):
    start_date, end_date = dates

    # Rename columns to make easier to read, NEED TO CHANGE IF CHANGES MADE IN GOOGLE FORM
    df.rename(columns={'How many tips did you use in total?':'Mosq Tips',
        'How many plates did you set up?':'Mosq Plates Set-up',
        'How many sitting drop plates did you set up? (Enter 0 if none)' : 'NumSDP',
        'How many hanging drop plates did you set up? (Enter 0 if none)' : 'NumHDP',
        'What is your name?':'Name',
        'Which lab are you from?':'Group',
        'Did you use your own plates or those from the common supply?':'Source of plates',
        'What day did you use the Mosquito?':'Self-report use date',
        'Which protocol did you use?':'Mosq Protocol Used',
        'Was this a training session?':'Training sesh?',
        'Did you park the Mosquito in position #3 and put everything back into the drawer after you were finished?':'Cleaned-up?',
        'How long did you use the Mosquito? (1 hr and 15 min → 1.25)':'Mosq Dur (hr)'}, inplace=True)
    # Look at 'Timestamp' column and convert to datetime type
    df.index = pd.to_datetime(df.pop('Timestamp'))
    df = df[df.index.notnull()]
    mask = (df.index > start_date) & (df.index <= end_date)
    df = df.loc[mask]
    df['Group'] = df['Group'].map(lambda s: s.lower()) #make group name lowercase to standardize
    df['Mosq Tips'] = pd.to_numeric(df['Mosq Tips'], errors='coerce').fillna(0) # convert objects to numeric
    df['NumSDP'] = pd.to_numeric(df['NumSDP'], errors='coerce').fillna(0) # convert objects to numeric
    df['NumHDP'] = pd.to_numeric(df['NumHDP'], errors='coerce').fillna(0) # convert objects to numeric
    df['Mosq Dur (hr)'] = pd.to_numeric(df['Mosq Dur (hr)'], errors='coerce').fillna(0) # convert objects to numeric
    return df
    
def clean_DF_MosquitoLCP(df,dates):
    start_date, end_date = dates
    # Rename columns to make easier to read, NEED TO CHANGE IF CHANGES MADE IN GOOGLE FORM
    df.rename(columns={'How many tips did you use in total?':'Mosq Tips',
        'How many LCP plates did you set up?':'Mosq LCP Plates Set-up',
        'How many sitting drop plates did you set up? (Enter 0 if none)' : 'NumSDP',
        'How many hanging drop plates did you set up? (Enter 0 if none)' : 'NumHDP',
        'What is your name?':'Name',
        'Which lab are you from?':'Group',
        'What day did you use the Mosquito LCP?':'Self-report use date',
        'Which protocol did you use?':'Mosq Protocol Used',
        'Was this a training session?':'Training sesh?',
        'Did you park the Mosquito in position #3 and put everything back into the drawer after you were finished?':'Cleaned-up?',
        'How long did you use the Mosquito LCP? (1 hr and 15 min → 1.25)':'Mosq Dur (hr)'}, inplace=True)
    
    df.index = pd.to_datetime(df.pop('Timestamp'))
    df = df[df.index.notnull()]
    mask = (df.index > start_date) & (df.index <= end_date)
    df = df.loc[mask]
    df['Group'] = df['Group'].map(lambda s: s.lower()) #make group name lowercase to standardize
    # convert columns to numeric values as they should be
    df['Mosq LCP Plates Set-up'] = pd.to_numeric(df['Mosq LCP Plates Set-up'], errors='coerce')
    df['NumSDP'] = pd.to_numeric(df['NumSDP'], errors='coerce').fillna(0)
    df['NumHDP'] = pd.to_numeric(df['NumHDP'], errors='coerce').fillna(0)
    a = pd.to_numeric(df['Mosq Tips'].ix[:,0], errors='coerce').fillna(0)
    b = pd.to_numeric(df['Mosq Tips'].ix[:,1], errors='coerce').fillna(0)
    df['Cum Mosq Tips'] = a.add(b)
    return df

def clean_DF_Dragonfly(df,dates):
    start_date, end_date = dates

    df.rename(columns={
        'What lab are you from?' : 'Group',
        'What is your name?' : 'Name',
        'How many NEW tips/plungers did you use?' : 'DFly New Tips',
        'How many NEW reservoirs did you use?' : 'DFly New Reservoirs',
        'How many NEW MXOne tip arrays did you use?' : 'DFly New Mixers',
        'What kind of plates did you use?' : 'DFly Plate Type',
        'How many NEW plates did you use?' : 'DFly New Plates',
        # 'How many screens did you set up?' : 'DFly New Plates Set-up'
        }, inplace=True)
    df.index = pd.to_datetime(df.pop('Timestamp'))
    df = df[df.index.notnull()]
    mask = (df.index > start_date) & (df.index <= end_date)
    df = df.loc[mask]
    df['Group'] = df['Group'].map(lambda s: s.lower()) #make group name lowercase to standardize
    df['DFly New Tips'] = pd.to_numeric(df['DFly New Tips'], errors='coerce').fillna(0) # convert objects to numeric
    df['DFly New Reservoirs'] = pd.to_numeric(df['DFly New Reservoirs'], errors='coerce').fillna(0) # convert objects to numeric
    df['DFly New Mixers'] = pd.to_numeric(df['DFly New Mixers'], errors='coerce').fillna(0) # convert objects to numeric
    df['DFly New Plates'] = pd.to_numeric(df['DFly New Plates'], errors='coerce').fillna(0) # convert objects to numeric
    return df

# obtain PI and their use_types from Google drive spreadsheet
def getPITypes():
    wks_groups = gc.open_by_key('1Yom-H6j04TJ5_W1Ic2dlqKsVWrHQbdInQyq_Q7ZVnZA').sheet1
    row_data = wks_groups.get_all_values(include_tailing_empty=False,include_tailing_empty_rows=False)

    coreUsers = []
    associateUsers = []
    regUsers = []

    k = 1 #skip first row
    while (k < len(row_data)):
        pi, user_type = row_data[k]
        pi = pi.lower()
        if (user_type=='regular'):
            regUsers.append(pi)
        elif (user_type=='associate'):
            associateUsers.append(pi)
        else:
            coreUsers.append(pi)
        k +=1
    return [coreUsers,associateUsers,regUsers,coreUsers + associateUsers + regUsers]

def getRechargeConst():
    df_rechargeConst = gc.open_by_key('1d6GVWGwwrlh_lTKxVRI08xZSiE__Zieu3WWtwbmOMlE'
        ).worksheet_by_title('master').get_as_df()
    return df_rechargeConst

# obtain usage log data from Google drive forms/spreadsheets (mosquito, mosquitoLCP, dragonfly)
def getGDriveLogUsage(dates):
    df_mosquitoLCPLogRAW = gc.open_by_key('1MpwGvh6xlOb4mrn8BtJgs7Fux7hmlZRRdmhBUkhqRAY').sheet1.get_as_df()
    df_mosquitoLogRAW = gc.open_by_key('1demabrSE50t_euIpP3AhM8V64I3BQuaK1VRcNJCSJmA').sheet1.get_as_df()
    df_dragonflyLogRAW = gc.open_by_key('1JciEUj4dg1AZedcmi42InLIQs5XINQt5aaok4-vnUwg').sheet1.get_as_df()
    df_screenOrders = gc.open_by_key('1d6GVWGwwrlh_lTKxVRI08xZSiE__Zieu3WWtwbmOMlE'
        ).worksheet_by_title('completedOrders').get_as_df()


    df_mosquitoLCPLog = clean_DF_MosquitoLCP(df_mosquitoLCPLogRAW,dates)
    df_mosquitoLog = clean_DF_MosquitoCrystal(df_mosquitoLogRAW,dates)
    df_dragonflyLog = clean_DF_Dragonfly(df_dragonflyLogRAW,dates)
    df_screenOrdersLog = clean_DF_screenOrders(df_screenOrders,dates)

    return [df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog, df_screenOrdersLog]

def clean_DF_screenOrders(df, dates):
    start_date, end_date = dates
    df.rename(columns={
        'Lab' : 'Group',
        'Name' : 'Requested By',
        'SKU' : 'Item Name',
        'Price' : 'Unit Price',
        'Total price' : 'Screens Total Cost',
        }, inplace=True)
    df.index = pd.to_datetime(df.pop('Date'))
    df = df[df.index.notnull()]
    mask = (df.index > start_date) & (df.index <= end_date)
    # df = df.loc[start_date:end_date]
    df = df.loc[mask]
    
    df['Group'] = df['Group'].map(lambda s: s.lower()) #make group name lowercase to standardize
    return df

def getRockImagerUsage(dates):
    start_date, end_date = dates

    # Read from RockImagers usage log file
    while True:
        try:
            filename = askopenfilename(initialdir = 'C:/Users/loren/Google Drive/ljiang/xrayFacilityRecharge/equipmentLogs/RockImagerEventLogs',
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
    print("\nSelected:\n" + filename +"\n")
    # MAIN SCRIPT BELOW
    # Remove "groups" not in recharge system (administrators, etc.)
    df = df[df['Group'] != 'Administrators']
    # Look at 'Duration' column and extract numerical data inplace
    durStr = 'Rock Dur (min)'
    df.rename(columns={'Duration':durStr}, inplace=True)
    df[durStr] = df[durStr].str.replace(r'[^0-9.]','', regex=True).astype('float')

    # Look at 'Time' column and convert to datetime type
    df.index = pd.to_datetime(df.pop('Time')) # Change index to datetime index
    df = df[df.index.notnull()]
    mask = (df.index > start_date) & (df.index <= end_date)
    # df = df.loc[start_date:end_date]
    df = df.loc[mask]
    df['Group'] = df['Group'].astype(str).map(lambda s: s.lower()) #make group name lowercase to standardize
    return df

def getGL(dates):
    df = gc.open_by_key('1L610loj5s41wQFYFnzNafI2OeY9kxL0WuPwebtW5k3Q').sheet1.get_as_df()
    # df_sheet = gc.open_by_key('1L610loj5s41wQFYFnzNafI2OeY9kxL0WuPwebtW5k3Q').sheet1
    # import pdb;pdb.set_trace();
    # df = df_sheet.get_as_df(end=(df_sheet.rows, 4))
    df['Actual'] = pd.to_numeric(df['Actual'], errors='coerce').fillna(0)
    start_date, end_date = dates

    # categories of charges; should be lower case
    amortizedExpenses = ['voucher']
    voucherExceptions = ['airgas', 'vpl', 'vantage point logistics', 'cdw-government']
    monthlyExpenses = ['recharge']
    payroll = ['payroll']

    # checks if strings of lst1 is a substring of any string in lst2 and returns array of bool
    def substringInListOfStrings(x,lst):
        for s in lst:
            if (x.find(s) > -1):
                return True
        return False

    df.index = pd.to_datetime(df.pop('JrnlDate'))
    mask = (df.index >= start_date) & (df.index <= end_date)
    # df = df.loc[start_date:end_date]
    df = df.loc[mask]
    df['Recharge Category'] = df[['TrnsTyp', 'Description']].apply(lambda s: 'amortizedExpenses' 
        if substringInListOfStrings(str(s[0]).lower(), amortizedExpenses) and \
            not(substringInListOfStrings(str(s[1]).lower(), voucherExceptions)) \
            else ('payroll' if substringInListOfStrings(str(s[0]).lower(), payroll)
                else 'monthlyExpenses'), axis = 1)

    return df   


# Make the output files more Excel friendly
# assumes index in TimeStamp format categorized by month (level -0)
def makeExcelFriendly(df):
    df.index.set_names(['Month/Year'],[0],inplace=True)
    df.index.set_levels(df.index.levels[0].strftime('%m/%Y'),level=0,inplace=True)
    return df

#takes df and output path w/ specified fileExt
def exportToFile(df, pth, fileExt='.xlsx'):
    if not(fileExt == '.xlsx' or fileExt == '.csv'):
        print("fileExt needs to be '.csv' or '.xlsx'")
    else:
        if (fileExt == '.csv'):
            df.to_csv(pth+fileExt)
        else:
            df.to_excel(pth+fileExt,freeze_panes=(1,1)) 

#assume df_lst is a list multi-index df with 'Group' level = 1
def getFilesByPI(df_lst,pth_lst,PI_lst,direc):
    for df,pth in zip(df_lst,pth_lst):
        for PI in PI_lst:
            bools = df.index.get_level_values('Group')==PI
            if any(bools):
                a=df.loc[bools]
                pi_folder = direc+PI+'/'

                if not os.path.exists(pi_folder):
                    os.makedirs(pi_folder)

                exportToFile(makeExcelFriendly(a),pi_folder+pth)

#takes in currency as str and converts to float
def currencyToFloat(s):
    return np.float(s.replace('$', ''))

def getCellByRowCol(df, rowHeader, rowSelector, colSelector):
        return df.loc[df[rowHeader]==rowSelector][colSelector].to_numpy()[0]

def calculateRecharge(dfs,date_range):
    start_date,end_date = date_range[0],date_range[1]
    [df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog, df_RockImager_1, 
        df_RockImager_2, df_GL, df_screenOrders, df_rechargeConst] = dfs
    
    # rechargeDict = [
    #                 'Prealiquoted hanging drop commerical screen',
    #                 'Prealiquoted sitting drop commerical screen',
    #                 '96-well Greiner Hanging drop plate',
    #                 'MRC2 Sitting drop plate',
    #                 'HD Consumables',
    #                 'SD Consumables',
    #                 'Spool of mosquito tips 9 mm pitch',
    #                 'Pack of 100 dragonfly tips',
    #                 'Reservoir dragonfly',
    #                 'RockImager time',
    #                 'Mosquito time',
    #                 'Dragonfly time',
    #                 'Assoc use multiplier',
    #                 'Core use multiplier',
    #                 'Regular use multiplier',
    #                 'Assoc facility fee',
    #                 'Core facility fee',
    #                 'Regular facility fee']

    def findRechargeConst(rowSelector, colSelector):
        return np.float(df_rechargeConst.loc[df_rechargeConst['Item']==rowSelector][colSelector]\
                .values[0].replace('$', '',))

    # getCellByRowCol(df, rowHeader, rowSelector, colSelector)
    # coreMultl = currencyToFloat(getCellByRowCol(df_rechargeConst, "Item", "Core use multiplier", "Price"))
    # coreFacilityFee = currencyToFloat(getCellByRowCol(df_rechargeConst, "Item",'Core facility fee','Price')) #650
    # assocMult = currencyToFloat(getCellByRowCol(df_rechargeConst, "Item",'Assoc use multiplier','Price'))
    # assocFacilityFee = currencyToFloat(getCellByRowCol(df_rechargeConst, "Item",'Assoc facility fee','Price'))
    # regMult = currencyToFloat(getCellByRowCol(df_rechargeConst, "Item",'Regular use multiplier','Price'))
    # regFacilityFee = currencyToFloat(getCellByRowCol(df_rechargeConst, "Item",'Regular facility fee','Price'))

    coreMultl = findRechargeConst("Core use multiplier", "Price")
    coreFacilityFee = findRechargeConst('Core facility fee','Price') #650
    assocMult = findRechargeConst('Assoc use multiplier','Price')
    assocFacilityFee = findRechargeConst('Assoc facility fee','Price')
    regMult = findRechargeConst('Regular use multiplier','Price')
    regFacilityFee = findRechargeConst('Regular facility fee','Price')

    # Get data from facilitySuppliesPricing (https://docs.google.com/spreadsheets/d/1d6GVWGwwrlh_lTKxVRI08xZSiE__Zieu3WWtwbmOMlE/edit#gid=0)
    costHDP = findRechargeConst('96-well Greiner Hanging drop plate','Price/Qty') #greiner 96 well hanging drop plate
    costSDP = findRechargeConst('MRC2 Sitting drop plate','Price/Qty') #swis mrc2 96 well sitting drop plate

    costMosqTip = findRechargeConst('Spool of mosquito tips 9 mm pitch','Price/Qty')

    costMosquitoTime = findRechargeConst('Mosquito time','Price')
    costDragonflyTime = findRechargeConst('Dragonfly time','Price')
    costRockImagerTime = findRechargeConst('RockImager time','Price')
    costMXOne = findRechargeConst('MXone pin arrays','Price/Qty')
    costHDSConsume = findRechargeConst('HD Consumables','Price') #cost of consumables for mosquito hanging drop plate setup (optical seal, micro reservoirs)
    costSDSConsume = findRechargeConst('SD Consumables','Price') #cost of consumables for mosquito sitting drop plate setup (optical seal, micro reservoirs)

    costDFlyReservoir = findRechargeConst('Reservoir dragonfly','Price/Qty')
    costDFlyTip = findRechargeConst('Pack of 100 dragonfly tips','Price/Qty')

    # Extract relevant data from df_Dragonfly in given date range
    wantedDflyCol = ['Group','Name','DFly New Tips','DFly New Reservoirs','DFly New Mixers','DFly New Plates','DFly Plate Type']
    DFlyUsageMonthByPI = df_dragonflyLog.groupby([pd.Grouper(freq="M"),'Group'])
    DFlyUsageYearByPI = df_dragonflyLog.groupby([df_dragonflyLog.index.year,'Group'])

    itemizedDFlyMonthByPI = DFlyUsageMonthByPI.apply(lambda x: x.head(
        len(x.index)))[wantedDflyCol].loc[start_date:end_date]
    itemizedDFlyMonthByPI['Cost/plate'] = itemizedDFlyMonthByPI['DFly Plate Type'].apply(lambda x: costHDP if x.find('hanging') > 0 else costSDP) #NEEDS to be changed for price for specific plate type
    itemizedDFlyMonthByPI['Cost/tip'] = costDFlyTip
    itemizedDFlyMonthByPI['Cost/reservoir'] = costDFlyReservoir
    itemizedDFlyMonthByPI['Cost/mixer'] = costMXOne
    itemizedDFlyMonthByPI=itemizedDFlyMonthByPI.iloc[::-1] #reverse index (descending date)

    itemizedDFlyMonthByPI['DFly Total Cost'] = itemizedDFlyMonthByPI['Cost/plate']*itemizedDFlyMonthByPI['DFly New Plates'] \
    + itemizedDFlyMonthByPI['DFly New Mixers']*costMXOne + itemizedDFlyMonthByPI['DFly New Reservoirs']*costDFlyReservoir \
    + itemizedDFlyMonthByPI['DFly New Tips']*costDFlyTip

    itemizedDFlyMonthByPI = itemizedDFlyMonthByPI[['Group','Name','DFly New Tips','DFly New Reservoirs','DFly New Mixers',
    'DFly New Plates','DFly Plate Type','Cost/tip','Cost/reservoir','Cost/mixer','Cost/plate','DFly Total Cost']] #reorganize columns 
    # Extract relevant data from df_RockImager in the given date range
    wantedRockCol = ['User Name','Project','Experiment','Plate Type', 'Rock Dur (min)']

    rockImagerUsageMonthByPI_1 = df_RockImager_1.groupby([pd.Grouper(freq="M"),'Group'])
    rockImagerUsageYearByPI_1 = df_RockImager_1.groupby([df_RockImager_1.index.year,'Group'])
    itemizedRockMonthByPI_1 = rockImagerUsageMonthByPI_1.apply(lambda x: x.head(
        len(x.index)))[wantedRockCol].loc[start_date:end_date]

    itemizedRockMonthByPI_1 = itemizedRockMonthByPI_1[(itemizedRockMonthByPI_1['Rock Dur (min)'] > 0)] #remove rows if 'Dur (min)' <= 0
    itemizedRockMonthByPI_1['Cost/min'] = costRockImagerTime
    itemizedRockMonthByPI_1['RockImager Total Cost'] = costRockImagerTime * itemizedRockMonthByPI_1['Rock Dur (min)']

    itemizedRockMonthByPI_1=itemizedRockMonthByPI_1.iloc[::-1] #reverse index (descending date)

    rockImagerUsageMonthByPI_2 = df_RockImager_2.groupby([pd.Grouper(freq="M"),'Group'])
    rockImagerUsageYearByPI_2 = df_RockImager_2.groupby([df_RockImager_2.index.year,'Group'])
    itemizedRockMonthByPI_2 = rockImagerUsageMonthByPI_2.apply(lambda x: x.head(
        len(x.index)))[wantedRockCol].loc[start_date:end_date]

    itemizedRockMonthByPI_2 = itemizedRockMonthByPI_2[(itemizedRockMonthByPI_2['Rock Dur (min)'] > 0)] #remove rows if 'Dur (min)' <= 0
    itemizedRockMonthByPI_2['Cost/min'] = costRockImagerTime
    itemizedRockMonthByPI_2['RockImager Total Cost'] = costRockImagerTime * itemizedRockMonthByPI_2['Rock Dur (min)']

    itemizedRockMonthByPI_2=itemizedRockMonthByPI_2.iloc[::-1] #reverse index (descending date)

    itemizedRockMonthByPI = itemizedRockMonthByPI_1.add(itemizedRockMonthByPI_2)

    # Extract relevant data from df_mosquitoLog in the given date range
    wantedMosqCol = ['Name','NumHDP','NumSDP','Mosq Protocol Used','Mosq Tips','Mosq Dur (hr)']

    mosqUsageMonthByPI = df_mosquitoLog.groupby([pd.Grouper(freq="M"),'Group'])
    mosqCrystalUsageYearByPI = df_mosquitoLog.groupby([df_mosquitoLog.index.year,'Group'])
    itemizedMosqCryMonthByPI = mosqUsageMonthByPI.apply(lambda x: x.head(
        len(x.index)))[wantedMosqCol].loc[start_date:end_date]

    itemizedMosqCryMonthByPI['Cost/tip'] = costMosqTip

    itemizedMosqCryMonthByPI['Mosquito Total Cost'] = costMosqTip*itemizedMosqCryMonthByPI['Mosq Tips'] \
    + costHDSConsume*itemizedMosqCryMonthByPI['NumHDP'] + costSDSConsume*itemizedMosqCryMonthByPI['NumSDP'] #needs to be changed eventually 

    itemizedMosqCryMonthByPI=itemizedMosqCryMonthByPI.iloc[::-1] #reverse index (descending date)
    wantedSOCol = ['Group','Requested By','Item Name','Qty','Unit Price','Screens Total Cost']
  
    itemizedSOMonthByPI = df_screenOrders.groupby([pd.Grouper(freq="M"),'Group']).apply(lambda x: x.head(
        len(x.index)))[wantedSOCol].loc[start_date:end_date]
    itemizedSOMonthByPI = itemizedSOMonthByPI.iloc[::-1]

    users = []
    usersFee = []
    for user in allUsers:
        if user in coreUsers:
            users.append(user)
            usersFee.append(coreFacilityFee)
        if user in associateUsers:
            users.append(user)
            usersFee.append(assocFacilityFee)


    monthIndex = pd.date_range(start_date,end_date,freq='M')
    numMonths = len(monthIndex)
    lenUsers = len(users)
    users = users *numMonths
    usersFee = usersFee *numMonths
    ind = []
    for m in monthIndex:
        start = datetime(m.year, m.month, 1) #dummy date
        end = datetime(m.year,m.month,10) #dummy date

        td = end - start
        delta = td // len(users)

        for k in range(lenUsers):
            ind.append(start)
            start += delta

    df_facFee = pd.DataFrame(index = ind, columns = ['Group','Facility Fee'])
    df_facFee.index = pd.to_datetime(df_facFee.index)
    df_facFee['Group'] = users
    df_facFee['Facility Fee'] = usersFee
    df_facFee = df_facFee.groupby([pd.Grouper(freq="M"),'Group']).sum(level=[0,1],numeric_only=True)


    a = itemizedMosqCryMonthByPI.sum(level=[0,1],numeric_only=True)[['NumHDP','NumSDP','Mosq Tips','Mosquito Total Cost']]
    b_1 = itemizedRockMonthByPI_1.sum(level=[0,1],numeric_only=True)[['Rock Dur (min)','RockImager Total Cost']]
    b_2 = itemizedRockMonthByPI_2.sum(level=[0,1],numeric_only=True)[['Rock Dur (min)','RockImager Total Cost']]
    c = itemizedSOMonthByPI.sum(level=[0,1],numeric_only=True)['Screens Total Cost']
    d = itemizedDFlyMonthByPI.sum(level=[0,1],numeric_only=True)['DFly Total Cost']

    b = (b_1.add(b_2,fill_value=0))

    monthlyRechargeTotal = pd.concat([a,b,c,d,df_facFee], axis=1).fillna(0)
    monthlyRechargeTotal['Raw Usage'] = (monthlyRechargeTotal['NumSDP']+monthlyRechargeTotal['NumHDP'])*10 \
    + monthlyRechargeTotal['Rock Dur (min)']
     # + monthlyRechargeTotal['Rock Dur (min)'] + monthlyRechargeTotal['DFly New Plates Set-up']*25 #might need to redefine to better definition

    a1 = monthlyRechargeTotal.groupby(level=[0,1]).sum().groupby(level=0)

    rawUsagePercent = a1.apply(lambda x: x['Raw Usage'] / x['Raw Usage'].sum())
    monthlyRechargeTotal['Usage prop'] = pd.Series.ravel(rawUsagePercent)

    df_GL_monthlyExpenses=df_GL[df_GL['Recharge Category'] == 'monthlyExpenses'].groupby(
        [pd.Grouper(freq="M")]).sum().loc[start_date:end_date]
    
    df_GL_payroll = df_GL[df_GL['Recharge Category'] == 'payroll'].groupby(
        [pd.Grouper(freq="M")]).sum().loc[start_date:end_date]

    monthlyRechargeTotal['Use Multiplier'] = regMult

    # Set Use Multiplier column and payments for Core and Assoc users
    monthlyRechargeTotal.loc[((monthlyRechargeTotal.index.levels[0],
        coreUsers),'Use Multiplier')] = coreMultl

    monthlyRechargeTotal.loc[((monthlyRechargeTotal.index.levels[0],
        associateUsers),'Use Multiplier')] = assocMult

    lst_monthlyExpenses = []
    lst_payroll = []
    for index, row in monthlyRechargeTotal.iterrows():
        if (index[0] in df_GL_monthlyExpenses.index):
            # do the stuff
            lst_monthlyExpenses.append(row['Usage prop'] * df_GL_monthlyExpenses.loc[index[0]]['Actual'])
            sumMonthFacFee = monthlyRechargeTotal.loc[index[0]]['Facility Fee'].sum()
            diff = df_GL_payroll.loc[index[0]]['Actual'] - sumMonthFacFee
            if (diff <= 0): # if total facilty fees exceed payroll total, then charge 0 per lab
                diff = 0
            lst_payroll.append(row['Usage prop'] *diff)
        else:
            lst_monthlyExpenses.append(0)
            lst_payroll.append(0)
    monthlyRechargeTotal['Month Dist. Cost'] = lst_monthlyExpenses  #distributed costs include everything except pay-per-use consumables and base salary/benefits
    monthlyRechargeTotal['Payroll Cost'] = lst_payroll

    monthlyRechargeTotal['Total Charge'] = (monthlyRechargeTotal['Month Dist. Cost'] 
        + monthlyRechargeTotal['Mosquito Total Cost'] + monthlyRechargeTotal['RockImager Total Cost']
        + monthlyRechargeTotal['DFly Total Cost']) *monthlyRechargeTotal['Use Multiplier'] \
        + monthlyRechargeTotal['Screens Total Cost'] \
        + monthlyRechargeTotal['Payroll Cost'] + monthlyRechargeTotal['Facility Fee']

    outSummary = monthlyRechargeTotal[['Facility Fee','Screens Total Cost','Mosquito Total Cost','RockImager Total Cost', 'DFly Total Cost',
    'Raw Usage','Usage prop','Month Dist. Cost','Payroll Cost','Use Multiplier','Total Charge']]
    outSummary.index.set_names(names='Group',level=1,inplace=True)

    outSummary = outSummary.sort_index(ascending=False)
    daterange = str(start_date)[0:10] +'_TO_'+str(end_date)[0:10]
    fileOut = ['mosquitoUsage'+daterange,'4c_rockImagerUsage'+daterange,
            '20c_rockImagerUsage'+daterange, 'dragonflyUsage'+daterange,'screenOrders'+daterange]

    itemizedMosqCryMonthByPI.index.set_levels(itemizedMosqCryMonthByPI
        .index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)
    itemizedRockMonthByPI_1.index.set_levels(itemizedRockMonthByPI_1
        .index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)
    itemizedRockMonthByPI_2.index.set_levels(itemizedRockMonthByPI_2
        .index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)
    itemizedDFlyMonthByPI.index.set_levels(itemizedDFlyMonthByPI
        .index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)
    itemizedSOMonthByPI.index.set_levels(itemizedSOMonthByPI
        .index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)

    dfOut = [itemizedMosqCryMonthByPI, itemizedRockMonthByPI_1, 
            itemizedRockMonthByPI_2, itemizedDFlyMonthByPI,itemizedSOMonthByPI]

    return outSummary,fileOut,dfOut

if __name__ == '__main__':
    # Get start_date and end_date from user input
    start_date, end_date = '', ''
    while 1:
        try:    
            start_date = datetime.strptime(input('Enter Start date in the format yyyy-m-d: '), '%Y-%m-%d')
            end_date = datetime.strptime(input('Enter End date in the format yyyy-m-d: ')  + '-23-59-59'
                , '%Y-%m-%d-%H-%M-%S')
            if (start_date>end_date):
                raise ValueError
            break
        except ValueError as e:
            print('\nInvalid date range or wrong format. Please try again or ctrl+C and ENTER to exit.')


    dates = [str(start_date),str(end_date)]
    print('The dates selected are: ' + dates[0] + ' to ' + dates[1])
    gc = pyg.authorize(service_file='../msg-Recharge-24378e029f2d.json')
    coreUsers, associateUsers, regUsers, allUsers = getPITypes()
    df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog, df_screenOrders = getGDriveLogUsage(dates)
    df_RockImager_1 = getRockImagerUsage(dates)
    df_RockImager_2 = getRockImagerUsage(dates)
    # import pdb;pdb.set_trace();
    df_GL = getGL(dates)
    df_rechargeConst = getRechargeConst()
    # import pdb;pdb.set_trace();
    dfs_input = [df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog, df_RockImager_1,
                    df_RockImager_2, df_GL, df_screenOrders, df_rechargeConst]
    rechargeSummary, fileOut_lst, dfOut_lst = calculateRecharge(dfs_input,[start_date,end_date])
    directory = '../monthlyRechargesTemp/' + str(start_date)[0:10]+'_TO_'+str(end_date)[0:10] + '/'

    if not os.path.exists(directory):
        os.makedirs(directory)
        getFilesByPI(dfOut_lst,fileOut_lst,allUsers,directory)
        exportToFile(makeExcelFriendly(rechargeSummary),directory+'rechargeSummary' + 
            str(start_date)[0:10]+'_TO_'+str(end_date)[0:10])
