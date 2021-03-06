import sys
import os
import pandas as pd 
import pygsheets as pyg 
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from tkinter.filedialog import askopenfilename


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
    # Look at 'Time' column and convert to datetime type
    # df.index = pd.to_datetime(df.pop('What day did you use the Mosquito?')) # Change index to datetime index
    df.index = pd.to_datetime(df.pop('Timestamp'))
    df = df[df.index.notnull()]
    mask = (df.index > start_date) & (df.index <= end_date)
    # df = df.loc[start_date:end_date]
    df = df.loc[mask]
    df['Group'] = df['Group'].map(lambda s: s.lower()) #make group name lowercase to standardize
    df['Mosq Tips'] = pd.to_numeric(df['Mosq Tips'], errors='coerce').fillna(0) # convert objects to numeric
    df['NumSDP'] = pd.to_numeric(df['NumSDP'], errors='coerce').fillna(0) # convert objects to numeric
    df['NumHDP'] = pd.to_numeric(df['NumHDP'], errors='coerce').fillna(0) # convert objects to numeric
    df['Mosq Dur (hr)'] = pd.to_numeric(df['Mosq Dur (hr)'], errors='coerce').fillna(0) # convert objects to numeric
    
    # df.dropna(subset=['Group', 'Mosq Plates Set-up', 'Mosq Tips'], inplace=True) # Removes rows in the labels subset that are NaN
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
    # df = df.loc[start_date:end_date]
    df = df.loc[mask]
    df['Group'] = df['Group'].map(lambda s: s.lower()) #make group name lowercase to standardize
    # convert columns to numeric values as they should be
    df['Mosq LCP Plates Set-up'] = pd.to_numeric(df['Mosq LCP Plates Set-up'], errors='coerce')
    df['NumSDP'] = pd.to_numeric(df['NumSDP'], errors='coerce').fillna(0)
    df['NumHDP'] = pd.to_numeric(df['NumHDP'], errors='coerce').fillna(0)
    a = pd.to_numeric(df['Mosq Tips'].ix[:,0], errors='coerce').fillna(0)
    b = pd.to_numeric(df['Mosq Tips'].ix[:,1], errors='coerce').fillna(0)
    df['Cum Mosq Tips'] = a.add(b)

    # df.dropna(subset=['Group', 'Mosq Plates Set-up', 'Mosq Tips'], inplace=True) # Removes rows in the labels subset that are NaN
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
    # df = df.loc[start_date:end_date]
    df = df.loc[mask]
    df['Group'] = df['Group'].map(lambda s: s.lower()) #make group name lowercase to standardize
    df['DFly New Tips'] = pd.to_numeric(df['DFly New Tips'], errors='coerce').fillna(0) # convert objects to numeric
    df['DFly New Reservoirs'] = pd.to_numeric(df['DFly New Reservoirs'], errors='coerce').fillna(0) # convert objects to numeric
    df['DFly New Mixers'] = pd.to_numeric(df['DFly New Mixers'], errors='coerce').fillna(0) # convert objects to numeric
    df['DFly New Plates'] = pd.to_numeric(df['DFly New Plates'], errors='coerce').fillna(0) # convert objects to numeric
    # df['DFly New Plates Set-up'] = pd.to_numeric(df['DFly New Plates Set-up'], errors='coerce').fillna(value=0,inplace=True) # convert objects to numeric
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
        if (user_type=='regular'):
            regUsers.append(pi)
        elif (user_type=='associate'):
            associateUsers.append(pi)
        else:
            coreUsers.append(pi)
        k +=1
    return [coreUsers,associateUsers,regUsers,coreUsers + associateUsers + regUsers]

# obtain usage log data from Google drive forms/spreadsheets (mosquito, mosquitoLCP, dragonfly)
def getGDriveLogUsage(dates):
    df_mosquitoLCPLogRAW = gc.open_by_key('1MpwGvh6xlOb4mrn8BtJgs7Fux7hmlZRRdmhBUkhqRAY').sheet1.get_as_df()
    df_mosquitoLogRAW = gc.open_by_key('1demabrSE50t_euIpP3AhM8V64I3BQuaK1VRcNJCSJmA').sheet1.get_as_df()
    df_dragonflyLogRAW = gc.open_by_key('1JciEUj4dg1AZedcmi42InLIQs5XINQt5aaok4-vnUwg').sheet1.get_as_df()

    df_mosquitoLCPLog = clean_DF_MosquitoLCP(df_mosquitoLCPLogRAW,dates)
    df_mosquitoLog = clean_DF_MosquitoCrystal(df_mosquitoLogRAW,dates)
    df_dragonflyLog = clean_DF_Dragonfly(df_dragonflyLogRAW,dates)

    return [df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog]

def getScreenOrders(dates):
    start_date, end_date = dates

    def removeBadChar(x):
        if isinstance(x, str):
            x=x.replace('"','').replace('=','')
        return x

    while True:
        try:
            filename = askopenfilename(initialdir = 'C:/Users/loren/Documents/ucsf_sra/msg_recharge_automation/screenOrders',
                title = "Select screenOrders file (.csv or .tsv only)")
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
    # make 'Date Requested' the index
    df = df.applymap(removeBadChar) #removes unwanted " and = character
    df.rename(removeBadChar,axis='columns',inplace=True)
    df.rename(index=str, columns={"Spend Tracking Code": "Group",
        "Total Price":"Screens Total Cost"}
        ,inplace=True)
    df.index = pd.to_datetime(df.pop('Date Requested')) # Change index to datetime index
    df = df[df.index.notnull()]
    mask = (df.index > start_date) & (df.index <= end_date)
    # df = df.loc[start_date:end_date]
    df = df.loc[mask]
    df['Group'] = df['Group'].map(lambda s: s.lower()) #make group name lowercase to standardize
    df.sort_values(['Group'],inplace=True)
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
    return df

def getGL(dates):
    df = gc.open_by_key('1L610loj5s41wQFYFnzNafI2OeY9kxL0WuPwebtW5k3Q').sheet1.get_as_df()

    start_date, end_date = dates

    # categories of charges; should be lower case
    amortizedExpenses = ['voucher']
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
    df['Recharge Category'] = df['TrnsTyp'].apply(lambda s: 'amortizedExpenses' 
        if substringInListOfStrings(str(s).lower(), amortizedExpenses) else ('payroll' 
            if substringInListOfStrings(str(s).lower(), payroll) else 'monthlyExpenses'))

    return df   


# Make the output files more Excel friendly
# assumes index in TimeStamp format categorized by month (level -0)
def makeExcelFriendly(df):
    df.index.set_names(['Month/Year'],[0],inplace=True)
    df.index.set_levels(df.index.levels[0].strftime('%m/%Y'),level=0,inplace=True)
    # j = 0
    # while (j < len(df.index.levels)):
    #     if df.index.levels[j].dtype == 'datetime64[ns]':
    #         df.index.set_levels(df.index.levels[j]
    #             .strftime('%m/%d/%Y'),level=j,inplace=True)
    #     j+=1
    return df

#takes df and output path w/ specified fileExt
def exportToFile(df, pth, fileExt='.xlsx'):
    if not(fileExt == '.xlsx' or fileExt == '.csv'):
        print("fileExt needs to be '.csv' or '.xlsx'")
    else:
        if (fileExt == '.csv'):
            df.to_csv(pth+fileExt,freeze_panes=(1,1))
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

def calculateRecharge(dfs,date_range):
    start_date,end_date = date_range[0],date_range[1]
    df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog, df_RockImager_1, df_RockImager_2, df_GL, df_ScreenOrders = dfs
    coreMultl = 1.0
    coreFacilityFee = 650

    assocMult = 1.5
    assocFacilityFee = 175

    regMult = 2.0
    # Get data from facilitySuppliesPricing (https://docs.google.com/spreadsheets/d/1d6GVWGwwrlh_lTKxVRI08xZSiE__Zieu3WWtwbmOMlE/edit#gid=0)
    costHDP = 1.16 #greiner 96 well hanging drop plate
    costSDP = 9.20 #swis mrc2 96 well sitting drop plate

    costHDScreen = 25.00
    costSDScreen = 25.00
    costMosqTip = 0.07

    costMosquitoTime = 0.00 #per hr
    costDragonflyTime = 0.00 #per hr
    costRockImagerTime = 5.00 / 60.0 #per min

    costHDSConsume = 7.50 #cost of consumables for mosquito hanging drop plate setup (optical seal, micro reservoirs)
    costSDSConsume = 3.89 #cost of consumables for mosquito sitting drop plate setup (optical seal, micro reservoirs)

    costMXOne = 0
    costDFlyReservoir = 1.90
    costDFlyTip = 2.50

    # Extract relevant data from df_Dragonfly in given date range
    wantedDflyCol = ['Group','Name','DFly New Tips','DFly New Reservoirs','DFly New Mixers','DFly New Plates','DFly Plate Type']
    # df_dragonflyLog = df_dragonflyLog[wantedDflyCol]
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

    #finds the type of protocol used (sitting or hanging) and assigns approp. cost/plate
    # itemizedMosqCryMonthByPI['Cost/plate'] = itemizedMosqCryMonthByPI['Mosq Protocol Used'].apply(lambda x: 
    #     costHDSConsume if x.find('sitting')>0 else costSDSConsume)

    itemizedMosqCryMonthByPI['Cost/tip'] = costMosqTip

    itemizedMosqCryMonthByPI['Mosquito Total Cost'] = costMosqTip*itemizedMosqCryMonthByPI['Mosq Tips'] \
    + costHDSConsume*itemizedMosqCryMonthByPI['NumHDP'] + costSDSConsume*itemizedMosqCryMonthByPI['NumSDP'] #needs to be changed eventually 

    itemizedMosqCryMonthByPI=itemizedMosqCryMonthByPI.iloc[::-1] #reverse index (descending date)

    wantedSOCol = ['Group','Requested By','Item Name','Qty','Vendor','Unit Price','Screens Total Cost']
    
    itemizedSOMonthByPI = df_ScreenOrders.groupby([pd.Grouper(freq="M"),'Group']).apply(lambda x: x.head(
        len(x.index)))[wantedSOCol].loc[start_date:end_date]
    itemizedSOMonthByPI = itemizedSOMonthByPI.iloc[::-1]

    a = itemizedMosqCryMonthByPI.sum(level=[0,1],numeric_only=True)[['NumHDP','NumSDP','Mosq Tips','Mosquito Total Cost']]
    b_1 = itemizedRockMonthByPI_1.sum(level=[0,1],numeric_only=True)[['Rock Dur (min)','RockImager Total Cost']]
    b_2 = itemizedRockMonthByPI_2.sum(level=[0,1],numeric_only=True)[['Rock Dur (min)','RockImager Total Cost']]
    c = itemizedSOMonthByPI.sum(level=[0,1],numeric_only=True)['Screens Total Cost']
    d = itemizedDFlyMonthByPI.sum(level=[0,1],numeric_only=True)[['DFly Total Cost']]

    b = (b_1.add(b_2,fill_value=0))

    monthlyRechargeTotal = pd.concat([a,b,c,d], axis=1).fillna(0)
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

    lst_monthlyExpenses = []
    lst_payroll = []
    for index, row in monthlyRechargeTotal.iterrows():
        lst_monthlyExpenses.append(row['Usage prop'] * df_GL_monthlyExpenses.loc[index[0]]['Actual'])

    for index, row in monthlyRechargeTotal.iterrows():
        lst_payroll.append(row['Usage prop'] * df_GL_payroll.loc[index[0]]['Actual'])

    monthlyRechargeTotal['Month Dist. Cost'] = lst_monthlyExpenses  #distributed costs include everything except pay-per-use consumables and base salary/benefits
    monthlyRechargeTotal['Payroll Cost'] = lst_payroll #payroll costs for that month
    monthlyRechargeTotal['Use Multiplier'] = regMult
    monthlyRechargeTotal['Facility Fee'] = 0
    
    # Set Use Multiplier column and payments for Core and Assoc users
    monthlyRechargeTotal.loc[((monthlyRechargeTotal.index.levels[0],
        coreUsers),'Use Multiplier')] = coreMultl
    monthlyRechargeTotal.loc[((monthlyRechargeTotal.index.levels[0],
        coreUsers),'Facility Fee')] = coreFacilityFee

    monthlyRechargeTotal.loc[((monthlyRechargeTotal.index.levels[0],
        associateUsers),'Use Multiplier')] = assocMult
    monthlyRechargeTotal.loc[((monthlyRechargeTotal.index.levels[0],
        associateUsers),'Facility Fee')] = assocFacilityFee

    monthlyRechargeTotal['Total Charge'] = (monthlyRechargeTotal['Month Dist. Cost'] 
        + monthlyRechargeTotal['Mosquito Total Cost'] + monthlyRechargeTotal['RockImager Total Cost']
        + monthlyRechargeTotal['DFly Total Cost'])\
        *monthlyRechargeTotal['Use Multiplier'] + monthlyRechargeTotal['Screens Total Cost'] \
        + monthlyRechargeTotal['Payroll Cost'] + monthlyRechargeTotal['Facility Fee']

    outSummary = monthlyRechargeTotal[['Facility Fee','Screens Total Cost','Mosquito Total Cost','RockImager Total Cost', 'DFly Total Cost',
    'Raw Usage','Usage prop','Month Dist. Cost','Payroll Cost','Use Multiplier','Total Charge']]
    outSummary.index.set_names(names='Group',level=1,inplace=True)

    outSummary = outSummary.sort_index(ascending=False)
    daterange = str(start_date)[0:10] +'_TO_'+str(end_date)[0:10]
    fileOut = ['mosquitoUsage'+daterange,'rockImagerUsage'+daterange,
    'dragonflyUsage'+daterange,'screenOrders'+daterange]

    itemizedMosqCryMonthByPI.index.set_levels(itemizedMosqCryMonthByPI
        .index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)
    itemizedRockMonthByPI.index.set_levels(itemizedRockMonthByPI
        .index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)
    itemizedDFlyMonthByPI.index.set_levels(itemizedDFlyMonthByPI
        .index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)
    itemizedSOMonthByPI.index.set_levels(itemizedSOMonthByPI
        .index.levels[2].strftime('%m/%d/%Y %H:%M:%S'),level=2,inplace=True)

    dfOut = [itemizedMosqCryMonthByPI, itemizedRockMonthByPI, itemizedDFlyMonthByPI,itemizedSOMonthByPI]

    return outSummary,fileOut,dfOut

if __name__ == '__main__':

    # Get start_date and end_date from user input
    start_date, end_date = '', ''
    while 1:
        try:    
            start_date = datetime.strptime(input('Enter Start date in the format yyyy-m-d: '), '%Y-%m-%d')
            end_date = datetime.strptime(input('Enter End date in the format yyyy-m-d: '), '%Y-%m-%d')
            if (start_date>end_date):
                raise ValueError
            break
        except ValueError as e:
            print('\nInvalid date range or wrong format. Please try again or ctrl+C and ENTER to exit.')


    dates = [str(start_date),str(end_date)]
    print('The dates selected are: ' + dates[0] + ' to ' + dates[1])
    gc = pyg.authorize(service_file='msg-Recharge-24378e029f2d.json')
    coreUsers, associateUsers, regUsers, allUsers = getPITypes()
    df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog = getGDriveLogUsage(dates)
    df_RockImager_1 = getRockImagerUsage(dates)
    df_RockImager_2 = getRockImagerUsage(dates)

    df_GL = getGL(dates)
    df_ScreenOrders = getScreenOrders(dates)
    dfs_input = [df_mosquitoLog, df_mosquitoLCPLog, df_dragonflyLog, df_RockImager_1,df_RockImager_2, df_GL, df_ScreenOrders]
    rechargeSummary, fileOut_lst, dfOut_lst = calculateRecharge(dfs_input,[start_date,end_date])
    #Path to be added eventually --> C:\Users\loren\Google Drive\ljiang\xrayFacilityRecharge\monthlyRecharges 
    directory = 'monthlyRechargesTemp/' + str(start_date)[0:10]+'_TO_'+str(end_date)[0:10] + '/'
    rechargeSummary.to_pickle('testPickle.pkl')
    if not os.path.exists(directory):
        os.makedirs(directory)
        getFilesByPI(dfOut_lst,fileOut_lst,allUsers,directory)
        exportToFile(makeExcelFriendly(rechargeSummary),directory+'rechargeSummary' + 
            str(start_date)[0:10]+'_TO_'+str(end_date)[0:10])