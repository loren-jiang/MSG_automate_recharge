import getScreenOrders, getRockImagerUsage, getMosquitoCrystalUsage, getGL#, getDragonyflyUsage
import pandas as pd
import sys
import matplotlib.pyplot as plt
from datetime import datetime


coreMultl = 1.0
assocMult = 1.5
regMult = 2.0

coreUsers = ['Chou','DeGrado','Fraser','Jura','Manglik','Minor','Rosenberg','Stroud']
associateUsers = ['Agard','Craik','Gross','Narlikar','Shokat','Taunton','Tolani','Wells']


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

# Collate monthly recharge data from different files

# dfScreens = getScreenOrders.main()
dfRockImager = getRockImagerUsage.main()
dfMosquitoCrystal = getMosquitoCrystalUsage.main()
testdfGL = getGL.main()
dfGL = getGL.main().loc[start_date:end_date]
# dfMosquitoLCP = getMosquitoLCPUsage.main()
# dfDragonfly = getDragonyflyUsage.main()

# Extract relevant data from dfRockImager in the given date range
wantedRockCol = ['User Name','Project','Experiment','Plate Type', 'Dur (min)']

rockImagerUsageMonthByPI = dfRockImager.groupby([pd.Grouper(freq="M"),'Group'])
rockImagerUsageYearByPI = dfRockImager.groupby([dfRockImager.index.year,'Group'])
itemizedRockMonthByPI = rockImagerUsageMonthByPI.apply(lambda x: x.head(
    len(x.index)))[wantedRockCol].loc[start_date:end_date]

itemizedRockMonthByPI = itemizedRockMonthByPI[(itemizedRockMonthByPI['Dur (min)'] > 0)] #remove rows if 'Dur (min)' <= 0
itemizedRockMonthByPI['Cost/min'] = costRockImagerTime
itemizedRockMonthByPI['RockImager Cost'] = costRockImagerTime * itemizedRockMonthByPI['Dur (min)']

itemizedRockMonthByPI=itemizedRockMonthByPI.iloc[::-1] #reverse index (descending date)

# Extract relevant data from dfMosquitoCrystal in the given date range
wantedMosqCol = ['Name','Plates','Protocol Used','Tips','Dur (hr)']

mosqCrystalUsageMonthByPI = dfMosquitoCrystal.groupby([pd.Grouper(freq="M"),'Group'])
mosqCrystalUsageYearByPI = dfMosquitoCrystal.groupby([dfMosquitoCrystal.index.year,'Group'])
itemizedMosqCryMonthByPI = mosqCrystalUsageMonthByPI.apply(lambda x: x.head(
    len(x.index)))[wantedMosqCol].loc[start_date:end_date]

#finds the type of protocol used (sitting or hanging) and assigns approp. cost/plate
itemizedMosqCryMonthByPI['Cost/plate'] = itemizedMosqCryMonthByPI['Protocol Used'].apply(lambda x: 
    costHDSConsume if x.find('sitting')>0 else costSDSConsume)

itemizedMosqCryMonthByPI['Cost/tip'] = costPerTip

itemizedMosqCryMonthByPI['Mosquito Cost'] = costPerTip*itemizedMosqCryMonthByPI['Tips'] \
+ itemizedMosqCryMonthByPI['Cost/plate']*itemizedMosqCryMonthByPI['Plates']

itemizedMosqCryMonthByPI=itemizedMosqCryMonthByPI.iloc[::-1] #reverse index (descending date)


# >>> itemizedMosqCryMonthByPI.sum(level=[0,1],numeric_only=True)
#>>> a.xs('Minor',level=1).xs('2018-10-31',level=0)
# a = mosqCrystalUsageMonthByPI.sum()['Tips']*costPerTip
# b = rockImagerUsageMonthByPI.sum()['Dur (min)']*costRockImagerTime
a = itemizedMosqCryMonthByPI.sum(level=[0,1],numeric_only=True)[['Tips','Mosquito Cost']]
b = itemizedRockMonthByPI.sum(level=[0,1],numeric_only=True)[['Dur (min)','RockImager Cost']]

monthlyRechargeTotal = pd.concat([a,b], axis=1).fillna(0)
monthlyRechargeTotal['Raw Usage'] = monthlyRechargeTotal['Tips'] + monthlyRechargeTotal['Dur (min)']

a1 = monthlyRechargeTotal.groupby(level=[0,1]).sum().groupby(level=0)

rawUsagePercent = a1.apply(lambda x: x['Raw Usage'] / x['Raw Usage'].sum())
monthlyRechargeTotal['Usage prop'] = pd.Series.ravel(rawUsagePercent)

# g = monthlyRechargeTotal.apply(lambda x: )
lst = []
for index, row in monthlyRechargeTotal.iterrows():
    lst.append(row['Usage prop'] * dfGL.loc[index[0]][' Actual'])

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

fileOut = ['monthlyRecharges/recharge'+str(start_date)[0:10]+'_TO_'+str(end_date)[0:10],
'monthlyRecharges/mosquitoUsage'+str(start_date)[0:10]+'_TO_'+str(end_date)[0:10],
'monthlyRecharges/rockImagerUsage'+str(start_date)[0:10] +'_TO_'+str(end_date)[0:10]]

dfOut = [outSummary, itemizedMosqCryMonthByPI, itemizedRockMonthByPI]

# monthlyRechargeTotal.to_csv('monthlyRecharges/recharge'+str(start_date)[0:10]
#   +'_TO_'+str(end_date)[0:10]+'.csv')
# monthlyRechargeTotal.to_excel('monthlyRecharges/recharge'+str(start_date)[0:10]
#   +'_TO_'+str(end_date)[0:10]+'.xlsx')

# #export to CSV files
# itemizedMosqCryMonthByPI.to_csv('monthlyRecharges/mosquitoUsage'+str(start_date)[0:10]
#   +'_TO_'+str(end_date)[0:10]+'.csv')

# itemizedRockMonthByPI.to_csv('monthlyRecharges/rockImagerUsage'+str(start_date)[0:10]
#   +'_TO_'+str(end_date)[0:10]+'.csv')

# #export to Excel files
# itemizedMosqCryMonthByPI.to_excel('monthlyRecharges/mosquitoUsage'+str(start_date)[0:10]
#   +'_TO_'+str(end_date)[0:10]+'.xlsx')

# itemizedRockMonthByPI.to_excel('monthlyRecharges/rockImagerUsage'+str(start_date)[0:10]
#   +'_TO_'+str(end_date)[0:10]+'.xlsx')