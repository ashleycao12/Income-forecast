from numpy import append
import pandas as pd
import openpyxl
import datetime as dt
import calendar as cal

from pandas.tseries.offsets import BusinessDay
#### Import data ###
df = pd.read_excel("Contract_Details.xlsx", engine= 'openpyxl')
df['No. of Days Supplying'] = pd.to_numeric(df['No. of Days Supplying'],downcast='integer',errors='coerce')
df.fillna(0)


ctrt_start = df['Contract Start'].values
ctrt_end = df['Contract End'].values
costperday = df['Cost per day'].values
d_supplyperweek = df['No. of Days Supplying'].values


HolidayDf = pd.read_excel("Contract_Details.xlsx", sheet_name="Non-operational days", engine= 'openpyxl')
holidaylist = []
for i in range(len(HolidayDf)):
    holidaylist.append(HolidayDf['Non-operational days'][i])



year = 2021
firstmonth = 4 #April
monthnum = 12-firstmonth +1
bdays_inmonth = []


### Calculate ###
for i in range (firstmonth,13):
    firstdate = dt.datetime(year,i,1)
    datenum = cal.monthrange(year,i)[1] 
    lastdate = dt.datetime(year,i,datenum)
    bdays_inmonth.append(pd.bdate_range(firstdate,lastdate,holidays=holidaylist,freq='C'))



monthnamelist = list(cal.month_name)
monthlist = []  #get list of 12 months in the year
for i in range (firstmonth,len(monthnamelist)):
    monthlist.append(monthnamelist[i]+" "+ str(year))  #list of months to add in df



#create new column in df
for i in range(len(monthlist)):
    df[monthlist[i]]=pd.Series(dtype='int')


for irow in range(len(df)):  # i is for the month column index
    try:  #avoid blank row
        ctrtdate_list = pd.date_range(ctrt_start[irow],ctrt_end[irow])

        for imon in range(monthnum):  #j = row number 
            deliverydate_list = [x for x in ctrtdate_list if x in bdays_inmonth[imon]]
            deliverydatenum = len(deliverydate_list)
            df[monthlist[imon]].values[irow] = deliverydatenum*d_supplyperweek[irow]/5*costperday[irow]
    except ValueError:
        pass    

df2 = df.groupby(['Supplier'])[monthlist].sum().reset_index()



#Export result
writer =pd.ExcelWriter('Result.xlsx', engine='openpyxl')
df.to_excel(writer,sheet_name='Sheet1',index=False)
df2.to_excel(writer,sheet_name='Sheet2',index=False)
writer.save()







