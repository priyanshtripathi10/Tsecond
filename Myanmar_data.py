from csv import DictWriter
import re


print('')
print('')
print('Myanmar Datalog Code Snippet........!')
print('')
print('')
print('Data Analysis taking place...******')
print('')
data = open("igtl_logs_em_till_date.txt", "r").read()                                            #Load the raw txt file here
new_list = []
re1='\{.*\}'	
rg = re.compile(re1)
m = rg.findall(data)
for x in m:
    new_list.append(eval(x))

##print(new_list)

with open("output_excel.csv", "w") as outfile:                                                   #csv file getting produced from raw txt file
    writer = DictWriter(outfile, ('VART', 'Amps_I1', 'Amps_I2', 'Amps_I3', 'Power_Factor1', 'Power_Factor2', 'Power_Factor3', 'VARY', 'WH', 'VYB', 'VBR', 'WT', 'Genset_HZ', 'VARB', 'KVA1', 'KVA2', 'KVA3', 'AT_V', 'Genset_V2', 'Genset_V3', 'Genset_V1', 'BB', 'FL', 'VRY', 'VARR', 'LH', 'VLLA', 'Date', 'Average_AMPS', 'time', 'KW1', 'KW3', 'KW2', 'ET', 'ID', 'VAT', 'PFA', 'FF', 'SS', 'VLNA', 'Battery_Voltage', 'DG_AC', 'EB_AC','Engine_Hour', 'Engine_state', 'Coolant_Temp', 'Oil_pressure', 'Starting_Mode', 'Engine_RPMs', 'Genset_Alert', 'Avgerage_AMPS', 'oil_Temp','Oil_Temp'))
    writer.writeheader()                                                                         #Intializing the columns needed ie VART, Amps_I1, Amps_I2
    writer.writerows(new_list)



import pandas as pd
from pandas import Series,DataFrame
df=pd.read_csv('/home/priyanshtripathi10/Documents/output_excel.csv')                       


et1=df.loc[df['ET']==1]                                                                            #Segregating data based on ET=1     #Module Online
et2=df.loc[df['ET']==2]                                                                            #Segregating data based on ET=2     #Fuel Filling
et3=df.loc[df['ET']==3]                                                                            #Segregating data based on ET=3     #DG Staus
et4=df.loc[df['ET']==4]                                                                            #Segregating data based on ET=4     #EB Status
et5=df.loc[df['ET']==5]                                                                            #Segregating data based on ET=5     #Sensor Status
et6=df.loc[df['ET']==6]                                                                            #Segregating data based on ET=6     #Energy Parameters


et1=et1[['AT_V','BB','Date','time','ID','Battery_Voltage','DG_AC','EB_AC']]      
et1.to_csv('/home/priyanshtripathi10/Downloads/Tsecond/et1.csv')                                   #Converting et1 to et1.csv

et2=et2[['Date','time','ID','FF']]
et2.to_csv('/home/priyanshtripathi10/Downloads/Tsecond/et2.csv')                                   #Converting et2 to et2.csv

et3=et3[['Date','time','ID','DG_AC']]
et3.to_csv('/home/priyanshtripathi10/Downloads/Tsecond/et3.csv')                                   #Converting et3 to et3.csv

et4=et4[['Date','time','ID','EB_AC']]
et4.to_csv('/home/priyanshtripathi10/Downloads/Tsecond/et4.csv')                                   #Converting et4 to et4.csv

et5=et5[['Date','time','ID','SS']]
et5.to_csv('/home/priyanshtripathi10/Downloads/Tsecond/et5.csv')                                   #Converting et5 to et5.csv

et6=et6[['Date','time','ID','VART','Amps_I1','Amps_I2','Amps_I2','Amps_I3','Power_Factor1','Power_Factor2','Power_Factor3','VARY','WH','VYB','VBR','WT','Genset_HZ','VARB','KVA1','KVA2','KVA3','Genset_V1','Genset_V2','Genset_V3','VRY','VARR','LH','VLLA','Average_AMPS','KW1','KW3','KW2','VAT','PFA']]
et6.to_csv('/home/priyanshtripathi10/Downloads/Tsecond/et6.csv')                                   #Converting et6 to et6.csv





#Calculating total Fuel Filled
et2=et2[['ID','Date','time','FF']]
Total_FF = et2['FF'].sum()
et2['Total']=Total_FF
et2.to_csv('/home/priyanshtripathi10/Downloads/customer_et2.csv')
print ('Total Fuel filled: '+str(Total_FF))                                                        #Summing up total fuel filling






#Loading data from et3.csv to calculate total DG run hours
df3=pd.read_csv('/home/priyanshtripathi10/Downloads/Tsecond/et3.csv')                              #Reading data from et3.csv
df3.drop('Unnamed: 0', axis=1, inplace=True)

alpha=df3.Date.unique()
alpha
for i in range(len(alpha)):
    df3 = df3.append({'Date':alpha[i],'time':'23:55:55','ID':'b827ebf18950','DG_AC':1},ignore_index=True)
df3['DateTime']=df3['Date'] + ' ' + df3['time']
df3['DateTime'] = df3['DateTime'].astype('datetime64[ns]')
df3=df3.sort_values(by='DateTime')
df3=df3.reset_index()
del df3['index']
del df3['Date']
del df3['time']

list2=[]
list3=[]
list4=[]
a=10
for i in range(len(df3)):
    if(df3['DG_AC'][i]==0 and a==10):
        list2.append(df3['DateTime'][i])
        list3.append(df3['DG_AC'][i])
        list4.append(df3['ID'][i])
        a=20
    elif(df3['DG_AC'][i]==1 and a==20):
        list2.append(df3['DateTime'][i])
        list3.append(df3['DG_AC'][i])
        list4.append(df3['ID'][i])
        a=10
dfet3 = pd.DataFrame({'ID':list4,'DG_DateTime':list2,'DG':list3})
dfet3_a1=dfet3.loc[dfet3['DG'] ==0]                                                                 #Segregating rows with DG=0
dfet3_a2=dfet3.loc[dfet3['DG'] ==1]                                                                 #Segregating rows with DG=0
dfet3_a2 = dfet3_a2.reset_index()
del dfet3_a2['index']
dfet3_a1 = dfet3_a1.reset_index()
del dfet3_a1['index']
dfet3_a1['DG_DateTime'] = dfet3_a1['DG_DateTime'].astype('datetime64[ns]')
dfet3_a2['DG_DateTime'] = dfet3_a2['DG_DateTime'].astype('datetime64[ns]')
del dfet3_a1['DG']
del dfet3_a2['DG']
dfet3_a2['Total_Run_Hour']=dfet3_a2['DG_DateTime']-dfet3_a1['DG_DateTime']
del dfet3_a2['ID']
dfet3_a1=dfet3_a1.rename(columns = {'DG_DateTime':'DG_On_time'})
dfet3_a2=dfet3_a2.rename(columns = {'DG_DateTime':'DG_Off_time'})
bigdata3 = pd.concat([dfet3_a1,dfet3_a2],axis=1)
bigdata3.to_csv('/home/priyanshtripathi10/Downloads/Tsecond/customer_et3.csv')

Total3=bigdata3['Total_Run_Hour'].sum()                                                             #Summing up total DG run hours
print('Total DG run hours: '+str(Total3))




#Loading data from et4.csv to calculate total EB run hours
df4=pd.read_csv('/home/priyanshtripathi10/Downloads/Tsecond/et4.csv')                               #Reading data from et4.csv 
df4.drop('Unnamed: 0', axis=1, inplace=True)

DF_obj = None
if df4['EB_AC'][0]==0 and df4['EB_AC'][len(df4)-1]==0:
        DF_obj= DataFrame({'Date':[df4['Date'][0]],'time':['23:55:55'],'ID':['b827ebf18950'],'EB_AC':[1]}) 
df4=pd.concat([df4,DF_obj])
df4=df4.reset_index()
del df4['index']

alpha=df4.Date.unique()
for i in range(len(alpha)):
    df4 = df4.append({'Date':alpha[i],'time':'23:55:55','ID':'b827ebf18950','EB_AC':1},ignore_index=True)
df4['DateTime']=df4['Date'] + ' ' + df4['time']
df4['DateTime'] = df4['DateTime'].astype('datetime64[ns]')
df4=df4.sort_values(by='DateTime')
df4=df4.reset_index()
del df4['index']
del df4['Date']
del df4['time']

list_2=[]
list_3=[]
list_4=[]
a=10
for i in range(len(df4)):
    if(df4['EB_AC'][i]==0 and a==10):
        list_2.append(df4['DateTime'][i])
        list_3.append(df4['EB_AC'][i])
        list_4.append(df4['ID'][i])
        a=20
    elif(df4['EB_AC'][i]==1 and a==20):
        list_2.append(df4['DateTime'][i])
        list_3.append(df4['EB_AC'][i])
        list_4.append(df4['ID'][i])
        a=10

dfet4 = pd.DataFrame({'ID':list_4,'DG_DateTime':list_2,'EB':list_3})
dfet4['DG_DateTime'] = dfet4['DG_DateTime'].astype('datetime64[ns]')
dfet4=dfet4.reset_index()
del dfet4['index']
dfet4=dfet4.sort_values(by='DG_DateTime')
 
dfet4_a1=dfet4.loc[dfet4['EB'] ==0]                                                               #Segregating rows with EB=0                                              
dfet4_a2=dfet4.loc[dfet4['EB'] ==1]                                                               #Segregating rows with EB=1   
dfet4_a2 = dfet4_a2.reset_index()
del dfet4_a2['index']
dfet4_a1 = dfet4_a1.reset_index()
del dfet4_a1['index']
dfet4_a1['DG_DateTime'] = dfet4_a1['DG_DateTime'].astype('datetime64[ns]')
dfet4_a2['DG_DateTime'] = dfet4_a2['DG_DateTime'].astype('datetime64[ns]')
del dfet4_a1['EB']
del dfet4_a2['EB']
dfet4_a2['Total_Run_Hour']=dfet4_a2['DG_DateTime']-dfet4_a1['DG_DateTime']
del dfet4_a2['ID']
dfet4_a1=dfet4_a1.rename(columns = {'EB_DateTime':'EB_On_time'})
dfet4_a2=dfet4_a2.rename(columns = {'EB_DateTime':'EB_Off_time'})
bigdata4 = pd.concat([dfet4_a1,dfet4_a2],axis=1)
bigdata4.to_csv('/home/priyanshtripathi10/Downloads/Tsecond/customer_et4.csv')

Total4=bigdata4['Total_Run_Hour'].sum()                                                           #Summing up total EB run hours
print ('Total EB run hours: '+str(Total4))



#Loading data from et5.csv to calculate total SS hours
df5=pd.read_csv('/home/priyanshtripathi10/Downloads/Tsecond/et5.csv')                             #Reading data from et5.csv
df5.drop('Unnamed: 0', axis=1, inplace=True)

alpha=df5.Date.unique()
alpha
for i in range(len(alpha)):
    df5 = df5.append({'Date':alpha[i],'time':'23:55:55','ID':'b827ebf18950','SS':1},ignore_index=True)
df5['DateTime']=df5['Date'] + ' ' + df5['time']
df5['DateTime'] = df5['DateTime'].astype('datetime64[ns]')
df5=df5.sort_values(by='DateTime')
df5=df5.reset_index()
del df5['index']
del df5['Date']
del df5['time']

list__2=[]
list__3=[]
list__4=[]
a=10
for i in range(len(df5)):
    if(df5['SS'][i]==0 and a==10):
        list__2.append(df5['DateTime'][i])
        list__3.append(df5['SS'][i])
        list__4.append(df5['ID'][i])
        a=20
    elif(df5['SS'][i]==1 and a==20):
        list__2.append(df5['DateTime'][i])
        list__3.append(df5['SS'][i])
        list__4.append(df5['ID'][i])
        a=10

dfet5 = pd.DataFrame({'ID':list__4,'SS_DateTime':list__2,'SS':list__3})
dfet5_a1=dfet5.loc[dfet5['SS'] ==0]                                                              #Segregating rows with SS=0
dfet5_a2=dfet5.loc[dfet5['SS'] ==1]                                                              #Segregating rows with SS=1
dfet5_a2 = dfet5_a2.reset_index()
del dfet5_a2['index']
dfet5_a1 = dfet5_a1.reset_index()
del dfet5_a1['index']
dfet5_a1['SS_DateTime'] = dfet5_a1['SS_DateTime'].astype('datetime64[ns]')
dfet5_a2['SS_DateTime'] = dfet5_a2['SS_DateTime'].astype('datetime64[ns]')
del dfet5_a1['SS']
del dfet5_a2['SS']
dfet5_a2['Run_Hour']=dfet5_a2['SS_DateTime']-dfet5_a1['SS_DateTime']
del dfet5_a2['ID']
dfet5_a1=dfet5_a1.rename(columns = {'SS_DateTime':'SS_On_time'})
dfet5_a2=dfet5_a2.rename(columns = {'SS_DateTime':'SS_Off_time'})
bigdata5 = pd.concat([dfet5_a1,dfet5_a2],axis=1)
bigdata5.to_csv('/home/priyanshtripathi10/Downloads/Tsecond/customer_et5.csv')

Total_SS =bigdata5['Run_Hour'].sum()                                                             #Summing up total Sensor hours
print ('Sensor Status: '+str(Total_SS))





#Loading data from et6.csv to calculate Total power & fuel consumption
et6['WH'].max()
et6['WH'].min()
et6['Total_power_consumption']=et6['WH'].max()-et6['WH'].min()
print('Total_power_consumption: '+str(et6['Total_power_consumption'].mean())+' KWh')              #Total power consumption
print('Total_fuel_consumption: '+str(et6['Total_power_consumption'].mean()*0.4)+' litres')        #Total fuel consumption






print('')
print('')
print('Kindly wait for a moment............Do not Shutdown!!!!!!.............................')
print('Excel files are getting generated for above')
print('')
print('')



#Producing desired excel worksheet from the above data analysis
import xlsxwriter

writer = pd.ExcelWriter('grade.xlsx', engine='xlsxwriter')
et1.applymap(str).to_excel(writer, sheet_name='et1')
et2.applymap(str).to_excel(writer, sheet_name='et2')
et3.applymap(str).to_excel(writer, sheet_name='et3')
et4.applymap(str).to_excel(writer, sheet_name='et4')
et5.applymap(str).to_excel(writer, sheet_name='et5')
et6.applymap(str).to_excel(writer, sheet_name='et6')
bigdata3.applymap(str).to_excel(writer, sheet_name='DG')
bigdata4.applymap(str).to_excel(writer, sheet_name='EB')
bigdata5.applymap(str).to_excel(writer, sheet_name='SS')

workbook = writer.book




workbook = writer.book

header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'align': 'middle',
    'fg_color': '#FF6600',
    'border': 1})

worksheet1= workbook.add_worksheet('customer')
row=0
col=0
text ='Device ID: '+str(et1['ID'][2])+'\n\n'+'Total Fuel filled: '+str(Total_FF)+'\n\n'+'Total EB run hours: '+str(Total4)+'\n\n'+'Total DG run hours: '+str(Total3)+'\n\n'+'Sensor Status: '+str(Total_SS)+'\n\n'+'Total_power_consumption: '+str(et6['Total_power_consumption'].mean())+' Watt Hr'+'\n\n'+'Total_fuel_consumption: '+str(et6['Total_power_consumption'].mean()*0.4)+' litres'
options = {                                                                          #Printing text having total EB,DG,SS etc as output
    'width': 700,
    'height': 400,
    'gradient': {'colors': ['#DDEBCF',
                            '#9CB86E',
                            '#156B13']},
    #'align': {'horizontal': 'center'},
     'font': {#'bold': True
             'italic': True,
             'name': 'Arial',
             'color': 'black',
             'size': 20}
}
worksheet1.insert_textbox(row,col,text, options)
row += 10
col += 10


worksheet1 = writer.sheets['EB']                                                    #Writing sheet of EB run hours for client 
worksheet1.set_column('A:A', None, None, {'hidden': True})
for col_num, value in enumerate(bigdata4.columns.values):
    worksheet1.write(0, col_num + 1, value, header_format)
    worksheet1.set_column('B:B', 15)
    worksheet1.set_column('C:E', 20)


worksheet2 = writer.sheets['DG']                                                    #Writing sheet of DG run hours for client
worksheet2.set_column('A:A', None, None, {'hidden': True})
for col_num, value in enumerate(bigdata3.columns.values):
    worksheet2.write(0, col_num + 1, value, header_format)
    worksheet2.set_column('B:B', 15)
    worksheet2.set_column('C:E', 20)

worksheet3 = writer.sheets['et1']                                                   #Writing sheet with parameters of ET=1
worksheet3.set_column('A:A', None, None, {'hidden': True})
for col_num, value in enumerate(et1.columns.values):
    worksheet3.write(0, col_num + 1, value, header_format)
    worksheet3.set_column('D:G', 15)
    worksheet3.set_column('B:C', 10)

worksheet4 = writer.sheets['et2']                                                   #Writing sheet with prameters of ET=2
worksheet4.set_column('A:A', None, None, {'hidden': True})
for col_num, value in enumerate(et2.columns.values):
    worksheet4.write(0, col_num + 1, value, header_format)
    worksheet4.set_column('B:D', 15)
    worksheet4.set_column('E:F', 10)

worksheet5 = writer.sheets['et3']                                                   #Writing sheet with parameters of ET=3
worksheet5.set_column('A:A', None, None, {'hidden': True})
for col_num, value in enumerate(et3.columns.values):
    worksheet5.write(0, col_num + 1, value, header_format)
    worksheet5.set_column('B:D', 15)
    worksheet5.set_column('E:E', 10)

worksheet6 = writer.sheets['et4']                                                   #Writing sheet with parameters of ET=4
worksheet6.set_column('A:A', None, None, {'hidden': True})
for col_num, value in enumerate(et4.columns.values):
    worksheet6.write(0, col_num + 1, value, header_format)
    worksheet6.set_column('B:D', 15)
    worksheet6.set_column('E:E', 10)

worksheet7 = writer.sheets['et5']                                                   #Writing sheet with parameters of ET=5
worksheet7.set_column('A:A', None, None, {'hidden': True})
for col_num, value in enumerate(et5.columns.values):
    worksheet7.write(0, col_num + 1, value, header_format)
    worksheet7.set_column('B:D', 15)
    worksheet7.set_column('E:E', 25)
    

worksheet8 = writer.sheets['et6']                                                  #Writing sheet with parameters of ET=6
worksheet8.set_column('A:A', None, None, {'hidden': True})
for col_num, value in enumerate(et6.columns.values):
    worksheet8.write(0, col_num + 1, value, header_format)
    worksheet8.set_column('B:D', 15)
    worksheet8.set_column('E:I', 10)
    worksheet8.set_column('J:L',15)
    worksheet8.set_column('M:Y', 15)
    worksheet8.set_column('Z:AI', 15)
    worksheet8.set_column('AJ:AJ', 25)

worksheet9 = writer.sheets['SS']                                                   #Writing sheet of SS for client
worksheet9.set_column('A:A', None, None, {'hidden': True})
for col_num, value in enumerate(bigdata5.columns.values):
    worksheet9.write(0, col_num + 1, value, header_format)
    worksheet9.set_column('B:B', 15)
    worksheet9.set_column('C:E', 20)

    
    


writer.save()
print("\n\n\t\tAll coMmaNds ExeCutEd SucCessFulLy!!!!!!!!!!!")




























