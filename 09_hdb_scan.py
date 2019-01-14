import re 
from datetime import datetime
import numpy as np

for dat in range(22, 30):
    try:
        path = './HDB_Financial/b827eb6a07fb_' + str(dat) +'.txt'

        with open(path) as fd:
            data = fd.read().split()

        list1 = []
        new_list=[]
        for words in data:
           if words[0:7] == 'AHeader':
               list1.append(words)

        for string in list1:
           list1 = re.split('[=&]',string)

           d = iter(list1)
           d = dict(zip(d,d))


           #print ('\n')

           for key, value in d.items():
               if key == 'dateTime':
                   obj = datetime.strptime(value, '%Y/%m/%d_%H:%M:%S')
                   d[key] = obj
           
           

           new_list.append(d)



        from csv import DictWriter
        with open("output_excel" + str(dat) + ".csv", "w") as outfile:
            writer = DictWriter(outfile, ('CB', 'dateTime', 'AHeader', 'EB', 'VARR', 'PFR', 'VAR', 'PT', 'EB_OnTime', 'FREQ', 'DG_OnTime', 'SS', 'WR', 'CT', 'DG', 'EB_OffTime', 'WH', 'VB', 'VARY', 'VAB', 'VARB', 'ASmartDG_id', 'LAT', 'LH', 'PFY', 'VLNA', 'VY', 'LON', 'Event_Type', 'VBR', 'VLLA', 'WY', 'BB_V', 'PFA', 'WB', 'VART', 'CY', 'DLAT', 'WT', 'DLON', 'PFB', 'DG_V', 'volume', 'CR', 'VR', 'AT_V', 'FF_DateTime', 'VAY', 'DG_OffTime', 'VYB', 'VRY', 'VAT'))
            writer.writeheader()
            writer.writerows(new_list)



        import pandas as pd
        from pandas import Series,DataFrame
        df=pd.read_csv('output_excel' + str(dat) + '.csv')                       #Put the csv file address here  ie output_file.csv


        et1=df.loc[df['Event_Type']==1]                                                                     #Segregating data based on ET=1     #Module Online
        et2=df.loc[df['Event_Type']==2]                                                                     #Segregating data based on ET=2     #Fuel Filling
        et3=df.loc[df['Event_Type']==3]                                                                     #Segregating data based on ET=3     #DG Staus
        et4=df.loc[df['Event_Type']==4]                                                                     #Segregating data based on ET=4     #EB Status
        et5=df.loc[df['Event_Type']==5]                                                                     #Segregating data based on ET=5     #Sensor Status
        et6=df.loc[df['Event_Type']==6]
        et7=df.loc[df['Event_Type']==7] 




        et1=et1[['AT_V','BB_V','dateTime','ASmartDG_id','VB','DG','EB','DG_V']]
        et1.to_csv('./HDB_Financial/et1' + str(dat) + '.csv')  


        et2=et2[['dateTime','ASmartDG_id','volume']]
        et2.to_csv('./HDB_Financial/et2' + str(dat) + '.csv') 

        et3=et3[['dateTime','ASmartDG_id','DG']]
        et3.to_csv('./HDB_Financial/et3' + str(dat) + '.csv')

        et4=et4[['dateTime','ASmartDG_id','EB']]
        et4.to_csv('./HDB_Financial/et4' + str(dat) + '.csv')

        et5=et5[['dateTime','ASmartDG_id','PT']]
        et5.to_csv('./HDB_Financial/et5' + str(dat) + '.csv')  

        et6=et6[['dateTime','ASmartDG_id','SS']]
        et6.to_csv('./HDB_Financial/et6' + str(dat) + '.csv')


        et7=et7[['CB', 'VARR', 'PFR', 'VAR', 'FREQ', 'WR', 'CT', 'WH', 'VB', 'VARY', 'VAB', 'VARB', 'LH', 'PFY', 'VLNA', 'VY', 'VBR', 'VLLA', 'WY', 'PFA', 'WB', 'VART', 'CY', 'WT', 'PFB', 'CR', 'VR', 'VAY', 'VYB', 'VRY', 'VAT']]
        et7.to_csv('./HDB_Financial/et7' + str(dat) + '.csv')



        print('')
        print('')
        print('Code for Tech Mahindra!!!')
        print('')
        #Calculating total Fuel Filled
        et2=et2[['dateTime','ASmartDG_id','volume']]
        Total_FF = et2['volume'].sum()
        et2['Total']=Total_FF
        et2.to_csv('./HDB_Financial/customer_et2' + str(dat) + '.csv')
        print ('Total Fuel filled: '+str(Total_FF))





        #Loading data from et3.csv to calculate total DG run hours
        df3=pd.read_csv('./HDB_Financial/et3' + str(dat) + '.csv')                               #Loading data from et3.csv
        df3.drop('Unnamed: 0', axis=1, inplace=True)

        list2=[]
        list3=[]
        list4=[]
        a=10
        for i in range(len(et3)):
            if(df3['DG'][i]==0 and a==10):
                list2.append(df3['dateTime'][i])
                list3.append(df3['DG'][i])
                list4.append(df3['ASmartDG_id'][i])
                a=20
            elif(df3['DG'][i]==1 and a==20):
                list2.append(df3['dateTime'][i])
                list3.append(df3['DG'][i])
                list4.append(df3['ASmartDG_id'][i])
                a=10

        dfet3 = pd.DataFrame({'ID':list4,'DG_DateTime':list2,'DG':list3})
        dfet3_a1=dfet3.loc[dfet3['DG'] ==0]
        dfet3_a2=dfet3.loc[dfet3['DG'] ==1]
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
        bigdata3.to_csv('./HDB_Financial/customer_et3' + str(dat) + '.csv')

        Total3=bigdata3['Total_Run_Hour'].sum()                                                                    #Summing up total DG run hours
        print('Total DG run hours: '+str(Total3))




        #Loading data from et4.csv to calculate total EB run hours


        df4=pd.read_csv('./HDB_Financial/et4' + str(dat) + '.csv')                              #Loading data from et4.csv            
        df4.drop('Unnamed: 0', axis=1, inplace=True)

        DF_obj = None
        if df4['EB'][0]==0 and df4['EB'][len(df4)-1]==0:
           DF_obj= DataFrame({'dateTime':[df4['dateTime'][0][0:10]+' 23:55:55'],'ASmartDG_id':[df4['ASmartDG_id'][0]],'EB':[1]})
                 
           
           
           
        df4=pd.concat([df4,DF_obj]) 
        df4 = df4.reset_index()
        del df4['index']                   

        list_2=[]
        list_3=[]
        list_4=[]
        a=10
        for i in range(len(df4)):
            if(df4['EB'][i]==0 and a==10):
                list_2.append(df4['dateTime'][i])
                list_3.append(df4['EB'][i])
                list_4.append(df4['ASmartDG_id'][i])
                a=20
            elif(df4['EB'][i]==1 and a==20):
                list_2.append(df4['dateTime'][i])
                list_3.append(df4['EB'][i])
                list_4.append(df4['ASmartDG_id'][i])
                a=10
                
        dfet4 = pd.DataFrame({'ID':list_4,'DG_DateTime':list_2,'EB':list_3})
        dfet4_a1=dfet4.loc[dfet4['EB'] ==0]
        dfet4_a2=dfet4.loc[dfet4['EB'] ==1]
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
        dfet4_a1=dfet4_a1.rename(columns = {'DG_DateTime':'EB_On_time'})
        dfet4_a2=dfet4_a2.rename(columns = {'DG_DateTime':'EB_Off_time'})
        bigdata4 = pd.concat([dfet4_a1,dfet4_a2],axis=1)
        bigdata4.to_csv('./HDB_Financial/customer_et4' + str(dat) + '.csv')

        Total4=bigdata4['Total_Run_Hour'].sum()                                                                    #Summing up total DG run hours
        print ('Total EB run hours: '+str(Total4))





        #Loading data from et5.csv to calculate total PT run hours


        df5=pd.read_csv('./HDB_Financial/et5' + str(dat) + '.csv')                              #Loading data from et4.csv            
        df5.drop('Unnamed: 0', axis=1, inplace=True)

        list_pt1=[]
        list_pt2=[]
        list_pt3=[]
        a=10
        for i in range(len(et5)):
            if(df5['PT'][i]==1 and a==10):
                list_pt1.append(df5['dateTime'][i])
                list_pt2.append(df5['PT'][i])
                list_pt3.append(df5['ASmartDG_id'][i])
                a=20
            elif(df5['PT'][i]==0 and a==20):
                list_pt1.append(df5['dateTime'][i])
                list_pt2.append(df5['PT'][i])
                list_pt3.append(df5['ASmartDG_id'][i])
                a=10
                
        dfet5 = pd.DataFrame({'ID':list_pt3,'PT_DateTime':list_pt1,'PT':list_pt2})
        dfet5_a1=dfet5.loc[dfet5['PT'] ==0]
        dfet5_a2=dfet5.loc[dfet5['PT'] ==1]
        dfet5_a2 = dfet5_a2.reset_index()
        del dfet5_a2['index']
        dfet5_a1 = dfet5_a1.reset_index()
        del dfet5_a1['index']
        dfet5_a1['PT_DateTime'] = dfet5_a1['PT_DateTime'].astype('datetime64[ns]')
        dfet5_a2['PT_DateTime'] = dfet5_a2['PT_DateTime'].astype('datetime64[ns]')
        del dfet5_a1['PT']
        del dfet5_a2['PT']
        dfet5_a2['Total_Run_Hour']=dfet5_a1['PT_DateTime']-dfet5_a2['PT_DateTime']
        del dfet5_a2['ID']
        dfet5_a1=dfet5_a1.rename(columns = {'PT_DateTime':'PT_Off_time'})
        dfet5_a2=dfet5_a2.rename(columns = {'PT_DateTime':'PT_On_time'})
        bigdata5 = pd.concat([dfet5_a1,dfet5_a2],axis=1)
        bigdata5.to_csv('./HDB_Financial/customer_et5' + str(dat) + '.csv')

        Total5=bigdata5['Total_Run_Hour'].sum()                                                                    #Summing up total DG run hours
        print ('Total PT run hours: '+str(Total5))





        #Loading data from et6.csv to calculate Sensor Status
        df6=pd.read_csv('./HDB_Financial/et6' + str(dat) + '.csv')                              #Loading data from et5.csv
        df6.drop('Unnamed: 0', axis=1, inplace=True)

        list__2=[]
        list__3=[]
        list__4=[]
        a=10
        for i in range(len(et6)):
            if(df6['SS'][i]==0 and a==10):
                list__2.append(df6['dateTime'][i])
                list__3.append(df6['SS'][i])
                list__4.append(df6['ASmartDG_id'][i])
                a=20
            elif(df6['SS'][i]==1 and a==20):
                list__2.append(df6['dateTime'][i])
                list__3.append(df6['SS'][i])
                list__4.append(df6['ASmartDG_id'][i])
                a=10

        dfet6 = pd.DataFrame({'ID':list__4,'SS_DateTime':list__2,'SS':list__3})
        dfet6_a1=dfet6.loc[dfet6['SS'] ==0]
        dfet6_a2=dfet6.loc[dfet6['SS'] ==1]
        dfet6_a2 = dfet6_a2.reset_index()
        del dfet6_a2['index']
        dfet6_a1 = dfet6_a1.reset_index()
        del dfet6_a1['index']
        dfet6_a1['SS_DateTime'] = dfet6_a1['SS_DateTime'].astype('datetime64[ns]')
        dfet6_a2['SS_DateTime'] = dfet6_a2['SS_DateTime'].astype('datetime64[ns]')
        del dfet6_a1['SS']
        del dfet6_a2['SS']
        dfet6_a2['Total_Run_Hour']=dfet6_a2['SS_DateTime']-dfet6_a1['SS_DateTime']
        del dfet6_a2['ID']
        dfet6_a1=dfet6_a1.rename(columns = {'SS_DateTime':'SS_On_time'})
        dfet6_a2=dfet6_a2.rename(columns = {'SS_DateTime':'SS_Off_time'})
        bigdata6 = pd.concat([dfet6_a1,dfet6_a2],axis=1)
        bigdata6.to_csv('./HDB_Financial/customer_et6' + str(dat) + '.csv')

        Total_SS =bigdata6['Total_Run_Hour'].sum()                                                             #Summing up total Sensor hours
        print ('Sensor Status: '+str(Total_SS))



        #Loading data from et7.csv to calculate Total power & fuel consumption
        et7['WH'].max()
        et7['WH'].min()
        et7['Total_power_consumption']=et7['WH'].max()-et7['WH'].min()
        print('Total_power_consumption: '+str(et7['Total_power_consumption'].mean())+' Watt Hr')              #Total power consumption
        print('Total_fuel_consumption: '+str(et7['Total_power_consumption'].mean()*0.4)+' litres') 



        import xlsxwriter

        writer = pd.ExcelWriter('grade_'+ str(dat) + '.xlsx', engine='xlsxwriter')

        et1.applymap(str).to_excel(writer, sheet_name='et1')
        et2.applymap(str).to_excel(writer, sheet_name='et2')
        et3.applymap(str).to_excel(writer, sheet_name='et3')
        et4.applymap(str).to_excel(writer, sheet_name='et4')
        et5.applymap(str).to_excel(writer, sheet_name='et5')
        et6.applymap(str).to_excel(writer, sheet_name='et6')
        et7.applymap(str).to_excel(writer, sheet_name='et7')
        bigdata3.applymap(str).to_excel(writer, sheet_name='DG')
        bigdata4.applymap(str).to_excel(writer, sheet_name='EB')
        bigdata5.applymap(str).to_excel(writer, sheet_name='PT')
        bigdata6.applymap(str).to_excel(writer, sheet_name='SS')


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
        text ='Device ID: '+str(et1['ASmartDG_id'][2])+'\n\n'+'Total Fuel filled: '+str(Total_FF)+'\n\n'+'Total EB run hours: '+str(Total4)+'\n\n'+'Total DG run hours: '+str(Total3)+'\n\n'+'Total power tampering hours: '+str(Total5)+'\n\n'+'Sensor Status: '+str(Total_SS)+'\n\n'+'Total_power_consumption: '+str(et7['Total_power_consumption'].mean())+' Watt Hr'+'\n\n'+'Total_fuel_consumption: '+str(et7['Total_power_consumption'].mean()*0.4)+' litres'
        options = {
            'width': 700,
            'height': 500,
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


        worksheet2 = writer.sheets['DG']
        worksheet2.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(bigdata3.columns.values):
           worksheet2.write(0, col_num + 1, value, header_format)
           worksheet2.set_column('C:E', 20)
           worksheet2.set_column('B:B',15)


        worksheet20 = writer.sheets['EB']
        worksheet20.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(bigdata4.columns.values):
           worksheet20.write(0, col_num + 1, value, header_format)
           worksheet20.set_column('C:E', 20)
           worksheet20.set_column('B:B',15)
          

        worksheet30 = writer.sheets['PT']
        worksheet30.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(bigdata5.columns.values):
           worksheet30.write(0, col_num + 1, value, header_format)
           worksheet30.set_column('C:E', 20)
           worksheet30.set_column('B:B',15)



        worksheet40 = writer.sheets['SS']
        worksheet40.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(bigdata6.columns.values):
           worksheet40.write(0, col_num + 1, value, header_format)
           worksheet40.set_column('C:E', 20)
           worksheet40.set_column('B:B',15)

            

        worksheet3 = writer.sheets['et1']
        worksheet3.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(et1.columns.values):
            worksheet3.write(0, col_num + 1, value, header_format)
        worksheet3.set_column('B:C', 8)
        worksheet3.set_column('D:E', 20)
        worksheet3.set_column('F:I', 8)
            
            

        worksheet4 = writer.sheets['et2']
        worksheet4.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(et2.columns.values):
            worksheet4.write(0, col_num + 1, value, header_format)
        worksheet4.set_column('B:C', 20)
        worksheet4.set_column('D:D', 10)
        worksheet4.set_column('E:E', 20)
            


        worksheet5 = writer.sheets['et3']
        worksheet5.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(et3.columns.values):
            worksheet5.write(0, col_num + 1, value, header_format)
            worksheet5.set_column('B:C', 20)
            worksheet5.set_column('D:D', 10)
            
        worksheet6 = writer.sheets['et4']
        worksheet6.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(et4.columns.values):
            worksheet6.write(0, col_num + 1, value, header_format)
            worksheet6.set_column('B:C', 20)
            worksheet6.set_column('D:D', 10)


        worksheet7 = writer.sheets['et5']
        worksheet7.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(et5.columns.values):
            worksheet7.write(0, col_num + 1, value, header_format)
            worksheet7.set_column('B:C', 20)
            worksheet7.set_column('D:D', 10)

        worksheet8 = writer.sheets['et6']
        worksheet8.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(et6.columns.values):
            worksheet8.write(0, col_num + 1, value, header_format)
            worksheet8.set_column('B:C', 20)
            worksheet8.set_column('D:D', 10)


        worksheet9 = writer.sheets['et7']
        worksheet9.set_column('A:A', None, None, {'hidden': True})
        for col_num, value in enumerate(et7.columns.values):
            worksheet9.write(0, col_num + 1, value, header_format)
            worksheet9.set_column('B:AF', 10)
            worksheet9.set_column('AG:AG', 25)
            



        writer.save()
        print("\n\n\t\tAll coMmaNds ExeCutEd SucCessFulLy..!!!!!!!!!!!")
    except Exception as e:
        print (e)
