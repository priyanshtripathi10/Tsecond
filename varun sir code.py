import re 
from datetime import datetime

with open('b827eb6a07fb.txt') as fd:
   data = fd.read().split()

list1 = []
new_list=[]
for words in data:
   if words[0:7] == 'AHeader':
       list1.append(words)

for string in list1:
   list1 = re.split('[=&]',string) # Seperated with & which is a seperator to identify 

   d = iter(list1)
   d = dict(zip(d,d))

   # Use above Two lines or below line 

   '''
   d = dict( [ (k, v) for k,v in zip (list1[::2], list1[1::2]) ] )
   '''
   #print ('\n')

   for key, value in d.items():
       if key == 'dateTime':
           obj = datetime.strptime(value, '%Y/%m/%d_%H:%M:%S')
           d[key] = obj
   
   ##### Print or Write it into file
   print (d)
   ##print(type(d))

   new_list.append(d)
#print(new_list)


from csv import DictWriter
with open("output_excel.csv", "w") as outfile:
    writer = DictWriter(outfile, ('CB', 'dateTime', 'AHeader', 'EB', 'VARR', 'PFR', 'VAR', 'PT', 'EB_OnTime', 'FREQ', 'DG_OnTime', 'SS', 'WR', 'CT', 'DG', 'EB_OffTime', 'WH', 'VB', 'VARY', 'VAB', 'VARB', 'ASmartDG_id', 'LAT', 'LH', 'PFY', 'VLNA', 'VY', 'LON', 'Event_Type', 'VBR', 'VLLA', 'WY', 'BB_V', 'PFA', 'WB', 'VART', 'CY', 'DLAT', 'WT', 'DLON', 'PFB', 'DG_V', 'volume', 'CR', 'VR', 'AT_V', 'FF_DateTime', 'VAY', 'DG_OffTime', 'VYB', 'VRY', 'VAT'))
    writer.writeheader()
    writer.writerows(new_list)
   

