# Tsecond

This is the readme file consisting of the instructions regarding the use of python script devloped to perform the data visualization & analysis of the raw data getting generated from SmartDG located at various sites.The code is devloped keeping the data generated from Myanmar site,TechMahindra and Napino as input and can be easily manipulated with the upcoming formats.


The data getting generated from SmartDG consist of various parameters which are raw in nature and has to be organized in form of Excel Workbook or CSV file to easily read & draw conclusions from it.Below are the python scripts made to tackle the above problem and to automate the whole task.

The data getting generated from SmartDG located at various sites is in the form of continuous stream of text & characters which is getting stored in the form of text format for instance if the data is getting generated from the device located in Myanmar location it will be named as myanmar_datalog_2018_12_26.txt and if it is getting generated from TechMahindra it will be named as b827eb6a07fb_28.txt. This data can either be a continuous logs or can be segregiated in accordance with the dates as per the requirnment.

If the raw data is from the Myanmar site take use of Myanmar_data.py & if the data is getting generated from the TechMahindra site take use of Tech_Mahindra_data.py.
After opening the required python script insert the location of raw text file in the variable "data":

data = open("myanmar_datalog_2018_12_26.txt", "r").read()   

This will generate an outfile which will arrange data in csv format. After this we will segregate the data based on Event_Type and will make individual CSV files,we can insert path for this in this:

et1.to_csv('INSERT HERE THE PATH DIRECTORY')      

After this it will do the desired operations programmed in it and will produce the required CSV files for the DataFrame devloped and you can insert the path directory where you want it. For eg

bigdata4.to_csv('/INSERT PATH HERE/customer_et4.csv')

In the end we will produce the desired worksheet from the above performed segregations & analysis by means of XLSX WRITER module



