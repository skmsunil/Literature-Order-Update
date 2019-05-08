### Literature Order Update ###
# Initial Coding : Sunil Kumar Mishra
# Created On : 9/11/2018 
# Description: Will automate the task of downloading the file from Outlook
#              will prepare the Header Insert File using the downloaded file
#              will insert the header to salesforce Directory, retrieve the success id
#              prepare the insert file and retrive the success id and the
#              update the Literature Order.
import simple_salesforce
from simple_salesforce import Salesforce
import os
import sys
import pandas as pd
import win32com.client
import datetime
  
from datetime import timedelta
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
#import xlsxwriter



os.chdir('C:\\Users\\sunil.kumar.mishra.m\\Desktop\\Python')


### Downloading the File from Outlook using Subject 

#Fetching the today's date in day-month-year format(01-12-2017)
todayDate = datetime.date.today()
DateFormat = ('{:04d}'.format(todayDate.year)+'-'+'{:02d}'.format(todayDate.month)+'-'+'{:02d}'.format(todayDate.day))
todayfile=todayDate.strftime('%Y%m%d')
todayfile2=todayDate.strftime('%Y-%m-%d')
#sf = Salesforce(username="support1@regeneron.com.acn" , password="Regeneron2",security_token='', sandbox=True)


#object to access outlook application, Inbox, Item Inside the Inbox
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages=inbox.Items
subject=messages['Drop Shipment']

message= messages.GetLast()
body_content= message.body
file_name='Drop_Shipment_'+todayfile+'.xlsx'
for attachment in message.Attachments:
    attachment = message.Attachments
    break
attachment = attachment.Item(1)
fn=os.getcwd()+'\\'+file_name
attachment.SaveASFile(fn)
print('File Download Complete')
df= pd.read_excel('Drop_Shipment_'+todayfile+'.xlsx', 'Sheet1')

#Reading the downloaded file and creating the Header Insert file
sf = Salesforce(username="support1@regeneron.com.acn" , password="Regeneron2",security_token='', sandbox=True)
print('Login Successful')
print('Fetching Query')

#df= pd.read_excel('Drop_Shipment_'+todayfile+'.xlsx', 'Sheet1')
Id= []
Us_Id=[]
data= []
email = []
email = df['Email'].tolist()

n = len(email)
for i in range(n):
    try:
        
        s=sf.query("Select Id from User where Email = '%s'"% email[i])
        for record in s['records']:
            Us_Id.append(record['Id'])
        print("Query Fetched")
    except Exception as e:
           print(e)

User_id= {}
User_id['Us_Id'] = Us_Id
df1 = pd.DataFrame(User_id)
df2 =df1['Us_Id']

cell_value = df[['First Name', 'Last Name','Home Address','Home Address 2',
                 'Home City','Home State','Email']]
cell_value = cell_value.rename(columns = {'Home Address': 'Ship_to_Address_Line_1',
                                          'First Name':'First_Name',
                                           'Home Address 2':'Ship_to_Address_Line_2',
                                           'Home City':'Ship_to_City',
                                           'Home State':'Ship_to_State'
                                           })

cell_value['Order_For_User'] = df2 
cell_value['Order_Create_Date'] = DateFormat
cell_value['Ship_to_Country'] = "US"
cell_value['Delivery_Date']= "UPS Ground"
cell_value['Bulk_Order']="TRUE"
cell_value['REG_Address_Delivery']= "My Address"
cell_value['REG_Address_Type']= "TRUE"

cell_value.to_csv('IO_Header_Insert_'+todayfile+'.csv',index = False)

print('Values Copied')


#Insert the data of the IO Header Insert file to Literature Order(Inventory_order_vod__c )object

Id= []
Success = []
Read_Header_file = pd.read_csv('IO_Header_Insert_'+todayfile+'.csv')

success = pd.DataFrame(columns = ['First_Name','Ship_to_Address_Line_1','Ship_to_City',
                                  'Ship_to_State','Order_For_User','Order_Create_Date',
                                  'Ship_to_Country', 'Delivery_Date','Bulk_Order',
                                  'REG_Address_Delivery__c','REG_Address_Type__c' ])
Failure = pd.DataFrame(columns = ['First_Name', 'Error'])

userId = []
On_success = []
On_Failure = []
error = []
er = []
First_Name = []
Ship_to_Address_Line_1 = []
Ship_to_City = []
Ship_to_State = []
Order_For_User_2 = Read_Header_file['Order_For_User'].tolist()
Order_Create_Date=[]
Order_Create_Date= todayfile2
Order_Create_Date2=[]
Ship_to_Country =[] 
Delivery_Date = []
Bulk_Order = []
Show_Error = []
Error_type = []          
append_error = []
REG_Address_Delivery = []
REG_Address_Type= []
Order_For_User=[]
Order_For_User_3=[]
#Success = []
error= []
length =len(Read_Header_file)
for i in range(length):
    try:
             create_data = sf.Inventory_Order_vod__c.create({"REG_Ship_to_Address_Line_1__c":str(Read_Header_file.Ship_to_Address_Line_1[i]),
                                                         "Order_For_User_vod__c":Read_Header_file.Order_For_User[i],
                                                         "REG_Address_Delivery__c":str(Read_Header_file.REG_Address_Delivery[i]),
                                                         "REG_Address_Type__c":str(Read_Header_file.REG_Address_Type[i]),
                                                         "Order_Create_Date_vod__c":str(Read_Header_file.Order_Create_Date[i]),
                                                         "REG_Delivery_Date__c":str(Read_Header_file.Delivery_Date[i]),
                                                         "REG_Ship_to_City__c":str(Read_Header_file.Ship_to_City[i]),
                                                         "REG_Ship_to_State__c":str(Read_Header_file.Ship_to_State[i]),
                                                         "REG_Ship_Country__c":str(Read_Header_file.Ship_to_Country[i]),
                                                         })                  
             print('Data is Created')
             Id.append(create_data['id'])
             print(Id)
             print(Order_For_User_2[i])
             Order_For_User.append(Order_For_User_2[i])
             print(Order_For_User)
             Success.append(create_data['success'])
             error.append(create_data['errors'])
    except Exception as e:
             print(e)
             error_name = Read_Header_file.First_Name[i]
             error_Address= Read_Header_file.Ship_to_Address_Line_1[i],
             error_City= Read_Header_file.Ship_to_City[i],
             error_State= Read_Header_file.Ship_to_State[i],
             error_country= Read_Header_file.Ship_to_Country[i],
             error_order_create_date = Read_Header_file.Order_Create_Date[i],
             
             error_delivery_date=Read_Header_file.Delivery_Date[i],
             error_bulk_order= Read_Header_file.Bulk_Order[i]
             error_order_for_user=Order_For_User_2[i]
             print('Failed')
             First_Name.append(error_name)
             Ship_to_Address_Line_1.append(error_Address)
             Ship_to_City.append(error_City)
             Ship_to_State.append(error_State)
            
             Order_Create_Date2.append(error_order_create_date)
             Ship_to_Country.append(error_country)
             Delivery_Date.append(error_delivery_date)
             Bulk_Order.append(error_bulk_order)
             Order_For_User3.append(error_order_for_user)
             er.append(e)
             Error_type.append(type(e))
success_dict={}
success_dict['Order_for_User_Id'] = Order_For_User
success_dict['Id'] = Id
success_dict['Email'] = Read_Header_file.Email
success_dict['Success'] = Success


error_dict ={}
error_dict['First_Name'] = First_Name
error_dict['Ship_to_Address_Line_1'] = Ship_to_Address_Line_1
error_dict['Ship to City'] = Ship_to_City
error_dict['Ship to State'] = Ship_to_State

error_dict['Order Create Date'] = Order_Create_Date
error_dict['Ship to Country'] = Ship_to_Country
error_dict['Delivery Date'] = Delivery_Date
error_dict['Bulk Order'] = Bulk_Order
error_dict['Order_For_User'] = Order_For_User_3

success = pd.DataFrame(success_dict)
error = pd.DataFrame(error_dict)

#success file will be used for Preparation of Data Load Templates
success.to_excel('Io_Header_Success_'+todayfile+'.xlsx', index = False)
error.to_excel('Io_Header_Error_'+todayfile+'.xlsx', index = False)
print('DataFrame is Created and Saved')




# Preparing Literature Order Allocation inserting data and fetching id for Preparation of Data Load Templates
Id=[]
success_alloc=[]
Success_alloc=[]
result = []
error_alloc=[]


#preparing Bulk File after concateinating
Item = []
Bulk_Item=[]
Item_2 = df['Item']+'-Bulk'
df['Bulk_Item'] = Item_2
print(df['Bulk_Item'])
Bulk_Item=[]
Bulk_Item = df['Bulk_Item'].tolist()
QTY = []

#Adding Start Date and End Date
Order_Start_Date = todayDate.strftime('%Y-%m-%d')
print(Order_Start_Date)
end_date= todayDate + timedelta(days=15)
Order_End_Date=end_date.strftime('%Y-%m-%d')
print(Order_End_Date)
Allocation_Start_Date = Order_Start_Date
Allocation_End_Date = Order_End_Date

# Fetching the ID for REG_Alliance_Product_Code__c 
product = []
Product_id=[]
prod_id=[]
Product_ID=[]
QTY=[]
product = df['Item'].tolist()

n = len(product)
for i in range(n):
    try:
        
        s=sf.query("Select Id from Product_vod__c where REG_Alliance_Product_Code__c = '%s'"% product[i])
        for record in s['records']:
            prod_id.append(record['Id'])
        print("Query Fetched")
        print(prod_id)
    except Exception as e:
           print(e)

Product_id= {}
Product_id['prod_id'] = prod_id
df3 = pd.DataFrame(Product_id)
df['Product_ID'] = df3
df.to_excel('Drop_Shipment_'+todayfile+'.xlsx', index=False)

#Inserting In Inventory_Order_Allocation_vod__c

#product=set(product)
#product=list(product)
#print(product)

#item_in_list =len(product)
item_in_file = len(df)
print(item_in_file)
#for j in range(item_in_file):
    #print (j)
  #  key = 0 __getitem__(0)
         #try:
Allocation_data = sf.Inventory_Order_Allocation_vod__c.create({"Name":str(df.Bulk_Item[0]),                                                                 
                                                        "Product_Order_Allocation_Quantity_vod__c":str(df.QTY[0]),                                                                 
                                                        "Product_vod__c":str(df.Product_ID[0]),
                                                        "Order_Start_Date_vod__c": Order_Start_Date,
                                                        "Order_End_Date_vod__c":Order_End_Date,
                                                        "Allocation_Start_Date_vod__c":str(Allocation_Start_Date),
                                                        "Allocation_End_Date_vod__c":str(Allocation_End_Date)})
Allocation_data_Id=Allocation_data['id']
'''for record in Allocation_data[records]:
        Allocation_data_Id=record['id']
        print(Allocation_data_Id)
   Id.append(Allocation_data['id'])                          
    Success_alloc.append(Allocation_data['success'])
                            
         #except Exception as e:
           #  print(e)

'''
        
success_dict_alloc ={}
#success_dict_alloc['Product_ID']= prod_id[0]
success_dict_alloc['Email'] = df['Email']
Email= success_dict_alloc['Email']
suc_mail= len(Email)
print(prod_id)
print(suc_mail)
for p in range(suc_mail):
    print(p)
    success_dict_alloc['Id'] = Allocation_data_Id
    success_dict_alloc['Success_alloc'] =' Success'
    print(Allocation_data_Id)
success_dict_alloc['Product_ID']= df.Product_ID
#success_dict_alloc['Id'] = Allocation_data_Id    

success_alloc = pd.DataFrame(success_dict_alloc)


success_alloc.to_excel('Allocation_Success_'+todayfile+'.xlsx', index = False)
#error_alloc.to_excel('Allocation_Error_'+todayfile+'.xlsx', index = False)
print("Data Created")



#Preparing IO Insert File

Header_success = pd.read_excel('Io_Header_Success_'+todayfile+'.xlsx')
Allocation_success =pd.read_excel('Allocation_Success_'+todayfile+'.xlsx')
Read_Header_file = pd.read_csv('IO_Header_Insert_'+todayfile+'.csv')


Load_Data_temp=[]
Load_Data_temp1 = []
Load_Data_temp2=[]
Load_Data_temp3=[]

#concatenating all three Files

Load_Data_temp1=df[['QTY','Product_ID','Item', 'Email']]
Load_Data_temp3=Allocation_success[['Id', 'Email']]
Load_Data_temp = Load_Data_temp1.merge(Load_Data_temp3, on=['Email'], how ='left' )

Load_Data_temp.to_csv('test.csv')



Load_Data_temp= Load_Data_temp.rename(columns={'Order_quantity_vod':'ORDER_QUANTITY_UOM_VOD__C',
                                                'QTY':'ORDER_QUANTITY_VOD__C',
                                                'Id': "Inventory_Order_Allocation_vod__c",
                                                'Product_ID':'Product_vod__c',
                                                'Item':'ZVOD_PRODUCT_REG_PRODUCT_CODE__C'                                                
                                                })
Load_Data_temp ['ORDER_QUANTITY_UOM_VOD__C'] = "Each"
Load_Data_temp['Email']= Read_Header_file['Email']
#Load_Data_temp['Id']= "INVENTORY_ORDER_HEADER_VOD__C"
Io_file_insert = Load_Data_temp.merge(Header_success, on = 'Email', how = 'left')


Io_file_insert = Io_file_insert.rename(columns={'Id':'INVENTORY_ORDER_HEADER_VOD__C'})


Io_file_insert.to_csv('IO_Line_Insert_'+todayfile+'.csv',index = False)
#Io_file_insert = Io_file_insert[['Inventory_Order_Allocation_vod__c']]

Io_file_insert = pd.read_csv('IO_Line_Insert_'+todayfile+'.csv')

#Inserting Into Inventory_Order_Line_vod__c object
item_in_io_file =len(Io_file_insert)

for b in range(len(Io_file_insert)):
   print(Io_file_insert)
   try:
        Io_file_data = sf.Inventory_Order_Line_vod__c.create({"Inventory_Order_Allocation_vod__c":str(Io_file_insert.Inventory_Order_Allocation_vod__c[b]),                                                                 
                                                                "ORDER_QUANTITY_VOD__C":str(Io_file_insert.ORDER_QUANTITY_VOD__C[b]),                                                                 
                                                                "Product_vod__c":str(Io_file_insert.Product_vod__c[b]),
                                                                "ZVOD_PRODUCT_REG_PRODUCT_CODE__C": str(Io_file_insert.ZVOD_PRODUCT_REG_PRODUCT_CODE__C[b]),
                                                                "INVENTORY_ORDER_HEADER_VOD__C":str(Io_file_insert.INVENTORY_ORDER_HEADER_VOD__C[b]),#alloc contains header value
                                                                "ORDER_QUANTITY_UOM_VOD__C":str(Io_file_insert.ORDER_QUANTITY_UOM_VOD__C[b])})
                                                                                                                                
        Id.append(Allocation_data['id'])
        print(Id)
   except Exception as e:
            print(e)
print("Values has been Successfully updated In Io_Insert_file")

'''
# Updating Success File
Id = []
Insert_status =Io_file_insert['Inventory_Order_Allocation_vod__c']
for k in range(len(Io_file_insert)):
             print(Insert_status)
             insert_stat = Insert_status[k]
             Update_data = sf.Inventory_Order_vod__c.update(insert_stat,{"Order_Status_vod__c":"Submitted_vod"})                  
             print('Status is Updated')
             Id.append(['id'])
             print(Id)
'''   

