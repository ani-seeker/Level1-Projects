#!/usr/bin/env python
# coding: utf-8

# # Import Libraries

# In[238]:


import pandas as pd
import numpy as np
import os, sys, path
from datetime import date, timedelta


# In[239]:


today = date.today()
yesterday = today - timedelta(days = 1)
yesterday1 = today - timedelta(days = 1)

yesterday = yesterday.strftime('%Y-%m-%d')
yesterday1 = yesterday1.strftime('%Y%m%d')

path = os.path.join(r'\\aupdiice01\Transport\Stibo_Archive',yesterday)

filenames = os.listdir(path)

file = [name for name in filenames if name.startswith("Vendor_Part_Details") and name.endswith(".txt") and yesterday1 in name]

stibo_daily_file = os.path.join(path,file[0])


# In[240]:


print("Name of STIBO Daily File :", stibo_daily_file)


# # Import Files

# In[241]:


iice_full_file = r"\\aupdiice01\iICE_Export\BulkVendorItemsExport_AU.txt"

print("Name of IICE Daily File :", iice_full_file)

#stibo_daily_file = r"C:\Users\axg03\Documents\Anirudh\Repositories\sample data\Vendor_Part_Details_20200531T060009.txt"
#iice_full_file = r"C:\Users\axg03\Documents\Anirudh\Repositories\sample data\IICE_vendorBulk_sample_01062020.txt"

print("Name of STIBO Daily File :", stibo_daily_file)


# In[242]:


df_stibo = pd.read_csv(stibo_daily_file,sep='\t',header=(0),encoding='utf8',low_memory=False,converters={'WIS_Part_No': lambda x: str(x),'Alt_Part_1': lambda x: str(x),'Alt_Part_2': lambda x: str(x),'Alt_Part_3': lambda x: str(x),'Orig_Branch': lambda x: str(x),'Synonym': lambda x: str(x)})


# In[243]:


print(iice_full_file)
print(stibo_daily_file)

df_iice = pd.read_csv(iice_full_file,sep='\t',header=(0),encoding='iso8859-1',low_memory=False,converters={'WIS_Part_No': lambda x: str(x),'Alt_Part_1': lambda x: str(x),'Alt_Part_2': lambda x: str(x),'Alt_Part_3': lambda x: str(x),'Orig_Branch': lambda x: str(x)})


# In[245]:


print("number of records in IICE File :",df_iice.shape[0])
print("number of records in IICE File :",df_stibo.shape[0])


# #display 1st few rows

# #merge two files

# In[246]:


df_merge = pd.merge(left=df_iice,right=df_stibo,how='right',left_on=['WIS_Part_No','Vendor_No'],right_on=['WIS_Part_No','PPV'])


# In[247]:


df_merge = df_merge.replace(np.nan, '', regex=True)


# #Compare all the columns for these 2 files

# In[248]:


df_merge['is_equal'] = ''


# In[259]:


df_merge.loc[df_merge['Item_Number_x'] != df_merge['Item_Number_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Item_Number is not updated in iICE'
df_merge.loc[df_merge['Barcode_Available_x'] != df_merge['Barcode_Available_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Barcode_Available is not updated in iICE'
df_merge.loc[df_merge['Barcode_x'] != df_merge['Barcode_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Barcode is not updated in iICE'
df_merge.loc[df_merge['Ctry_Of_Origin_x'] != df_merge['Ctry_Of_Origin_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Ctry_Of_Origin is not updated in iICE'
df_merge.loc[df_merge['Purch_UOM_x'] != df_merge['Purch_UOM_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Purch_UOM is not updated in iICE'
df_merge.loc[df_merge['Purch_UOM_Conv_x'] != df_merge['Purch_UOM_Conv_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Purch_UOM_Conv is not updated in iICE'
df_merge.loc[df_merge['Min_Purch_Qty_x'] != df_merge['Min_Purch_Qty_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Min_Purch_Qty is not updated in iICE'
df_merge.loc[df_merge['Standard_Pack_Qty_x'] != df_merge['Standard_Pack_Qty_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Standard_Pack_Qty is not updated in iICE'
df_merge.loc[df_merge['PPV_Flag_x'] != df_merge['PPV_Flag_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'PPV_Flag is not updated in iICE'
df_merge.loc[df_merge['POA_x'] != df_merge['POA_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'POA is not updated in iICE'
df_merge.loc[df_merge['New_Invoice_Cost_x'] != df_merge['New_Invoice_Cost_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'New_Invoice_Cost is not updated in iICE'
df_merge.loc[df_merge['New_Cost_Effective_Date_x'] != df_merge['New_Cost_Effective_Date_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'New_Cost_Effective_Date is not updated in iICE'
df_merge.loc[df_merge['Cost_Load_Date_x'] != df_merge['Cost_Load_Date_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Cost_Load_Date is not updated in iICE'
#df_merge.loc[df_merge['Invoice_Cost_Variance_x'] != df_merge['Invoice_Cost_Variance_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Invoice_Cost_Variance is not updated in iICE'
df_merge.loc[df_merge['Invoice_Cost'] != df_merge['New_Invoice_Cost_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Invoice_Cost is not updated in iICE'
df_merge.loc[df_merge['Vendor_No'] != df_merge['PPV'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Vendor_No is not updated in iICE'
df_merge.loc[df_merge['Vendor_Name_x'] != df_merge['Vendor_Name_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Vendor_Name is not updated in iICE'
df_merge.loc[df_merge['Search_Seq_1_x'] != df_merge['Search_Seq_1_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Search_Seq_1 is not updated in iICE'
df_merge.loc[df_merge['DG_Flag_x'] != df_merge['DG_Flag_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'DG_Flag is not updated in iICE'
df_merge.loc[df_merge['Hazardous_x'] != df_merge['Hazardous_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Hazardous is not updated in iICE'
df_merge.loc[df_merge['MSDS_Available_x'] != df_merge['MSDS_Available_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'MSDS_Available is not updated in iICE'
df_merge.loc[df_merge['MSDS_No_x'] != df_merge['MSDS_No_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'MSDS_No is not updated in iICE'
df_merge.loc[df_merge['MSDS_Issue_Date_x'] != df_merge['MSDS_Issue_Date_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'MSDS_Issue_Date is not updated in iICE'
df_merge.loc[df_merge['Danger_No_x'] != df_merge['Danger_No_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Danger_No is not updated in iICE'
df_merge.loc[df_merge['Danger_Sub_Risk_x'] != df_merge['Danger_Sub_Risk_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Danger_Sub_Risk is not updated in iICE'
df_merge.loc[df_merge['Combustable_Liq_x'] != df_merge['Combustable_Liq_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Combustable_Liq is not updated in iICE'
df_merge.loc[df_merge['Packing_Group_x'] != df_merge['Packing_Group_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Packing_Group is not updated in iICE'
df_merge.loc[df_merge['HazChem_Code_x'] != df_merge['HazChem_Code_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'HazChem_Code is not updated in iICE'
df_merge.loc[df_merge['UN_No_x'] != df_merge['UN_No_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'UN_No is not updated in iICE'
df_merge.loc[df_merge['Proper_Shipping_Name_x'] != df_merge['Proper_Shipping_Name_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Proper_Shipping_Name is not updated in iICE'
df_merge.loc[df_merge['Weight_x'] != df_merge['Weight_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Weight is not updated in iICE'
df_merge.loc[df_merge['Weight_Unit_x'] != df_merge['Weight_Unit_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Weight_Unit is not updated in iICE'
df_merge.loc[df_merge['Overweight_Warning_x'] != df_merge['Overweight_Warning_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Overweight_Warning is not updated in iICE'
df_merge.loc[df_merge['Width_(m)_x'] != df_merge['Width_(m)_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Width_(m) is not updated in iICE'
df_merge.loc[df_merge['Depth_(m)_x'] != df_merge['Depth_(m)_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Depth_(m) is not updated in iICE'
df_merge.loc[df_merge['Height_(m)_x'] != df_merge['Height_(m)_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Height_(m) is not updated in iICE'
df_merge.loc[df_merge['Volume_x'] != df_merge['Volume_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Volume is not updated in iICE'
df_merge.loc[df_merge['Oversize_Warning_x'] != df_merge['Oversize_Warning_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Oversize_Warning is not updated in iICE'
df_merge.loc[df_merge['Shelf_Life_x'] != df_merge['Shelf_Life_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Shelf_Life is not updated in iICE'
df_merge.loc[df_merge['Max_Shelf_Life(Months)_x'] != df_merge['Max_Shelf_Life(Months)_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Max_Shelf_Life(Months) is not updated in iICE'
df_merge.loc[df_merge['NCM_Responsible_x'] != df_merge['NCM_Responsible_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'NCM_Responsible is not updated in iICE'
df_merge.loc[df_merge['NCM_Responsible_Email_x'] != df_merge['NCM_Responsible_Email_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'NCM_Responsible_Email is not updated in iICE'
#df_merge.loc[df_merge['Sourcing_Manager_x'] != df_merge['Sourcing_Manager_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Sourcing_Manager is not updated in iICE'
#df_merge.loc[df_merge['Sourcing_Manager_Email_x'] != df_merge['Sourcing_Manager_Email_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Sourcing_Manager_Email is not updated in iICE'
df_merge.loc[df_merge['Qty_Break_1_x'] != df_merge['Qty_Break_1_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Qty_Break_1 is not updated in iICE'
df_merge.loc[df_merge['Qty_Break_Cost_1_x'] != df_merge['Qty_Break_Cost_1_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Qty_Break_Cost_1 is not updated in iICE'
df_merge.loc[df_merge['Qty_Break_2_x'] != df_merge['Qty_Break_2_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Qty_Break_2 is not updated in iICE'
df_merge.loc[df_merge['Qty_Break_Cost_2_x'] != df_merge['Qty_Break_Cost_2_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Qty_Break_Cost_2 is not updated in iICE'
df_merge.loc[df_merge['Qty_Break_3_x'] != df_merge['Qty_Break_3_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Qty_Break_3 is not updated in iICE'
df_merge.loc[df_merge['Qty_Break_Cost_3_x'] != df_merge['Qty_Break_Cost_3_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Qty_Break_Cost_3 is not updated in iICE'
df_merge.loc[df_merge['Style_No_x'] != df_merge['Style_No_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Style_No is not updated in iICE'
df_merge.loc[df_merge['Style_Name_x'] != df_merge['Style_Name_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Style_Name is not updated in iICE'
df_merge.loc[df_merge['Cert_Expiry_Date_x'] != df_merge['Cert_Expiry_Date_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Cert_Expiry_Date is not updated in iICE'
df_merge.loc[df_merge['Cert_Code_x'] != df_merge['Cert_Code_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Cert_Code is not updated in iICE'
df_merge.loc[df_merge['Ex_Works_Days_x'] != df_merge['Ex_Works_Days_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Ex_Works_Days is not updated in iICE'
df_merge.loc[df_merge['Warranty'] != df_merge['Manuf_Warranty'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Warranty is not updated in iICE'
df_merge.loc[df_merge['Freight_Weight_(kg)_x'] != df_merge['Freight_Weight_(kg)_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Freight_Weight_(kg) is not updated in iICE'
df_merge.loc[df_merge['Brand_Prim_Web_x'] != df_merge['Brand_Prim_Web_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Brand_Prim_Web is not updated in iICE'
df_merge.loc[df_merge['Brand_Sec_Web_x'] != df_merge['Brand_Sec_Web_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Brand_Sec_Web is not updated in iICE'
df_merge.loc[df_merge['UTM_x'] != df_merge['UTM_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'UTM is not updated in iICE'
df_merge.loc[df_merge['Stackable_x'] != df_merge['Stackable_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Stackable is not updated in iICE'
df_merge.loc[df_merge['Stack_Factor_x'] != df_merge['Stack_Factor_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Stack_Factor is not updated in iICE'
df_merge.loc[df_merge['Max_Stack_x'] != df_merge['Max_Stack_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Max_Stack is not updated in iICE'
df_merge.loc[df_merge['Indigenous_Supplier?_x'] != df_merge['Indigenous_Supplier?_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Indigenous_Supplier? is not updated in iICE'


# In[260]:


df_export = df_merge[['WIS_Part_No','PPV','is_equal']]
#df_export.WIS_Part_No = df_export.WIS_Part_No.astype('str')


# In[261]:


df_export.WIS_Part_No = df_export.WIS_Part_No.apply('="{}"'.format)
#df_export.to_csv('modified_vendorpart.csv')


# In[262]:


df_export.rename(columns = {'is_equal':'ErrorMessage'}, inplace = True)
#df_export.rename(columns = {'NEW_CATEGORY_ID_x':'Category'}, inplace = True)

df_export1 = df_merge[['WIS_Part_No','is_equal']]
df_export1.rename(columns = {'is_equal':'ErrorMessage'}, inplace = True)


# In[263]:


df_export1['ErrorMessage'] = df_export1['ErrorMessage'].replace('', 'No Error Message, Data matches with STIBO')


# In[264]:


df_group = df_export1.groupby(['ErrorMessage']).count()

df_group.columns

df_group

#df_group.to_excel('IICE_Part_Level_Comparison_with_STIBO_Daily_test.xlsx')

# Create a Pandas Excel writer using XlsxWriter as the engine.

writer = pd.ExcelWriter('IICE_Vendor_Part_Level_Comparison_with_STIBO_Daily.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.

df_group.to_excel(writer,sheet_name='Summary')
df_export.to_excel(writer,index=False,sheet_name='Error Message Details', startrow=1, header=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Summary']
worksheet1 = writer.sheets['Error Message Details']

# Add a header format.
header_format = workbook.add_format({
    'bold': True,
    #'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})

# Write the column headers with the defined format.
for col_num, value in enumerate(df_group.columns.values):
    worksheet.write(0, col_num+1, value, header_format)

# Set the column width and format.
worksheet.set_column('A:A', 100)
worksheet1.set_column('B:B', 50)
worksheet1.set_column('C:C', 100)

for col_num, value in enumerate(df_export.columns.values):
    worksheet1.write(0, col_num, value, header_format)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
