#!/usr/bin/env python
# coding: utf-8

# # Import Libraries

# In[409]:


import pandas as pd
import numpy as np
import os, sys, path
from datetime import date, timedelta


# In[410]:


today = date.today()
yesterday = today - timedelta(days = 1)
yesterday1 = today - timedelta(days = 1)

yesterday = yesterday.strftime('%Y-%m-%d')
yesterday1 = yesterday1.strftime('%Y%m%d')

path = os.path.join(r'\\aupdiice01\Transport\Stibo_Archive',yesterday)

filenames = os.listdir(path)

file = [name for name in filenames if name.startswith("Part_Details") and name.endswith(".txt") and yesterday1 in name]

stibo_daily_file = os.path.join(path,file[0])


# In[411]:


print("Name of STIBO Daily File :", stibo_daily_file)


# # Import Files

# In[412]:


iice_full_file = r"\\aupdiice01\iICE_Export\All_WISAU_Items.txt"

print("Name of IICE Daily File :", iice_full_file)


# In[413]:


df_iice = pd.read_csv(iice_full_file,sep='\t',header=(0),encoding='iso8859-1',low_memory=False,converters={'Part_No': lambda x: str(x),'Alt_Part_1': lambda x: str(x),'Alt_Part_2': lambda x: str(x),'Alt_Part_3': lambda x: str(x),'Orig_Branch': lambda x: str(x)})


# In[416]:


df_stibo = pd.read_csv(stibo_daily_file,sep='\t',header=(0),encoding='utf8',low_memory=False,converters={'Part_No': lambda x: str(x),'Alt_Part_1': lambda x: str(x),'Alt_Part_2': lambda x: str(x),'Alt_Part_3': lambda x: str(x),'Orig_Branch': lambda x: str(x),'Synonym': lambda x: str(x),'NATO_Stock_No': lambda x: str(x)})

print(iice_full_file)
print(stibo_daily_file)


# In[417]:



print("number of records in IICE File :",df_iice.shape[0])
print("number of records in IICE File :",df_stibo.shape[0])


# In[418]:


df_stibo.Orig_Branch = df_stibo.Orig_Branch.astype(str)


# #display 1st few rows

# In[419]:


df_iice.tail(5)


# In[369]:


df_stibo.head(5)


# #merge two files

# In[420]:


df_merge = pd.merge(df_iice,df_stibo,on='Part_No',how='right')


# In[421]:


df_merge = df_merge.replace(np.nan, '', regex=True)


# #Compare all the columns for these 2 files

# In[422]:



print(df_merge.shape)
print(df_merge.head(5))


# In[423]:


df_merge['is_equal'] = ''


# In[424]:


df_merge.loc[df_merge['Alt_Part_1_x'] != df_merge['Alt_Part_1_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Alt_Part_1 is not updated in iICE'
df_merge.loc[df_merge['Alt_Part_2_x'] != df_merge['Alt_Part_2_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Alt_Part_2 is not updated in iICE'
df_merge.loc[df_merge['Alt_Part_3_x'] != df_merge['Alt_Part_3_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Alt_Part_3 is not updated in iICE'
df_merge.loc[df_merge['Alt_Part_Ratio_1_x'] != df_merge['Alt_Part_Ratio_1_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Alt_Part_Ratio_1 is not updated in iICE'
df_merge.loc[df_merge['Alt_Part_Ratio_2_x'] != df_merge['Alt_Part_Ratio_2_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Alt_Part_Ratio_2 is not updated in iICE'
df_merge.loc[df_merge['Alt_Part_Ratio_3_x'] != df_merge['Alt_Part_Ratio_3_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Alt_Part_Ratio_3 is not updated in iICE'
df_merge.loc[df_merge['AwkwardGoods_x'] != df_merge['AwkwardGoods_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'AwkwardGoods is not updated in iICE'
df_merge.loc[df_merge['BI_Web_Ranged_x'] != df_merge['BI_Web_Ranged_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'BI_Web_Ranged is not updated in iICE'
#df_merge.loc[df_merge['Brand_WIS_x'] != df_merge['Brand_WIS_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Brand_WIS is not updated in iICE'
df_merge.loc[df_merge['BRI1_T200_x'] != df_merge['BRI1_T200_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'BRI1_T200 is not updated in iICE'
df_merge.loc[df_merge['BW_Web_x'] != df_merge['BW_Web_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'BW_Web is not updated in iICE'
df_merge.loc[df_merge['BWX_Video_x'] != df_merge['BWX_Video_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'BWX_Video is not updated in iICE'
df_merge.loc[df_merge['Description_x'] != df_merge['Description_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Description is not updated in iICE'
df_merge.loc[df_merge['Display_Pref_x'] != df_merge['Display_Pref_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Display_Pref is not updated in iICE'
df_merge.loc[df_merge['GM_Cat_Code'] != df_merge['GM_CAT_CODE'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'GM_CAT_CODE is not updated in iICE'
df_merge.loc[df_merge['GM_Category'] != df_merge['GM_CATEGORY'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'GM_CATEGORY is not updated in iICE'
df_merge.loc[df_merge['GS_Bonus_x'] != df_merge['GS_Bonus_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'GS_Bonus is not updated in iICE'
df_merge.loc[df_merge['GS_Reporting_x'] != df_merge['GS_Reporting_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'GS_Reporting is not updated in iICE'
df_merge.loc[df_merge['GTP_x'] != df_merge['GTP_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'GTP is not updated in iICE'
df_merge.loc[df_merge['GWR_Category_x'] != df_merge['GWR_Category_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'GWR_Category is not updated in iICE'
df_merge.loc[df_merge['Indent_Vendor_x'] != df_merge['Indent_Vendor_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Indent_Vendor is not updated in iICE'
df_merge.loc[df_merge['IPD_Ctlg_Page_x'] != df_merge['IPD_Ctlg_Page_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'IPD_Ctlg_Page is not updated in iICE'
df_merge.loc[df_merge['IPD_MSQ_x'] != df_merge['IPD_MSQ_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'IPD_MSQ is not updated in iICE'
df_merge.loc[df_merge['IPD_MSQ_x'] != df_merge['IPD_MSQ_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'IPD_MSQ is not updated in iICE'
df_merge.loc[df_merge['IPD_MSQ_ORF_x'] != df_merge['IPD_MSQ_ORF_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'IPD_MSQ_ORF is not updated in iICE'
df_merge.loc[df_merge['IPD_OMQ_x'] != df_merge['IPD_OMQ_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'IPD_OMQ is not updated in iICE'
df_merge.loc[df_merge['IPD_OMQ_ORF_x'] != df_merge['IPD_OMQ_ORF_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'IPD_OMQ_ORF is not updated in iICE'
df_merge.loc[df_merge['IPD_Stock_Mod_Max_x'] != df_merge['IPD_Stock_Mod_Max_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'IPD_Stock_Mod_Max is not updated in iICE'
df_merge.loc[df_merge['IPD_Stock_Mod_Min_x'] != df_merge['IPD_Stock_Mod_Min_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'IPD_Stock_Mod_Min is not updated in iICE'
df_merge.loc[df_merge['IPD_Stream_Flag_x'] != df_merge['IPD_Stream_Flag_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'IPD_Stream_Flag is not updated in iICE'
df_merge.loc[df_merge['MOA_Stream_Flag_x'] != df_merge['MOA_Stream_Flag_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'MOA_Stream_Flag is not updated in iICE'
df_merge.loc[df_merge['National_PPV_x'] != df_merge['National_PPV_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'National_PPV is not updated in iICE'
df_merge.loc[df_merge['NATO_Stock_No_x'] != df_merge['NATO_Stock_No_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'NATO_Stock_No is not updated in iICE'
df_merge.loc[df_merge['NCM_Cat_Code'] != df_merge['NCM_CAT_CODE'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'NCM_CAT_CODE is not updated in iICE'
df_merge.loc[df_merge['NCM_Category'] != df_merge['NCM_CATEGORY'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'NCM_CATEGORY is not updated in iICE'
df_merge.loc[df_merge['NEW_CAT_ID_Code'] != df_merge['NEW_CAT_ID_CODE'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'NEW_CAT_ID_CODE is not updated in iICE'
df_merge.loc[df_merge['NEW_CATEGORY_ID_x'] != df_merge['NEW_CATEGORY_ID_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'NEW_CATEGORY_ID is not updated in iICE'
df_merge.loc[df_merge['New_Part_x'] != df_merge['New_Part_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'New_Part is not updated in iICE'
df_merge.loc[df_merge['Ob_Approved_By_x'] != df_merge['Ob_Approved_By_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Ob_Approved_By is not updated in iICE'
df_merge.loc[df_merge['Ob_Business_Stream_x'] != df_merge['Ob_Business_Stream_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Ob_Business_Stream is not updated in iICE'
df_merge.loc[df_merge['Ob_Change_Date_x'] != df_merge['Ob_Change_Date_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Ob_Change_Date is not updated in iICE'
df_merge.loc[df_merge['Obsolescence_Code_x'] != df_merge['Obsolescence_Code_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Obsolescence_Code is not updated in iICE'
df_merge.loc[df_merge['Orig_Branch_x'] != df_merge['Orig_Branch_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Orig_Branch is not updated in iICE'
df_merge.loc[df_merge['Other_Conv_1_x'] != df_merge['Other_Conv_1_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Other_Conv_1 is not updated in iICE'
df_merge.loc[df_merge['Other_Conv_2_x'] != df_merge['Other_Conv_2_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Other_Conv_2 is not updated in iICE'
df_merge.loc[df_merge['Other_Conv_3_x'] != df_merge['Other_Conv_3_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Other_Conv_3 is not updated in iICE'
df_merge.loc[df_merge['Other_UOM_1_x'] != df_merge['Other_UOM_1_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Other_UOM_1 is not updated in iICE'
df_merge.loc[df_merge['Other_UOM_2_x'] != df_merge['Other_UOM_2_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Other_UOM_2 is not updated in iICE'
df_merge.loc[df_merge['Other_UOM_3_x'] != df_merge['Other_UOM_3_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Other_UOM_3 is not updated in iICE'
df_merge.loc[df_merge['Part_Creation_Approver_x'] != df_merge['Part_Creation_Approver_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Part_Creation_Approver is not updated in iICE'
df_merge.loc[df_merge['Part_Creation_Date_x'] != df_merge['Part_Creation_Date_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Part_Creation_Date is not updated in iICE'
df_merge.loc[df_merge['Part_Creation_Reason_x'] != df_merge['Part_Creation_Reason_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Part_Creation_Reason is not updated in iICE'
df_merge.loc[df_merge['Part_No'] != df_merge['Part_No'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Part_No is not updated in iICE'
df_merge.loc[df_merge['Partial_Supply_x'] != df_merge['Partial_Supply_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Partial_Supply is not updated in iICE'
df_merge.loc[df_merge['PER1_T200_x'] != df_merge['PER1_T200_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'PER1_T200 is not updated in iICE'
df_merge.loc[df_merge['PPV_x'] != df_merge['PPV_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'PPV is not updated in iICE'
df_merge.loc[df_merge['PPV_Part_No_x'] != df_merge['PPV_Part_No_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'PPV_Part_No is not updated in iICE'
df_merge.loc[df_merge['Preferred_Prod_Equiv_x'] != df_merge['Preferred_Prod_Equiv_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Preferred_Prod_Equiv is not updated in iICE'
df_merge.loc[df_merge['Preferred_Prod_Equiv_Ratio_x'] != df_merge['Preferred_Prod_Equiv_Ratio_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Preferred_Prod_Equiv_Ratio is not updated in iICE'
df_merge.loc[df_merge['Pricing_Conv_x'] != df_merge['Pricing_Conv_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Pricing_Conv is not updated in iICE'
df_merge.loc[df_merge['Pricing_UOM_x'] != df_merge['Pricing_UOM_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Pricing_UOM is not updated in iICE'
df_merge.loc[df_merge['Product_Department_x'] != df_merge['Product_Department_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Product_Department is not updated in iICE'
df_merge.loc[df_merge['Product_Group_x'] != df_merge['Product_Group_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Product_Group is not updated in iICE'
df_merge.loc[df_merge['Product_Sub_Group_x'] != df_merge['Product_Sub_Group_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Product_Sub_Group is not updated in iICE'
df_merge.loc[df_merge['Returnable_x'] != df_merge['Returnable_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Returnable is not updated in iICE'
df_merge.loc[df_merge['Sales_Tax_Percent_x'] != df_merge['Sales_Tax_Percent_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Sales_Tax_Percent is not updated in iICE'
df_merge.loc[df_merge['Search_Seq_1_x'] != df_merge['Search_Seq_1_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Search_Seq_1 is not updated in iICE'
df_merge.loc[df_merge['Sequence_No_x'] != df_merge['Sequence_No_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Sequence_No is not updated in iICE'
df_merge.loc[df_merge['SLOBS_New_Part_x'] != df_merge['SLOBS_New_Part_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'SLOBS_New_Part is not updated in iICE'
df_merge.loc[df_merge['Standard_Pack_Qty_x'] != df_merge['Standard_Pack_Qty_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Standard_Pack_Qty is not updated in iICE'
df_merge.loc[df_merge['Status_Change_Reason_x'] != df_merge['Status_Change_Reason_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Status_Change_Reason is not updated in iICE'
df_merge.loc[df_merge['Status_Review_Date_x'] != df_merge['Status_Review_Date_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Status_Review_Date is not updated in iICE'
df_merge.loc[df_merge['Stock_Status_x'] != df_merge['Stock_Status_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Stock_Status is not updated in iICE'
df_merge.loc[df_merge['Stock_UOM_x'] != df_merge['Stock_UOM_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Stock_UOM is not updated in iICE'
df_merge.loc[df_merge['Strat_Review_Per_x'] != df_merge['Strat_Review_Per_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Strat_Review_Per is not updated in iICE'
df_merge.loc[df_merge['Superseded_By'] != df_merge['Superseded_by'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Superseded_by is not updated in iICE'
df_merge.loc[df_merge['IPD_Supplementary_Part_1'] != df_merge['Supplimentary_Part_1'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Supplimentary_Part_1 is not updated in iICE'
df_merge.loc[df_merge['IPD_Supplementary_Part_2'] != df_merge['Supplimentary_Part_2'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Supplimentary_Part_2 is not updated in iICE'
df_merge.loc[df_merge['IPD_Supplementary_Part_3'] != df_merge['Supplimentary_Part_3'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Supplimentary_Part_3 is not updated in iICE'
df_merge.loc[df_merge['Supply_Type_x'] != df_merge['Supply_Type_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Supply_Type is not updated in iICE'
df_merge.loc[df_merge['SYD1_T200_x'] != df_merge['SYD1_T200_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'SYD1_T200 is not updated in iICE'
df_merge.loc[df_merge['Synonym_x'] != df_merge['Synonym_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Synonym is not updated in iICE'
df_merge.loc[df_merge['UNSPSC_x'] != df_merge['UNSPSC_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'UNSPSC is not updated in iICE'
df_merge.loc[df_merge['VA_Base_Part_No_x'] != df_merge['VA_Base_Part_No_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'VA_Base_Part_No is not updated in iICE'
df_merge.loc[df_merge['Web_Matrix_x'] != df_merge['Web_Matrix_y'], 'is_equal'] = df_merge['is_equal'] + ' || ' + 'Web_Matrix is not updated in iICE'


# In[425]:


df_export = df_merge[['Part_No','NEW_CATEGORY_ID_x','is_equal']]
df_export.Part_No = df_export.Part_No.apply('="{}"'.format)


# In[462]:


df_export.rename(columns = {'is_equal':'ErrorMessage'}, inplace = True)
df_export.rename(columns = {'NEW_CATEGORY_ID_x':'Category'}, inplace = True)


# In[469]:


df_export1 = df_merge[['Part_No','is_equal']]
df_export1.rename(columns = {'is_equal':'ErrorMessage'}, inplace = True)


# In[470]:


df_export1


# In[471]:


df_export1['ErrorMessage'] = df_export1['ErrorMessage'].replace('', 'No Error Message, Data matches with STIBO')


# In[472]:


df_export1


# In[473]:


df_group = df_export1.groupby(['ErrorMessage']).count()


# In[474]:


df_group.columns


# In[476]:


df_group


# In[395]:


#df_group.to_excel('IICE_Part_Level_Comparison_with_STIBO_Daily_test.xlsx')


# # Create a Pandas Excel writer using XlsxWriter as the engine.

# In[430]:


writer = pd.ExcelWriter('IICE_Part_Level_Comparison_with_STIBO_Daily.xlsx', engine='xlsxwriter')


# # Write each dataframe to a different worksheet.

# In[431]:


df_group.to_excel(writer,sheet_name='Summary')
df_export.to_excel(writer,index=False,sheet_name='Error Message Details', startrow=1, header=False)


# In[432]:


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


# In[433]:


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
