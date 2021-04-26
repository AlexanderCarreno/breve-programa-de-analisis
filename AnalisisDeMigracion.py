#!/usr/bin/env python
# coding: utf-8

# In[49]:


# Load pandas
import pandas as pd
import xlsxwriter
import numpy
import os 

main_path = "your_path_here"


# In[50]:


#function to Load excel data into DataFrame
def load_data(main_path,file_name): 
    full_path = os.path.join(main_path,file_name)
    sheet1 = pd.read_excel(full_path, index_col=None, header=None, skiprows=2, sheet_name=0)
    sheet2 = pd.read_excel(full_path, index_col=None, header=None, skiprows=1, sheet_name=1) 
    sheet3 = pd.read_excel(full_path, index_col=None, header=None, skiprows=0, sheet_name=2)
    return sheet1, sheet2, sheet3


# In[51]:


df1, df2, df3=load_data(main_path,'Sample-file-1.xlsm')
df4, df5, df6=load_data(main_path,'Sample-file-2.xlsm')

#append information in a DataFrame
df1 = df1.append(df4, ignore_index = True)  


# In[52]:


# Show dataframe 1
df1.head(10)


# In[53]:


# Show dataframe 1
df2.head(10)


# In[ ]:


'''
left_join_df= pd.merge(df1, df2, on='Name', how='left')
print(left_join_df)
'''


# In[ ]:


'''
left_join_df_array = left_join_df.to_numpy()
print(left_join_df_array)
'''


# In[ ]:


'''
# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('C:/data/Expenses01.xlsx',{'nan_inf_to_errors': True})
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.
expenses = (left_join_df_array)

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for Name,year,Carrer  in (expenses):
    worksheet.write(row, col,  Name)
    worksheet.write(row, col + 1,  year)
    worksheet.write(row, col + 2,  Carrer)
    row += 1
    
workbook.close()
'''


# In[ ]:





# In[ ]:




