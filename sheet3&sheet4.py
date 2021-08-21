#!/usr/bin/env python
# coding: utf-8

# In[87]:


import pandas as pd

sheet3 = pd.read_excel('./input.xlsx', sheet_name='人效底稿数据')
sheet3_dnum = sheet3.shape[0]
this_month_start = pd.datetime(2021,8,1)
for row_num in range(2,sheet3_dnum):
    data = sheet3.loc[row_num].values
    output = 0
    if data[7]: #当月在职tag
        data[13] = 1 #本月有效人力
    elif data[8]: #当月离职tag
        data[13] = ((data[6] - this_month_start).days)/30 #计算天数
    else:
        data[13] = 0


# In[88]:


import numpy as np

def get_pdata(data):#获取地区和人效
    if data[1] == "北京":
        return 0, data[12]
    if data[1] == "河南":
        return 1, data[12]
    if data[1] == "河北":
        return 2, data[12]
    

def get_month(begin,end):
    begin_year, end_year = begin.year, end.year
    begin_month, end_month = begin.month, end.month
    if begin_year == end_year:
        months = end_month -  begin_month
    else:
        months = (end_year - begin_year) * 12 + end_month - begin_month
    return months

sheet4 = pd.read_excel('./input.xlsx', sheet_name='人效需要呈现数据')
pwrite_numpy = np.zeros((3, 4), dtype = float)
nwrite_numpy = np.zeros((3, 4), dtype = float)
job_type = ["城市经理", "城市经理（进店）"]
this_month_end = pd.datetime(2021,8,31)
for row_num in range(2,sheet3_dnum):
    data = sheet3.loc[row_num].values
    if data[0] == False and (data[3] in job_type): #在职
        if data[8]:
            months = get_month(data[5], data[6])
        else:
            months = get_month(data[5], this_month_end)
        (w_row, pdata) = get_pdata(data)
        #print(row, months, pdata)
        if months <= 2:
            pwrite_numpy[w_row,0] += pdata
            nwrite_numpy[w_row,0] += 1
        elif months <= 6:
            pwrite_numpy[w_row,1] += pdata
            nwrite_numpy[w_row,1] += 1
        elif months <= 12:
            pwrite_numpy[w_row,2] += pdata
            nwrite_numpy[w_row,2] += 1
        else:
            pwrite_numpy[w_row,3] += pdata
            nwrite_numpy[w_row,3] += 1
            
for row_num in range(5,8):
    for col_num in range(6,10):
        sheet4.iloc[row_num,col_num] = pwrite_numpy[row_num-5, col_num-6]
    for col_num in range(10,14):
        sheet4.iloc[row_num,col_num] = nwrite_numpy[row_num-5, col_num-10] * sheet4.iloc[row_num,1] / pwrite_numpy[row_num-5, col_num-10]

   


# In[89]:


writer = pd.ExcelWriter("./output.xlsx")
sheet3.to_excel(writer, sheet_name = '人效底稿数据')
sheet4.to_excel(writer, sheet_name = '人效需要呈现数据')
writer.save()
writer.close()


# In[ ]:




