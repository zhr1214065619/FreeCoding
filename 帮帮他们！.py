#!/usr/bin/env python
# coding: utf-8

# In[22]:


import pandas as pd

sheet3 = pd.read_excel('./input.xlsx', sheet_name='人效底稿数据')
sheet3_dnum = sheet3.shape[0]
this_month_start = pd.datetime(2021,8,1)
for row_num in range(2,sheet3_dnum):
    data = sheet3.loc[row_num].values
    output = 0
    if data[7] and (((data[5] - this_month_start).days) < 0) : #当月在职tag
        data[12] = 1 #本月有效人力
    elif data[7]: #当月入职
        data[12] = ((data[5] - this_month_start).days) / 30
    elif data[8]: #当月离职tag
        data[12] = ((data[6] - this_month_start).days) / 30 #计算天数
    else:
        data[12] = 0


# In[23]:


import numpy as np
import math

def get_data(data, tar_col):#获取地区和需要的数据列
    if data[1] == "北京":
        return 0, data[tar_col]
    if data[1] == "河南":
        return 1, data[tar_col]
    if data[1] == "河北":
        return 2, data[tar_col]
    

def get_month(begin,end):
    begin_year, end_year = begin.year, end.year
    begin_month, end_month = begin.month, end.month
    if begin_year == end_year:
        months = end_month -  begin_month
    else:
        months = (end_year - begin_year) * 12 + end_month - begin_month
    return months

sheet4 = pd.read_excel('./input.xlsx', sheet_name='人效需要呈现数据')
pwrite_numpy = np.zeros((3, 4), dtype = float) #有效人力部分
nwrite_numpy = np.zeros((3, 4), dtype = float) #统计人数
fwrite_numpy = np.zeros((3, 4), dtype = float) #前面部分[起租单量, 上月人效, 当月起租金额总和, 上月起租单量]
job_type = ["城市经理", "城市经理（进店）"]
this_month_end = pd.datetime(2021,8,31)
for row_num in range(2,sheet3_dnum):
    data = sheet3.loc[row_num].values
    if data[3] in job_type: #是城市经理*
        #计算起租数量等
        (w_row, qdata) = get_data(data, 9) #获得当月起租单量总和
        if math.isnan(qdata):
            qdata = 0
        fwrite_numpy[w_row, 0] += qdata
        (w_row, qdata) = get_data(data, 11) #获得起租金额总和
        if math.isnan(qdata):
            qdata = 0
        fwrite_numpy[w_row, 2] += qdata
        
        if data[7]: #计算脑瘫要求的上个月人效环比
            (w_row, qdata) = get_data(data, 13) #获得上月有效人力总和
            if math.isnan(qdata):
                qdata = 0
            fwrite_numpy[w_row, 1] += qdata
            (w_row, qdata) = get_data(data, 10) #获得上月起租单量总和
            if math.isnan(qdata):
                qdata = 0
            fwrite_numpy[w_row, 3] += qdata
        
        if data[0] == False: #满足非当月离职
            if data[8]:
                months = get_month(data[5], data[6])
            else:
                months = get_month(data[5], this_month_end)
            (w_row, pdata) = get_data(data, 12) #获得当月人效
            #print(row, months, pdata)
            if months <= 2:
                pwrite_numpy[w_row, 0] += pdata
                nwrite_numpy[w_row, 0] += 1
            elif months <= 6:
                pwrite_numpy[w_row, 1] += pdata
                nwrite_numpy[w_row, 1] += 1
            elif months <= 12:
                pwrite_numpy[w_row, 2] += pdata
                nwrite_numpy[w_row, 2] += 1
            else:
                pwrite_numpy[w_row, 3] += pdata
                nwrite_numpy[w_row ,3] += 1
#print(sum(pwrite_numpy[0, :]))
#写入有效人力和人效均值            
for row_num in range(5,8):
    sheet4.iloc[row_num, 5] = sum(pwrite_numpy[row_num-5, :])
    for col_num in range(6,10):
        sheet4.iloc[row_num, col_num] = pwrite_numpy[row_num-5, col_num-6]
    for col_num in range(10,14):
        sheet4.iloc[row_num, col_num] = nwrite_numpy[row_num-5, col_num-10] * sheet4.iloc[row_num,1] / pwrite_numpy[row_num-5, col_num-10]

#写入前面
for row_num in range(5,8):
    sheet4.iloc[row_num, 1] = int(fwrite_numpy[row_num-5, 0]) #写入本月起租数量
    sheet4.iloc[row_num, 2] = sheet4.iloc[row_num, 1] / sheet4.iloc[row_num, 5] #写入当月人效
    sheet4.iloc[row_num, 3] = sheet4.iloc[row_num, 2] / (fwrite_numpy[row_num-5, 3] / fwrite_numpy[row_num-5, 1]) - 1
    sheet4.iloc[row_num, 4] = fwrite_numpy[row_num-5, 2] / sheet4.iloc[row_num, 1]
   


# In[89]:


def get_week(input_date):
    this_month_start = pd.datetime(int(input_date.year), int(input_date.month), 1)
    start_day_of_week = pd.to_datetime(this_month_start).dayofweek
    day_of_week_result = math.ceil((input_date.day + start_day_of_week) / 7)
    return day_of_week_result

sheet1 = pd.read_excel('./input.xlsx', sheet_name='渠道底稿数据')
sheet1_dnum = sheet1.shape[0]
for row_num in range(1, sheet1_dnum):
    data = sheet1.loc[row_num].values
    if pd.isna(data[2]): #如果表格没有值，那么值为nan
        sheet1.iloc[row_num, 4] = 0
        sheet1.iloc[row_num, 6] = 0
    if pd.isna(data[3]):
        sheet1.iloc[row_num, 5] = 0
        sheet1.iloc[row_num, 7] = 0
    if not pd.isna(data[2]) and (not pd.isna(data[3])):
        sheet1.iloc[row_num, 4] = pd.to_datetime(data[2]).month
        sheet1.iloc[row_num, 6] = get_week(pd.to_datetime(data[2]))
        sheet1.iloc[row_num, 5] = pd.to_datetime(data[3]).month
        sheet1.iloc[row_num, 7] = get_week(pd.to_datetime(data[3]))        


# In[116]:


sheet2 = pd.read_excel('./input.xlsx', sheet_name='渠道需呈现内容')
row_num_total = 52
this_month = 8
qresult_output = np.zeros((row_num_total, 8), dtype = int) #起租月份计数
qwresult_output = np.zeros((row_num_total, 6), dtype = int) #起租周计数
bresult_output = np.zeros((row_num_total, 8), dtype = int) #报单月份计数
bwresult_output = np.zeros((row_num_total, 6), dtype = int) #报单周计数
for row_num in range(0, row_num_total):
    data = sheet2.loc[5 + row_num].values
    this_index = data[2] #渠道名称
    for sheet1_row_num in range(1, sheet1_dnum):
        sheet1_data = sheet1.loc[sheet1_row_num].values
        if sheet1_data[0] == this_index and sheet1_data[1]:
            if not pd.isna(sheet1_data[2]) and sheet1_data[4] <= 8:
                qresult_output[row_num, sheet1_data[4] - 1] += 1
                if sheet1_data[4] == this_month:
                    #if sheet1_data[6] == 5:
                        #print(sheet1_data)
                    qwresult_output[row_num, sheet1_data[6] - 1] += 1
            if not pd.isna(sheet1_data[3]) and sheet1_data[5] <= 8:
                bresult_output[row_num, sheet1_data[5] - 1] += 1
                if sheet1_data[5] == this_month:
                    #print(sheet1_data)
                    #print(sheet1_data)
                    bwresult_output[row_num, sheet1_data[7] - 1] += 1
#辅助列
f1result_output = np.zeros((row_num_total,1), dtype = float)
for row_num in range(0, row_num_total):
    f1result_output[row_num,0] = (qresult_output[row_num, 5] + qresult_output[row_num, 6])/2
#拼接结果 
qoutput = np.concatenate((qresult_output,qwresult_output[:,:5]), axis = 1) #不显示6周
boutput = np.concatenate((bresult_output,bwresult_output[:,:5]), axis = 1) #不显示6周
output = np.concatenate((qoutput, boutput, f1result_output), axis = 1)
for row_num in range(0, output.shape[0]):
    for col_num in range(0,output.shape[1]):
        sheet2.iloc[5 + row_num, 3 + col_num] = output[row_num, col_num]
    if f1result_output[row_num,0] == 0:
        sheet2.iloc[5 + row_num, 3 + output.shape[1]] = 'D'
    elif f1result_output[row_num,0] < 3:
        sheet2.iloc[5 + row_num, 3 + output.shape[1]] = 'C'
    elif f1result_output[row_num,0] < 5:
        sheet2.iloc[5 + row_num, 3 + output.shape[1]] = 'B'
    else:
        sheet2.iloc[5 + row_num, 3 + output.shape[1]] = 'A'


# In[117]:


writer = pd.ExcelWriter("./output.xlsx")
sheet1.to_excel(writer, sheet_name = '渠道底稿数据')
sheet2.to_excel(writer, sheet_name = '渠道需呈现内容')
sheet3.to_excel(writer, sheet_name = '人效底稿数据')
sheet4.to_excel(writer, sheet_name = '人效需要呈现数据')
writer.save()
writer.close()

