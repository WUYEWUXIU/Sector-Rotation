#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import matplotlib as plt
pd.set_option('display.max_rows', None)
pd.set_option('max_columns', None)
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']#显示中文字体

import warnings
warnings.filterwarnings("ignore")
import time 


# In[2]:


dui_zhao_biao = pd.read_excel('对照表.xlsx')
dui_zhao_biao.set_index('代码',inplace=True)
dui_zhao_biao = pd.Series(dui_zhao_biao['简称'],dui_zhao_biao.index)


# In[3]:


def cal_dfs(series):
    Week_series = pd.Series()
    for i in interval:
        Week_series[series.index[i]] = series[i-4:i].sum()
    return Week_series

def cal_dfs_prod(series):
    Week_series = pd.Series()
    for i in interval:
        Week_series[series.index[i]] = (1+series[i-4:i]/100).prod()-1
    return Week_series

def import_data():
    #从Wind导入申万104个二级行业的月度数据【涨跌幅，成交量，成交额，换手率】 剔除801215和801144
    if Freq == 'm':
        temp_dfs = pd.read_excel("Month.xlsx",sheet_name=None,index_col=False)
        for sheet in temp_dfs.values():
            sheet.set_index('Unnamed: 0',inplace=True)
            sheet.dropna(axis=1,how="all",inplace=True) 
        dfs = pd.Series(temp_dfs)
    if Freq == 'w':
        temp_dfs = pd.read_excel("Week.xlsx",sheet_name=None,index_col=False)
        for sheet in temp_dfs.values():
            sheet.set_index('Unnamed: 0',inplace=True)
            sheet.dropna(axis=1,how="all",inplace=True) 
        dfs = pd.Series(temp_dfs)
    return dfs

def Guai_Li_Cal_Month(series):
    shift_2 = series.shift(2)
    shift_1 = series.shift(1)
    Mean = (shift_1 + shift_2)/2
    Final_series = (series - Mean)/Mean
    return  Final_series

def Guai_Li_Cal_Week(series):
    shift_2 = series.shift(2*4)
    shift_1 = series.shift(4)
    Mean = (shift_1 + shift_2)/2
    Final_series = (series - Mean)/Mean
    Week_series = pd.Series()
    for i in interval:
        Week_series[Final_series.index[i]] = Final_series[i-4:i].sum()
    return  Week_series

def cal_guai_li():
    #计算4-6指标：乖离率
    guai_li_lv_values = []
    for indexer in ['成交量','成交额','换手率']:
        temp_df = dfs[indexer].copy()
        if Freq == 'm':
            temp_df = temp_df.apply(Guai_Li_Cal_Month)
        if Freq == 'w':
            temp_df = temp_df.apply(Guai_Li_Cal_Week)
        guai_li_lv_values.append(temp_df)
    guai_li_lv = pd.Series(guai_li_lv_values,index=['月度成交量乖离率','月度成交额乖离率','月度换手率乖离率'])
    return guai_li_lv

def CrowdList(num):
    Crowd_List = pd.Series()
    six_index = pd.DataFrame()
    n = 0
    for df in dfs[['成交量','成交额','换手率']]:
        n = n + 1
        for industry in df.columns:
            series = df.loc[:,industry]
            series = series.rank()
            six_index.loc[n,industry] = series[df.index[-num]]/len(series)
            
    n = 3
    for df in guai_li_lv:
        n = n + 1
        for industry in df.columns:
            series = df.loc[:,industry]
            series = series.rank()
            six_index.loc[n,industry] = series[df.index[-num]]/len(series)
            
    judge = (six_index > 0.6).sum() > 2
    BlackList = judge[judge==True].index.tolist()
    WhiteList = [a for a in six_index.columns if a not in BlackList]
    Crowd_List['黑名单'] = BlackList
    Crowd_List['白名单'] = WhiteList
    Crowd_List['拥挤度排序'] = six_index.mean().sort_values(ascending=False).index.tolist()
    Crowd_List['平均值'] = six_index.mean()[six_index.mean().sort_values(ascending=False).index].tolist()
    return Crowd_List

def VolatilityList(num):
    waiting_to_be_picked = Zhang_Die.loc[Zhang_Die.index[-num],:].sort_values(ascending=False).index[0:40]
    Vol_List = waiting_to_be_picked.tolist()
    for industry in waiting_to_be_picked[10:]:
        if Zhang_Die.loc[Zhang_Die.index[-num],industry] < -0.05:
            Vol_List.remove(industry)
    return Vol_List

if __name__ == "__main__":
    Freq = input("Frequency:")
    num = 1
    dfs = import_data()
    interval = [i for i in range(len(dfs['成交量'].index)-1,11,-4)]
    interval.reverse()
    Zhang_Die = dfs['涨跌幅']
    dfs.drop('涨跌幅')
    if Freq == 'w':
        Zhang_Die = Zhang_Die.apply(cal_dfs_prod)
        for df in dfs:
            df = df.apply(cal_dfs)
    guai_li_lv = cal_guai_li()
    CrowdList_Month_os = CrowdList(num)
    VolatilityList_Month_os = VolatilityList(num)
    All_Data = pd.concat([pd.DataFrame({'动量组合': dui_zhao_biao[VolatilityList_Month_os].tolist()}),
                    pd.DataFrame({'拥挤度黑名单':dui_zhao_biao[CrowdList_Month_os['黑名单']].tolist()}),
                    pd.DataFrame({'拥挤度白名单':dui_zhao_biao[CrowdList_Month_os['白名单']].tolist()}),
                    pd.DataFrame({'拥挤度排序':dui_zhao_biao[CrowdList_Month_os['拥挤度排序']].tolist()}),
                    pd.DataFrame({'拥挤度百分位平均值':CrowdList_Month_os['平均值']}),
                    pd.DataFrame({'复合策略组合':dui_zhao_biao[[a for a in VolatilityList_Month_os if a in CrowdList_Month_os['白名单']]].tolist()})],
                         axis=1)
    display(All_Data)
    if Freq == 'w':
        All_Data.to_excel('（周）复合策略统计.xlsx')
    if Freq == 'm':
        All_Data.to_excel('（月）复合策略统计.xlsx')
    

