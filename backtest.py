# %%
import collections
from math import isnan
from os import write
from pandas.core.dtypes.missing import isna
from scipy.stats import rankdata
import pandas as pd
import numpy as np
import matplotlib as plt
from pathlib import Path
import matplotlib.pyplot as plt
from scipy import stats
from datetime import datetime
import matplotlib.dates as mdates
pd.set_option('display.max_rows', None)
pd.set_option('max_columns', None)
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # 显示中文字体

# %%
dfs = pd.read_excel('Month.xlsx', sheet_name=['涨跌幅', '成交量', '成交额', '换手率'])
Return = dfs['涨跌幅'].set_index('Unnamed: 0')
Index = []
for key, df in dfs.items():
    if not (key == '涨跌幅'):
        Index.append(df)
Index = pd.Series(Index)
Index_idx = Index.apply(lambda x: x.set_index('Unnamed: 0'))
Deviation = Index_idx.apply(lambda x: (x / (x.shift(1)/2 + x.shift(2)/2) - 1))
Together = pd.concat([Index_idx, Deviation])
Percent = Together.apply(lambda x: x.rolling(
    x.shape[0], min_periods=24).apply(lambda y: rankdata(y)[-1]/len(y)))

n = 0.6
Base = Return.apply(np.mean, axis=1)
Percent_n = Percent.apply(lambda x: (x > 0.6).astype(np.int))
Grade = Percent_n.sum()

# %%
# Backtest
month = 1  # 为什么之前是1？
num_of_industries = 40
num_of_signals = 3
bound = -0.05  # 本身就是以5计数的？


Volatility_Group = Return.apply(lambda x: x.sort_values(
    ascending=False)[0:num_of_industries].index.tolist(), axis=1)
Volatility_Group_filtered = pd.Series()
for i, group in enumerate(Volatility_Group):
    tmp = Return.loc[Volatility_Group.index[i], group]
    Volatility_Group_filtered[Volatility_Group.index[i]] = \
        tmp[tmp > bound].index.tolist()
Crowd_Group = Grade.apply(
    lambda x: x[x < num_of_signals].index.tolist(), axis=1)
Position = pd.Series()
for date in Return.index:
    Position[date] = [x for x in Volatility_Group_filtered[date]
                      if x in Crowd_Group[date]]
Position = Position['2015-11':]
Position = Position.shift(1)
Position = Position.dropna()
Portfolio_return = pd.Series()
for date in Position.index:
    if not isnan(Return.loc[date, Position[date]].mean()):
        Portfolio_return[date] = Return.loc[date, Position[date]].mean()
    else:
        Portfolio_return[date] = 0
Base = Base[Position.index[0]:]
Net_return = Portfolio_return - Base
Net_value = (Net_return / 100 + 1).cumprod()
# annulized return
annu_r = (Net_value[-1] ** (12/len(Net_value)) - 1)*100
# annulized volatility
annu_vol = Net_return.std() * (12 ** 1/2)
# max_drawdown
max_drawdown = ((Net_value.cummax() - Net_value) /
                Net_value.cummax()).max()*100
Performance = pd.DataFrame(np.array([annu_r, annu_vol, max_drawdown]).reshape(
    1, 3), columns=['年化超额', '年化波动', '最大回撤'], index=['复合策略表现'])
Performance


# %%
ref = pd.read_excel('对照表.xlsx')
ref.set_index('代码', inplace=True)
ref_dict = pd.Series()
for date in Position.index:
    ref_dict[date] = ref.loc[Position[date], '简称'].tolist()
# Position.to_excel('Position.xlsx')
# ref_dict.to_excel('Position(中文).xlsx')

# %%
# 最新一期动量组合
last_volatility_group = Volatility_Group_filtered[-2]
last_volatility_group_df = pd.Series(
    ref.loc[last_volatility_group, '简称'].values)
# 最新一期拥挤度白名单
last_crowd_group_good = Crowd_Group[-2]
last_crowd_group_good_df = pd.Series(
    ref.loc[last_crowd_group_good, '简称'].values)
# 最新一期拥挤度黑名单
last_crowd_group_bad = [
    a for a in Return.columns.tolist() if a not in last_crowd_group_good]
last_crowd_group_bad_df = pd.Series(ref.loc[last_crowd_group_bad, '简称'].values)
# 最新一期持仓
last_positon = Position[-1]
last_positon_df = pd.Series(ref.loc[last_positon, '简称'].values)
# 最新一期拥挤读平均值
order = (Percent.sum() / 6).iloc[-2].sort_values(ascending=False)
order.index = ref.loc[order.index, '简称']

writer = pd.ExcelWriter('【广发金工】行业拥挤度跟踪20211207.xlsx')
df1 = pd.concat([last_volatility_group_df, last_crowd_group_good_df,
                last_crowd_group_bad_df, last_positon_df], axis=1)
df1.columns = ['动量组合', '拥挤读白名单', '拥挤度黑名单', '复合策略持仓']
df1.to_excel(writer, "考虑拥挤度动量组合")

df2 = order
df2.to_excel(writer, "细分行业拥挤度排序")

writer.save()


# %%
performance = []
# 每月超额
Net_return_check = Net_return.copy()
for date in Net_return.index:
    if not len(Position[date]):
        Net_return_check[date] = 0
performance.append(Net_return_check)
# 策略每月
performance.append(Portfolio_return)
# 基准每月
performance.append(Base)
# 累计超额
Net_value_check = (Net_return_check / 100 + 1).cumprod()
performance.append(Net_value_check)
# 策略净值
performance.append((Portfolio_return/100+1).cumprod())
# 基准净值
performance.append((Base/100+1).cumprod())

perform_stats = pd.concat(performance, axis=1)
perform_stats.columns = ['每月超额', '每月策略', '每月基准', '累计超额', '策略净值', '基准净值']
perform_stats.to_excel('表现统计.xlsx')
# %%
