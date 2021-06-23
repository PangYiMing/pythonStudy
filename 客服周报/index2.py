# 导入模块
# encoding='utf-8'
import pandas as pd
import numpy as np
import datetime
# 读取表格
df1 = pd.read_excel(
    '/Users/lyndonpang/Downloads/热线工作情况.xlsx', sheet_name='呼入工作量',)
df2 = pd.read_excel(
    '/Users/lyndonpang/Downloads/热线工作情况.xlsx', sheet_name='呼出工作量',)
df3 = pd.read_excel(
    '/Users/lyndonpang/Downloads/热线工作情况.xlsx', sheet_name='工作效率',)
df4 = pd.read_excel(
    '/Users/lyndonpang/Downloads/热线工作情况.xlsx', sheet_name='质检报表',)
# 时间格式转换
TimeSeries = pd.to_datetime(df4['开始时间'], format="%Y-%m-%d %H:%M:%S")
df4['开始时间'] = TimeSeries.dt.strftime("%Y/%m/%d %H:%M:%S")
# 表格连接
df = pd.merge(df1, df2, on=['姓名', '开始时间'], how='outer')
df = pd.merge(df, df3, on=['姓名', '开始时间'], how='left')
df = pd.merge(df, df4, on=['姓名', '开始时间'], how='left')
# 数据清洗整理
df = df.replace(['--', '-'], ['0%', 0])
df = df.loc[df['姓名'].str.contains('热')]
df['日期'] = df['开始时间'].str[5:10]
df['工作量'] = df['呼入量']+df['呼出量']
df['在线时长/h'] = df['在线时长']
df['邀评率'] = (df['邀评数_x']+df['邀评数_y'])/(df['呼入量']+df['呼出量'])
df['满意度_x'] = df['满意度_x'].map(lambda x: float(x.strip('%'))/100.0)
df['满意度_y'] = df['满意度_y'].map(lambda x: float(x.strip('%'))/100.0)
df['满意度'] = (df['满意度_x']*df['参评数_x'] + df['满意度_y']
             * df['参评数_y'])/(df['参评数_x']+df['参评数_y'])
df = df[['日期', '姓名', '工作量', '在线时长/h', '质检总得分', '邀评率', '满意度']]
df['质检总得分'] = round(df['质检总得分'].astype('float'), 1)
# 时间格式转换为小时数


def func(x):
    vals = x.strip().split(':')
    if len(vals) < 3:
        print('****', vals)
        return 0
    h, m, s = vals
    return round(int(h) + int(m)/60 + int(s)/3600, 2)


df['在线时长/h'] = df['在线时长/h'].map(func)
# 数据透视
table = pd.pivot_table(df, index=["姓名"], values=[
                       "工作量", "在线时长/h", "邀评率", "满意度"], columns=["日期"], aggfunc=[np.sum], fill_value=0, margins=True)
# 据表格式转换
table1 = table.stack(level=1)
table2 = table1.unstack(level=1)


def translateFloatToBaifenbi(x):
    if str(x) == 'nan':
        return '0%'
    return format(x, '.2%')


for index in range(table2.shape[1]):
    if '满意度' in table2.iloc[:, index].name or '邀评率' in table2.iloc[:, index].name:
        print(table2.iloc[:, index].name)
        table2.iloc[:, index] = table2.iloc[:, index].map(
            translateFloatToBaifenbi)
        pass
    pass
swapA = []
for index in range(table2.shape[1]):
    if '在线时长/h' in table2.iloc[:, index].name or '满意度' in table2.iloc[:, index].name:
        swapA = table2.iloc[:, index]
        table2.iloc[:, index] = table2.iloc[:, index+1]
        table2.iloc[:, index+1] = swapA
        pass
    pass
print(table2)
table2.to_excel('/Users/lyndonpang/Downloads/18.xlsx')
