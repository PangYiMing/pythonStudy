# 导入绘图模块
# encoding='utf-8'
import pandas as pd
import numpy as np
import datetime
# df1 = pd.read_excel('D:/work/客服中心/热线工作情况.xlsx', sheet_name='呼入工作量',)
# df2 = pd.read_excel('D:/work/客服中心/热线工作情况.xlsx', sheet_name='呼出工作量',)
# df3 = pd.read_excel('D:/work/客服中心/热线工作情况.xlsx', sheet_name='工作效率',)
# df4 = pd.read_excel('D:/work/客服中心/热线工作情况.xlsx', sheet_name='质检报表',)
df1 = pd.read_excel(
    '/Users/lyndonpang/Downloads/热线工作情况.xlsx', sheet_name='呼入工作量',)
df2 = pd.read_excel(
    '/Users/lyndonpang/Downloads/热线工作情况.xlsx', sheet_name='呼出工作量',)
df3 = pd.read_excel(
    '/Users/lyndonpang/Downloads/热线工作情况.xlsx', sheet_name='工作效率',)
df4 = pd.read_excel(
    '/Users/lyndonpang/Downloads/热线工作情况.xlsx', sheet_name='质检报表',)


df = pd.merge(df1, df2, on=['姓名', '开始时间'], how='outer')
df = pd.merge(df, df3, on=['姓名', '开始时间'], how='left')
df = pd.merge(df, df4, on=['姓名', '开始时间'], how='left')
df = df.replace("--", "0")
df['日期'] = df['开始时间'].str[5:10]
df['工作量'] = df['呼入量']+df['呼出量']
df['在线时长/h'] = df['在线时长']
df['邀评率'] = (df['邀评数_x']+df['邀评数_y'])/(df['呼入量']+df['呼出量'])
df['满意度_x'] = df['满意度_x'].map(lambda x: float(x.strip('%'))/100.0)
df['满意度_y'] = df['满意度_y'].map(lambda x: float(x.strip('%'))/100.0)
df['满意度'] = (df['满意度_x']*df['参评数_x'] + df['满意度_y']
             * df['参评数_y'])/(df['参评数_x']+df['参评数_y'])
df = df[['日期', '姓名', '工作量', '在线时长/h', '质检总得分', '邀评率', '满意度']]


def func(x):
    vals = x.strip().split(':')
    if len(vals) < 3:
        print('****', vals)
        return 0
    h, m, s = vals
    return round(int(h) + int(m)/60 + int(s)/3600, 2)


df['在线时长/h'] = df['在线时长/h'].map(func)


def translateFloatToBaifenbi(x):
    if str(x) == 'nan':
        return '0%'
    return format(x, '.2%')


# # #     # return str(round(x*100, 2))+'%'
# df['邀评率'] = df['邀评率'].map(translateFloatToBaifenbi)
# df['满意度'] = df['满意度'].map(translateFloatToBaifenbi)


# def BaifenbiTranlateFloat(x):
#     return float(x.split('%')[0])/100.0


# def mysum(a, axis=None, dtype=None, out=None, keepdims=np._NoValue,
#           initial=np._NoValue, where=np._NoValue):

#     # if isinstance(a, _gentype):
#     #     # 2018-02-25, 1.15.0
#     #     # warnings.warn(
#     #     #     "Calling np.sum(generator) is deprecated, and in the future will give a different result. "
#     #     #     "Use np.sum(np.fromiter(generator)) or the python sum builtin instead.",
#     #     #     DeprecationWarning, stacklevel=3)
#     # if type(a[0]) == 'str':
#     #     a.str.strip('%').astype(float)/100.0  # Serie
#     #     pass
#     # b = a
#     res = 0.0
#     for item in a:
#         print('\n\n--------item-------------\n\n', item, isinstance(item, str),
#               '\n\n-------------item-----------\n\n')
#         if isinstance(item, str) and item != '--':
#             # b = a.str.strip('%').astype(float)/100.0  # Serie
#             res = res + 1
#             continue
#         res = res + item
#         pass
#     return res
#     # res = sum(a)
#     # a = a.map(translateFloatToBaifenbi)
#     # # print('\n\n--------b-------------\n\n', b,
#     # #       '\n\n-------------b-----------\n\n')
#     # print('\n\n--------a-------------\n\n', a,
#     #       '\n\n-------------a-----------\n\n')
#     # print('\n\n--------res-------------\n\n', res,
#     #       '\n\n-------------res-----------\n\n')
#     # if out is not None:
#     #     out[...] = res
#     #     return out
#     # return res

#     # return _wrapreduction(a, np.add, 'sum', axis, dtype, out, keepdims=keepdims,
#     #                       initial=initial, where=where)


df = df.loc[df['姓名'].str.contains('热')]
table = pd.pivot_table(df, index=["姓名"], values=[
                       "工作量", "在线时长/h", "邀评率", "满意度"], columns=["日期"], aggfunc=[np.sum], fill_value=0, margins=True)
# table = pd.pivot_table(df, index=["姓名"], values=[
#                        "工作量", "在线时长/h", "邀评率", "满意度"], columns=["日期"],
#                        aggfunc={'工作量': np.sum, '在线时长/h': np.sum, '邀评率': np.sum, '满意度': np.sum, }, fill_value=0, margins=True)
# table.groupby(['06/01', '邀评率']).apply(translateFloatToBaifenbi).unstack()

table1 = table.stack(level=1)
table2 = table1.unstack(level=1)
for index in range(table2.shape[1]):
    # if (index+1) % 3 == 0:
    #     print('****', table2.iloc[-1, index])
    #     table2.iloc[:, index] = table2.iloc[:, index].map(
    #         translateFloatToBaifenbi)
    #     print(table2.iloc[:, index].name)
    #     continue
    # if (index+1) % 4 == 0:
    # table2.iloc[:, index] = table2.iloc[:, index].map(
    #     translateFloatToBaifenbi)
    if '满意度' in table2.iloc[:, index].name or '邀评率' in table2.iloc[:, index].name:
        print(table2.iloc[:, index].name)
        table2.iloc[:, index] = table2.iloc[:, index].map(
            translateFloatToBaifenbi)
        pass
    pass
# table2.to_excel('D:/work/客服中心/1.xlsx')
# for item in range:
#     pass


table2.to_excel('/Users/lyndonpang/Downloads/15.xlsx')
