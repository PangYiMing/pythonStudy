# import numpy as np
import pandas as pd
# import os
# import openpyxl
# from openpyxl import Workbook
df = pd.read_excel('/Users/lyndonpang/workspace/pythonTest/pythonStudy/execExcel/周报/sample.xlsx')
# a=lambda x: str(x).split(',')
# df[df.columns[1]]=df[df.columns[1]].astype(str)
def columns_map(x):
    return str(x)
#注意这里传入的是函数名，不带括号
# df["工单编号"] = df["工单编号"].map(columns_map)
# df["订单号"] = df["订单号"].map(columns_map)
# print(df)

