
# coding:utf-8
import xlrd
import numpy as np
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # window 可能要注掉
# plt.rcParams['font.sans-serif'] = ['SimHei'] # window 可能要解开
# excel read
# from matplotlib import pyplot as plt
# from datetime import date,datetime

xlsPath = 'test03.xlsx'  # 文件路径
# 你要定义的x的字段
xField = '年级2'
# 你要定义的y的字段 不填写默认数量
yField = ''


xAxis = '年级'  # x轴显示字体
yAxis = '数量'  # y轴显示字体
title = '每年级的学生数量'  # title显示字体，要英文


# 如果类型（x轴）是数字后面可以补充字段 要英文
xField2 = '年级1'

# 打开excel
book = xlrd.open_workbook(xlsPath)
# print("The number of worksheets is {0}".format(book.nsheets))
# print("Worksheet name(s): {0}".format(book.sheet_names()))

# sheet1 sheet2
sheet1 = book.sheet_by_index(0)  # 通过索引获取表格

rows = sheet1.row_values(2)  # 获取行内容 0是第一行/列

cols = sheet1.col_values(3)  # 获取列内容 0是第一行/列

targetRow = 1


xIndex = 0
yIndex = 0
# ncols是总列数
for i in range(0, sheet1.ncols):
    # print('sheet name:',sheet1.name, ' , 第',targetRow,'条记录:',sheet1.cell_value(0,i),':',sheet1.cell_value(targetRow,i))
    #  cell_value（y,x） y是列数 x 是行数 从零开始
    if sheet1.cell_value(0, i) == xField:
        xIndex = i
        pass
    if yField != '' and sheet1.cell_value(0, i) == yField:
        yIndex = i
        pass

# xSet = set()
xDict = {}
# ncols是总行数
for i in range(0, sheet1.nrows):
    if i != 0:
        xOne = sheet1.cell_value(i, xIndex)
        if xDict.get(str(xOne)) == None:
            xDict[str(xOne)] = 1
        else:
            xDict[str(xOne)] += 1
            pass
        # xSet = xSet.add(sheet1.cell_value(i,xIndex))
        pass
    pass
# print(xDict)
xList = []
yList = []
# 从字典构造数组
for key in xDict:
    # print(str(key)+':'+ str(xDict[key]))
    if key != '':
        xList += [str(key) + xField2]
        yList += [xDict[str(key)]]
    else:
        xList += ['unknown']
        yList += [xDict[str(key)]]
        pass

# x = np.array(xList)
# y = np.array(yList)


plt.xlabel(xAxis)
plt.ylabel(yAxis)
plt.title(title)
# 线性
# plt.plot(x,y)
# # 离散 ro 红色 bo 蓝色 go 绿色
# # plt.plot(x,y,'go')
# # 柱状
plt.bar(xList, yList)

plt.show()
