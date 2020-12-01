#!/usr/bin/env python3.7
# coding:utf-8
import openpyxl
import os
import sys
import time
import shutil

# 1 先把目标文件复制到目标文件夹下面
# 2 处理需求1 分列
# 3 处理需求2 过滤 删除
# 4 处理需求2 覆盖


def getTargetPathByName(name):
    targetFile = '周报/' + name
    path = sys.argv[0].replace('perfectWeekReport.py', targetFile)
    return path


def getBuildPath():
    bPath = getTargetPathByName(
        '') + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    # 格式化成2016-03-20 11:45:39形式
    return bPath

# 如果存在直接返回true，不存在就创建，返回创建结果


def mkdir(path):
    import os
    #
    # path = path.strip()
    # path = path.rstrip('\\')
    isExists = os.path.exists(path)

    if not isExists:
        os.makedirs(path)
        isExists = os.path.exists(path)
        resolve = '成功' if isExists else '失败'
        print(path+'创建'+resolve)
        return isExists
    return True

# 获取整行


def getTargetRow(ws, id, amount=None):
    nextId = id+1
    rows = ws.iter_rows(id, nextId)
    for row in rows:
        return row
    return None

# 获取整列


def getTargetCol(ws, id, amount=None):
    nextId = id+1
    cols = ws.iter_cols(id, nextId)
    for col in cols:
        return col
    return None


def getColByTitle(ws, title):
    col = None
    row = getTargetRow(ws, 0)
    for cell in row:
        if cell.value == title:
            col = getTargetCol(ws, cell.column)
            break
        pass
    return col


def execData(path, path2):
    book = openpyxl.load_workbook(path)

    ws = book.active

    # 拿到工单分类名称列 根据名字获取整列
    col = getColByTitle(ws, '工单分类名称')
    # 业务类型的column的id
    ywlxId = col[0].column+5
    # 工单分类名称列后加4列 4列问题 1列业务类型 增加2列 省/战区
    ws.insert_cols(col[0].column+1, 7)
    # 新增col以后，更新row
    row = getTargetRow(ws, 0)

    num = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
    secondProbleColIndex = row[col[0].column+1].column
    provinceIndex = row[col[0].column+5].column

    # 设置 这四列到名称
    for index in range(0, 7):
        if index == 4:
            row[col[0].column+index].value = '业务类型'
            pass
        if index == 5:
            row[col[0].column+index].value = '省份'
            pass
        if index == 6:
            row[col[0].column+index].value = '战区'
            pass
        row[col[0].column+index].value = num[index+1]+'级问题'
        pass

    # print(str(ws['A1'].column)+','+str(ws['A1'].row)) # 1,1

    for cell in col:
        if str(cell.value) != 'None':
            sArr = cell.value.split('/')
            index = 0
            for string in sArr:
                if(index >= 4):
                    break
                myrow = getTargetRow(ws, cell.row)

                myrow[cell.column+index].value = string
                # 业务类型 字段为次日达：二级问题为 次日达
                if index == 1 and string == '次日达':
                    myrow[cell.column+4].value = '次日达'
                    pass
                else:
                    myrow[cell.column+4].value = '及时达'
                    pass
                if index == 2 and len(sArr) == 3:
                    myrow[cell.column+index+1].value = string
                    pass
                index += 1

                pass
            pass
        pass

    # 获取为 创建人部门 的列
    col = getColByTitle(ws, '创建人部门')

    # 当 创建人部门 字段为 营销部(15000089)/客服中心二线组(15035424)
    # 删除 二级问题 为 无效咨询 的数据
    # 删除 业态 为 永辉生活 全球潮物 的数据
    for cell in col[::-1]:
        if cell.value == '营销部(15000089)/客服中心二线组(15035424)':
            # 更新
            secondProbleStr = ws.cell(cell.row, secondProbleColIndex).value
            # print('cell.value'+cell.value +
            #       ', secondProbleStr:'+str(secondProbleStr))
            if secondProbleStr == '无效咨询':
                # 删除这一行
                ws.delete_rows(cell.row)
                pass
            ytcol = getColByTitle(ws, '业态')
            print('ytcol[cell.row] val:'+str(ytcol[cell.row-1].value) +
                  ', row:'+str(ytcol[cell.row-1].row)+' and cell.row:'+str(cell.row))
            if '永辉生活' in str(ytcol[cell.row-1].value) or '全球潮物' in str(ytcol[cell.row-1].value):
                ws.delete_rows(cell.row)
                pass
            pass
        pass
    # 业务类型 字段为 全球购：业务中含有 全球潮购
    ytcol = getColByTitle(ws, '业态')
    for ytcell in ytcol[::-1]:
        if '全球潮物' in str(ytcell.value):
            print('ytcell.value:'+str(ytcell.value))
            ws.cell(ytcell.row, ywlxId).value = '全球购'
            pass
        pass
    # TODO 获取另外的book 城市
    book2 = openpyxl.load_workbook(path2)

    ws2 = book2.active
    cityCol = getColByTitle(ws, '城市')
    cityCol2 = getColByTitle(ws2, '城市-简')
    province = getColByTitle(ws2, '省份')
    area = getColByTitle(ws2, '战区')
    for cell in cityCol:
        for idx, cell2 in enumerate(cityCol2):
            if(cell2.value in str(cell.value)):
                ws.cell(cell.row, provinceIndex).value = province[idx].value
                ws.cell(cell.row, provinceIndex+1).value = area[idx].value
                pass
            pass
        pass
    book2.close()
    # ws.append(['1', '2', '3'])
    # print(ws['c3'].column)
    # 删除一行 从0开始
    # ws.delete_rows(0)

    # 添加N列 idx从1开始 amount 指数量
    # ws.insert_cols(2, 2)

    # 获取整行
    # row = getTargetRow(ws, 0)
    # for cell in row:
    #     print(cell.value)
    #     pass

    # 获取整列
    # col = getTargetCol(ws, 1)
    # for cell in col:
    #     print(cell.value)
    #     pass

    # print(rows.values())
    # 分割字符串
    # sArr = 'a/ab/abc'.split('/')
    # print(sArr)

    # print(book.sheetnames)
    # print(ws['A1'].value)
    # 保存文件
    bPath = getBuildPath()
    bfile = bPath + '/' + path.split('/')[-1]
    # 创建文件夹
    mkdir(bPath)
    # 创建文件
    book.save(bfile)
    pass


def copyFileToDefaultPathWithExec(name, name2):
    #  通过周报下面的文件名字，获取目标文件
    path = getTargetPathByName(name)
    path2 = getTargetPathByName(name2)

    # 处理数据
    execData(path, path2)

    pass


copyFileToDefaultPathWithExec('客诉工单.xlsx', '区域信息.xlsx')
# copyFileToDefaultPathWithExec('客诉工单.xlsx')


# #   通过周报具体的路径，获取目标文件
# def getTargetPathByPath(path):

#     pass


# # mkdir('aa')
# # def read_excel_xlrd():
# #     '''Read Excel with xlrd'''
# #     # file
# #     TC_workbook = xlrd.open_workbook(r"NewCreateWorkbook.xls")

# #     # sheet
# #     all_sheets_list = TC_workbook.sheet_names()
# #     print("All sheets name in File:", all_sheets_list)

# #     first_sheet = TC_workbook.sheet_by_index(0)
# #     print("First sheet Name:", first_sheet.name)
# #     print("First sheet Rows:", first_sheet.nrows)
# #     print("First sheet Cols:", first_sheet.ncols)

# #     second_sheet = TC_workbook.sheet_by_name("SheetName_test")
# #     print("Second sheet Rows:", second_sheet.nrows)
# #     print("Second sheet Cols:", second_sheet.ncols)

# #     first_row = first_sheet.row_values(0)
# #     print("First row:", first_row)
# #     first_col = first_sheet.col_values(0)
# #     print("First Column:", first_col)

# #     # cell
# #     cell_value = first_sheet.cell(1, 0).value
# #     print("The 1th method to get Cell value of row 2 & col 1:", cell_value)
# #     cell_value2 = first_sheet.row(1)[0].value
# #     print("The 2th method to get Cell value of row 2 & col 1:", cell_value2)
# #     cell_value3 = first_sheet.col(0)[1].value
# #     print("The 3th method to get Cell value of row 2 & col 1:", cell_value3)


# # if __name__ == "__main__":
# #     read_excel_xlrd()
