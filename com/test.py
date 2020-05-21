#! /usr/bin/python3
# -*- coding: utf-8 -*-
# @Time : 2020/05/21 15:35
# @File : mergeExcel
# @Software: PyCharm


"""========================
1. 将同一个Excel中的不同的sheet页，汇总到一个新的表格中
================================"""


import xlrd
import xlsxwriter

#源文件
filename = 'C:\\Users\\yibu\\PycharmProjects\\MergeExcel\\input\\testexcel.xlsx'
#目标文件
tarfile = 'C:\\Users\\yibu\\PycharmProjects\\MergeExcel\\output\\out.xlsx'

#读文件
data = xlrd.open_workbook(filename)

#获取sheet页名称;
sheet_names = data.sheet_names()


#指定sheetname
# sheet_names = ['sheet1', 'sheet2']

print(sheet_names)

#新建目标文件
wh = xlsxwriter.Workbook(tarfile)
wadd = wh.add_worksheet('total')


#预设标题加粗
bold = wh.add_format({'bold':1})

tar = []

#读取源文件数据
for sheet_name in sheet_names:
    #获取sheet页的名称
    sheet = data.sheet_by_name(sheet_name)
    #获取表头
    sh_title = data.sheet_by_index(0).row_values(0)
    wadd.write_row('A1',sh_title,bold)
    #获取表的行数
    nrows = sheet.nrows
    #循环打印
    for i in range(nrows):
        #跳过第一行
        if i == 0:
            continue
        # print(sheet.row_values(i))
        tar.append(sheet.row_values(i))

for row_num,row_data in enumerate(tar):
    wadd.write_row(row_num+1,0,row_data)


wh.close()

