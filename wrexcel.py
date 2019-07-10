#!/usr/bin/python
#-*-coding:utf-8-*-

import xlrd,json
def group(filename,row):
    sheet_group = xlrd.open_workbook(filename).sheet_by_name(u'分组')
    row_group = sheet_group.nrows
    if row < row_group:
        result = json.dumps(sheet_group.row_values(row),encoding = 'UTF-8', ensure_ascii=False)
        return result
    else:
        print u'输入行号大于“分组列表”最大行数'
def account(filename,row):
    sheet_account = xlrd.open_workbook(filename).sheet_by_name(u'分组')
    row_group = sheet_account.nrows
    if row < row_group:
        result = json.dumps(sheet_account.row_values(row),encoding = 'UTF-8', ensure_ascii=False)
        return result
    else:
        print u'输入行号大于“分组列表”最大行数'


def Main():
    filename = u'C:\\Users\\lvzc\\Desktop\\跳板机24X登录分配20190625test.xlsx'  #文件绝对路径
    a = group(filename,200)
    if a != None:
        print a
    # excel1_sheet1 = xlrd.open_workbook(filename).sheet_by_name(u'分组')          #打开文件
    # excel1_sheet3 = xlrd.open_workbook(filename).sheet_by_name(u'账号信息')  # 打开文件
    # row1 = excel1_sheet1.nrows
    # row2 = excel1_sheet3.nrows
    # print '“分组”表内容如下：'
    # for i in range(row1):
    #     rows = json.dumps(excel1_sheet1.row_values(i),encoding = 'UTF-8', ensure_ascii=False)
    #     print rows
    # print '“帐号信息”表内容如下：'
    # for i in range(row2):
    #     rows = json.dumps(excel1_sheet3.row_values(i), encoding='UTF-8', ensure_ascii=False)
    #     print rows
    # cell = excel1.sheets()[0].cell(0,1).value
    # cols = json.dumps(excel1.sheets()[0].col_values(1),encoding = 'UTF-8', ensure_ascii=False)
    # for i in len(excel1.sheets()[0].row(1)):
    #     cell1 = excel1.sheets()[0].cell(i,1).value
    #     print cell1
    #print json.dumps(excel1.sheet_names(),encoding = 'UTF-8', ensure_ascii=False)
    #解决打印中文乱码
    # print cell
    # print cols
if __name__=='__main__':
    Main()