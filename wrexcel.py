#!/usr/bin/python
#-*-coding:utf-8-*-
import xlrd,json
def content(filename,sheet,row): #说明：filename为文件的绝对路径；sheet为表格名，row为行号
    sheet_group = xlrd.open_workbook(filename).sheet_by_name(sheet)
    row_group = sheet_group.nrows
    if row < row_group:
        result = json.dumps(sheet_group.row_values(row),encoding = 'UTF-8', ensure_ascii=False)
        return result
    else:
        print u'输入行号大于“%s”列表最大行数' %sheet

def test():#例子，测试用
    filename = u'C:\\Users\\lvzc\\Desktop\\跳板机24X登录分配20190625test.xlsx'  #文件绝对路径
    a = content(filename,u'账号信息',400)
    if a != None:
        print a
if __name__=='__main__':
    test()