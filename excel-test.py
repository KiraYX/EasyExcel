import xlrd
import xlwt
from xlutils import copy

workbook = xlrd.open_workbook("阿尔泰.xls")
names=workbook.sheet_names()
print(names)
worksheet=workbook.sheet_by_index(0)
print(worksheet.name)

import win32com.client
excel = win32com.client.Dispatch('Excel.Application')
workbook = excel.Workbooks.open(r"C:\Users\think\OneDrive\Kira-Code\EasyExcel\test.xls")
print(workbook)
sheet = workbook.WorkSheets('Sheet2')
print(sheet)
## get the cell value
row,col = 1,1
sheet.Cells(row,col).value


# excel_path='test.xls'#文件路径
# #excel_path=unicode('D:\\测试.xls','utf-8')#识别中文路径
# rbook = xlrd.open_workbook(excel_path,formatting_info=True)#打开文件
# wbook = copy.copy(rbook)#复制文件并保留格式
# w_sheet = wbook.get_sheet(1)#索引sheet表
# row=6
# col=3
# value=20180803
# w_sheet.write(row,col,value)
# wbook.save(excel_path)#保存文件

