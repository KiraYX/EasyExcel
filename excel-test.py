import xlrd
import xlwt
from xlutils import copy

workbook = xlrd.open_workbook("阿尔泰.xls")
names = workbook.sheet_names()
print(names)
worksheet=workbook.sheet_by_index(0)
print(worksheet.name)
for r in range(6):
    for c in range(6):
        taxpayer = worksheet.cell_value(r, c)
        print(taxpayer)
        print(r,c)


excel_path='财务报表报送与信息采集（适用未执行新金融准则、新收入准则和新租赁准则的一般企业）.xls'#文件路径
#excel_path=unicode('D:\\测试.xls','utf-8')#识别中文路径
rbook = xlrd.open_workbook(excel_path,formatting_info=True)#打开文件
wbook = copy.copy(rbook)#复制文件并保留格式
w_sheet = wbook.get_sheet(0)#索引sheet表
row=4
col=1
value='G201808033234'
# w_sheet.write(row,col,value)
# wbook.save(excel_path)#保存文件

