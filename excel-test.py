import xlrd
import xlwt
from xlutils import copy
import pandas as pd
import numpy as np
import os
from fuzzywuzzy import fuzz as fz
from fuzzywuzzy import process

def main():
    srcsheet = srcRead()
    destbook = destRead()
    destsheet = assetLiabSheet(destbook)
    writebook = copy.copy(destbook)
    writesheet = writebook.get_sheet(1)
    print(writesheet)
    for i in range(1,53):
        [row,col] = findInTable(srcsheet,str(i))
        if isValueExist(srcsheet,row,col):
            # print(srcsheet.cell_value(row,col+1),srcsheet.cell_value(row,col+2))
            srclabel = getSrcLabel(srcsheet,row,col)
            print('行次='+str(i))
            print('原科目名称：'+srclabel)
            periodEnd = srcsheet.cell_value(row,col+1)
            yearStart = srcsheet.cell_value(row,col+2)
            print('periodEnd = '+str(periodEnd),'yearStart = '+str(yearStart))
            [xpos,ypos] = findInTable(destsheet,srclabel)
            destlabel = destsheet.cell_value(xpos,ypos)
            print('目标表科目：'+destlabel)
            print(xpos,ypos)
            writesheet.write(xpos,ypos+2,periodEnd)
            writesheet.write(xpos,ypos+3,yearStart)
            print('--------------------------')
    writebook.save('output.xls')

def getSrcLabel(sheet,row,col):
    label = sheet.cell_value(row,col-1)
    label = label.strip()
    return label

def isValueExist(sheet,row,col):
    cell1 = sheet.cell(row,col+1)
    cell2 = sheet.cell(row,col+2)
    # value1 = sheet.cell_value(row,col+1)
    # value2 = sheet.cell_value(row,col+2)
    # print(cell1,cell2)
    # print(value1,value2)
    if cell1.ctype == xlrd.XL_CELL_BLANK and cell2.ctype == xlrd.XL_CELL_BLANK:  
        return False
    else:
        return True
    
def findInTable(sheet,key):
    srcrows = sheet.nrows
    srccols = sheet.ncols
    for r in range(0,srcrows):
        for c in range(0,srccols):
            similarity = fz.ratio(str(sheet.cell_value(r,c)),key)
            if similarity > 80:
                print('similarity is '+str(similarity))
                # print('cell value next to num = '+str(sheet.cell_value(r,c+1)))
                # print('found label '+key+' position is '+str(r)+' '+str(c))
                return [r,c]
            else:
                continue
    print('(!) cannot find same label in dest')
    return [0,0]
def destRead():
    destfile = '财务报表报送与信息采集（适用未执行新金融准则、新收入准则和新租赁准则的一般企业）.xls'
    destbook = xlrd.open_workbook(destfile, formatting_info=True)
    return destbook

def assetLiabSheet(destbook):
    destsheet = destbook.sheet_by_name("资产负债表")
    return destsheet

def srcRead():
    srcfile = '阿尔泰.xls'
    srcbook = xlrd.open_workbook(srcfile, formatting_info=True)
    srcsheet = srcbook.sheet_by_index(0)
    return srcsheet

def oldmain():
    pd.set_option('display.max_rows', None)
    sdf = cleanData()
    destFile = '财务报表报送与信息采集（适用未执行新金融准则、新收入准则和新租赁准则的一般企业）.xls'
    rbook = xlrd.open_workbook(destFile, formatting_info=True)
    names = rbook.sheet_names()
    print(names)
    ssheet = rbook.sheet_by_name("资产负债表")
    nrows = ssheet.nrows
    ncols = ssheet.ncols
    print(ssheet.cell_value(5,1))
    print(nrows,ncols)
    for i in range(0,nrows):
        for j in range(0,ncols):
            if fz.ratio(str(ssheet.cell_value(i,j)),"货币资金") > 99:
                print(fz.ratio(ssheet.cell_value(i,j),"货币资金"))
                print("find")
                print(i,j)
                r = i
                c = j
                print(ssheet.cell_value(i,j))
    print('out')
    print(r,c)
    test = sdf.at['货币资金','periodEnd']
    print(test)
    wbook = copy.copy(rbook)
    wsheet = wbook.get_sheet(1)
    wsheet.write(r,c+2,test)
    print(wsheet)
    wbook.save('output.xls')

    # print(sdf)
    # test = sdf.iloc[1]
    # print(test)

    
def cleanData():
    srcFile = '阿尔泰.xls'
    df = pd.read_excel(srcFile)
    findByKey(df, "编制单位")
    findByKey(df, "编制日期")
    df = pd.read_excel(srcFile, header=5)
    df = df.fillna(0)
    df.columns = df.columns.str.replace(' ','')
    df.columns = df.columns.str.strip()
    assetDf = pd.DataFrame(df, columns=["资产","期末余额","年初余额"])
    assetDf.columns = ['item','periodEnd','yearStart']
    liabDf = pd.DataFrame(df, columns=["负债和所有者权益","期末余额.1","年初余额.1"])
    liabDf.columns = ['item','periodEnd','yearStart']
    # print(assetDf)
    # print(liabDf)
    dictDF = pd.concat([assetDf,liabDf])
    # print(dictDF)
    sdf = dictDF
    sdf['item'] = sdf['item'].str.replace(' ','')
    sdf = dictDF.set_index('item')
    # print(sdf)
    return sdf

def findByKey(df, key):
    data = df.values
    for i in data:
        for j in i:
            if key in str(j):
               print(j)
               print(j[j.index('：')+1:])
    return j

def pdread():
    pass

if __name__ == "__main__":
    main()