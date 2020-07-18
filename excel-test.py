import xlrd
import xlwt
from xlutils import copy
import pandas as pd
import numpy as np
import os

def main():
    cleanData()

def cleanData():
    # destFile = '财务报表报送与信息采集（适用未执行新金融准则、新收入准则和新租赁准则的一般企业）.xls'#文件路径
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
    sdf = dictDF.set_index('item')
    print(sdf)

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