import xlrd
import xlwt
from xlutils import copy
import pandas as pd
import numpy as np

def main():
    exportFileName = '财务报表报送与信息采集（适用未执行新金融准则、新收入准则和新租赁准则的一般企业）.xls'#文件路径
    originFileName = '阿尔泰.xls'
    df = pd.read_excel(originFileName)
    # print(origin)
    cols = list(df.columns)
    print(cols)


def pdread():
    pass

if __name__ == "__main__":
    main()