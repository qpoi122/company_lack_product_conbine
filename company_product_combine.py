# -*- coding: UTF-8 -*-

'''
__author__="zf"
__mtime__ = '2016/11/8/21/38'
__des__: 简单的读取文件
__lastchange__:'2016/11/16'
'''
from __future__ import division
import os
import math
from xlwt import Workbook, Formula
import xlrd
import copy
import types
import time

def is_num(unum):
    try:
        unum + 1
    except TypeError:
        return 0
    else:
        return 1


# 不带颜色的读取
def open_file(content):
    # 打开文件
    global workbook, file_excel
    file_excel = str(content)
    if '.xl' not in file_excel:
        file = (file_excel + '.xls')  # 文件名及中文合理性
        if not os.path.exists(file):  # 判断文件是否存在
            file = (file_excel + '.xlsx')
            if not os.path.exists(file):
                print("文件不存在")
    else:
        file = file_excel
        if not os.path.exists(file):
            print("文件不存在")
    workbook = xlrd.open_workbook(file)
    print('suicce')


def read_allmesg(file_name):
    allmesg = {}
    open_file(file_name)
    Sheetname = workbook.sheet_names()

    for name in range(len(Sheetname)):
        table = workbook.sheets()[name]
        nrows = table.nrows
        # 获取标题行数据
        title = table.row_values(1)
        for n in range(2, nrows):
            a = table.row_values(n)

            allmesg[n] = {}
            for i in range(len(a)):
                if is_num(a[i]) == 1:
                    if math.modf(a[i])[0] == 0 or a[i] == 0:  # 获取数字的整数和小数
                        a[i] = int(a[i])  # 将浮点数化成整数
                allmesg[n][title[i]] = a[i]

        return allmesg



def deal_allmesg(allmesg):
    company_lack = {}
    for no, item in allmesg.items():
        if isinstance(item['欠订单型号'],int):
            item['欠订单型号'] = str(item['欠订单型号'])
        if '/' in item['欠订单型号']:
            lack_product_compose = item['欠订单型号'].split('/')
        else:
            lack_product_compose = [item['欠订单型号']]
        # print (lack_product_compose)

        for x in lack_product_compose:

            company_lack.setdefault(item['客户'], {x:0})
            company_lack[item['客户']].setdefault(x,0)

            company_lack[item['客户']][x] += item['原订单数量']

    return company_lack

def output_mesg(company_lack):
    book = Workbook()
    sheet1 = book.add_sheet(u'1')
    i = 0
    num = 1
    for key, value in company_lack.items():
        for s, d in value.items():

            sheet1.write(i, 0, key)
            sheet1.write(i, num, s)
            sheet1.write(i, num+1, d)
            i = i + 1
        num = 1


    book.save('4.xls')  # 存储excel
    book = xlrd.open_workbook('4.xls')

    print('----------------------------------------------------------------------------------------')
    print('----------------------------------------------------------------------------------------')
    print(u'计算完成')

    print('----------------------------------------------------------------------------------------')

    print('----------------------------------------------------------------------------------------')

    time.sleep(10)




if __name__ == "__main__":
    allmesg = read_allmesg('4.12起欠单')
    company_lack = deal_allmesg(allmesg)
    output_mesg(company_lack)