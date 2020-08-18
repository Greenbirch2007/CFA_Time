#! -*- coding:utf-8 -*-

import xlrd
import pandas as pd
import os
import datetime
import re
import time
import pandas as pd
import openpyxl
import pymysql
import requests
from lxml import etree
from requests.exceptions import RequestException
import os
import xlrd
import sys
from selenium import webdriver


def read_xlrd(excelFile):
    data = xlrd.open_workbook(excelFile)
    table = data.sheet_by_index(0)
    dataFile = []
    for rowNum in range(table.nrows):
        dataFile.append(table.row_values(rowNum))

    # # if 去掉表头
    # if rowNum > 0:

    return dataFile


def text_save(filename, data):  # filename为写入CSV文件的路径，data为要写入数据列表.
    file = open(filename, 'a')
    for i in range(len(data)):
        s = str(data[i]).replace('[', '').replace(']', '')  # 去除[],这两行按数据不同，可以选择
        s = s.replace("'", '').replace(',', '') + '\n'  # 去除单引号，逗号，每行末尾追加换行符
        file.write(s)
    file.close()
    print("保存文件成功")


def call_page(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return response.text
        return None
    except RequestException:
        return None

    return html

def RemoveDot(item):
    f_l = []
    for it in item:
        f_str = "".join(it.split(","))
        f_l.append(f_str)

    return f_l

def parse_html(html):

    print(item)
def remove_block(items):
    new_items = []
    for it in items:
        f = "".join(it.split())
        new_items.append(f)
    return new_items





# Your search found 1381 results.
if __name__ == '__main__':
    csv_l = []
    # 默认访问本目录
    excelFile = 'gl.xlsx'
    full_items = read_xlrd(excelFile=excelFile)
    for i in full_items:
        # ['Yemen', '也门']
        one_page = 'https://www.cfainstitute.org/community/membership/directory/pages/results.aspx?FirstName=&LastName=&City=&State=&Province=&Country={0}&cfacharterholderonly=False&financialadvisoronly=False&WPId=g_100a348c_3193_4073_aeb3_9e578436345f&postalcode=&society=&employer=&YearCharterEarned=0&Radius=0&units=&sort=Province'.format(i[0])
        html = call_page(one_page)
        selector = etree.HTML(html)
        conNum = selector.xpath('//*[@id="directory-results-container"]/div[2]/div[1]/span/text()')
        csv_l.append((i[0],i[1],conNum))
        print(i[0],i[1],conNum)

    columns = ["name1", "name2","num"]
    allC_dt = pd.DataFrame(csv_l,columns=columns)
    allC_dt.to_excel("C:\\Users\\Administrator\\Desktop\\CFA_Time\\全球分布分析\\ff.xlsx", index=0)


#
#  name  //a/div/span
#
# url  //a/@href
#
# city //a/ul/li[1]/span