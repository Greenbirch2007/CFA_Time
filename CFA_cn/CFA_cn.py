#! -*- coding:utf-8 -*-


import datetime
import re
import time

import pymysql
import requests
from lxml import etree
from requests.exceptions import RequestException
import os
import xlrd
import sys


def read_xlrd(excelFile):
    data = xlrd.open_workbook(excelFile)
    table = data.sheet_by_index(0)
    dataFile = []
    for rowNum in range(table.nrows):
        dataFile.append(table.row_values(rowNum))

       # # if 去掉表头
       # if rowNum > 0:


    return dataFile


def text_save(filename, data):#filename为写入CSV文件的路径，data为要写入数据列表.
    file = open(filename,'a')
    for i in range(len(data)):
        s = str(data[i]).replace('[','').replace(']','')#去除[],这两行按数据不同，可以选择
        s = s.replace("'",'').replace(',','') +'\n'   #去除单引号，逗号，每行末尾追加换行符
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





def RemoveDot(item):
    f_l = []
    for it in item:

        f_str = "".join(it.split(","))
        f_l.append(f_str)

    return f_l




def remove_block(items):
    new_items = []
    for it in items:
        f = "".join(it.split())
        new_items.append(f)
    return new_items





def insertDB(content):
    connection = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='123456', db='cfa',
                                 charset='utf8mb4', cursorclass=pymysql.cursors.DictCursor)

    cursor = connection.cursor()
    try:


        cursor.executemany('insert into cfa_cn (name,city,GetYear,oneUrl) values (%s,%s,%s,%s)', content)
        connection.commit()
        connection.commit()
        connection.close()
        print('向MySQL中添加数据成功！')
    except TypeError :
        pass




# Your search found 1381 results.
if __name__ == '__main__':
    name_list = []
    city_list = []
    f_url = []
    csv_l =[]

    # 默认访问本目录
    excelFile = 'cn.xlsx'
    full_items = read_xlrd(excelFile=excelFile)
    for i in full_items:
        name_list.append(i[0])
        city_list.append(i[1])

        if i[2][:5] == '?uid=':
            f_url.append("https://www.cfainstitute.org/community/membership/directory/pages/results.aspx" + i[2])
        else:
            pass

    for i1, i2, i3 in zip(name_list, city_list, f_url):
        html = call_page(i3)
        patt = re.compile('<span>Awarded the CFA charter on <strong>(.*?)</strong></span>', re.S)
        year = re.findall(patt, html)
        try:
        
            tu = tuple((i1,i2,year[0],i3))
            print(tu)
            csv_l.append(tu)
        except:
            pass
    print(csv_l)
    insertDB(csv_l)







# name,city,GetYear,oneUrl
# create table cfa_cn(id int not null primary key auto_increment,
# name text,
# city text,
#  GetYear text,
#  oneUrl text
# ) engine=InnoDB  charset=utf8;
