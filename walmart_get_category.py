#encoding=utf-8

import requests
import re
import xlrd
from pyExcelerator import *
# import time

ip = []
data = xlrd.open_workbook(r'ID.xlsx')
table = data.sheets()[0]
nrows = table.nrows
ncols = table.ncols
for i in range(1,nrows):
    a = table.cell(i,0).value
    ip.append(a)

url_header = 'http://walmart.com/ip/'
categorys = []
for key in ip:
    url = url_header+'%d'%key
    print url
    html = requests.get(url)
    content = html.content
    categorysp = re.findall('<span itemprop=“name”>(.*?)</span>',content,re.S)
    C = ''
    for cla in categorysp:
        C = C+'>'+cla
    C = C.replace('amp;','').lstrip('>')
    categorys.append(C)


w = Workbook()     #创建一个工作簿
ws = w.add_sheet('sheet1')
ws.write(0,0,'walmart ID')
ws.write(0,1,'category')
for i in range(1,nrows):
    ws.write(i,0,ip[i-1])
    ws.write(i,1,categorys[i-1])

w.save(r"category.xls")