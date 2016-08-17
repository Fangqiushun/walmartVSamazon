# coding: utf-8

import requests
import re
import xlrd
from pyExcelerator import *
import time

def get_keywords():
    '''获取关键词'''
    words = []              
    #打开数据表导入关键词
    w = xlrd.open_workbook('words.xlsx')
    table = w.sheets()[0]
    nrows = table.nrows
    for row in range(1,nrows):
        words.append(table.cell(row,0).value)
    return words

def ensure_place():
    '''选择平台'''
    place = raw_input("Please input the place:")
    while place != "walmart" and place != "amazon":
        print "The place can't be found!!! You can only choose 'walmart' or 'amazon'!"
        place = raw_input("Please input the new place:")
    return place


def build_url(words,place):
    '''生成链接'''
    urls = []
    les = []

    for k in range(len(words)):
        if place == "walmart":
            word = words[k].replace("&","%26").replace(",","%2C").replace(" ","%20")
            url = 'http://www.walmart.com/search/?query=%s'%word
        elif place =="amazon":
            word = words[k].replace("&","%26").replace(",","%2C").replace(" ","+")
            url = 'https://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%%3Daps&field-keywords=%s'%word
        urls.append(url)
    return urls
    
def spider(urls,place):
    '''抓取listing数量'''
    results = []
    count = 0
    for url in urls:
        header = {'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:47.0) Gecko/20100101 Firefox/47.0'}
        req = requests.get(url,headers=header)
        try:
            if place == "walmart":
                result = re.findall('Showing.*?of (.*?) results',req.content)[0]
            else:
                try:
                    result = re.findall('of (.*?) results for',req.content)[0]
                except:
                    result = re.findall('>(.*?) results for',req.content)[0]
        except:
            result = 0
        results.append(result)
        count = count + 1
        print "Finish %d"%(count)
    return results

def savedata(words,urls,results,place):
    '''建立excel表存放结果'''
    w = Workbook()     #创建一个工作簿
    ws = w.add_sheet('sheet1')     #创建一个工作表
    ws.write(0,0,'Key_word')   #在k+2行1列写入links[k]
    ws.write(0,1,'Listing')
    ws.write(0,2,'Link')
    ws.write(0,3,'Time')
    for k in range(len(words)):
        ws.write(k+1,0,words[k])   #在k+2行1列写入links[k]
        ws.write(k+1,1,results[k])
        ws.write(k+1,2,urls[k])
        ws.write(k+1,3,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    now = time.strftime("%Y-%m-%d %H_%M_%S", time.localtime())
    w.save(r"listing\%s_key_word_listings_%s.xls"%(place,now)) 
    
if __name__ == "__main__":
    start = time.clock()
    print "Let's go !!!"
    words = get_keywords()
    place = ensure_place()
    urls = build_url(words,place)
    results = spider(urls,place)
    savedata(words,urls,results,place)
    print "Well done!!!"
    end = time.clock()
    print "The function run time is : %.03f seconds" %(end-start)
