# coding: utf-8

import pickle
import requests
import re
import xlrd
from pyExcelerator import *
import time

def get_keywords():
    words = []              #定义words用于存关键词 
    #打开数据表导入关键词
    w = xlrd.open_workbook('words.xlsx')
    table = w.sheets()[0]
    nrows = table.nrows
    for row in range(1,nrows):
        words.append(table.cell(row,0).value)
    return words

def ensure_place():
    place = raw_input("Please input the place:")
    while place != "walmart" and place != "amazon":
        print "The place can't be found!!! You can only choose 'walmart' or 'amazon'!"
        place = raw_input("Please input the new place:")
    return place

def get_level(place):
    '''读取amazon上search栏上的分类选项及其ID'''
    with open(r'%slevel.pickle'%place,'rb') as d:
        level = pickle.load(d)
    return level

def build_url(words,level,place):
    '''生成链接'''
    urls = []
    les = []
    lee = raw_input("Please input the level you want:")
    
    for le in level:
        if le == lee:
            for k in range(len(words)):
                if place == "walmart":
                    word = words[k].replace("&","%26").replace(",","%2C").replace(" ","%20")
                    url = 'http://www.walmart.com/search/?query=%s&cat_id=%s'%(word,level[le])
                    
                else:
                    word = words[k].replace("&","%26").replace(",","%2C").replace(" ","+")
                    url = 'https://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%%3D%s&field-keywords=%s'%(level[le],word)
        
                urls.append(url)
                les.append(le)
    return urls,les
    
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

def savedata(words,urls,les,results,place):
    '''建立excel表存放结果'''
    w = Workbook()     #创建一个工作簿
    ws = w.add_sheet('sheet1')     #创建一个工作表
    ws.write(0,0,'Key_word')   #在k+2行1列写入links[k]
    ws.write(0,1,'Listing')
    ws.write(0,2,'Link')
    ws.write(0,3,'Level')
    ws.write(0,4,'Time')
    for k in range(len(words)):
        ws.write(k+1,0,words[k])   #在k+2行1列写入links[k]
        ws.write(k+1,1,results[k])
        ws.write(k+1,2,urls[k])
        ws.write(k+1,3,les[k])
        ws.write(k+1,4,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    now = time.strftime("%Y-%m-%d %H_%M_%S", time.localtime())
    w.save(r"listing\%s_key_word_listings_%s.xls"%(place,now)) 
    
if __name__ == "__main__":
    start = time.clock()
    print "Let's go !!!"
    words = get_keywords()
    place = ensure_place()
    level = get_level(place)
    urls,les = build_url(words,level,place)
    results = spider(urls,place)
    savedata(words,urls,les,results,place)
    print "Well done!!!"
    end = time.clock()
    print "The function run time is : %.03f seconds" %(end-start)
