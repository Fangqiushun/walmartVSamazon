#encoding=utf-8

from lxml import etree
import requests
import re
import xlrd
from pyExcelerator import *
import time

#导入需要搜索的链接
urls = []
data = xlrd.open_workbook(r"keyword_link.xlsx")
table = data.sheets()[0]
nrows = table.nrows
for row in range(1,nrows):
    urls.append(table.cell(row,1).value)

requests.adapters.DEFAULT_RETRIES = 5
print urls
for url in urls:
    '''每个链接获取的资料自成一个excel文件'''
    #定义需要获取的资料
    links = []
    WMIDs = []
    locations = []
    rankings = []
    sellers = []
    prices = []
    reviews = []
    stars = []
    stocks = []
    excel_name = table.cell(urls.index(url)+1,0).value+time.strftime("%Y-%m-%d %H_%M_%S", time.localtime())    #为excel文件命名

    split = url.split('page=2')
    url = split[0]+'page=1'+split[1]
    req = requests.get(url)
    n = int(re.findall('Showing.*?of (.+?) results',req.content)[0].replace(',',''))
    show_num = int(re.findall('Showing (.*?) of.+?results',req.content)[0].replace(',',''))
    page_sum = n/show_num+1
    page_limit = 1000/show_num
    if page_sum>page_limit:
        page_sum = page_limit
    

    for i in range(page_sum):
        url = split[0]+'page=%d'%(i+1)+split[1]
        html = requests.get(url)
        selector = etree.HTML(html.content)
        for j in range(show_num):
            try:
                flink = selector.xpath('//*[@id="tile-container"]/ul/li[%d]/div/a[1]/@href'%(j+1))[0]
                link = 'http://www.walmart.com'+flink
                links.append(link)
                WMID = re.findall('(\d+)',link)[-1]
                WMIDs.append(WMID)
                location = '%d-%d'%((i+1),(j+1))
                locations.append(location)
                ranking = i*show_num+(j+1)
                rankings.append(ranking)
            except:
                continue

            try:
                seller_content = selector.xpath('//*[@id="tile-container"]/ul/li[%d]/div/div[4]/ul/li[2]/text()'%(j+1))[0]
                seller = re.findall('Shipped by (.*)',seller_content)[0]
                if not seller:
                    seller = 'Walmart.com'
            except:
                seller = 'Walmart.com'
            sellers.append(seller)

            path = '//*[@id="tile-container"]/ul/li[%d]/div/div[2]/div/span[@class="price price-display"]'%(j+1)
            path_bak = '//*[@id="tile-container"]/ul/li[%d]/div/div[2]/div/span[@class="price price-display price-not-available"]'%(j+1)
            content0 = selector.xpath(path+'/text()')
            content1 = selector.xpath(path+'/span[2]/text()')
            content2 = selector.xpath(path+'/span[3]/text()')
            if not content0:
                content0 = selector.xpath(path_bak+'/text()')
                content1 = selector.xpath(path_bak+'/span[2]/text()')
                content2 = selector.xpath(path_bak+'/span[3]/text()')
            if content0:
                try:
                    content00 = content0[1].split(',')
                    content0 = content00[0]+content00[1]
                    price = float(content0+content1[0]+content2[0])
                except:
                    price = float(content0[1]+content1[0]+content2[0])
            else:
                price = 'In stores only'
            prices.append(price)

            try:
                review_content = selector.xpath('//*[@id="tile-container"]/ul/li[%d]/div/div[3]/span/span[2]/text()'%(j+1))[0]
                review = int(re.findall('.*?(\d+).*?',review_content)[0])
                print review
            except:
                review = 0
                print review
            reviews.append(review)

            try:
                star_content = selector.xpath('//*[@id="tile-container"]/ul/li[%d]/div/div[3]/span/span[1]/text()'%(j+1))[0]
                star = float(star_content.strip(' stars'))
            except:
                star = 0
            stars.append(star)

            try:
                stock = selector.xpath('//*[@id="tile-container"]/ul/li[%d]/div/div[2]/div/span[@class="price-auxblock"]/div/text()'%(j+1))[0]
            except:
                stock = ''
            stocks.append(stock)


    w = Workbook()     #创建一个工作簿
    ws = w.add_sheet('sheet1')     #创建一个工作表
    ws.write(0,0,'Link')
    ws.write(0,1,'WMID')
    ws.write(0,2,'Time')
    ws.write(0,3,'Locate')
    ws.write(0,4,'Ranking')
    ws.write(0,5,'Seller')
    ws.write(0,6,'Price')
    ws.write(0,7,'Reviews')
    ws.write(0,8,'Star')
    ws.write(0,9,'Search_result_num')
    ws.write(1,9,n)
    ws.write(0,10,'Stock')
    for k in range(len(links)):
        ws.write(k+1,0,links[k])    #在k+2行1列写入links[k]
        ws.write(k+1,1,WMIDs[k])
        ws.write(k+1,2,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        ws.write(k+1,3,locations[k])
        ws.write(k+1,4,rankings[k])
        ws.write(k+1,5,sellers[k])
        ws.write(k+1,6,prices[k])
        ws.write(k+1,7,reviews[k])
        ws.write(k+1,8,stars[k])
        ws.write(k+1,10,stocks[k])
    w.save(r"ranking\%s.xls"%excel_name)     #保存
    print "Next one!"
print "Game over!!!"
