#encoding:utf-8
from __future__ import unicode_literals

import urllib

from bs4 import BeautifulSoup
import xlwt
from selenium import webdriver

import response


TOP250 = 'http://movie.douban.com/top250?start=%d&filter=&format=text'
INDEX = 'http://index.baidu.com/?tpl=trend&word=%s'

def createxls(result):
    wbk = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = wbk.add_sheet('sheet 1', cell_overwrite_ok=True)
    for i, row in enumerate(result):
        for j, col in enumerate(row.values()):
            sheet.write(i, j, col)
    wbk.save('baiduindex.xls')

def index(title):
    #soup = BeautifulSoup(response.get_source(INDEX % urllib.quote(title.encode('utf-8'))))
    driver = webdriver.Firefox()
    driver.get(INDEX % urllib.quote(title.encode('gbk')))


def top250():
    result = []
    for i in [0, 50, 100, 150, 200, 250]:
        soup = BeautifulSoup(response.get_source(TOP250 % i)).findAll(attrs={'class': 'item'})
        for movie in soup:
            result.append({
                'movie_id' : movie.find(attrs={'class':'m_order'}).text.strip(),
                'url' : movie.find('a').get('href'),
                'title' : movie.a.text,
                'year' : movie.span.text,
                'rating' : movie.em.text,
                'comments' : movie.find(attrs={'headers':'m_rating_num'}).text.strip()
            })
            #index(movie.a.text.split('/')[0])
    createxls(result)

#top250()
index(u'肖生克的救赎')
