#encoding:utf-8
from __future__ import unicode_literals

import urllib

from bs4 import BeautifulSoup
import xlwt
import yaml
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import response

global gl_driver
gl_driver = webdriver.Firefox()

TOP250 = 'http://movie.douban.com/top250?start=%d&filter=&format=text'
INDEX = 'http://index.baidu.com/?tpl=trend&word=%s'

def createxls(result):
    wbk = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = wbk.add_sheet('sheet 1', cell_overwrite_ok=True)
    title = ['ID', '电影名', '年份', '评分', '评论数', '地址', '整体搜索指数',
            '移动搜索指数', '整体同比', '整体环比', '移动同比', '移动环比'
    ]
    for i, t in enumerate(title):
        sheet.write(1, i, t)
    for i, row in enumerate(result):
        sheet.write(i+2, 0, row['movie_id'])
        sheet.write(i+2, 1, row['title'])
        sheet.write(i+2, 2, row['year'])
        sheet.write(i+2, 3, row['rating'])
        sheet.write(i+2, 4, row['comments'])
        sheet.write(i+2, 5, row['url'])
        sheet.write(i+2, 6, row['ztsszs'])
        sheet.write(i+2, 7, row['ydsszs'])
        sheet.write(i+2, 8, row['zttb'])
        sheet.write(i+2, 9, row['zthb'])
        sheet.write(i+2, 10, row['ydtb'])
        sheet.write(i+2, 11, row['ydhb'])
    wbk.save('baiduindex.xls')

def login(times, title):
    gl_driver.implicitly_wait(10)
    gl_driver.get(INDEX % urllib.quote(title.encode('gbk')))
    if times == 1:
        elem = gl_driver.find_element_by_name("userName")
        elem.send_keys(yaml.load(open('config.yaml').read())['username'])
        elem = gl_driver.find_element_by_name("password")
        elem.send_keys(yaml.load(open('config.yaml').read())['password'])
        elem.send_keys(Keys.RETURN)

def index(times, title):
    index = {}
    login(times, title)
    try:
        elem = gl_driver.find_element_by_class_name('gColor1')
        elem.send_keys(Keys.RETURN) # 选中30天
        for i, elem in enumerate(gl_driver.find_elements_by_class_name("ftlwhf")):
            index[i] = elem.text
            print elem.text
        return index
    except:
        print '未找到电影'
        return [None for i in range(12)]

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
            #(登陆次数, 电影名)
            print 'id:',result[-1]['movie_id']
            index_result = index(int(movie.find(attrs={'class':'m_order'}).text.strip()), movie.a.text.split('/')[0].strip())
            result[-1]['ztsszs'] = index_result[6]
            result[-1]['ydsszs'] = index_result[7]
            result[-1]['zttb'] = index_result[8]
            result[-1]['zthb'] = index_result[9]
            result[-1]['ydtb'] = index_result[10]
            result[-1]['ydhb'] = index_result[11]
    createxls(result)

top250()
