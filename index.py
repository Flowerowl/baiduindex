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
    for i, row in enumerate(result):
        for j, col in enumerate(row.values()):
            sheet.write(i, j, col)
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
    for i, elem in enumerate(gl_driver.find_elements_by_class_name("ftlwhf")):
        index[i] = elem.text
        print elem.text
    return index

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
            result[-1].update(index(int(movie.find(attrs={'class':'m_order'}).text.strip()), movie.a.text.split('/')[0].strip()))
    createxls(result)

top250()
