# -*- coding: utf-8 -*-
import bs4
import requests
import time  # 引入time，计算下载时间
import xlwings as xw
import os

os.chdir("E:/code/python/file")  # 存放文件位置

a = xw.App(visible=True, add_book=False)
wb = a.books.add()
sht = wb.sheets[0]

sht.range('a1').expand('table').value = ['位置', '总价', '单价', '房屋户型', '所在楼层', '建筑面积', '户型结构', '套内面积', '建筑类型', '房屋朝向',
                                         '建筑结构', '装修情况', '梯户比例', '供暖方式', '配备电梯', '产权年限', '链接']


def open_url(url):
    # 设置重连次数
    requests.adapters.DEFAULT_RETRIES = 15
    # 设置连接活跃状态为False
    s = requests.session()
    s.keep_alive = False  # 在连接时关闭多余连接
    return requests.get(url, headers=
    {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.109 Safari/537.36'}
                        , timeout=10)


host = 'https://zz.lianjia.com/ershoufang/jinshui/pg'
afx = 'co21sf1a4a5/'

detailurl = set()

count = 1  # 初始页
start = time.time()
size = 0
q = 100  # 爬取页数

while count <= q:
    url = host + str(count) + afx
    r = open_url(url)
    soup = bs4.BeautifulSoup(r.text, 'html.parser')
    targets = soup.find_all('a', class_="img")
    for i in targets:
        detailurl.add(i['href'])
    print('\r' + "正在下载：第" + str(count) + '页,' + "已经下载：" + int(count / q * 100) * "█" + "【" + str(
        round(float(count / q) * 100, 2)) + "%" + "】", end="")
    count += 1

count1 = 0
chunk_size = 1024  # 每次块大小为1024
content_size = int(len(detailurl))

line = 1
for i in detailurl:
    line += 1
    soup1 = bs4.BeautifulSoup(open_url(i).text, 'html.parser')
    s = soup1.find("title").text
    title = [s[s.find('郑州') + 4:-6]]
    price = [soup1.find("span", class_="total").text + '万']
    ym2 = [soup1.find("span", class_="unitPriceValue").text]
    IntroContent = [i[4:] for i in list(filter(None, soup1.find_all("div", class_="content")[2].text.split('\n')))]
    sht.range('a' + str(line)).expand('table').value = title + price + ym2 + IntroContent + [i]
    size = size + 1
    print('\r' + "已经下载：" + int(size / content_size * 100) * "█" + "【" + str(
        round(float(size / content_size) * 100, 2)) + "%" + "】",
          end="")

wb.save('金水区.xlsx')  # 文件名
end = time.time()
print("总耗时:" + str(end - start) + "秒")
