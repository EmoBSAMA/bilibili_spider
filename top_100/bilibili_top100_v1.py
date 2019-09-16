import requests
import re
from bs4 import BeautifulSoup
import xlwt
import xlrd
from xlutils.copy import copy
import time
import datetime


'''
更新流程，先把rank-item提取出来，然后再把每一个item即每一排名单位组个分析
mun--->排名
title--->标题
av--->av号
play--->播放量
danmu--->弹幕量
up_name--->up主名字
'''
#初始化
def restart():
    global num_l,title_l,av_l,play_l,danmu_l,up_name_l,porn_l
    num_l=[]
    title_l=[]
    av_l=[]
    play_l=[]
    danmu_l=[]
    up_name_l=[]
    porn_l=[]

now = time.time()
h=0

#网页选择
headers = {'user-agent': 'my-app/0.0.1'}
name_t =['每小时','每三天','每七天']
ti=[1,3,7]
url = []
xls_name = []
t=3
for i in range(0,len(name_t)):
    xls_name.append('Bilibili_全站_'+name_t[i]+'_TOP100.xls')
    url.append("https://www.bilibili.com/ranking/all/0/1/" + str(ti[i]))

#网页获取
def get_html():
    global headers,url,soup,t
    r = requests.get(url[t],headers=headers)
    r.encoding = "UTF-8"
    html = r.text
    soup = BeautifulSoup(html, 'lxml')
    print("-----------------------")
    print(url[t])
    #print(soup)

#excel读取或创建
def xls_r_or_w():
    global soup,book_w,sheet,t
    try:
        print("xls文件读取中")
        book_r = xlrd.open_workbook(xls_name[t],formatting_info=True)
    except:
        print("读取失败")
        book_w = xlwt.Workbook(encoding="utf-8",style_compression=0)
        pd = False
        print('excel文档生成中')
    else:
        book_w = copy(book_r)
        pd = True

    dt = datetime.datetime.now()
    dt = dt.strftime('%Y-%m-%d %H-%M-%S') 
    print("现在时间： " + dt)

    sheet = book_w.add_sheet(str(dt),cell_overwrite_ok=True)
    sheet.write(0,0,'排名')
    sheet.write(0,1,'标题')
    sheet.write(0,2,'av号')
    sheet.write(0,3,'播放量')
    sheet.write(0,4,'弹幕量')
    sheet.write(0,5,'UP主')
    sheet.write(0,6,'综合得分')


#网页分析
def analyze_html():
    global soup,num_l,title_l,av_l,play_l,danmu_l,up_name_l,porn_l
    rank_item = soup.find_all('li',class_='rank-item')

    for i in range(0,len(rank_item)):
        
        num = rank_item[i].find('div',class_= 'num').string
        
        title = rank_item[i].find('a',class_= 'title').string
        
        av = rank_item[i].find('a',class_='title')
        av = str(av)
        av = re.search('(?<=av)\d*',av)
        av = av.group(0)
        
        play = rank_item[i].find('i',class_='b-icon play').parent
        play = re.search('(?<=i\>).*[^<](?=\<\/span)',str(play))
        play = play.group(0)

        danmu = rank_item[i].find('i',class_='b-icon view').parent
        danmu = re.search('(?<=i\>).*[^<](?=\<\/span)',str(danmu))
        danmu = danmu.group(0)

        up_name = rank_item[i].find('i',class_='b-icon author').parent
        up_name = re.search('(?<=i\>).*[^<](?=\<\/span)',str(up_name))
        up_name = up_name.group(0)

        porn = rank_item[i].find('div',class_='pts')
        porn = porn.contents[0]
        porn = porn.string

        num_l.append(num)
        title_l.append(title)
        av_l.append(av)
        play_l.append(play)
        danmu_l.append(danmu)
        up_name_l.append(up_name)
        porn_l.append(porn)
        #print(num_l[i],title_l[i],av_l[i],play_l[i],danmu_l[i],up_name_l[i],porn_l[i]) 


#键入数值到excel
def xls_input():
    global soup,book_w,xls_name,sheet,num_l,title_l,av_l,play_l,danmu_l,up_name_l,porn_l,t
    print("数据输出中，文件名：" + xls_name[t])
    for i in range(len(num_l)):
        sheet.write(i+1,0,num_l[i])
        sheet.write(i+1,1,title_l[i])
        sheet.write(i+1,2,av_l[i])
        sheet.write(i+1,3,play_l[i])
        sheet.write(i+1,4,danmu_l[i])
        sheet.write(i+1,5,up_name_l[i])
        sheet.write(i+1,6,porn_l[i])

    book_w.save(xls_name[t])

    print("输出完成")
    print("-----------------------\n\n\n")
    soup = BeautifulSoup("",'lxml')


for tim in range(0,2311597363):
    t=3
    restart()
    '''if time.time() - now >= 3600 and time.time() - now < 86400:
        t = 0
        h += 1
        now = time()
    elif  h % 24 == 0:
        t = 1
    elif h % 72 ==0:
        t = 2

    if t == 0 or t == 1 or t == 2:
        get_html()
        xls_r_or_w()
        analyze_html()
        xls_input()
        t = 4
    '''
    if t == 3:
        t = 0
        for fir in range(0,3):
            get_html()
            xls_r_or_w()
            analyze_html()
            xls_input()
            restart()
            h += 1
            t += 1
        t = 4
    
    time.sleep(1800)











