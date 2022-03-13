# -*- coding: utf-8 -*-
"""
Created on Wed Mar  9 16:16:49 2022

@author: Administrator

#用于快手小店数据导出
"""

import requests,time,calendar,glob,os,xlrd,sys
from datetime import date, timedelta

os.environ['REQUESTS_CA_BUNDLE'] =  os.path.join(os.path.dirname(sys.argv[0]), 'cacert.pem')
Today = date.today().strftime("%Y-%m-%d")
#获得月初月末日期
def get_current_month_start_and_end(date):
    """
    年份 date(2017-09-08格式)
    :param date:
    :return:本月第一天日期和本月最后一天日期
    """
    if date.count('-') != 2:
        raise ValueError('- is error')
    year, month = str(date).split('-')[0], str(date).split('-')[1]
    end = calendar.monthrange(int(year), int(month))[1]
    start_date = '%s-%s-01' % (year, month)
    end_date = '%s-%s-%s' % (year, month, end)
    return start_date, end_date
start_date,end_date = get_current_month_start_and_end(Today)
Yesterday = (date.today() + timedelta(days=-1)).strftime("%Y-%m-%d")
yesterday = Yesterday + ' 23:59:59'
yesterday = time.strptime(yesterday, "%Y-%m-%d %H:%M:%S")
yesterday = int(time.mktime(yesterday))*1000+999 #昨天时间戳
MonthEnd = end_date +  ' 23:59:59'
MonthEnd = time.strptime(MonthEnd, "%Y-%m-%d %H:%M:%S")
MonthEnd = int(time.mktime(MonthEnd))*1000+999 #月末时间戳
YesterdayStart = Yesterday+ ' 00:00:00'
YesterdayStart = time.strptime(YesterdayStart, "%Y-%m-%d %H:%M:%S")
YesterdayStart = int(time.mktime(YesterdayStart))*1000
MonthEnd_Start = end_date + ' 00:00:00'
MonthEnd_Start = time.strptime(MonthEnd_Start, "%Y-%m-%d %H:%M:%S")
MonthEnd_Start = int(time.mktime(MonthEnd_Start))*1000


#获取Cookie
def read_cookie_file():
    CookieList = []
    f = open('KuaiShouCookie.txt',encoding = 'utf-8')
    f = f.readlines()
    for i in f:
        if len(i) >200:
            CookieList.append(i)
    return CookieList

def getshopname(cookie):
    InfoUrl = 'https://s.kwaixiaodian.com/rest/app/tts/seller/login/info'
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36',
        'cookie' : cookie,
        'kpf': 'PC_WEB',
        'Referer': 'https://s.kwaixiaodian.com/zone/im/dashboard'}
    res = requests.get(InfoUrl,headers=headers)
    resjson = res.json()
    shopId = resjson['brandShopStatus']['shopId']
    if str(shopId) == '5000029050':
        ShopName = 'OPPO商城官方旗舰店'
        print('当前店铺为：%s'%ShopName)
        shopname = '欢太'
    elif str(shopId) == '5000094778':
        ShopName = 'realme旗舰店'
        print('当前店铺为：%s'%ShopName)
        shopname = '真我'
    elif str(shopId) == '5000175467':
        ShopName = '一加手机旗舰店'
        print('当前店铺为：%s'%ShopName)
        shopname = '一加'
    elif str(shopId) == '1659573367':
        ShopName = 'OPPO旗舰店'
        print('当前店铺为：%s'%ShopName)
        shopname = 'OPPO'
    else :
        shopname = str(shopId)
        print('懒得兼容其它店铺~')
    return shopname

def downloadfiles():
    CookieList = read_cookie_file()
    for cookie in CookieList:
        cookie=cookie.replace('\n', '')
        headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36',
            'cookie' : cookie}
        #满意度-昨日
        myd_zr = 'https://s.kwaixiaodian.com/rest/app/tts/cs/data/v2/satisfaction/detail/download?dateType=1&dateParam=%s&queryType=2'%yesterday
        #满意度-近七天
        myd_qt = 'https://s.kwaixiaodian.com/rest/app/tts/cs/data/v2/satisfaction/detail/download?dateType=2&dateParam=%s&queryType=2'%yesterday
        #满意度-本月
        myd_by = 'https://s.kwaixiaodian.com/rest/app/tts/cs/data/v2/satisfaction/detail/download?dateType=5&dateParam=%s&queryType=2'%MonthEnd
        #回复率-昨日
        hfl_zr = 'https://s.kwaixiaodian.com/rest/pc/cs/b/data/reply/detail/download?dateType=1&dateParam=%s&elementId=&sortField=fiveMinutesRat&ascending=false&queryType=2'%YesterdayStart
        #回复率-近七天
        hfl_qt = 'https://s.kwaixiaodian.com/rest/pc/cs/b/data/reply/detail/download?dateType=2&dateParam=%s&elementId=&sortField=fiveMinutesRat&ascending=false&queryType=2'%YesterdayStart
        #回复率-本月
        hfl_by = 'https://s.kwaixiaodian.com/rest/pc/cs/b/data/reply/detail/download?dateType=5&dateParam=%s&elementId=&sortField=fiveMinutesRat&ascending=false&queryType=2'%MonthEnd_Start
        shopname = getshopname(cookie)
        urldic = {
            myd_zr:'满意度-昨日.xls',
            myd_qt:'满意度-近七天.xls',
            myd_by:'满意度-本月.xls',
            hfl_zr:'回复率-昨日.xls',
            hfl_qt:'回复率-近七天.xls',
            hfl_by:'回复率-本月.xls'
            }
        for url in urldic:
            time.sleep(5)
            res = requests.get(url,headers=headers)
            with open(shopname+'_'+urldic[url],'wb') as f:
                f.write(res.content)
downloadfiles()

myd_zr_List = glob.glob(os.path.join('','*满意度-昨日.xls'))
myd_qt_List = glob.glob(os.path.join('','*满意度-近七天.xls'))
myd_by_List = glob.glob(os.path.join('','*满意度-本月.xls'))
hfl_zr_List = glob.glob(os.path.join('','*回复率-昨日.xls'))
hfl_qt_List = glob.glob(os.path.join('','*回复率-近七天.xls'))
hfl_by_List = glob.glob(os.path.join('','*回复率-本月.xls'))

#满意度-昨日合并
count=0
for myd_zr_xls in myd_zr_List:
    ShopName = myd_zr_xls.split('_')[0]
    book = xlrd.open_workbook(myd_zr_xls)
    sheet = book.sheet_by_index(0)
    for i in range(sheet.nrows):
        if i>0:
            l = sheet.row_values(i)
            l.append(ShopName)
        else:
            if count==0:
                l = sheet.row_values(i)
                count+=1
            else:
                continue
        l = str(l).replace('[', '')
        l = str(l).replace(']', '')
        l = str(l).replace("'", '')
        with open('满意度-昨日合并.csv','a') as f:
            f.write(l+'\n')
#满意度-近七日合并
count=0
for myd_qt_xls in myd_qt_List:
    ShopName = myd_qt_xls.split('_')[0]
    book = xlrd.open_workbook(myd_qt_xls)
    sheet = book.sheet_by_index(0)
    for i in range(sheet.nrows):
        if i>0:
            l = sheet.row_values(i)
            l.append(ShopName)
        else:
            if count==0:
                l = sheet.row_values(i)
                count+=1
            else:
                continue
        l = str(l).replace('[', '')
        l = str(l).replace(']', '')
        l = str(l).replace("'", '')
        with open('满意度-近七日合并.csv','a') as f:
            f.write(l+'\n')
#满意度-本月合并
count=0
for myd_by_xls in myd_by_List:
    ShopName = myd_by_xls.split('_')[0]
    book = xlrd.open_workbook(myd_by_xls)
    sheet = book.sheet_by_index(0)
    for i in range(sheet.nrows):
        if i>0:
            l = sheet.row_values(i)
            l.append(ShopName)
        else:
            if count==0:
                l = sheet.row_values(i)
                count+=1
            else:
                continue
        l = str(l).replace('[', '')
        l = str(l).replace(']', '')
        l = str(l).replace("'", '')
        with open('满意度-本月合并.csv','a') as f:
            f.write(l+'\n')

#回复率-昨日合并
count=0
for hfl_zr_xls in hfl_zr_List:
    ShopName = hfl_zr_xls.split('_')[0]
    book = xlrd.open_workbook(hfl_zr_xls)
    sheet = book.sheet_by_index(0)
    for i in range(sheet.nrows):
        if i>0:
            l = sheet.row_values(i)
            if l[2] =='--':
                break
            else:
                l.append(ShopName)
        else:
            if count==0:
                l = sheet.row_values(i)
                count+=1
            else:
                continue
        l = str(l).replace('[', '')
        l = str(l).replace(']', '')
        l = str(l).replace("'", '')
        with open('回复率-昨日合并.csv','a') as f:
            f.write(l+'\n')
#回复率-近七日合并
count=0
for hfl_qt_xls in hfl_qt_List:
    ShopName = hfl_qt_xls.split('_')[0]
    book = xlrd.open_workbook(hfl_qt_xls)
    sheet = book.sheet_by_index(0)
    for i in range(sheet.nrows):
        if i>0:
            l = sheet.row_values(i)
            if l[2] =='--':
                break
            else:
                l.append(ShopName)
        else:
            if count==0:
                l = sheet.row_values(i)
                count+=1
            else:
                continue
        l = str(l).replace('[', '')
        l = str(l).replace(']', '')
        l = str(l).replace("'", '')
        with open('回复率-近七日合并.csv','a') as f:
            f.write(l+'\n')

#回复率-本月合并
count=0
for hfl_by_xls in hfl_by_List:
    ShopName = hfl_by_xls.split('_')[0]
    book = xlrd.open_workbook(hfl_by_xls)
    sheet = book.sheet_by_index(0)
    for i in range(sheet.nrows):
        if i>0:
            l = sheet.row_values(i)
            if l[2] =='--':
                break
            else:
                l.append(ShopName)
        else:
            if count==0:
                l = sheet.row_values(i)
                count+=1
            else:
                continue
        l = str(l).replace('[', '')
        l = str(l).replace(']', '')
        l = str(l).replace("'", '')
        with open('回复率-本月合并.csv','a') as f:
            f.write(l+'\n')
print('10秒后即将删除.xls文件，取消请直接关闭！')
time.sleep(10)
xlsFiles = glob.glob(os.path.join('','*.xls'))
for xlsFile in xlsFiles:
    os.remove(xlsFile)
print('运行完毕！')