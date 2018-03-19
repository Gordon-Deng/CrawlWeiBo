#-*-coding:utf-8-*- 
#2017-11-08 Gordon gordondeng@foxmail.com

import urllib, requests, re, xlrd, xlwt
from requests.exceptions import RequestException


def getOnePage(url, encoding = 'utf-8', header = {}, cookie = {}, json = False):
    try:
        response = requests.get(url, headers = header, cookies = cookie)
        if response.status_code == 200:
            response.encoding = encoding
            if json:
                return response.json()
            else:
                return response.text
        return None
    except RequestException:
        print('ER:getOnePage', url)
        return None

def getWeiBoName(rdFile, colNum):
    nameList = []
    colCount = 0
    while colCount <= colNum:
        name = readFromXLS(rdFile, colCount)
        nameList.append(name)
        colCount += 2 
    return nameList

def getWeiboURL(name, header, cookie):
    #get weiboUID weiboURL
    url = 'http://s.weibo.com/weibo/' + name
    html = getOnePage(url, header, cookie)
    weiboUID = re.findall(r'action-data=\\\"uid=(\d+)', html)[0]
    weiboURL = 'https://weibo.com/u/' + weiboUID + '?is_hot=1'
    
    #get weiboIDCN, weiboCNURL
    items = re.findall(r'star_detail.*?href=\\"http:\\/\\/weibo\.com\\/([_\w]+)\?refer', html)
    if items != []:
        weiboIDCN = items[0]
    else:
        weiboIDCN = -1
    weiboCNURL = 'https://weibo.com/' + str(weiboIDCN) + '?is_hot=1'
    
    #get weiboPageID
    if weiboIDCN == -1:
        html = getOnePage(weiboURL, header, cookie)
    else:
        html = getOnePage(weiboCNURL, header, cookie)
    items = re.findall("page_id']='(\d+)'", html)
    if items == []:
        weiboPageID = -1
    else:
        weiboPageID = items[0]
    return weiboURL, weiboCNURL, weiboPageID

def readFromXLS(file, col):
    xlsFile = xlrd.open_workbook(file)
    table = xlsFile.sheets()[0]
    items = table.col_values(col)
    while True:
        if items[-1] == '':
            items = items[:-2]
        else:
            break
    return items

def writeToXLS(content, filename):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    booksheet = workbook.add_sheet('Sheet 1', cell_overwrite_ok = True)
    for i, rowVal in enumerate(content):
        for j, colVal in enumerate(rowVal):
            booksheet.write(i, j, colVal)
    workbook.save('%s.xls' % filename)
    return

header = {"User-Agent": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
          "Host": "weibo.com"}
cookie = {"Cookie": 'SINAGLOBAL=375435575531.87726.1508677024809; UOR=www.baidu.com,data.weibo.com,tech.ifeng.com; YF-Page-G0=b9004652c3bb1711215bacc0d9b6f2b5; SSOLoginState=1510125785; SCF=AjUbm3PQrcHv1ElVam7QscuEqmniuTgE4Qyl3l_6NHdRfk9O3qaxnf3S8M79loI_bOp14UAKj8zaGvEcGcTgV-w.; SUB=_2A253BsCKDeRhGeBO7VIT9CbIwzmIHXVUdbVCrDV8PUNbmtANLRXikW8Rssszp--LhPJB-FXeLXnO1e2A-g..; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WFrb8b65iCT0ifykZ.qvATx5JpX5KMhUgL.Foq7So5EShnX1h-2dJLoI7yZdcH4BcyfIBtt; SUHB=0cSXg7f-uqCcpZ; ALF=1541661785; httpsupgrade_ab=SSL; _s_tentry=-; Apache=7480308592542.213.1510125793216; ULV=1510125793224:11:4:1:7480308592542.213.1510125793216:1509798965202; YF-V5-G0=7fb6f47dfff7c4352ece66bba44a6e5a; wb_cusLike_6060248485=N'}

'''
startAccount , startFlag = '', False
rdFile = 'WeiBoData.xlsx'
nameList = getWeiBoName(rdFile, 24)
print(nameList)
for event in nameList:
    for account in event[1:]:
        if account == startAccount:
            startFlag = True
        
        if startFlag:
            weiboURL, weiboCNURL, weiboPageID = getWeiboURL(account, header, cookie)
            print(weiboURL, weiboCNURL, weiboPageID)
    print('hahahha')
'''
col = 26
rdFile = 'WeiBoData.xlsx'
weiboPageIDAccount = []
event = readFromXLS(rdFile, col)
print(event)

eventPageIDAccount = []
eventPageIDAccount.append([event[0], 'Null', 'Null', 'Null'])

for account in event[1:]:
    weiboURL, weiboCNURL, weiboPageID = getWeiboURL(account, header, cookie)
    eventPageIDAccount.append([account, weiboPageID, weiboURL, weiboCNURL])
    print([account, weiboPageID, weiboURL, weiboCNURL])
writeToXLS(eventPageIDAccount, col)
