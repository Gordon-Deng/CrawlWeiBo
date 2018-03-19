# -*-coding:utf-8-*- 
# 2017-11-08 Gordon gordondeng@foxmail.com

import time, urllib, requests, re, xlrd, xlwt
from requests.exceptions import RequestException


def getOnePage(url, encoding = 'utf-8', header = {}, cookie = {}, json = False):
    try:
        response = requests.get(url, headers = header, cookies = cookie)
        time.sleep(10)
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

class SelfException(Exception):
       def __init__(self):
           super(SelfException, self).__init__()

def reHTML(patternStr, html):
    pattern = re.compile(patternStr, re.S)
    items = re.findall(pattern, html)
    if items == []:
        return None
    return items[0]

def getWeiboURL(name, header, cookie):
    url = 'http://s.weibo.com/user/' + name
    html = getOnePage(url, header, cookie)
    if 'noresult_tit' in html:
        raise SelfException

    patternStr = r'person_detail.*?href=\\"\\/\\/(weibo\.com[\\/u]*\\/[_\w\d]+)\?refer_flag'
    reResult = reHTML(patternStr, html) # 'weibo.com\\/u\\/5111239333', 'weibo.com\\/abtcxtw'
    if 'weibo.com\\/u\\/' in reResult:
        weiboID = reResult[len('weibo.com\\/u\\/'):]
        weiboURL = 'https://weibo.com/u/' + str(weiboID) + '?is_hot=1'
    else:
        weiboID = reResult[len('weibo.com\\/'):]
        weiboURL = 'https://weibo.com/' + weiboID + '?is_hot=1'
    return weiboID, weiboURL

def getWeiboPageID(url, header, cookie):
    html = getOnePage(url, header, cookie) 
    # weiboUID = reHTML(r'\'uid\'\]=\'(\d+)\'', html)
    weiboOID = reHTML(r'\[\'oid\'\]=\'(\d+)\'', html)
    weiboPageID = reHTML(r'\[\'page_id\'\]=\'(\d+)\'', html)
    return weiboOID, weiboPageID

def rdOldWeiboAccountTable(oldFile):
    accountTable = []
    for ii in range(0, 5):
        aCol = rdXLSByCol(oldFile, ii)
        accountTable.append(aCol)
    return accountTable

def findAccountData(account, accountTable):
    for ii in range(len(accountTable[0])):        
        if accountTable[0][ii] == account:
            return ii
    return -1

def rdWeiboAccount(file):
    accountList = []
    xlsFile = xlrd.open_workbook(file)
    table = xlsFile.sheets()[0]
    ncols = table.ncols
    for ii in range(0, ncols, 2):
        account = rdXLSByCol(rdFile, ii)
        for ele in account[1:]:
            if ele not in accountList:
                accountList.append(ele)
    return accountList

def rdXLSByCol(file, col):
    xlsFile = xlrd.open_workbook(file)
    table = xlsFile.sheets()[0]
    aCol = table.col_values(col)
    while aCol[-1] == '':
        aCol = aCol[:-1]
    return aCol

def wtToXLS(content, filename):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    booksheet = workbook.add_sheet('Sheet 1', cell_overwrite_ok = True)
    for i, rowVal in enumerate(content):
        for j, colVal in enumerate(rowVal):
            booksheet.write(i, j, colVal)
    workbook.save('%s.xls' % filename)
    return


header = {"User-Agent": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
          "Host": "weibo.com"}
cookie = {"Cookie": 'SINAGLOBAL=375435575531.87726.1508677024809; httpsupgrade_ab=SSL; wvr=6; SSOLoginState=1510977708; _s_tentry=sass.weibo.com; Apache=6047910936032.892.1510977906168; ULV=1510977906229:13:6:2:6047910936032.892.1510977906168:1510906033827; ULOGIN_IMG=15109821787652; cross_origin_proto=SSL; UOR=www.baidu.com,data.weibo.com,www.baidu.com; SCF=AjUbm3PQrcHv1ElVam7QscuEqmniuTgE4Qyl3l_6NHdRGCbiRHXaWjBi3GPRCXdLp29gMcS9IG5EoZWaOnTjPKw.; SUHB=0vJ48kY7zMHfq7; ALF=1542542812; SUB=_2A253FFPbDeRhGeBO7VIT9CbIwzmIHXVUYMITrDV8PUNbn9BeLWbXkW9p_3fWchrVlJLd6p27rUJD6g84jQ..; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WFrb8b65iCT0ifykZ.qvATx5JpX5KzhUgL.Foq7So5EShnX1h-2dJLoI7yZdcH4BcyfIBtt'}

oldFile = 'FirstWeiboPageID.xls'
rdFile = 'SecondWeiboAccount.xlsx'

accountDict = {}
accountListToWT = []

oldAccountTable = rdOldWeiboAccountTable(oldFile)
accountList = rdWeiboAccount(rdFile)


while accountList != []:
    account = accountList.pop()
    print('***************** start to crawl', account)

    pos = findAccountData(account, oldAccountTable)
    if pos >= 0:
        # ccount, weiboID, weiboOID, weiboPageID, weiboURL = oldAccountTable[0][pos], oldAccountTable[1][pos], accountTable[2][pos], accountTable[3][pos], accountTable[4][pos]
        # accountListToWT.append([account, weiboID, weiboOID, weiboPageID, weiboURL])
        print('Found in oldFile!')
        continue

    try:
        weiboID, weiboURL = getWeiboURL(account, header, cookie)
        weiboOID, weiboPageID = getWeiboPageID(weiboURL, header, cookie)
        print('OK ', weiboID, weiboURL, weiboOID, weiboPageID)
        accountListToWT.append([account, weiboID, weiboOID, weiboPageID, weiboURL])
    except SelfException:
        print('#####$$$$$ noresult_tit !!!!!!!')
        continue
    except TypeError:
        accountList.append(account)
        print('$$$$$$$$$$ waiting (10S)')
        time.sleep(10)
        continue

wtToXLS(accountListToWT, 'SecondWeiboData.xlsx')
