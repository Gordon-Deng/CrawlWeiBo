#-*-coding:utf-8-*- 
#2017-11-08 Gordon gordondeng@foxmail.com

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
    #weiboUID = reHTML(r'\'uid\'\]=\'(\d+)\'', html)
    weiboOID = reHTML(r'\[\'oid\'\]=\'(\d+)\'', html)
    weiboPageID = reHTML(r'\[\'page_id\'\]=\'(\d+)\'', html)
    return weiboOID, weiboPageID

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
cookie = {"Cookie": 'SINAGLOBAL=375435575531.87726.1508677024809; httpsupgrade_ab=SSL; _s_tentry=passport.weibo.com; Apache=7186570091724.176.1510906033799; ULV=1510906033827:12:5:1:7186570091724.176.1510906033799:1510125793224; login_sid_t=17df4c829663f4a6bbd26c9b133bbc33; cross_origin_proto=SSL; ULOGIN_IMG=15109311133851; crossidccode=CODE-gz-1EfIdp-14MIPV-CojPptgMOFEHm8oc7b00c; SSOLoginState=1510931170; SCF=AjUbm3PQrcHv1ElVam7QscuEqmniuTgE4Qyl3l_6NHdRWIzLvbrECk4_6dXnb9spZUng6vola03rz_SMGeX6nfc.; SUB=_2A253CotUDeRhGeBO7VIT9CbIwzmIHXVUYfucrDV8PUNbmtBeLRahkW8-T0BFMIZ08qHpp_Qz6BBEc6gegQ..; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WFrb8b65iCT0ifykZ.qvATx5JpX5o275NHD95Qcehq7eoBRShnfWs4DqcjzCJvLw-SQUg.t; SUHB=0102F2n1ba-YNI; ALF=1542467204; wvr=6; UOR=www.baidu.com,data.weibo.com,login.sina.com.cn'}


rdFile = 'SecondWeiboAccount.xlsx'

accountDict = {}
bigAccountList = []
accountList = rdWeiboAccount(rdFile)

while accountList != []:
    account = accountList.pop()
    try:
        print('***************** start to crawl ', account)
        weiboID, weiboURL = getWeiboURL(account, header, cookie)
        weiboOID, weiboPageID = getWeiboPageID(weiboURL, header, cookie)
        print('OK ', weiboID, weiboURL, weiboOID, weiboPageID)
        bigAccountList.append([account, weiboID, weiboOID, weiboPageID, weiboURL])
    except SelfException as e:
        print('#####$$$$$noresult_tit', e)
        continue
    except TypeError:
        accountList.append(account)
        print('$$$$$$$$$$waiting')
        time.sleep(10)
        continue

wtToXLS(bigAccountList, 'SecondWeiboData.xlsx')
