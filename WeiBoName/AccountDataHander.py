#-*-coding:utf-8-*-
#2017-11-08 Gordon gordondeng@foxmail.com


import time, urllib, requests, re, xlrd, xlwt
from requests.exceptions import RequestException

class SelfException(Exception):
       def __init__(self):
           super(SelfException, self).__init__()

def getWeiboPageID(account, oldPageIDTable):
    for ii in range(len(oldPageIDTable[0])):
        if oldPageIDTable[0][ii] == account:
            return account, oldPageIDTable[1][ii]
    return account, None

def rdWeiboAccountOriginal(file):
    accountTable = []
    xlsFile = xlrd.open_workbook(file)
    table = xlsFile.sheets()[0]
    ncols = table.ncols
    for ii in range(0, ncols, 2):
        account = rdXLSByCol(file, ii)
        accountTable.append(account)
    return accountTable

def rdWeiboPageID(file):
    pageidTable = []
    for ii in [0, 3]:
        account = rdXLSByCol(file, ii)
        pageidTable.append(account)
    return pageidTable

def rdWeiboAccount(file):
    accountList = []
    xlsFile = xlrd.open_workbook(file)
    table = xlsFile.sheets()[0]
    ncols = table.ncols
    for ii in range(0, ncols, 2):
        account = rdXLSByCol(file, ii)
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

rdAccountFile = 'SecondWeiboAccount.xlsx'
rdPageIDFile = 'FirstWeiboPageID-UN-128-177.xls'

'''
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
'''
accountTableOriginal = rdWeiboAccountOriginal(rdAccountFile)
oldPageIDTable = rdWeiboPageID(rdPageIDFile)

allResultTable = []

for event in accountTableOriginal:
    allResultTable.append([event[0], 'PageID'])
    for account in event[1:]:
        result1, result2 = getWeiboPageID(account, oldPageIDTable)
        allResultTable.append([result1, result2])
    print(allResultTable)
wtToXLS(allResultTable, 'TheFuck.xlsx')
