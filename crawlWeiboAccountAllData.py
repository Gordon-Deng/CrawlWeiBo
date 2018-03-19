#-*-coding:utf-8-*-
#2017-11-08 Gordon gordondeng@foxmail.com

import os, time, requests, re, xlrd, xlwt, random, csv
from requests.exceptions import RequestException


def getOnePage(url, encoding = 'UTF-8', headers = {}, cookies = {}, json = False):
    try:
        response = requests.get(url, headers = headers, cookies = cookies)
        response.encoding = encoding
        if response.status_code != 200:
            print("ER:getOnePage", "status_code is " + str(response.status_code), url)
            return None
        if json:
            return response.json()
        return response.text
    except RequestException:
        print("ER:getOnePage", url)
        return None

def reHTML(patternStr, html, first = False, function = None, functionArguments = ()):
    pattern = re.compile(patternStr, re.S)
    items = re.findall(pattern, html)
    if items == []:
        return '' if first else []
    for ii in range(len(items)):
        if type(items[ii]) == tuple:
            itemStrip = []
            for ele in items[ii]:
                itemStrip.append(str(ele).strip())
            items[ii] = itemStrip
        else:
            items[ii] = str(items[ii]).strip()
    if first:
        return items[0]
    if function != None:
        return function(items, functionArguments)
    return items

def mergeList(a, b):
    for ele in b:
        a.append(ele)
    return a

def wtToCSV(biglist, file, printFlag = True, info = ''):
    with open(file, 'w', newline = '') as csvFile:
        writer = csv.writer(csvFile, dialect = 'excel')
        for row in biglist:
            writer.writerow(row)
        if printFlag:
            if info != '':
                info = ' ' + info
            print("write successfully:" + file + info)
        return

def rdCSVTableByRow(file, firstRowFlag = True, printFlag = True, info = ''):
    if not ".csv" in file:
        file += ".csv"
    with open(file,'r') as csvFile:
        lines = csv.reader(csvFile)
        biglist = []
        for row in lines:
            if row != []:
                biglist.append(row)
    if printFlag:
        if info != '':
            info = ' ' + info
        print("read successfully:" + file + info)
    return biglist

def wtToXLS(biglist, file, sheetName = 'Sheet1', printFlag = True, info = ''):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    booksheet = workbook.add_sheet(sheetName, cell_overwrite_ok = True)
    for i, rowVal in enumerate(biglist):
        for j, colVal in enumerate(rowVal):
            booksheet.write(i, j, colVal)
    file = str(file)
    if not ".xls" in file:
        file += ".xls"
    workbook.save(file)
    if printFlag:
        if info != '':
            info = ' ' + info
        print("write successfully:" + file + info)
    return

def rdXLSTableByRow(file, sheetNum = 0, firstRowFlag = True, printFlag = True, info = '', returnDictListFlag = False):
    if not ".xls" in file:
        file += ".xls"
    xlsFile = xlrd.open_workbook(file)
    sheet = xlsFile.sheets()[sheetNum]
    biglist = []
    if returnDictListFlag:
        title = sheet.row_values(0)
        titleLen = len(title)
        info = "title:" + str(title)
    else:
        if firstRowFlag:
            biglist.append(sheet.row_values(0))
    for ii in range(1, sheet.nrows):
        row = sheet.row_values(ii)
        if row != []:
            if returnDictListFlag:
                rowDcit = {}
                for j in range(titleLen):
                    rowDcit[title[j]] = row[j]
                row = rowDcit
            biglist.append(row)
    if printFlag:
        if info != '':
            info = ' ' + info
        print("read successfully:" + file + info)
    return biglist

def searchWeiboURL(account, headers, cookies):
    url = 'http://s.weibo.com/user/' + account
    html = getOnePage(url, headers, cookies).encode('utf-8').decode('unicode_escape')
    if "pl_common_sassfilter" in html:
        print(account)
        raise canNotFind
    if 'noresult_tit' in html:
        return ''
    pat = 'person_name.*?href="(.*?)\?.*?title="(.*?)"'
    items = reHTML(pat, html) # 'weibo.com\\/u\\/5111239333', 'weibo.com\\/abtcxtw'
    if items != []:
        for item in items:
            if item[1] == account:
                weiboURL = "https://" + item[0].replace("\\/", "/")[2:] + "?is_all=1"
                return weiboURL
    return ''

def findFromDataTable(account, dataTable):
    for row in dataTable:
        if row[0] == account:
            return row
    return []

def getAccountThreeNum(pageID, headers, cookies):
    # https://weibo.com/p/1001062464171992/home?is_all=1
    weiboPageIDURL = "https://weibo.com/p/" + str(pageID) + "/home?is_all=1"
    print("getAccountThreeNum", weiboPageIDURL)
    html = getOnePage(weiboPageIDURL, headers, cookies)
    items = reHTML(r"class=\\\"W_f1\d\\\">(\d+)<", html)
    if items == []:
        followingNum, followerNum, bowenNum = '', '', ''
    else:
        followingNum, followerNum, bowenNum = items[0], items[1], items[2]
    return followingNum, followerNum, bowenNum

def getAccountBasicData(account, headers, cookies, weiboURL = ''):
    if weiboURL == '':
        weiboURL = searchWeiboURL(account, headers, cookies)
    if weiboURL != '':
        html = getOnePage(weiboURL, headers, cookies)
        #print(weiboURL)
        #print(html)
        #exit()
        oID = reHTML('\[\'oid\'\]=\'(\d+)\'', html, first = True)
        pageID = reHTML('\[\'page_id\'\]=\'(\d+)\'', html, first = True)
        items = reHTML(r"class=\\\"W_f1\d\\\">(\d+)<", html)
        if items == []:
            '''
            if pageID != []:
                followingNum, followerNum, bowenNum = getAccountThreeNum(pageID, headers, cookies)
            else:
                followingNum, followerNum, bowenNum = '', '', ''
            '''
            followingNum, followerNum, bowenNum = getAccountThreeNum(pageID, headers, cookies)
        else:
            followingNum, followerNum, bowenNum = items[0], items[1], items[2]
        return [account, weiboURL, oID, pageID, followingNum, followerNum, bowenNum]
    return [account, weiboURL, '', '', '', '', '']

class canNotFind(Exception):
       def __init__(self):
           super(canNotFind, self).__init__()

def getOneFileAccountAllData(account_File, account_allData_lib_File, headers, cookies):
    account_allData_File = account_File[:account_File.find(".xls")] + "_AllData.xls"
    account_table = rdXLSTableByRow(account_File)
    try:
        account_allData_table = rdXLSTableByRow(account_allData_File)
    except FileNotFoundError:
        wtToXLS([["account", "uID", "institution", "province", "provinceID", "weiboURL", "oID", "pageID", "followingNum", "followerNum", "bowenNum"]], account_allData_File, info = 'new file')
        account_allData_table = rdXLSTableByRow(account_allData_File)
    account_alldata_lib_table = rdXLSTableByRow(account_allData_lib_File)

    for accountInfo in account_table[1:]:
        print(accountInfo[0])
        accountInfo[0] = accountInfo[0].strip()
        account, weiboURL = accountInfo[0], accountInfo[5]
        accountAllData = findFromDataTable(account, account_allData_table)
        if accountAllData == []:

            accountAllData = findFromDataTable(account, account_alldata_lib_table)
            if accountAllData == []:
                accountBasicData = getAccountBasicData(account, headers, cookies, weiboURL = weiboURL)
                accountAllData = mergeList(accountInfo[:5], accountBasicData[1:])
                if accountAllData[6] != '' and accountAllData[7] != '' and accountAllData[8] != '' and accountAllData[9] != '' and accountAllData[10] != '':
                    account_alldata_lib_table.append(accountAllData)
                    wtToXLS(account_alldata_lib_table, account_allData_lib_File, printFlag = False)

            print(accountAllData)
            #if accountAllData[7] != '':
            account_allData_table.append(accountAllData)
            wtToXLS(account_allData_table, account_allData_File, printFlag = False)

    print("OK-->" + account_File)


headers = {"User-Agent": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
          "Host": "weibo.com"}
cookies = {
"Cookie": "login_sid_t=01382ab0b6510f2c0dcf9f25af91d08b; cross_origin_proto=SSL; _s_tentry=passport.weibo.com; Apache=19731464414.383114.1517303735394; SINAGLOBAL=19731464414.383114.1517303735394; ULV=1517303735400:1:1:1:19731464414.383114.1517303735394:; SSOLoginState=1517303758; SCF=ArIlp8Ka1-FXhws57Lbj05FHQlAkOJgzFZExdtXA1HD32qYsFp1iPdrSXyZAlaI3phPz6N-Y_BwLrL9ehKMmoAM.; SUB=_2A253dEefDeRhGeBK7FYR9SrOyDiIHXVUAD5XrDV8PUNbmtBeLUKskW9NR5tPTyjzg7twbgzoACS6izjgeZtGP-Tu; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9Whk0FgS-iar-.KeDQ0SLHBy5JpX5K2hUgL.FoqXS0B7SKBEe0B2dJLoI7DAUcvbMcHEUJHE; SUHB=0jBYeKFn6GFhDP; ALF=1548839758; un=zwobserver@sina.com; wvr=6"
}

'''
account_File = '北京暴雨事件官方微博汇总1207.xls'
account_allData_lib_File = 'account_alldata_lib.xls'

getOneFileAccountAllData(account_File, account_allData_lib_File, headers, cookies)
'''


account = "北京朝阳劲松健康教育"
weiboURL = "https://weibo.com/u/3486199513?is_all=1"
accountBasicData = getAccountBasicData(account, headers, cookies, weiboURL = weiboURL)
print(accountBasicData)


# https://weibo.com/jiuzhaigou?is_all=1
# https://weibo.com/u/2117508734?is_all=1
# https://weibo.com/p/1001061803921393/home?is_all=1


# https://weibo.com/u/2649761901?is_all=1
