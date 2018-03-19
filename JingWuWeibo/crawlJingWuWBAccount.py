# -*-coding:utf-8-*- 
# 2017-12-09 Gordon gordondeng@foxmail.com


import requests, re, xlwt
from requests.exceptions import RequestException

def postAllAccountHTML(url, postType, date):
    html = ""
    headers = {"Accept": "application/json",
              "User-Agent": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
              "Accept-Language": "zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",
              "Host": "bang.weibo.com",
              "Origin": "http://bang.weibo.com",
              "Accept-Encoding": "gzip, deflate",
              "Content-Type": "application/x-www-form-urlencoded",
              #"Referer": "http://bang.weibo.com/zhengwuwb/gongan/month?date={}".format(date)}
              "Referer": url.format(date)}
    url = "http://bang.weibo.com/aj/getvuser"
    for page in range(1, 4):
        data = {"provinceid": "3", "cityid": "0", "date": "{}".format(date), "space": "month", "type": "{}".format(postType), "page": "{}".format(page), "pagesize": "20"}
        html = html + str(requests.post(url, data = data, headers = headers).json()["data"]["html"])
    return html

def reHTML(patternStr, html, first = False, function = None, functionArguments = ()):
    pattern = re.compile(patternStr, re.S)
    items = re.findall(pattern, html)
    if items == []:
        return []
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

def wtToXLS(content, file, sheetName = 'Sheet1', printFlag = True, info = ''):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    booksheet = workbook.add_sheet(sheetName, cell_overwrite_ok = True)
    for i, rowVal in enumerate(content):
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

def getAllMonth():
    monthList = []
    years = ["2017", "2016", "2015"]
    months = ["12", "11", "10", "09", "08", "07", "06", "05", "04", "03", "02", "01"]
    for year in years:
        for month in months:
            monthList.append(year + month + "01")
    return monthList

def getAccountFormHTML(html):
    pat = "data-uid=\"(\d+)\".*?\"name\">(.*?)<.*?class=\"bio\">(.*?)<"
    items = reHTML(pat, html)
    accountList = []
    for item in items:
        accountList.append([item[1], item[0], item[2], "province", "provinceID"]) # account, uid, institution, province, provinceID
    #print(accountList, len(accountList))
    return accountList

def isnew(account, accountList):
    for item in accountList:
        if item[0] == account:
            return False
    return True


bigList = [["account", "uid", "institution", "province", "provinceID"]]

monthList = getAllMonth()

url, postType = "http://bang.weibo.com/zhengwuwb/gongan/month?date={}", "1"

for month in monthList:
    print(month)
    html = postAllAccountHTML(url, postType, month)
    accountList = getAccountFormHTML(html)
    for item in accountList:
        if isnew(item[0], bigList):
            bigList.append(item)
            print(item)
    wtToXLS(bigList, "警务微博账号")


url, postType = "http://bang.weibo.com/zhengwuzs/gongan/month?date={}", "0"

for month in monthList:
    print(month)
    html = postAllAccountHTML(url, postType, month)
    accountList = getAccountFormHTML(html)
    for item in accountList:
        if isnew(item[0], bigList):
            bigList.append(item)
            print(item)
    wtToXLS(bigList, "警务微博账号")


'''
html = postAllAccountHTML("20170301")
accountList = getAccountFormHTML(html)
for item in accountList:
    print(item[0])
'''
# http://bang.weibo.com/zhengwuzs/gongan/month


# url, postType = "http://bang.weibo.com/zhengwuzs/gongan/month?date={}", "0"

