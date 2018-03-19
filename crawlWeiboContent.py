#-*-coding:utf-8-*-
#2017-11-08 Gordon gordondeng@foxmail.com

import os, time, requests, re, xlrd, xlwt, random, json, csv, codecs
from lxml import etree
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

def mergeList(a, b):
    for ele in b:
        a.append(ele)
    return a

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

def rdXLSTableByRow(file, sheetNum = 0, firstRowFlag = True, printFlag = True, info = ''):
    if not ".xls" in file:
        file += ".xls"
    xlsFile = xlrd.open_workbook(file)
    table = xlsFile.sheets()[sheetNum]
    nRows = table.nrows
    xlsTable = []
    if firstRowFlag:
        start = 0
    else:
        start = 1
    for ii in range(start, nRows):
        row = table.row_values(ii)
        if row != []:
            xlsTable.append(row)
    if printFlag:
        if info != '':
            info = ' ' + info
        print("read successfully:" + file + info)
    return xlsTable

def wtToCSV(biglist, file, printFlag = True, info = ''):
    with open(file, 'w+', newline = '', encoding = 'gb18030') as csvFile:
        writer = csv.writer(csvFile, dialect = 'excel')
        for row in biglist:
            writer.writerow(row)
        if file[-4:] != ".csv":
            file += ".csv"
        if printFlag:
            if info != '':
                info = ' ' + info
            print("write successfully:" + file + info)
        return

#    except Exception as e:
#        print("Write an CSV file to path: %s, Case: %s" % (file, e))

def rdCSVTableByRow(file, firstRowFlag = True, printFlag = True, info = ''):
    if file[-4:] != ".csv":
        file += ".csv"
    with open(file,'r') as csvFile:
        lines = csv.reader(csvFile)
        biglist = []
        for row in lines:
            if row != []:
                biglist.append(row)
    if not firstRowFlag:
        biglist = biglist[1:]
    if printFlag:
        if info != '':
            info = ' ' + info
        print("read successfully:" + file + info)
    return biglist

def getOnePagebarBowen(url, headers, cookies):
    onePagebarBowenList = []
    html = getOnePage(url, headers = headers, cookies = cookies, json = True)
    htmlCode, htmlData = html["code"], html["data"]
    if htmlCode == "100000" and len(htmlData) < 100:
        return []
    selector = etree.HTML(htmlData)
    info = selector.xpath("//div[@action-type='feed_list_item']")
    for ii in range(len(info)):

        bowenURL = "https://weibo.com" + info[ii].xpath(".//div[@class='WB_from S_txt2']/a/@href")[0]
        if '?' in bowenURL:
            bowenURL = bowenURL[:bowenURL.find('?')]
        bowendate = info[ii].xpath(".//div[@class='WB_from S_txt2']/a/@title")[0]
        bowenContent = info[ii].xpath(".//div[@class='WB_text W_f14']")[0].xpath("string()").strip()
        numbers = info[ii].xpath(".//div[@class='WB_handle']//em[2]")
        forwardNum = numbers[1].xpath("string()")
        commentNum = numbers[2].xpath("string()")
        likeNum = numbers[3].xpath("string()")
        try:
            deviceSource = info[ii].xpath(".//div[@class='WB_from S_txt2']/a[2]/text()")[0]
        except:
            deviceSource = ''

        if "...展开全文" in bowenContent:
            bowenID = info[ii].xpath(".//div[@class='WB_from S_txt2']/a/@name")[0]
            bowenIDURL = 'https://weibo.com/p/aj/mblog/getlongtext?ajwvr=6&mid={}'.format(bowenID)
            bowenIDHTML = getOnePage(bowenIDURL, headers = headers, cookies = cookies, json = True)
            if bowenIDHTML["code"] == "100000":
                bowenIDSelector = etree.HTML(bowenIDHTML['data']['html'])
                bowenContent = bowenIDSelector.xpath("string()")
        bowenContent.replace("????????", '')
        sBowenContent = ""
        isForward = info[ii].xpath(".//@isforward")
        if isForward == []:
            isForward = "否"
        elif isForward == ['1']:
            isForward = "是"
            try:
                sBowenContent = info[ii].xpath(".//div[@class='WB_text']")[0].xpath("string()").strip()
            except:
                pass
        else:
            isForward = str(isForward)
        #print(["xvhao", bowenContent, isForward, "zhuanfaneirong", bowendate, forwardNum, commentNum, likeNum, deviceSource, bowenURL])

        oneBowenDict = {}
        oneBowenDict["isForward"] = isForward
        oneBowenDict["sBowenContent"] = sBowenContent
        oneBowenDict["bowenURL"] = bowenURL
        oneBowenDict["bowendate"] = bowendate
        oneBowenDict["bowenContent"] = bowenContent
        oneBowenDict["forwardNum"] = forwardNum
        oneBowenDict["commentNum"] = commentNum
        oneBowenDict["likeNum"] = likeNum
        oneBowenDict["deviceSource"] = deviceSource
        onePagebarBowenList.append(oneBowenDict)
    return onePagebarBowenList

def getOnePageBowen(pageID, page, month, headers, cookies):
    onePageBowenList = []
    pagebar = ["", "&pagebar=0", "&pagebar=1"]
    prePage = [int(page) - 1, page, page]
    for ii in range(3):
        url = "https://weibo.com/p/aj/v6/mblog/mbloglist?ajwvr=6&domain={}&is_all=1{}&id={}&page={}&pre_page={}&stat_date={}"
        url = url.format(str(pageID)[:6], pagebar[ii], pageID, page, prePage[ii], month)
        #print(url)
        onePagebarBowenList = getOnePagebarBowen(url, headers, cookies)
        #if onePagebarBowenList == []:
        #    break
        onePageBowenList = mergeList(onePageBowenList, onePagebarBowenList)
    return onePageBowenList

def getOneMonthBowen(pageID, month, headers, cookies):
    oneMonthBowenList = []
    page = 1
    while True:
        onePageBowenList = getOnePageBowen(pageID, page, month, headers, cookies)
        print(len(onePageBowenList))
        #print(len(onePageBowenList), onePageBowenList)
        if onePageBowenList == []:
            break
        oneMonthBowenList = mergeList(oneMonthBowenList, onePageBowenList)
        page += 1
    return oneMonthBowenList

# 年月按时间倒序排列 ["201712", "201708"]
def getAccountDates(url, headers, cookies):
    html = getOnePage(url, headers = headers, cookies = cookies)
    months = reHTML("action-data=..is_all=1&stat_date=(\d{6})", html)
    return months

def dictlistToBiglist(items):
    biglist = []
    for ii, item in enumerate(items):
        itemList = [str(ii), item["bowenContent"], item["isForward"], item["sBowenContent"], item["bowendate"], item["forwardNum"], item["commentNum"], item["likeNum"], item["deviceSource"], item["bowenURL"]]
        biglist.append(itemList)
    return biglist

def oneAccountAllMonthsBowen(pageID, months, headers, cookies):
    allMonthBowenList = []
    for month in months:
        oneMonthBowenList = getOneMonthBowen(pageID, month, headers, cookies)
        allMonthBowenList = mergeList(allMonthBowenList, oneMonthBowenList)
    allMonthBowenList = dictlistToBiglist(allMonthBowenList)
    return allMonthBowenList


headers = {"User-Agent": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
          "Host": "weibo.com"}
cookies = {"Cookie": "login_sid_t=01382ab0b6510f2c0dcf9f25af91d08b; cross_origin_proto=SSL; _s_tentry=passport.weibo.com; Apache=19731464414.383114.1517303735394; SINAGLOBAL=19731464414.383114.1517303735394; ULV=1517303735400:1:1:1:19731464414.383114.1517303735394:; SSOLoginState=1517303758; SCF=ArIlp8Ka1-FXhws57Lbj05FHQlAkOJgzFZExdtXA1HD32qYsFp1iPdrSXyZAlaI3phPz6N-Y_BwLrL9ehKMmoAM.; SUB=_2A253dEefDeRhGeBK7FYR9SrOyDiIHXVUAD5XrDV8PUNbmtBeLUKskW9NR5tPTyjzg7twbgzoACS6izjgeZtGP-Tu; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9Whk0FgS-iar-.KeDQ0SLHBy5JpX5K2hUgL.FoqXS0B7SKBEe0B2dJLoI7DAUcvbMcHEUJHE; SUHB=0jBYeKFn6GFhDP; ALF=1548839758; un=zwobserver@sina.com; wvr=6"}

#url = "https://weibo.com/p/aj/v6/mblog/mbloglist?ajwvr=6&domain=100306&is_all=1&pagebar=&id=1003061669879400&page=1&pre_page=0&stat_date=201712"

'''
url = "https://weibo.com/u/2649762111?is_all=1"
pageID = 1001062649762111

months = getAccountDates(url, headers, cookies)
print(months)

for month in months:
    oneMonthBowenList = getOneMonthBowen(pageID, month, headers, cookies)
    biglist = dictlistToBiglist(oneMonthBowenList)
    wtToXLS(biglist, month)



#print(["xvhao", bowenContent, isForward, "zhuanfaneirong", bowendate, forwardNum, commentNum, likeNum, deviceSource, bowenURL])
'''


months = ["201209", "201208", "201207"]

account_allData_File = "北京暴雨事件官方微博汇总1207_AllData - 剩余 - 100-6.xls"
account_allData_table = rdXLSTableByRow(account_allData_File, firstRowFlag = False)

path = account_allData_File[:account_allData_File.find("_AllData.csv")] + "_相关微博\\"
if not os.path.isdir(path):
    os.mkdir(path)
    print(str(path))

for pageIDRow in account_allData_table:
    #biglist = [["number", "bowenContent", "isForward", "sBowenContent", "bowendate", "forwardNum", "commentNum", "likeNum", "deviceSource", "bowenURL"]]
    biglist = [["序号", "微博内容", "是否原创", "转发内容", "发布时间", "转发数", "评论数", "点赞数", "设备源", "微博ID"]]
    allMonthBowenList = oneAccountAllMonthsBowen(pageIDRow[7], months, headers, cookies)
    biglist = mergeList(biglist, allMonthBowenList)
    accountPath = path + pageIDRow[0] + "\\"
    if not os.path.isdir(accountPath):
        os.mkdir(accountPath)
    wtToCSV(biglist, accountPath + "{}.csv".format(pageIDRow[0]))


