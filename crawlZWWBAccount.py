# -*-coding:utf-8-*- 
# 2017-12-09 Gordon gordondeng@foxmail.com

# crawlZWWBAccount

import requests, re, xlwt
from requests.exceptions import RequestException

def getOnePage(url, encoding = 'UTF-8', headers = {}, cookies = {}):
    try:
        response = requests.get(url, headers = headers, cookies = cookies)
        if response.status_code == 200:
            response.encoding = encoding
            return response
        return None
    except RequestException:
        print("ER:getOnePage", url)
        return None

def reHTML(patternStr, html, first = False):
    pattern = re.compile(patternStr, re.S)
    items = re.findall(pattern, html)
    if items == []:
        return None
    if first:
        return items[0]
    return items

def getProvinceDomain():
    url = "http://bang.weibo.com/zhengwuwb/shengfen/month"
    html = getOnePage(url).text
    pat = "data-url=\"([/\w]+)\".*?\"name\">(.*?)<"
    provinceDomain = reHTML(pat, html)
    return provinceDomain

def postAllAccountHTML(domain, provinceID):
    html = ""
    headers = {"Accept": "application/json",
              "User-Agent": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
              "Accept-Language": "zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",
              "Host": "bang.weibo.com",
              "Origin": "http://bang.weibo.com",
              "Accept-Encoding": "gzip, deflate",
              "Content-Type": "application/x-www-form-urlencoded",
              "Referer": "http://bang.weibo.com".format(domain)}
    url = "http://bang.weibo.com/aj/getvuser"
    for page in range(1, 6):
        data = {"provinceid": "{}".format(provinceID), "cityid": "0", "date": "20170601", "space": "month", "type": "1", "page": "{}".format(page), "pagesize": "20"}
        html = html + str(requests.post(url, data = data, headers = headers).json()["data"]["html"])
    return html

def getAccountOfProvince(domain):
    provinceURL = "http://bang.weibo.com" + domain
    html = getOnePage(provinceURL).text
    pat = "\'provinceid\'\] = (\d+)"
    provinceID = reHTML(pat, html)[0]

    html = postAllAccountHTML(domain, provinceID)
    pat = "data-uid=\"(\d+)\".*?\"name\">(.*?)<.*?class=\"bio\">(.*?)<"
    items = reHTML(pat, html)
    accountList = [["account", "uid", "institution", "province", "provinceID"]]
    for item in items:
        province = domain[11:domain.find('/month')]
        accountList.append([item[1].strip(), item[0].strip(), item[2].strip(), province, provinceID]) # account, uid, institution, province, provinceID
    return accountList

def wtToXLS(content, filename):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    booksheet = workbook.add_sheet('Sheet 1', cell_overwrite_ok = True)
    for i, rowVal in enumerate(content):
        for j, colVal in enumerate(rowVal):
            booksheet.write(i, j, colVal)
    workbook.save('%s.xls' % filename)
    return

def crawlZWWBAccount():
    provinceDomain = getProvinceDomain()
    #print(provinceDomain)

    bigList = [["account", "uid", "institution", "province", "provinceID"]]
    for province in provinceDomain:
        accountOfProvince = getAccountOfProvince(province[0])
        wtToXLS(accountOfProvince, province[1] + "账号")
        print("OK-->" + province[1] + "账号", len(accountOfProvince))
        for item in accountOfProvince[1:]:
            bigList.append(item)
    wtToXLS(bigList, "all_province_accounts")
    print("------>> All is done!", len(bigList))


crawlZWWBAccount()
