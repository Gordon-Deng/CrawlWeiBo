#coding = utf-8
import json
from json import JSONDecodeError
import requests
import time
import re
from lxml import etree
import traceback
import os
import csv
import xlrd

cookie123 = 'SINAGLOBAL=1272989567732.712.1510386272810; httpsupgrade_ab=SSL; wvr=6; YF-Ugrow-G0=ea90f703b7694b74b62d38420b5273df; SSOLoginState=1512356322; SCF=Ap3ymQJKyCKaFZm5pFfjTB4NpAX8n2twckKaT6qorqY_VxvtaxGY4aZlHZDI7Jm3Gx6djM38UmwB6HQcxv630Uw.; SUB=_2A253IMmzDeRhGeRP41sS8SzFyT-IHXVUV7x7rDV8PUNbmtBeLVmikW9NUBQtzDCDYMdFlXs_57bdvi_rOTC3BTFE; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WF4bidpGZac3VOCPAB1wwy-5JpX5KMhUgL.Fozp1h.0eKz4eoe2dJLoI7D_qg4.BHiDdJ-0; SUHB=04WvWW6Uc6UrEk; ALF=1543892320; _s_tentry=login.sina.com.cn; UOR=,,login.sina.com.cn; Apache=1461117417073.8096.1512356329100; ULV=1512356329174:20:3:2:1461117417073.8096.1512356329100:1512284451878; YF-V5-G0=69afb7c26160eb8b724e8855d7b705c6; wb_cusLike_2189312923=N; YF-Page-G0=061259b1b44eca44c2f66c85297e2f50'
cookie = {'Cookie': cookie123}
a_headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36'}

class Event:
    def __init__(self, nextdate, curdate, lastdate, eventKeyword):
        self.keyword = eventKeyword
        self.lastdate = lastdate
        self.nextdate = nextdate
        self.curdate = curdate
        self.dict = {}
        for id in event_govWeiboID:
            self.dict[id] = {'weibos':[],'likesNum':[],'forwardsNum':[],'yuanchuang':[],'zhuanfaneirong':[],
                            'commentNum':[],'created_at':[],'weiboID':[],'device':[]}

    #获取所有用户在事件前后3个月的微博内容
    def getEventWeibo(self):
        #依次遍历该事件下的每个账号
        for id in event_govWeiboID:
            # 每个账号的domain
            domain = (int)(id/10000000000)
            #每个账号抓取三个月的微博内容
            for stat_date in [self.nextdate,self.curdate,self.lastdate]:
                page = 1
                pageout = False
                section = 0
                jumpMonth = False
                while pageout != True:
                    # 尝试读取
                    try:
                        uniformUrl = 'https://weibo.com/p/aj/v6/mblog/mbloglist?ajwvr=6&domain={}'.format(domain) + '&is_all=1{}&id={}&page={}{}&stat_date={}'
                        # 分3次读取该页微博内容，并判断是否已到最后一页
                        for times in [1,2,3]:
                            # 3次的url
                            if times == 1:
                                urlpage = uniformUrl.format('', id, page, '&pre_page={}'.format(page - 1), stat_date)
                            elif times == 2:
                                urlpage = uniformUrl.format('&pagebar=0', id, page, '&pre_page={}'.format(page), stat_date)
                            elif times == 3:
                                urlpage = uniformUrl.format('&pagebar=1',id,page,'&pre_page={}'.format(page),stat_date)

                            section = times

                            blockCrawled = False # 用于判断该block是否已经读取
                            jsonPage = requests.get(urlpage,cookies=cookie).json()
                            # 判断该账号是否有相关内容
                            HtmlCont = jsonPage['data']
                            selector = etree.HTML(HtmlCont)
                            info = selector.xpath("//div[@action-type='feed_list_item']")
                            # 如果该block未曾读取，采集微博内容
                            for i in range(0, len(info)):

                                DotPattern = r'(?<=\d)\?\?(?=\d\d)'
                                #微博ID(同时进行判断该page页是否在某一次曾经读取)
                                try:
                                    weiboIdPattern = r'/\d+/(.*)?\?'
                                    href = info[i].xpath(".//div[@class='WB_from S_txt2']/a/@href")
                                    weiboid = re.findall(weiboIdPattern, href[0])[0]
                                    if (i == 0) and (weiboid in self.dict[id]['weiboID']):
                                        blockCrawled = True
                                    else:
                                        pass
                                    if blockCrawled == True:
                                        break # 跳出for i in range(0,len(info))
                                    self.dict[id]['weiboID'].append(weiboid)
                                except IndexError:
                                    self.dict[id]['weiboID'].append(-1)

                                #微博内容
                                try:
                                    str_t = info[i].xpath(".//div[@class='WB_text W_f14']")
                                    weibos = str_t[0].xpath("string(.)")
                                    #查询是否有展开全文
                                    thisWbName_t = info[i].xpath(".//div[@class='WB_from S_txt2']/a/@name")
                                    thisWbName = thisWbName_t[0]
                                    if '​​​​...展开全文c' in weibos:
                                        # 构件获取全文url
                                        FullContUrl = 'https://weibo.com/p/aj/mblog/getlongtext?ajwvr=6&mid={}'.format(thisWbName)
                                        FullContJsData = requests.get(FullContUrl, cookies=cookie,headers=a_headers).json()
                                        FullContSelector = etree.HTML(FullContJsData['data']['html'])
                                        FullString = FullContSelector.xpath("string(.)")
                                        weibos = FullString
                                    weibos = re.sub(DotPattern,'.',str(weibos))
                                    self.dict[id]['weibos'].append(weibos)
                                except IndexError:
                                    self.dict[id]['weibos'].append(-1)

                                #转发数
                                try:
                                    str_forwards = info[i].xpath(".//a[contains(@action-data,'allowForward=1')]//em[2]")
                                    forwards = str_forwards[0].xpath("string(.)")
                                    self.dict[id]['forwardsNum'].append(forwards)
                                except IndexError:
                                    self.dict[id]['forwardsNum'].append(-1)

                                #评论数
                                try:
                                    str_commentNum = info[i].xpath(".//a[@action-type='fl_comment']//em[2]")
                                    commentNum = str_commentNum[0].xpath("string(.)")
                                    self.dict[id]['commentNum'].append(commentNum)
                                except IndexError:
                                    self.dict[id]['commentNum'].append(-1)

                                #点赞数
                                try:
                                    str_likes = info[i].xpath(".//a[@action-type='fl_like']//em[2]")
                                    likes = str_likes[0].xpath("string(.)")
                                    self.dict[id]['likesNum'].append(likes)
                                except IndexError:
                                    self.dict[id]['likesNum'].append(-1)

                                #发布时间
                                try:
                                    created_at = info[i].xpath(".//div[@class='WB_from S_txt2']/a[1]/text()")[0]
                                    self.dict[id]['created_at'].append(created_at)
                                except IndexError:
                                    self.dict[id]['created_at'].append(-1)

                                #发布源
                                try:
                                    deviceSource = info[i].xpath(".//div[@class='WB_from S_txt2']/a[2]/text()")[0]
                                    self.dict[id]['device'].append(deviceSource)
                                except IndexError:
                                    self.dict[id]['device'].append(-1)

                                #判断微博内容是否原创
                                try:
                                    # 是否有转发
                                    yuanchuang = info[i].xpath(".//div[@class='WB_feed_expand']")
                                    if yuanchuang == []:
                                        self.dict[id]['yuanchuang'].append('是')
                                        self.dict[id]['zhuanfaneirong'].append('无')
                                    # 有转发
                                    else:
                                        self.dict[id]['yuanchuang'].append('否')
                                        str_zhuanfa = info[i].xpath(".//div[@class='WB_feed_expand']//div[@class='WB_text']")
                                        # 转发是否有内容
                                        if str_zhuanfa == []: #无内容的话
                                            str_empty = info[i].xpath(".//div[@class='WB_feed_expand']//div[@class='WB_empty']")
                                            # 是否删除
                                            if str_empty != []:
                                                self.dict[id]['zhuanfaneirong'].append('微博已被原作者删除')
                                            else:
                                                self.dict[id]['zhuanfaneirong'].append(-1)
                                        else: #转发的有内容
                                            zhuanfaneirong = str_zhuanfa[0].xpath("string(.)")
                                            if '展开全文' in zhuanfaneirong:
                                                forwardSelec = str_zhuanfa[0].xpath("./a[@action-data]/@action-data")
                                                WbIdPattern = r'mid=(.*)'
                                                forwardWbName = re.findall(WbIdPattern, forwardSelec[0])[0]
                                                forwardFullUrl = 'https://weibo.com/p/aj/mblog/getlongtext?ajwvr=6&mid={}'.format(forwardWbName)
                                                forwardFullContJsData = requests.get(forwardFullUrl, cookies=cookie,headers=a_headers).json()
                                                forwardFullContSelector = etree.HTML(forwardFullContJsData['data']['html'])
                                                forwardFullString = forwardFullContSelector.xpath("string(.)")
                                                zhuanfaneirong = forwardFullString
                                            zhuanfaneirong = re.sub(DotPattern,'.',zhuanfaneirong)
                                            self.dict[id]['zhuanfaneirong'].append(zhuanfaneirong)
                                except IndexError:
                                    self.dict[id]['yuanchuang'].append('-1')


                            # 若该block已经读取，跳至下一个block
                            if blockCrawled == True:
                                continue


                            # 尝试判断是否已到最后一页
                            allDone = selector.xpath("//a[@action-type='fl_nextTimeBase']")
                            if allDone == []:
                                pass
                            else:
                                pageout = True
                            if pageout == True:
                                break
                        print('{}_账号“{}”在时间段{}的第{}页微博内容已完成采集'.format(self.keyword,event_govWeiboID[id],stat_date,page))
                        # 若该页微博内容全部正确读取，更新page
                        page += 1
                    except (SyntaxError,requests.RequestException,JSONDecodeError) as err1:
                        print('获取用户{}在时间段{}的第{}页第{}次读取错误'.format(event_govWeiboID[id],stat_date,page,section))
                        print('休息十分钟之后重新获取该页微博')
                        traceback.print_exc()
                        time.sleep(600)
                        continue

                    #若该账号无相关时间的微博
                    except AttributeError as attriErr:
                        if jsonPage['data']=='    ':
                            jumpMonth = True
                        if jumpMonth == True:
                            print('账号“{}”没有在时间段{}的微博，即将查找下一月份'.format(event_govWeiboID[id],stat_date))
                            break

            self.writeCSV(id)

    # write各账号微博
    def writeCSV(self,id):
        # for id in event_govWeiboID:
        try:
            # 测试
            # print(len(self.dict[id]['weibos']),len(self.dict[id]['yuanchuang']),len(self.dict[id]['zhuanfaneirong']),len(self.dict[id]['created_at']),
            #       len(self.dict[id]['forwardsNum']),len(self.dict[id]['commentNum']),len(self.dict[id]['likesNum']),len(self.dict[id]['device']),
            #       len(self.dict[id]['weiboID']))
            accountName = event_govWeiboID[id]
            eventDirName = self.keyword + '_相关微博'
            if os.path.isdir(eventDirName) == False:
                os.mkdir(eventDirName)
            if os.path.isdir(eventDirName + '/' + accountName) == False:
                os.mkdir(eventDirName + '/' + accountName)
            file = open(eventDirName + '/' + accountName + '/' + accountName + '.csv','w+',newline='',encoding='gb18030')
            csv_write = csv.writer(file,dialect='excel')
            csv_write.writerow(['序号','微博内容','是否原创','转发内容','发布时间','转发数','评论数','点赞数','设备源','微博ID'])
            for i in range(len(self.dict[id]['weibos'])):
                csv_write.writerow([i+1,self.dict[id]['weibos'][i],self.dict[id]['yuanchuang'][i],self.dict[id]['zhuanfaneirong'][i],self.dict[id]['created_at'][i],self.dict[id]['forwardsNum'][i],
                                       self.dict[id]['commentNum'][i],self.dict[id]['likesNum'][i],self.dict[id]['device'][i],self.dict[id]['weiboID'][i]])
            file.close()
        except (Exception,UnicodeEncodeError) as writeErr:
            print('write Error:' + str(writeErr))
            traceback.print_exc()

    # 主程序
    def start(self):
        try:
            Event.getEventWeibo(self)
            # Event.writeCSV(self) #已被替代
            print('信息抓取完毕')
            print('------------------------------')
        except Exception as err2:
            print('start err:' + str(err2))

#定义事件账号字典
eventsDict = {
              '北京暴雨': {'account': {},'date':[201208,201207,201206]},
              '九寨沟7.0级地震': {'account': {},'date':[201709,201708,201707]},
              '上海禽流感': {'account': {},'date':[201304,201303,201302]},
              '雅安地震': {'account': {},'date':[201305,201304,201303]},
              '昆明火车站': {'account': {},'date':[201404,201403,201402]},
              '广州火车站暴力袭击': {'account': {},'date':[201406,201405,201404]},
              '东方之星': {'account': {},'date':[201507,201506,201505]},
              '天津滨海新区爆炸': {'account': {},'date':[201509,201508,201507]},
              '深圳山体滑坡': {'account': {},'date':[201601,201512,201511]},
              '江苏盐城龙卷风': {'account': {},'date':[201607,201606,201605]},
              '秦岭隧道重大交通事故': {'account': {},'date':[201709,201708,201707]},
              '山东非法疫苗案': {'account': {},'date':[201604,201603,201602]},
              '上海踩踏事件': {'account': {},'date':[201501,201412,201411]},
              '青岛黄岛中石化输油管爆炸事件':{'account':{},'date':[201312,201311,201310]},
              '昆山铝粉尘爆炸': {'account': {}, 'date': [201409,201408,201407]},
              '四川茂县山体滑坡':{'account':{},'date':[201707,201706,201705]}
              }

#填充一个事件账号的字典内容
def getAccountInfo(eventsName):
    data = xlrd.open_workbook('PageIDSecond.xls')
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    colEvent = 0
    for i in range(ncols):
        if eventsName == table.row(0)[i].value:
            colEvent = i
            break
    for j in range(1,nrows):
        if table.row(j)[colEvent].value == '':
            break
        id = (int)(table.row(j)[colEvent+1].value)
        accountName = table.row(j)[colEvent].value
        eventsDict[eventsName]['account'][id] = accountName


#依次补充完整事件账号字典
getAccountInfo('北京暴雨')
getAccountInfo('九寨沟7.0级地震')
getAccountInfo('上海禽流感')
getAccountInfo('雅安地震')
getAccountInfo('昆明火车站')
getAccountInfo('广州火车站暴力袭击')
getAccountInfo('东方之星')
getAccountInfo('天津滨海新区爆炸')
getAccountInfo('深圳山体滑坡')
getAccountInfo('江苏盐城龙卷风')
getAccountInfo('秦岭隧道重大交通事故')
getAccountInfo('山东非法疫苗案')
getAccountInfo('上海踩踏事件')
getAccountInfo('青岛黄岛中石化输油管爆炸事件')
getAccountInfo('昆山铝粉尘爆炸')
getAccountInfo('四川茂县山体滑坡')


for order in eventsDict:
    event_govWeiboID = eventsDict[order]['account']
    theEvent = Event(eventsDict[order]['date'][0], eventsDict[order]['date'][1], eventsDict[order]['date'][2],order)
    theEvent.start()

