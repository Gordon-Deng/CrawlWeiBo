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



cookie123 = 'SINAGLOBAL=9084902443125.73.1507170861801; wb_cmtLike_2189312923=1; httpsupgrade_ab=SSL; YF-Ugrow-G0=ea90f703b7694b74b62d38420b5273df; YF-V5-G0=572595c78566a84019ac3c65c1e95574; YF-Page-G0=5c7144e56a57a456abed1d1511ad79e8; _s_tentry=login.sina.com.cn; Apache=2176051169746.9531.1510111970777; ULV=1510111970834:9:9:9:2176051169746.9531.1510111970777:1510017154078; wb_cusLike_2189312923=N; login_sid_t=bcd5f937c1e7a949f33e5260b45627b6; cross_origin_proto=SSL; UOR=,,login.sina.com.cn; SSOLoginState=1510140587; SCF=AkbJVAxwLZYISuX21R5-XTBnMiRAYeQNiIrt8LSjLBE0oyNHX4VHJNv-5LGM7Gl1RP_WGCfx7ojmU47scak_noM.; SUB=_2A253Bpr7DeThGeNG61IQ9ijFwz-IHXVUdYszrDV8PUNbmtBeLUnBkW9Okud4udIUUFzMPyBa9PkOrpGvHg..; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9W5kSRNb69ZNoCxUBdiObh605JpX5K2hUgL.Fo-Reh5pSoq41he2dJLoIXnLxKML1-eL1-qLxKqL1KnLB-qLxKBLBonL12BLxK.L1-BLBK5LxKnLBo5LBo2LxKML1-2L1hBLxK.LB-2L1K2LxK-LBKeLB-zt; SUHB=0ElKL-_vr8Zzbj; ALF=1541676586; un=caribbeannotes@126.com; wvr=6; wb_cusLike_5800166983=N'
cookie = {'Cookie':cookie123}
a_headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36'}

class Event:
    def __init__(self,nextdate,curdate,lastdate,eventKeyword):
        self.keyword = eventKeyword
        self.lastdate = lastdate
        self.nextdate = nextdate
        self.curdate = curdate
        self.dict = {}
        for id in event_govWeiboID:
            self.dict[id] = {'weibos':[],'likesNum':[],'forwardsNum':[],
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
                            HtmlCont = jsonPage['data']
                            selector = etree.HTML(HtmlCont)
                            info = selector.xpath("//div[@action-type='feed_list_item']")
                            # 如果该block未曾读取，采集微博内容
                            for i in range(0, len(info)):

                                #微博ID(同时进行判断该page页是否在某一次曾经读取)
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

                                #微博内容
                                str_t = info[i].xpath(".//div[@class='WB_text W_f14']")
                                weibos = str_t[0].xpath("string(.)")
                                self.dict[id]['weibos'].append(weibos)

                                #转发数
                                try:
                                    str_forwards = info[i].xpath(".//a[contains(@action-data,'allowForward=1')]//em[2]")
                                    forwards = str_forwards[0].xpath("string(.)")
                                    self.dict[id]['forwardsNum'].append(forwards)
                                except IndexError:
                                    self.dict[id]['forwardsNum'].append(-1)

                                #评论数
                                str_commentNum = info[i].xpath(".//a[@action-type='fl_comment']//em[2]")
                                commentNum = str_commentNum[0].xpath("string(.)")
                                self.dict[id]['commentNum'].append(commentNum)

                                #点赞数
                                str_likes = info[i].xpath(".//a[@action-type='fl_like']//em[2]")
                                likes = str_likes[0].xpath("string(.)")
                                self.dict[id]['likesNum'].append(likes)

                                #发布时间
                                created_at = info[i].xpath(".//div[@class='WB_from S_txt2']/a[1]/text()")[0]
                                self.dict[id]['created_at'].append(created_at)

                                #发布源
                                deviceSource = info[i].xpath(".//div[@class='WB_from S_txt2']/a[2]/text()")[0]
                                self.dict[id]['device'].append(deviceSource)


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
                    except (Exception,SyntaxError,requests.RequestException,JSONDecodeError) as err1:
                        print('获取用户{}在时间段{}的第{}页第{}次读取错误'.format(event_govWeiboID[id],stat_date,page,section))
                        print('休息十分钟之后重新获取该页微博')
                        traceback.print_exc()
                        time.sleep(600)
                        continue

    # write各账号微博
    def writeCSV(self):
        for id in event_govWeiboID:
            try:
                accountName = event_govWeiboID[id]
                eventDirName = self.keyword + '_相关微博'
                if os.path.isdir(eventDirName) == False:
                    os.mkdir(eventDirName)
                if os.path.isdir(eventDirName + '/' + accountName) == False:
                    os.mkdir(eventDirName + '/' + accountName)
                file = open(eventDirName + '/' + accountName + '/' + accountName + '.csv','w+',newline='',encoding='gb18030')
                csv_write = csv.writer(file,dialect='excel')
                csv_write.writerow(['序号','微博内容','发布时间','转发数','评论数','点赞数','设备源','微博ID'])
                for i in range(len(self.dict[id]['weibos'])):
                    csv_write.writerow([i+1,self.dict[id]['weibos'][i],self.dict[id]['created_at'][i],self.dict[id]['forwardsNum'][i],
                                           self.dict[id]['commentNum'][i],self.dict[id]['likesNum'][i],self.dict[id]['device'][i],self.dict[id]['weiboID'][i]])
                file.close()
            except (Exception,UnicodeEncodeError) as writeErr:
                print('write Error:' + str(writeErr))
                traceback.print_exc()

    # 主程序
    def start(self):
        try:
            Event.getEventWeibo(self)
            Event.writeCSV(self)
            print('信息抓取完毕')
            print('------------------------------')
        except Exception as err2:
            print('start err:' + str(err2))

#定义事件账号字典
eventsDict = {'九寨沟7.0级地震':{'account':{},'date':[201709,201708,201707]},
              '北京暴雨': {'account': {},'date':[201208,201207,201206]},
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
              '上海踩踏事件': {'account': {},'date':[201501,201412,201411]}}

#填充一个事件账号的字典内容
def getAccountInfo(eventsName):
    data = xlrd.open_workbook('test2.xlsx')
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
getAccountInfo('九寨沟7.0级地震')
getAccountInfo('北京暴雨')
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


for order in eventsDict:
    event_govWeiboID = eventsDict[order]['account']
    theEvent = Event(eventsDict[order]['date'][0],eventsDict[order]['date'][1],eventsDict[order]['date'][2],order)
    theEvent.start()







