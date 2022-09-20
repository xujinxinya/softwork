import os
import asyncio
import re
import openpyxl
#import time
from pyppeteer import launch
from bs4 import BeautifulSoup
#import cProfile


timeList = ['日期']
newList = ['本土新增']
noList = ['本土无症状']
xList = ['香港新增病例']
aList = ['澳门新增病例']
tList = ['台湾新增病例']
Anhui = ['安徽的新增']
Beijing = ['北京的新增']
Chongqing = ['重庆的新增']
Fujian = ['福建的新增']
Gansu = ['甘肃的新增']
Guangdong = ['广东的新增']
Guangxi = ['广西的新增']
Guizhou = ['贵州的新增']
Hainan = ['海南的新增']
Hebei = ['河北的新增']
Henan = ['河南的新增']
Heilongjiang = ['黑龙江的新增']
Hubei = ['湖北的新增']
Hunan = ['湖南的新增']
Jilin = ['吉林的新增']
Jiangsu = ['江苏的新增']
Jiangxi = ['江西的新增']
Liaoning = ['辽宁的新增']
Neimenggu = ['内蒙古的新增']
Ningxia = ['宁夏的新增']
Qinghai = ['青海的新增']
Shandong = ['山东的新增']
Shanxi = ['山西的新增']
Shanxii = ['陕西的新增']
Shanghai = ['上海的新增']
Sichuan = ['四川的新增']
Tianjin = ['天津的新增']
Xizang = ['西藏的新增']
Xinjiang = ['新疆的新增']
Yunnan = ['云南的新增']
Zhejiang = ['浙江的新增']
Anhui2 = ['安徽无症状']
Beijing2 = ['北京无症状']
Chongqing2 = ['重庆无症状']
Fujian2 = ['福建无症状']
Gansu2 = ['甘肃无症状']
Guangdong2 = ['广东无症状']
Guangxi2 = ['广西无症状']
Guizhou2 = ['贵州无症状']
Hainan2 = ['海南无症状']
Hebei2 = ['河北无症状']
Henan2 = ['河南无症状']
Heilongjiang2 = ['黑龙江无症状']
Hubei2 = ['湖北无症状']
Hunan2 = ['湖南无症状']
Jilin2 = ['吉林无症状']
Jiangsu2 = ['江苏无症状']
Jiangxi2 = ['江西无症状']
Liaoning2 = ['辽宁无症状']
Neimenggu2 = ['内蒙古无症状']
Ningxia2 = ['宁夏无症状']
Qinghai2 = ['青海无症状']
Shandong2 = ['山东无症状']
Shanxi2 = ['山西无症状']
Shanxii2 = ['陕西无症状']
Shanghai2 = ['上海无症状']
Sichuan2 = ['四川无症状']
Tianjin2 = ['天津无症状']
Xizang2 = ['西藏无症状']
Xinjiang2 = ['新疆无症状']
Yunnan2 = ['云南无症状']
Zhejiang2 = ['浙江无症状']
prov_list = [Anhui, Beijing, Chongqing, Fujian, Gansu, Guangdong, Guangxi, Guizhou, Hainan, Hebei, Henan, Heilongjiang, Hubei, Hunan, Jilin, Jiangsu, Jiangxi, Liaoning, Neimenggu, Ningxia, Qinghai, Shandong, Shanxi, Shanxii, Shanghai, Sichuan, Tianjin, Xizang, Xinjiang, Yunnan, Zhejiang]
prov_list2 = [Anhui2, Beijing2, Chongqing2, Fujian2, Gansu2, Guangdong2, Guangxi2, Guizhou2, Hainan2, Hebei2, Henan2, Heilongjiang2, Hubei2, Hunan2, Jilin2, Jiangsu2, Jiangxi2, Liaoning2, Neimenggu2, Ningxia2, Qinghai2, Shandong2, Shanxi2, Shanxii2, Shanghai2, Sichuan2, Tianjin2, Xizang2, Xinjiang2, Yunnan2, Zhejiang2]

#将 pyppeteer 的操作封装成 fetchUrl 函数，用于发起网络请求，获取网页源码
async def pyppteer_fetchUrl(url):
    browser = await launch({'headless': False, 'dumpio': True, 'autoClose':True})
    page = await browser.newPage()

    await page.goto(url)
    await asyncio.wait([page.waitForNavigation()])
    str = await page.content()
    await browser.close()
    return str

def fetchUrl(url):
    try:
        return asyncio.get_event_loop().run_until_complete(pyppteer_fetchUrl(url))
    except:
        return asyncio.get_event_loop().run_until_complete(pyppteer_fetchUrl(url))


#把当天新增或者无症状的省份数据添加到对于省份
def pro_new(place,p_list,people):
    k = 0
    for i in place:
        for j in p_list:
            s = j[0]
            if i == s[:-3]:
                j.append(people[k])
                k = k + 1
                break

#当天没有新增或者无症状的省份列表添加'0'
def pro_full(p_list, num):
    for p in p_list:
        if len(p) < num:
            p.append('0')

#通过 getPageUrl 函数构造每一页的 URL 链接
def getPageUrl():
    for page in range(1, 13):
        if page == 1:
            yield 'http://www.nhc.gov.cn/xcs/yqtb/list_gzbd.shtml'
        else:
            url = 'http://www.nhc.gov.cn/xcs/yqtb/list_gzbd_' + str(page) + '.shtml'
            yield url

#通过 getTitleUrl 函数，获取某一页的文章列表中的每一篇文章的标题，链接，和发布日期。
def getTitleUrl(html):
    bsobj = BeautifulSoup(html, 'html.parser')
    titleList = bsobj.find('div', attrs={"class": "list"}).ul.find_all("li")
    for item in titleList:
        link = "http://www.nhc.gov.cn" + item.a["href"]
        title = item.a["title"]
        date = item.span.text
        yield title, link, date


#正则匹配需要的数据
findtime = re.compile("(.*?)0—24时")#日期
findnew = re.compile("新增确诊病例.*?本土病例(.*?)例（")#本土新增
findno = re.compile("新增无症状感染者.*?本土(.*?)例（")#本土无症状
findx = re.compile("香港特别行政区(.*?)例")#香港累计病例
finda = re.compile("澳门特别行政区(.*?)例")#澳门累计病例
findt = re.compile("台湾地区(.*?)例")#台湾累计病例
findprovince = re.compile("新增确诊病例.*?本土病例\d+例（(.*?)）")#各省新增
findprovince2 = re.compile("新增无症状感染者.*?本土\d+例（(.*?)）")#各省无症状
#通过 getContent 函数，获取某一篇文章的正文中的数据
def getContent(html):
    bsobj = BeautifulSoup(html, 'html.parser')
    #正文
    cnt = bsobj.find('div', attrs={"id": "xw_box"}).find_all("p")
    s = ""
    if cnt:
        for item in cnt:
            s += item.text
        time = re.findall(findtime, s)[0]
        new = re.findall(findnew, s)[0]
        no = re.findall(findno, s)[0]
        xiangGang = re.findall(findx, s)[0]
        aoMen = re.findall(finda, s)[0]
        taiWan = re.findall(findt, s)[0]
        province = re.findall(findprovince, s)[0]
        province2 = re.findall(findprovince2, s)[0]
        #各省数据分析统计
        province = province.replace('例', '')
        people = re.findall(r"\d+", province)
        place = re.findall('[\u4e00-\u9fa5]+', province)
        province2 = province2.replace('例', '')
        people2 = re.findall(r"\d+", province2)
        place2 = re.findall('[\u4e00-\u9fa5]+', province2)

        yield new, no, xiangGang, aoMen, taiWan, time, place, people, place2, people2

    return "爬取失败！"

#这次没用到
def saveFile(path, filename, content):
    if not os.path.exists(path):
        os.makedirs(path)

    # 保存文件
    with open(path + filename + ".txt", 'w', encoding='utf-8') as f:
        f.write(content)

#按行导入excel
def create_excel():
    # 创建工作簿
    workbook = openpyxl.Workbook()
    # 创建工作表
    mysheet = workbook.create_sheet("Sheet2")
    # 获取当前工作表（活跃工作表（当前编辑的工作表））
    worksheet = workbook.active
    # 写入数据
    for jj in range(0, len(timeList)):
        worksheet.append([timeList[jj], newList[jj], str(xList[jj]), str(aList[jj]), str(tList[jj]), Anhui[jj], Beijing[jj], Chongqing[jj], Fujian[jj], Gansu[jj] ,Guangdong[jj], Guangxi[jj], Guizhou[jj], Hainan[jj], Hebei[jj], Henan[jj], Heilongjiang[jj], Hubei[jj], Hunan[jj],Jilin[jj], Jiangsu[jj], Jiangxi[jj], Liaoning[jj], Neimenggu[jj], Ningxia[jj], Qinghai[jj], Shandong[jj], Shanxi[jj], Shanxii[jj], Shanghai[jj], Sichuan[jj], Tianjin[jj], Xizang[jj], Xinjiang[jj], Yunnan[jj], Zhejiang[jj]])

    for jj in range(0, len(timeList)):
        mysheet.append([timeList[jj], noList[jj], Anhui2[jj], Beijing2[jj], Chongqing2[jj], Fujian2[jj], Gansu2[jj], Guangdong2[jj], Guangxi2[jj], Guizhou2[jj], Hainan2[jj], Hebei2[jj],Henan2[jj], Heilongjiang2[jj], Hubei2[jj], Hunan2[jj], Jilin2[jj], Jiangsu2[jj], Jiangxi2[jj], Liaoning2[jj], Neimenggu2[jj], Ningxia2[jj], Qinghai2[jj], Shandong2[jj], Shanxi2[jj], Shanxii2[jj], Shanghai2[jj], Sichuan2[jj], Tianjin2[jj], Xizang2[jj], Xinjiang2[jj], Yunnan2[jj], Zhejiang2[jj]])
    # 保存数据
    workbook.save("personal.xlsx")

#主函数
if "__main__" == __name__:
    try:
        for url in getPageUrl():
            s =fetchUrl(url)
            for title, link, date in getTitleUrl(s):
                print(title, link)
                html =fetchUrl(link)
                for new, no, xiangGang, aoMen, taiWan, time, place, people, place2, people2 in getContent(html):
                    #把获取到的数据放入对应的列表
                    newList.append(new)
                    noList.append(no)
                    xList.append(xiangGang)
                    aList.append(aoMen)
                    tList.append(taiWan)
                    timeList.append(time)
                    num_list = len(timeList)
                    pro_new(place, prov_list, people)
                    pro_new(place2, prov_list2, people2)
                    pro_full(prov_list, num_list)
                    pro_full(prov_list2, num_list)
                #输出验证数据是否正确
                print(Anhui, Beijing, Chongqing, Fujian, Gansu, Guangdong, Guangxi, Guizhou, Hunan, Hebei, Henan, Heilongjiang, Hubei, Hunan, Jilin, Jiangsu, Jiangxi, Liaoning, Neimenggu, Ningxia, Qinghai, Shandong, Shanxi, Shanxii, Shanghai, Sichuan, Tianjin, Xizang, Xinjiang, Yunnan, Zhejiang)
                print(timeList)
                print(newList)
                print(noList)
                print(xList)
                print(aList)
                print(tList)
                #saveFile("./infor/", title, content)
                print("-----"*20)
        #算港澳台的单天新增
        for i in range(1, len(xList) - 1):
            xList[i] = int(xList[i]) - int(xList[i + 1])
        for i in range(1, len(aList) - 1):
            aList[i] = int(aList[i]) - int(aList[i + 1])
        for i in range(1, len(tList) - 1):
            tList[i] = int(tList[i]) - int(tList[i + 1])
    #导入excel
        create_excel()
    except:
        create_excel()
