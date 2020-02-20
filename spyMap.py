import requests
from bs4 import BeautifulSoup
import bs4
import os
import json
import xlwt
from urllib.parse import urlencode
import xlrd
import time
import xlwings as xw

# 设置请求头，模拟浏览器访问
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36',
}

def getJsonText(area,kw,pn):
    now = int(time.time()*1000)
    data = {
        'newmap': '1',
        'reqflag': 'pcmap',
        'biz': '1',
        'from': 'webmap',
        'da_par': 'direct',
        'pcevaname': 'pc4.1',
        'qt': 's',
        'da_src': 'searchBox.button',
        'wd': area+kw,  # 修改关键字
        'c': '289',
        'src': '0',
        'wd2': '',
        'pn': '0',
        'sug': '0',
        'l': '13',
        'addr':'0',
        'pl_data_type':'hotel',
        'pl_sub_type':'酒店',
        'pl_price_section':'0,+',
        'pl_sort_type':'default',
        'pl_sort_rule':'0',
        'pl_discount2_section':'0,+',
        'pl_groupon_section':'0,+',
        'pl_cater_book_pc_section':'0,+',
        'pl_hotel_book_pc_section':'0,+',
        'pl_ticket_book_flag_section':'0,+',
        'pl_movie_book_section':'0,+',
        'pl_business_type':'hotel',
        'pl_business_id':'',
        'da_src':'pcmappg.poi.page',
        'on_gel':'1',
        'b': '(12194927,4759185;12248303,4805969)',
        'biz_forward': '{"scaler":1,"styles":"pl"}',
        'sug_forward': '',
        'auth': 'MCR5b1yEGc7vgvcCXcx1618LaYIL66T0uxHHxVxVBTTtDpnSCE@@B1GgvPUDZYOYIZuVt1cv3uVtGccZcuVtPWv3GuxtVwi04960vyACFIMOSU7ucEWe1GD8zv7u@ZPuxtfvyudw8E62qvyyuoqFmqE2552b2b3=Z1vzXX3hJrZZWuV',
        'device_ratio': '1',
        'tn': 'B_NORMAL_MAP',
        'nn': '0',
        'u_loc': '12434702,4957777',
        'ie': 'utf-8',
        't':now,

        'b':'(12194927,4759185;12248303,4805969)',
         'pl_data_type':'hotel',#取值有： hotel（宾馆）；cater（餐饮）；life（生活娱乐）
        'pl_sub_type':'酒店',
        'pl_price_section':'0,+',
        'pl_sort_type':'default',#default（默认）；price（价格）；total_score（好评）；level（星级）；health_score（卫生）；distance（距离排序，只有圆形区域检索有效）
        'pl_sort_rule':'0',#排序规则：0（从高到低），1（从低到高）
        'pl_discount2_section':'0,+',
        'pl_groupon_section':'0,+',
        'pl_cater_book_pc_section':'0,+',
        'pl_hotel_book_pc_section':'0,+',
        'pl_ticket_book_flag_section':'0,+',
        'pl_movie_book_section':'0,+',
        'pl_business_type':'hotel',
        'pl_business_id':'',
    }

    data2={
        'newmap':'1',
        'reqflag':'pcmap',
        'biz':'1',
        'from':'webmap',
        'da_par':'direct',
        'pcevaname':'pc4.1',
        'qt':'spot',#query_type??  con,s,spot
        'from':'webmap',
        'c':'283',#from <<BaiduMap_cityCode_1102.txt>>
        'wd': area+kw,  # 修改关键字
        'wd2':'',
       # 'pn':'0',#page_num??无效
        'nn':pn,#??
        'rn':'50',#单页条数，返回<=50
        'db':'0',
        'sug':'0',
        'addr':'0',
        'da_src':'pcmappg.poi.page',
        'on_gel':'1',
        'src':'7',
        'gr':'3',
        'l':'15',
        'tn':'B_NORMAL_MAP',
        'auth':'x4JcO0V4BOLVMTY8HP0w3wLv7W2RbxTfuxHRxRLTBxTtComRB199Ay1uVt1GgvPUDZYOYIZuxHtPqIVH42Iff=fxXwPWv3GuLNEtZh62eGUvhgMZSguxzBEHLNRTVtcEWe1GD8zv7u@ZPuxBTtuLmSfU2K4O3pFHJHEf0wd0vyISyIMI7yIuswVVHd52Ejjg2Je7',
        'u_loc':'12434702,4957777',
        'ie':'utf-8',
         't':now,
        }
    # 把字典对象转化为url的请求参数
    url = 'https://map.baidu.com/?' + urlencode(data2)
    #print(url)
    try:
        r = requests.get(url,timeout = 30,headers=headers)
        r.raise_for_status()
        r.encoding = 'ascii'
        return r.text
    except:
        return ""

##从excel文件中获取一列数据
def getColumFromExcel(afile,sheet,c):
    rlist = []
    try:
        if  os.path.exists(afile):
            #打开文件
            wordbook = xlrd.open_workbook(afile)
            #获取sheet1
            Sheet1 = wordbook.sheet_by_name(sheet)
            #获取sheet1的第一列
            cols = Sheet1.col_values(c)
            for col in cols:
                if col.strip('')!='':
                    rlist.append(col.strip(''))
    except Exception as e:
        print(e)
    return rlist

def analysisJson2Info(jsoninfo):
    rlist = []
    jContentList = []
    j={}
    try:
        j = json.loads(jsoninfo)
        jContentList = j['content'] #content[alias,name,tel,tag]
    except Exception as e:
        print(e)
    print("total:"+str(j["result"]['total']))
    
    for i in range(len(jContentList)):
        tmp =[]
       #print("===alias==="+jContentList[i]['alias'])
        if  ('name' in jContentList[i]):
            #print("===name==="+jContentList[i]['name'])
            tmp.append(jContentList[i]['name'])
        if  ('tel' in jContentList[i]):
            #print("===tel===="+jContentList[i]['tel'])
            tmp.append(jContentList[i]['tel'])
        else:
            tmp.append("")
        if  ('addr' in jContentList[i]):
            #print("====adr==="+jContentList[i]['addr'])
            tmp.append(jContentList[i]['addr'])
        #if  ('tag' in jContentList[i]):
            #print("===tag===="+jContentList[i]['tag'])
            #tmp.append(jContentList[i]['tag'])
        rlist.append(tmp)
    rlist.append(j["result"]['total'])
    return rlist

def write2excel(rootpath,area,kw,jsonInfo):
    #判断是否存在excel
    if not os.path.exists(rootpath+area+'.xlsx'):
        book = xlwt.Workbook()
    else:
        book = xlrd.open_workbook(rootpath+area+'.xlsx')
    #判断是否存在sheet
    sheet = book.add_sheet(kw)#新建sheet酒店，商铺，写字楼
    hang = 0
    for i in range(len(jsonInfo)-1):
        lie = 0
        sheet.write(hang, lie, jsonInfo[i][0])#名称
        lie += 1
        sheet.write(hang, lie, jsonInfo[i][1])#电话
        lie += 1
        sheet.write(hang, lie, jsonInfo[i][2])#地址
        hang += 1
    book.save(rootpath+area+'.xls')

def write2Excel(rootpath,area,kw,jsonInfo):
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = True
    app.screen_updating = True
    if not os.path.exists(rootpath+area+'.xlsx'):
        #新建工作簿
        book = app.books.add()
    else:
        book = app.books.open(rootpath+area+'.xlsx')
            
    try:
        #新增sheet表
        book.sheets.add(kw)    
        #打开sheet页
        sht = book.sheets[kw]

        #获取写入范围
        rg = sht.range('A1:C'+str(len(jsonInfo)+1))
        #设定格式
        rg.number_format = '@'
        #写标题行,设定字体为加粗，12号
        sht.range('A1').value=['名称','电话','地址']
        sht.range('A1:C1').api.Font.Bold = True
        sht.range('A1:C1').api.Font.Size = 12
        #获取A列第一个空单元格
        #r = sht.used_range.last_cell.row+1
        #添加信息
        #for i in range(len(jsonInfo)-1):
        sht.range('A2').value = jsonInfo
        #设定列自适应
        sht.autofit('c')
        #保存文件
        book.save(rootpath+area+'.xlsx')
        #关闭文件
        book.close()
        #退出excel应用
        app.quit()
    except Exception as e:
        print(e)

def saveFile(data,rpath,fn):
    rootdir = rpath
    path = rootdir + fn+".json"
    
    try:
        if not os.path.exists(rootdir):
            os.mkdir(rootdir)
        if not os.path.exists(path):
            with open(path,'w') as   f:
                f.write(data)
                f.close
                print("file saves successfully  "+path)
        else:
            print("file has exist")
    except Exception as e:
        print("failed"+e)


def main():
    
    areas= getColumFromExcel("f:/EErDuoSi_addr/areas.xlsx","Sheet1",0)#获取区域列表
    keywords = getColumFromExcel("f:/EErDuoSi_addr/areas.xlsx","Sheet1",1)#获取关键字列表
    startTime = time.time()
    for area in areas:
        for kw in keywords:
            print("a",area,"k",kw,sep=':')
            cnt = 0#已获取的数据条数
            alldata = []#获取的某个关键字的所有结果
            while True:
                json = getJsonText(area,kw,cnt)
                pdata = analysisJson2Info(json)
                alldata+=pdata[0:-1]
                
                tps = len(pdata)-1
                print("this:",tps)
                cnt += tps
                
                total = pdata[-1]#获取到查询的total
                saveFile(json,"d:/a/",area+kw+str(int(time.time()*1000) ))
                
                if cnt>= total:
                    break
            #if len(alldata)>=1:print(len(alldata),alldata[:2],sep='\n')
                
            write2Excel("f:/EErDuoSi_addr/v0.3/",area,kw,alldata)
    entTime = time.time()
    print('finished in :',float(entTime-startTime),"Sec.",sep=':')


main()
