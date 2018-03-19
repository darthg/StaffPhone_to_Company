import urllib
import urllib.parse
import time
import  guoxiaoyi.xiciProxies
import re
import csv
import os
import xlrd

def nextpage(browser):#查找下一页
     time.sleep(1)
     nextpage=browser.find_elements_by_xpath('//*[@id="1"]/div/div[5]/p/span[7]')
     nextpage.click()
     return browser



page_no=0
header=('User-Agent',"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.167 Safari/537.36")
IPfp=r'E:\GitHub\测试文件\IP_Pool.xls'
opener=urllib.request.build_opener()
opener.addheaders=[header]
ip=guoxiaoyi.xiciProxies.getPoxiesRand(IPfp)
proxy=urllib.request.ProxyHandler({"http":ip})#设立代理IP
opener=urllib.request.build_opener(proxy,urllib.request.HTTPHandler)#封装opener



company_name=[]
boss_name=[]
set_money=[]
set_date=[]
phone=[]
email=[]
address=[]

dicInput={'公司名称': '', '法定代表人': '', '注册资本': '','成立时间': '','电话': '','邮箱': '','地址': ''}

with open("E:\GitHub\测试文件\排查结果\员工手机号企查查_开办公司.csv",'a+',newline='') as csvfile:
     fieldnames=['公司名称','法定代表人','注册资本','成立时间','电话','邮箱','地址']
     writer=csv.DictWriter(csvfile,fieldnames=fieldnames)
     writer.writeheader()




name_pat='"iname":"(.*?)"'
ID_pat='"cardNum":"(.*?)"'
duty_pat='"duty":"(.*?)"'
caseCode_pat='"caseCode":"(.*?)"'
publishDate_pat='"publishDate":"(.*?)"'

try:
#workbook=xlrd.open_workbook(r"E:\GitHub\测试文件\2016年保险明细报表.xls")
     workbook=xlrd.open_workbook(r"E:\GitHub\测试文件\查询目标\员工办企业.xls")
except Exception as e:
     print("读取文件出现异常！请确认文件路径！")
sheets=workbook.sheet_by_index(0)
rows=sheets.nrows


for i in range(1,rows):
     try:
          name_input=sheets.cell_value(i,0)
          #ID_input=sheets.cell_value(i,1)
          url='https://sp0.baidu.com/8aQDcjqpAAV3otqbppnN2DJv/api.php?resource_id=6899&query=%E5%A4%B1%E4%BF%A1%E8%A2%AB%E6%89%A7%E8%A1%8C%E4%BA%BA%E5%90%8D%E5%8D%95&iname='+urllib.parse.quote(name_input)+'&ie=utf-8&oe=utf-8&format=json'
          print("正在查询"+url)
          data=opener.open(url,timeout=1).read().decode('utf-8','ignore')
          name=re.compile(name_pat).findall(data)
          ID_no=re.compile(ID_pat).findall(data)
          duty=re.compile(duty_pat).findall(data)
          caseCode=re.compile(caseCode_pat).findall(data)
          publishDate=re.compile(publishDate_pat).findall(data)
          if len(name)!=0:
               with open("E:\GitHub\测试文件\排查结果\员工办企业_失信人.csv",'a+',newline='') as csvfile:
                    fieldnames=['失信人名称','证件号','判决详情','文案号','发布时间']
                    writer=csv.DictWriter(csvfile)
                    for j in range(0,len(name)):
                         dicInput['失信人名称']=name[j]
                         dicInput['证件号']=ID_no[j]
                         dicInput['判决详情']=duty[j]
                         dicInput['文案号']=caseCode[j]
                         dicInput['发布时间']=publishDate[j]
                         writer.writerow(dicInput)
                    print("已经写入第"+str(j)+"条数据")
               time.sleep(1)
          else:
               print("未找到关于"+str(name_input)+"的失信人信息!!")
               continue
     except Exception as e:
          print("采集失败")
          continue