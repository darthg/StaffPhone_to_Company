import urllib
import urllib.parse
import urllib.request
import time
import  guoxiaoyi.xiciProxies
import csv
import xlrd
from lxml import html
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

page_no=0
headers=('User-Agent',"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.167 Safari/537.36")
IPfp=r'E:\GitHub\测试文件\IP_Pool.xls'
opener=urllib.request.build_opener()





dicInput={'公司名称': '', '法定代表人': '', '注册资本': '','成立时间': '','电话': '','邮箱': '','地址': ''}

with open("E:\GitHub\测试文件\排查结果\员工手机号企查查_开办公司.csv",'a+',newline='') as csvfile:
     fieldnames=['公司名称','法定代表人','注册资本','成立时间','电话','邮箱','地址','员工姓名','单位名称','部门名称','职位全称']
     writer=csv.DictWriter(csvfile,fieldnames=fieldnames)
     writer.writeheader()



'''
name_pat='"iname":"(.*?)"'
ID_pat='"cardNum":"(.*?)"'
duty_pat='"duty":"(.*?)"'
caseCode_pat='"caseCode":"(.*?)"'
publishDate_pat='"publishDate":"(.*?)"'
'''

try:
#workbook=xlrd.open_workbook(r"E:\GitHub\测试文件\2016年保险明细报表.xls")
     workbook=xlrd.open_workbook(r"E:\GitHub\测试文件\查询目标\江西全辖在职人员花名册.xls")
except Exception as e:
     print("读取文件出现异常！请确认文件路径！")
sheets=workbook.sheet_by_index(0)
rows=sheets.nrows
print("已经打开查询目标文件")

for i in range(112,rows):#循环控制从哪条开始
     try:
          phone_input=sheets.cell_value(i,24)
          staff_name=sheets.cell_value(i,3)
          staff_com=sheets.cell_value(i,0)
          staff_depart=sheets.cell_value(i,1)
          staff_duty=sheets.cell_value(i,5)
          #ID_input=sheets.cell_value(i,1)

          browser=webdriver.Chrome#设置浏览器类型
          chrome_options = webdriver.ChromeOptions()
          ip=guoxiaoyi.xiciProxies.getPoxiesRand(IPfp)
          print(ip)
          #chrome_options.add_argument('--proxy-server=http://114.226.128.205:6666')
          chrome_options.add_argument('user-agent="Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.167 Safari/537.36"')
          browser = webdriver.Chrome(chrome_options=chrome_options)


          url='http://www.qichacha.com/'
          browser.get(url)
          print("正在查询"+url)
          WebDriverWait(browser,30).until(
                    EC.presence_of_element_located((By.CLASS_NAME,"input-group"))
          )#设置等待时间
          elem1=browser.find_element_by_xpath('//*[@id="searchkey"]')
          print("已定位元素")
          elem1.clear()
          elem1.send_keys(phone_input)
          searchbutton=browser.find_element_by_id("V3_Search_bt")#中国裁判文书网找这个搜索按钮一定要用
          searchbutton.click()
          WebDriverWait(browser,30).until(
                    EC.presence_of_element_located((By.ID,"countOld"))
          )#设置等待时间
          data=browser.page_source
          xpathdata=html.etree.HTML(data)
          count=xpathdata.xpath('//*[@id="countOld"]/span/text()')
          if int(count[0])!=0:
               company_name=xpathdata.xpath('//*[@id="searchlist"]/table/tbody/tr/td[2]/a/text()')
               boss_name=xpathdata.xpath('//*[@id="searchlist"]/table/tbody/tr/td[2]/p[1]/a/text()')
               set_money=xpathdata.xpath('//*[@id="searchlist"]/table/tbody/tr/td[2]/p[1]/span[1]/text()')
               set_date=xpathdata.xpath('//*[@id="searchlist"]/table/tbody/tr/td[2]/p[1]/span[2]/text()')
               phone=xpathdata.xpath('//*[@id="searchlist"]/table/tbody/tr/td[2]/p[2]/text()')
               email=xpathdata.xpath('//*[@id="searchlist"]/table/tbody/tr/td[2]/p[2]/span/text()')
               address=xpathdata.xpath('//*[@id="searchlist"]/table/tbody/tr/td[2]/p[3]/text()')

               with open("E:\GitHub\测试文件\排查结果\员工手机号企查查_开办公司.csv",'a+',newline='') as csvfile:
                    fieldnames=['公司名称','法定代表人','注册资本','成立时间','电话','邮箱','地址','员工姓名','单位名称','部门名称','职位全称']
                    writer=csv.DictWriter(csvfile,fieldnames=fieldnames)
                    for j in range(0,len(company_name)):
                         dicInput['公司名称']=company_name[j]
                         print(company_name[j])
                         dicInput['法定代表人']=boss_name[j]
                         dicInput['注册资本']=set_money[j]
                         dicInput['成立时间']=set_date[j]
                         dicInput['电话']=phone[j]
                         dicInput['邮箱']=email[j]
                         dicInput['地址']=address[j]
                         dicInput['员工姓名']=staff_name
                         dicInput['单位名称']=staff_com
                         dicInput['部门名称']=staff_depart
                         dicInput['职位全称']=staff_duty
                         writer.writerow(dicInput)
                         print("已经写入第"+str(j+1)+"条数据")
               time.sleep(1)
               browser.close()
          elif int(count[0])==0:
               print("未找到关于"+str(staff_name)+"私自开办企业信息!!")
               browser.close()
               continue

     except Exception as e:
          print(e)
          print("采集"+str(staff_name)+"失败")
          browser.close()
          continue