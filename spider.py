# -*- coding: utf-8 -*-
from requests_html import HTMLSession
import xlsxwriter
import json
from retrying import retry
import time

@retry(wait_random_min=1000,wait_random_max=5000)
def get_html(url):
  'url'
  session = HTMLSession(
    browser_args=[
				'--no-sand',
				'--user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36"'
			]
    )#设置浏览器user-agent
  response = session.get(url)
  response.html.render(timeout=30)
  return response
  
if __name__ == "__main__":
  fightID = input("输入WCL识别符：")
  boss_id_url = "https://cn.classic.warcraftlogs.com/reports/fights-and-participants/"+ fightID +"/0"

  '获取详细boss战ID 请求URL 返回json'
  r = get_html(boss_id_url)
  fightData = json.loads(r.text)
  fight_data = []
  for i in fightData['fights']:
    if 'kill' in i and i['boss'] != '0' and i['kill']:
      #生产详细战斗boss及战斗id列表
      temp = {'name':str(i['name']),'fightID':str(i['id'])}
      fight_data.append(temp)
  

  epgp = [
    {'name':'鲁西弗隆','point':40},{'name':'玛格曼达','point':40},{'name':'基赫纳斯','point':40},{'name':'迦顿男爵','point':40},{'name':'沙斯拉尔','point':40},{'name':'加尔','point':40},{'name':'萨弗隆先驱者','point':40},{'name':'焚化者古雷曼格','point':40},{'name':'管理者埃克索图斯','point':40},{'name':'熔岩爆发','point':60}
  ]
  dkp = [
    {'name':'鲁西弗隆','point':2},{'name':'玛格曼达','point':2},{'name':'基赫纳斯','point':2},{'name':'迦顿男爵','point':2},{'name':'沙斯拉尔','point':2},{'name':'加尔','point':2},{'name':'萨弗隆先驱者','point':2},{'name':'焚化者古雷曼格','point':2},{'name':'管理者埃克索图斯','point':2},{'name':'熔岩爆发','point':4},{'name':'狂野的拉佐格尔','point':3},{'name':'堕落的瓦拉斯塔兹','point':3},{'name':'勒什雷尔','point':3},{'name':'费尔默','point':3},{'name':'埃博诺克','point':3},{'name':'弗莱格尔','point':3},{'name':'克洛玛古斯','point':3},{'name':'奈法利安','point':6}
  ]
  #遍历boss ID列表匹配friendlies列表
  reslot = []
  for i in fight_data:
    temp_name = []
    for j in fightData['friendlies']:
      if '.'+i['fightID']+'.' in j['fights']:
        temp_name.append(j['name'])
    epgpScore = 0
    dkpScore = 0
    for k in epgp:
      if i['name'] == k['name']:
        epgpScore = k['point']
    for l in dkp:
      if i['name'] == l['name']:
        dkpScore = l['point']
    temp = {'boss':i['name'],'name':temp_name,'epgp':epgpScore,'dkp':dkpScore}
    reslot.append(temp)

#遍历reslot生产分数数组
EpgpPoint = {}
DKPPoint = {}
for i in reslot:
  for j in i['name']:
    if j in EpgpPoint:
        EpgpPoint[j] += i['epgp']
        DKPPoint[j] += i['dkp']
    else:
      EpgpPoint[j] = i['epgp']
      DKPPoint[j] = i['dkp']
        
      
#生成excel
  workbook = xlsxwriter.Workbook('./' + time.strftime("%Y-%m-%d", time.localtime()) + '活动汇总.xlsx')
  worksheet = workbook.add_worksheet(name="出勤统计")
  worksheet1 = workbook.add_worksheet(name="加分情况")
  col = 0
  worksheet.set_column('A:Z',12)  
  bold = workbook.add_format({'bold':True,'align':'center'})
  font_center = workbook.add_format({'align':'center'})
  for i in reslot:
    row = 1
    worksheet.write(0,col,i['boss'],bold)
    for j in i['name']:
      worksheet.write(row,col,j,font_center)
      row += 1
    col += 1
  #在sheet2里写入数据
  worksheet1.set_column('A:A',12) 
  worksheet1.set_column('C:C',12) 
  worksheet1.merge_range("A1:B1","EPGP加分情况",bold)
  worksheet1.merge_range("C1:D1","DKP加分情况",bold)
  n = 2
  for key,values in  EpgpPoint.items():
    worksheet1.write("A{}".format(n),key,font_center)
    worksheet1.write("B{}".format(n),values,font_center)
    n += 1
  n = 2
  for key,values in  DKPPoint.items():
    worksheet1.write("C{}".format(n),key,font_center)
    worksheet1.write("D{}".format(n),values,font_center)
    n += 1

  workbook.close()  