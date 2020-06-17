# -*- coding: utf-8 -*-
from requests_html import HTMLSession
import xlsxwriter
import json
from retrying import retry


@retry(wait_random_min=1000,wait_random_max=5000)
def get_html(url):
  'url'
  print('try')
  session = HTMLSession(
    browser_args=[
				'--no-sand',
				'--user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36"'
			]
    )#设置浏览器user-agent
  response = session.get(url)
  response.html.render(timeout=30)
  print('done'+ url )
  return response
  
if __name__ == "__main__":
  fightID = "4qXdLBTcPgFAbwHM"
  boss_id_url = "https://cn.classic.warcraftlogs.com/reports/fights-and-participants/"+ fightID +"/0"

  '获取详细boss战ID 请求URL 返回json'
  r = get_html(boss_id_url)
  fightData = json.loads(r.text)
  fight_data = []
  for i in fightData['fights']:
    if 'kill' in i and i['boss'] != '0':
      #生产详细战斗boss及战斗id列表
      temp = {'name':str(i['name']),'fightID':str(i['id'])}
      fight_data.append(temp)

  #遍历boss ID列表匹配friendlies列表
  reslot = []
  for i in fight_data:
    temp_name = []
    for j in fightData['friendlies']:
      if '.'+i['fightID']+'.' in j['fights']:
        temp_name.append(j['name'])
    temp = {'boss':i['name'],'name':temp_name}
    reslot.append(temp)

#生成excel
  workbook = xlsxwriter.Workbook('./击杀汇总.xlsx')
  worksheet = workbook.add_worksheet()
  col = 0
  bold = workbook.add_format({'bold':True})
  for i in reslot:
    row = 1
    worksheet.write(0,col,i['boss'],bold)
    for j in i['name']:
      worksheet.write(row,col,j)
      row += 1
    col += 1

  workbook.close()  