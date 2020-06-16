# -*- coding: utf-8 -*-
from requests_html import HTMLSession

def get_html(url):
  'url'
  session = HTMLSession()
  response = session.get(url)
  response.html.render(timeout=30)
  return response
  
if __name__ == "__main__":
  a = get_html("https://www.baidu.com/")
  print(a.text)