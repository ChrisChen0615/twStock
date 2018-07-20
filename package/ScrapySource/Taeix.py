# 每日收盤指數、漲跌、漲跌指數、漲跌百分比、成交量
import requests
from bs4 import BeautifulSoup
import pandas as pd
from package.Infrastructure import DateObj


class Taeix:
    def __init__(self, obj):
        self.dataDate = obj

    def GetTaeixList(self):
        url = "http://www.twse.com.tw/exchangeReport/MI_INDEX"
        data = {
            'response': 'json',
            'date': self.dataDate.strdate,
            'type': 'MS'
        }
        r = requests.post(url, data=data)
        data_json = r.json()
        taeix = data_json["data1"][1]
        total = int((data_json["data3"][13][1]).replace(
            ',', '')) / 100000000  # 成交量(億元)
        taeix[2] = BeautifulSoup(taeix[2], 'html.parser').find('p').text
        taeix[2] = taeix[2] + taeix[3]
        taeix.pop(3)
        taeix[3] = taeix[3] + '%'  # 漲跌百分比加符號
        taeix.append("%.2f" % total)
        return taeix[1:]
