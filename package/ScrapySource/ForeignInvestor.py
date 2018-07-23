# 每日外資(含自營商)買賣超金額
import requests
from bs4 import BeautifulSoup
import pandas as pd
from package.Infrastructure import DateObj
from decimal import Decimal

class ForeignInvestor:
    def __init__(self, obj):
        self.dataDate = obj

    def GetCount(self):
        url = "http://wwwc.twse.com.tw/fund/BFI82U"
        data = {
            'response': 'json',
            'dayDate': self.dataDate.strdate,
            'weekDate': '',
            'monthDate': '',
            'type': 'day'
        }
        r = requests.post(url, data=data)
        data_json = r.json()
        foreignWithoutDealar = data_json["data"][3]

        dealr = data_json["data"][4]
        # 取代千分位
        intForeignWithoutDealar = int(foreignWithoutDealar[3].replace(',', ''))
        intDealr = int(dealr[3].replace(',', ''))
        count = intForeignWithoutDealar + intDealr
        #return '%.2f' %(count/100000000)
        total = count/100000000
        return total