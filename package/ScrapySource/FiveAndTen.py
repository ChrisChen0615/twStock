#前五大交易人留倉部位(所有), 前10大交易人留倉部位(所有)
import requests
from bs4 import BeautifulSoup
from package.Infrastructure import DateObj

class FiveAndTen:
    def __init__(self,obj):
        self.dataDate = obj
        self.five_all = 0 #前五大交易人留倉部位(所有期約)
        self.ten_all = 0 #前十大交易人留倉部位(所有期約)

    def CalCount(self):
        url = 'http://www.taifex.com.tw/cht/3/getLargeTradersFutContract'
        params = {
            'queryDate': self.dataDate.dateSlash
        }
        result = requests.post(url, params=params)
        c = result.content  # text
        soup = BeautifulSoup(c, "html.parser")
        table = soup.find_all('table', "table_f")
        rows = table[0].find_all('tr')
        five_buy = 0  # 前五大交易人(買方)
        five_sell = 0  # 前五大交易人(賣方)        
        ten_buy = 0  # 前十大交易人(買方)
        ten_sell = 0  # 前十大交易人(賣方)
        

        # 前五大交易人
        cols = rows[5].find_all('td')
        div = cols[1].find_all('div')
        five_buy = int(div[0].text.split('(')[0].replace(',', ''))
        div = cols[5].find_all('div')
        five_sell = int(div[0].text.split('(')[0].replace(',', ''))
        self.five_all = five_buy - five_sell

        # 前十大交易人
        cols = rows[5].find_all('td')
        div = cols[3].find_all('div')
        ten_buy = int(div[0].text.split('(')[0].replace(',', ''))
        div = cols[7].find_all('div')
        ten_sell = int(div[0].text.split('(')[0].replace(',', ''))
        self.ten_all = ten_buy - ten_sell
