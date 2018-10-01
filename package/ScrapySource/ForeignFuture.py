# 外資未平倉口數(大台), 外資未平倉口數(小台), 外資未平倉口數(大小台合計) = 大+(小/4)
import requests
from bs4 import BeautifulSoup
from package.Infrastructure import DateObj


class ForeignFuture:
    def __init__(self, obj):
        self.dataDate = obj
        self.big_future = 0  # 外資未平倉(口數):大台指期
        self.small_future = 0  # 外資未平倉(口數):小台指期
        self.future = 0  # 外資未平倉(口數):大+小/4

    def CalCount(self):
        url = 'http://www.taifex.com.tw/cht/3/futContractsDate'
        params = {
            'commodityId': '',
            'dateaddcnt': '',
            'doQuery': '1',
            'goday': '',
            'queryDate': self.dataDate.dateSlash,
            'queryType': '1'
        }
        result = requests.post(url, data=params)
        c = result.content  # text
        soup = BeautifulSoup(c, "html.parser")
        table = soup.find_all('table', "table_f")
        rows = table[0].find_all('tr')
        cols = rows[5].find_all('td')
        fonts = cols[11].find_all('div')[0].find_all('font')
        self.big_future += int(fonts[0].text.replace(',', ''))

        cols = rows[14].find_all('td')
        fonts = cols[11].find_all('div')[0].find_all('font')
        self.small_future += int(fonts[0].text.replace(',', ''))
        self.future = self.big_future + (self.small_future//4) #//相除取整數，無條件捨去?