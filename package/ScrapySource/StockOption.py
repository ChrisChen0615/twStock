# 自營(選)買權, 自營(選)賣權, 外資(選)買權, 外資(選)賣權
import requests
from bs4 import BeautifulSoup
from package.Infrastructure import DateObj


class StockOption:
    def __init__(self, obj):
        self.dataDate = obj
        self.dealer_call = 0  # 自營商_買權 金額
        self.dealer_put = 0  # 自營商_賣權
        self.foreign_call = 0  # 外資_買權
        self.foreign_put = 0  # 外資_賣權

    def CalCount(self):
        url = 'http://www.taifex.com.tw/cht/3/callsAndPutsDate'
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

        # 自營商_買權 金額
        cols = rows[3].find_all('td')
        self.dealer_call += int(cols[15].text.replace(',', ''))

        # 外資_買權 金額
        cols = rows[5].find_all('td')
        self.foreign_call += int(cols[12].text.replace(',', ''))

        # 自營商_賣權 金額
        cols = rows[6].find_all('td')
        self.dealer_put += int(cols[13].text.replace(',', ''))

        # 外資_賣權 金額
        cols = rows[8].find_all('td')
        self.foreign_put += int(cols[12].text.replace(',', ''))
