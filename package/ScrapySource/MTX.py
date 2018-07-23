#小台指未沖銷契約量, 散戶看多, 散戶看空, 散戶多空比
import requests
from bs4 import BeautifulSoup
from package.Infrastructure import DateObj


class MTX:
    def __init__(self, obj):
        self.dataDate = obj
        self.mtx_notsell = 0  # 小台指未沖銷契約量
        self.bull_side = 0  # 未平倉口數(多方):外資+自營商+投信
        self.bear_side = 0  # 未平倉口數(空方):外資+自營商+投信
        self.bull = 0  # 看多口數
        self.bear = 0  # 看空口數
        self.ratio = 0  # 多空比

    def CalCount(self):
        self.CalNotsell()
        self.CalBullAndBear()

        self.bull = self.mtx_notsell - self.bull_side
        self.bear = self.mtx_notsell - self.bear_side
        self.ratio = (self.bull - self.bear) / self.mtx_notsell

    def CalNotsell(self):
        url = "https://www.taifex.com.tw/chinese/3/3_1_1.asp"
        params = {
            'commodity_id': 'MTX',
            'commodity_id2': '',
            'commodity_id2t': '',
            'commodity_id2t2': '',
            'commodity_idt': 'MTX',
            'DATA_DATE_D': self.dataDate.dataDay,
            'DATA_DATE_M': self.dataDate.dataMonth,
            'DATA_DATE_Y': self.dataDate.dataYear,
            'dateaddcnt': '0',
            'datestart': self.dataDate.dateSlash,
            'goday': '',
            'market_code': '0',
            'MarketCode': '0',
            'qtype': '2',
            'sday': self.dataDate.dataDay,
            'smonth': self.dataDate.dataMonth,
            'syear': self.dataDate.dataYear
        }
        result = requests.post(url, data=params)
        c = result.content  # text
        soup = BeautifulSoup(c, "html.parser")
        table = soup.find_all('table', "table_f")
        rows = table[0].find_all('tr')
        last_row = len(rows) - 1  # 最末行
        cols = rows[last_row].find_all('td')
        self.mtx_notsell = int(cols[12].text)# 小台指未沖銷契約量

    def CalBullAndBear(self):
        url = "https://www.taifex.com.tw/chinese/3/7_12_3.asp"
        params = {
            'COMMODITY_ID': 'MXF',
            'DATA_DATE_D': self.dataDate.dataDay,
            'DATA_DATE_M': self.dataDate.dataMonth,
            'DATA_DATE_Y': self.dataDate.dataYear,
            'datestart': self.dataDate.dateSlash,
            'goday': '',
            'sday': self.dataDate.dataDay,
            'smonth': self.dataDate.dataMonth,
            'syear': self.dataDate.dataYear
        }
        result = requests.post(url, data=params)
        c = result.content  # text
        soup = BeautifulSoup(c, "html.parser")
        table = soup.find_all('table', "table_f")
        rows = table[0].find_all('tr')
        # 自營商 因為td colspan = 3
        cols = rows[3].find_all('td')
        fonts = cols[9].find_all('div')[0].find_all('font')
        self.bull_side += int(fonts[0].text.replace(',', ''))
        fonts = cols[11].find_all('div')[0].find_all('font')
        self.bear_side += int(fonts[0].text.replace(',', ''))

        for row in rows[4:6]:
            cols = row.find_all('td')
            fonts = cols[7].find_all('div')[0].find_all('font')
            self.bull_side += int(fonts[0].text.replace(',', ''))
            fonts = cols[9].find_all('div')[0].find_all('font')
            self.bear_side += int(fonts[0].text.replace(',', ''))
