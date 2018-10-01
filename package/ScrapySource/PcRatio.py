# 擷取每日選擇權put/call比
import requests
from bs4 import BeautifulSoup
import datetime
from package.Infrastructure import DateObj


class PCRatio:
    def __init__(self, obj):
        self.dataDate = obj

    def GetRatio(self):
        url = "http://www.taifex.com.tw/cht/3/pcRatio"        
        data = {
            'queryEndDate': self.dataDate.dateSlash,
            'queryStartDate': self.dataDate.dateSlash,
            'down_type': ''
        }        
        r = requests.post(url, data=data)
        c = r.content  # text
        soup = BeautifulSoup(c, "html.parser")

        table = soup.find_all('table', "table_a")
        rows = table[0].find_all('tr')
        pcRatio = []
        for row in rows[1:]:
            cols = row.find_all('td')
            col_date = datetime.datetime.strptime(
                cols[0].text, '%Y/%m/%d').strftime('%Y/%m/%d')
            if col_date == self.dataDate.dateSlash:
                pcRatio = cols
                break
        if(len(pcRatio) == 0):
            return 0
        else:
            return(float(pcRatio[6].text) / 100)  # 百分比數字
