# -*- coding: utf-8 -*-
# 上櫃外資 投信買賣超前三十
import requests
from package.Infrastructure import DateObj
from package.Infrastructure import CommonMethon as _cm
from package.ScrapyBuySell import BuySell
import numpy as np


class OTC(BuySell.BuySell):
    """證券櫃檯買賣中心物件"""

    def __init__(self, dateObj):
        super().__init__(dateObj)
        self.name = "上櫃"
        self.EnName = "OTCBuySell"

    def ScrapyData(self, dateObj):
        """爬取資料        
        dataObj:日期資料物件
        """
        url = 'http://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php'
        data = {
            'l': 'zh-tw',
            'se': 'EW',  # AL:全部、EW:不含權證、牛熊證
            't': 'D',
            'd': dateObj.twSlashDate
        }
        r = requests.get(url, params=data)
        data_json = r.json()
        self.DataSourceByScrapy = data_json["aaData"]
        if len(self.DataSourceByScrapy) == 0:
            self.IsGetSource = False
        # self.DataSourceByScrapy = np.array(data_json['aaData'])
        # if self.DataSourceByScrapy.size == 0:
        #     self.IsGetSource = False

    def FilterScrapyData(self):
        """
        資料整理
        從初始資料過濾取得所需欄位 並指派給 FilterData
        """
        self.FilterData = []
        for x in self.DataSourceByScrapy:
            one_data = [str(x[0]).strip(), x[1].strip(),
                        _cm.formatNo(x[10]),  # 外陸資買賣超股數
                        _cm.formatNo(x[13])]  # 投信買賣超股數
            self.FilterData.append(one_data)
        # self.FilterData = np.array([])
        # for x in self.DataSourceByScrapy:
        #     one_data = [str(x[0]).strip(), x[1].strip(),
        #                 _cm.formatNo(x[10]),  # 外陸資買賣超股數
        #                 _cm.formatNo(x[13])]  # 投信買賣超股數
        #     self.FilterData = np.concatenate(self.FilterData, one_data)
