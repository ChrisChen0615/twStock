# -*- coding: utf-8 -*-
# 上市外資 投信買賣超前三十
import requests
from package.Infrastructure import DateObj
from package.Infrastructure import CommonMethon as _cm
from package.ScrapyBuySell import BuySell


class TWSE(BuySell.BuySell):
    """台灣證券交易所物件"""

    def __init__(self, dateObj):
        super().__init__(dateObj)
        self.name = "上市"
        self.EnName = "TWSEBuySell"

    def ScrapyData(self, dateObj):
        """爬取資料
        從台灣證券交易所網站上爬取資料
        dataObj:日期資料物件
        """
        url = 'https://wwwc.twse.com.tw/fund/T86'
        data = {
            'response': 'json',
            'date': dateObj.strdate,
            'selectType': 'ALLBUT0999'  # 全部(不含權證、牛熊證、可展延牛熊證)
            # 'selectType': 'ALL'
        }
        r = requests.post(url, data=data)
        data_json = r.json()
        if data_json["stat"] == "OK":
            self.DataSourceByScrapy = data_json["data"]            
        else:
            self.IsGetSource = False

    def FilterScrapyData(self):
        """
        資料整理
        從初始資料過濾取得所需欄位 並指派給 FilterData
        """
        self.FilterData = []
        for x in self.DataSourceByScrapy:
            foreign = _cm.formatNo(x[4])  # 外陸資買賣超股數(不含外資自營商)
            foreign_dealer = _cm.formatNo(x[7])  # 外資自營商買賣超股數
            one_data = [str(x[0]).strip(), x[1].strip(),
                        (foreign + foreign_dealer),  # 外陸資買賣超股數
                        _cm.formatNo(x[10])]  # 投信買賣超股數
            self.FilterData.append(one_data)
