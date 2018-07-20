# 上市外資 投信買賣超前三十
import os
import csv
import requests
from bs4 import BeautifulSoup
import pandas as pd
import copy
from package.Infrastructure import DateObj

"""
使用copy.deepcopy 完整複製一份reference object
因為使用list.copy() 假設新list pop時 會異動到舊list
(只有clear()不會 append()亦會影響)
"""


def formatNo(noStr):
    noOrg = noStr.replace(',', '')
    noInt = int(noOrg)
    return int(noInt / 1000)


def main(dList):
    # 日期list
    dateList = dList
    for d in dateList:
        dateObj = DateObj.DateObj(d)
        ExportExcel(dateObj)


def ExportExcel(obj):
    resultList_tw = [
        '代號', '公司', '外資買超張數',
        '代號', '公司', '外資賣超張數',
        '代號', '公司', '投信買超張數',
        '代號', '公司', '投信賣超張數']

    url = 'http://wwwc.twse.com.tw/fund/T86'
    data = {
        'response': 'json',
        'date': obj.strdate,
        'selectType': 'ALLBUT0999'  # 全部(不含權證、牛熊證、可展延牛熊證)
        # 'selectType': 'ALL'
    }
    r = requests.post(url, data=data)
    data_json = r.json()
    if data_json["stat"] != "OK":
        print("資料讀取失敗")
    else:
        basic_data = data_json["data"]
        adjust_data = []
        for x in basic_data:
            foreign = formatNo(x[4])  # 外陸資買賣超股數(不含外資自營商)
            foreign_dealer = formatNo(x[7])  # 外資自營商買賣超股數
            one_data = ["'" + str(x[0]), x[1].strip(),
                        (foreign + foreign_dealer),  # 外陸資買賣超股數
                        formatNo(x[10])]  # 投信買賣超股數
            adjust_data.append(one_data)

        # 外資買超前三十
        adjust_data.sort(key=lambda x: x[2], reverse=True)
        foreign_Buy = copy.deepcopy(adjust_data)[:30]

        # 外資賣超
        adjust_data.sort(key=lambda x: x[2], reverse=False)
        foreign_Sell = copy.deepcopy(adjust_data)[:30]

        # 投信買超
        adjust_data.sort(key=lambda x: x[3], reverse=True)
        local_Buy = copy.deepcopy(adjust_data)[:30]

        # 投信賣超
        adjust_data.sort(key=lambda x: x[3], reverse=False)
        local_Sell = copy.deepcopy(adjust_data)[:30]

        final_data = []
        for x in foreign_Buy:
            idx = foreign_Buy.index(x)
            row_data = []
            # 外資買超
            foreign_Buy[idx].pop(3)
            row_data.extend(foreign_Buy[idx])
            # 外資賣超
            foreign_Sell[idx].pop(3)
            row_data.extend(foreign_Sell[idx])
            # 投信買超
            local_Buy[idx].pop(2)
            row_data.extend(local_Buy[idx])
            # 投信賣超
            local_Sell[idx].pop(2)
            row_data.extend(local_Sell[idx])

            final_data.append(row_data)

        SaveDirectory = os.getcwd()  # 印出目前工作目錄
        SaveAs = os.path.join(SaveDirectory, 'daily',
                              '上市買賣超.csv')  # 組合路徑，自動加上兩條斜線 "\\"

        """
        open parameter:mode
        r - 讀取(檔案需存在)
        w - 新建檔案寫入(檔案可不存在，若存在則清空)
        a - 資料附加到舊檔案後面(游標指在EOF)
        r+ - 讀取舊資料並寫入(檔案需存在且游標指在開頭)
        w+ - 清空檔案內容，新寫入的東西可在讀出(檔案可不存在，會自行新增)
        a+ - 資料附加到舊檔案後面(游標指在EOF)，可讀取資料
        b - 二進位模式
        """
        with open(SaveAs, 'a', newline='') as csvfile:
            # 以空白分隔欄位，建立 CSV 檔寫入器
            writer = csv.writer(csvfile)
            writer.writerow(resultList_tw)
            for l in final_data:
                writer.writerow(l)
