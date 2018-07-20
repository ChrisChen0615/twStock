# my datetime class
import datetime


class DateObj:
    """自定義時間物件"""

    def __init__(self, d):
        """"d:時間字串 ex.20180716"""
        datetimeObj = datetime.datetime.strptime(d, '%Y%m%d')

        self.strdate = d  # 時間字串
        self.datetimeObj = datetimeObj  # 時間字串轉datetime object
        self.dataDay = datetimeObj.strftime("%d")  # 日字串
        self.dataMonth = datetimeObj.strftime("%m")  # 月字串
        self.dataYear = datetimeObj.strftime("%Y")  # 年字串
        self.dateSlash = datetimeObj.strftime('%Y/%m/%d')  # 2018/07/16
        twYear = str(int(self.dateSlash[0:4])-1911)  # 民國年
        self.twSlashDate = twYear + self.dateSlash[4:]  # 107/07/16
