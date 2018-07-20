# 主程式
import os
import csv
from package.Infrastructure import DateObj
from package.ScrapySource import Taeix, ForeignInvestor, PcRatio, MTX
from package.ScrapySource import ForeignFuture, StockOption, FiveAndTen


def formatNum(no):
    """格式化數字，加千分位"""
    if type(no) != int:
        return 0
    else:
        return format(no, ',')


def main(dList):
    # 日期list
    dateList = dList
    # 最後輸出結果標題列
    resultList_tw = [
        '日期',
        '發行量加權股價指數', '漲跌指數', '漲跌趴數', '成交量(億)',
        '外資買賣超(億)', 'P/C ratio',
        #'小台指未沖銷契約量', '散戶看多', '散戶看空', '散戶多空比',
        '散戶多空比(口數)',
        '外資未平倉口數(大小台合計)',
        '自營(選)買權', '自營(選)賣權', '外資(選)買權', '外資(選)賣權',
        '前五大交易人留倉部位(所有)', '前10大交易人留倉部位(所有)']
    # 最後輸出結果
    resultList = []

    for d in dateList:
        dateObj = DateObj.DateObj(d)
        singleDayList = []
        singleDayList.append(dateObj.dateSlash)

        #'發行量加權股價指數', '漲跌指數', '漲跌趴數', '成交量'
        t = Taeix.Taeix(dateObj)
        tList = t.GetTaeixList()
        singleDayList.extend(tList)

        # '外資買賣超'
        fi = ForeignInvestor.ForeignInvestor(dateObj)
        fiStr = fi.GetCount()
        singleDayList.append(fiStr)

        #'P/C ratio'
        p = PcRatio.PCRatio(dateObj)
        pStr = p.GetRatio()
        singleDayList.append(pStr)

        #'小台指未沖銷契約量, 散戶看多, 散戶看空, 散戶多空比'
        m = MTX.MTX(dateObj)
        m.CalCount()
        # singleDayList.append(formatNum(m.mtx_notsell))
        # singleDayList.append(formatNum(m.bull))
        # singleDayList.append(formatNum(m.bear))
        singleDayList.append(format(m.ratio, '.2%'))

        ff = ForeignFuture.ForeignFuture(dateObj)
        ff.CalCount()
        singleDayList.append(formatNum(ff.future))

        #'自營(選)買權', '自營(選)賣權', '外資(選)買權', '外資(選)賣權'
        s = StockOption.StockOption(dateObj)
        s.CalCount()
        singleDayList.append(formatNum(s.dealer_call))
        singleDayList.append(formatNum(s.dealer_put))
        singleDayList.append(formatNum(s.foreign_call))
        singleDayList.append(formatNum(s.foreign_put))

        #'前五大交易人留倉部位(所有)', '前10大交易人留倉部位(所有)'
        fat = FiveAndTen.FiveAndTen(dateObj)
        fat.CalCount()
        singleDayList.append(formatNum(fat.five_all))
        singleDayList.append(formatNum(fat.ten_all))

        resultList.append(singleDayList)

    SaveDirectory = os.getcwd()  # 印出目前工作目錄
    SaveAs = os.path.join(SaveDirectory, 'daily',
                          'TWSE.csv')  # 組合路徑，自動加上兩條斜線 "\\"

    # csv file exist or not
    # os.path.isfile('test.txt') #如果不存在就返回False
    # os.path.exists(directory) #如果目錄不存在就返回False
    csvExist = os.path.isfile(SaveAs)

    if csvExist:
        # 開啟 CSV 檔案
        rows_list = []
        with open(SaveAs, newline='') as csvfile:
            # 讀取 CSV 檔案內容
            rows = csv.reader(csvfile)
            rows_list.extend(rows)

        with open(SaveAs, 'w', newline='') as csvfile:
            # 以空白分隔欄位，建立 CSV 檔寫入器
            writer = csv.writer(csvfile)
            for l in resultList:
                rows_list.insert(1, l)

            for l in rows_list:
                writer.writerow(l)
    else:
        with open(SaveAs, 'w', newline='') as csvfile:
            # 以空白分隔欄位，建立 CSV 檔寫入器
            writer = csv.writer(csvfile)
            writer.writerow(resultList_tw)
            for l in resultList:
                writer.writerow(l)

    # with open(SaveAs, 'a', newline='') as csvfile:
    #     # 以空白分隔欄位，建立 CSV 檔寫入器
    #     writer = csv.writer(csvfile)
    #     writer.writerow(resultList_tw)
    #     for l in resultList:
    #         writer.writerow(l)
