# 主程式
import os
import csv
from package.Infrastructure import DateObj
from package.ScrapySource import Taeix, ForeignInvestor, PcRatio, MTX
from package.ScrapySource import ForeignFuture, StockOption, FiveAndTen
from openpyxl import Workbook
from openpyxl import load_workbook
# 使用openpyxl内置的格式
from openpyxl.styles import numbers
from decimal import Decimal


def formatNumType(no, t):
    """格式化數字，加千分位"""
    if t == "int":
        # if type(no) != int:
        #     return 0
        # else:
        #     return format(no, ',')
        return int(no)
    elif t == "float":
        return float(no)
    elif t == "decimal":
        return Decimal(no).quantize(Decimal('0.00'))


def main(dList):
    # 日期list
    dateList = dList
    # 最後輸出結果標題列
    resultList_tw = [
        '日期',
        '加權指數', '漲跌指數', '漲跌趴數', '成交量(億)',
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

        #'加權指數', '漲跌指數', '漲跌趴數', '成交量'
        t = Taeix.Taeix(dateObj)
        tList = t.GetTaeixList()
        for idx in range(0, len(tList[:-1])):
            tList[idx] = tList[idx].replace(',', '')
            tList[idx] = formatNumType(tList[idx], "float")  # string to float
        tList[-1] = formatNumType(tList[-1], "decimal")  # float to decimal
        singleDayList.extend(tList)

        # '外資買賣超'
        fi = ForeignInvestor.ForeignInvestor(dateObj)
        fiStr = formatNumType(fi.GetCount(), "decimal")
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

        #singleDayList.append(format(m.ratio, '.2%'))
        singleDayList.append(m.ratio)

        ff = ForeignFuture.ForeignFuture(dateObj)
        ff.CalCount()
        singleDayList.append(formatNumType(ff.future, "int"))

        #'自營(選)買權', '自營(選)賣權', '外資(選)買權', '外資(選)賣權'
        s = StockOption.StockOption(dateObj)
        s.CalCount()
        singleDayList.append(formatNumType(s.dealer_call, "int"))
        singleDayList.append(formatNumType(s.dealer_put, "int"))
        singleDayList.append(formatNumType(s.foreign_call, "int"))
        singleDayList.append(formatNumType(s.foreign_put, "int"))

        #'前五大交易人留倉部位(所有)', '前10大交易人留倉部位(所有)'
        fat = FiveAndTen.FiveAndTen(dateObj)
        fat.CalCount()
        singleDayList.append(formatNumType(fat.five_all, "int"))
        singleDayList.append(formatNumType(fat.ten_all, "int"))

        resultList.append(singleDayList)

    SaveDirectory = os.getcwd()  # 印出目前工作目錄
    SaveAs = os.path.join(SaveDirectory, 'daily',
                          'TWSE.xlsx')  # 組合路徑，自動加上兩條斜線 "\\"

    # csv file exist or not
    # os.path.isfile('test.txt') #如果不存在就返回False
    # os.path.exists(directory) #如果目錄不存在就返回False
    csvExist = os.path.isfile(SaveAs)

    row_idx = 1  # 列 起始:1
    col_idx = 1  # 行 起始:1

    # cell value type皆是float,decimal，再formatting成指定格式
    if csvExist:
        wb = load_workbook(filename=SaveAs)
        sheet = wb['大盤']

        # 內容
        for row in resultList:
            sheet.insert_rows(2)
            col_idx = 1
            for r in row:
                # 離最近日期開始排序，所以從標題下一列開始寫入資料
                sheet.cell(row=2, column=col_idx).value = r
                if col_idx > 1:
                    if (col_idx == 4) or (col_idx == 7) or (col_idx == 8):
                        sheet.cell(
                            row=2, column=col_idx).number_format = numbers.FORMAT_PERCENTAGE_00
                    elif col_idx >=9:
                        sheet.cell(
                            row=2, column=col_idx).number_format = '#,##0'
                    else:
                        sheet.cell(
                            row=2, column=col_idx).number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
                col_idx += 1

        # 保存一個文檔
        wb.save(SaveAs)

    else:
        # 創建一個工作薄
        wb = Workbook()

        # 創建一個工作表(注意是一個屬性)
        sheet = wb.active

        # excel創建的工作表名默認為sheet1,一下代碼實現了給新創建的工作表創建一個新的名字
        sheet.title = '大盤'

        # 向工作表中輸入內容1-標題
        sheet.append(resultList_tw)
        row_idx += 1

        # 內容
        for row in resultList:
            sheet.insert_rows(2)
            col_idx = 1
            for r in row:
                # 離最近日期開始排序，所以從標題下一列開始寫入資料
                sheet.cell(row=2, column=col_idx).value = r
                col_idx += 1

        # 保存一個文檔
        wb.save(SaveAs)


# #2 import csv
#     csvExist = os.path.isfile(SaveAs)
#     if csvExist:
#         # 開啟 CSV 檔案
#         rows_list = []
#         with open(SaveAs, newline='') as csvfile:
#             # 讀取 CSV 檔案內容
#             rows = csv.reader(csvfile)
#             rows_list.extend(rows)

#         with open(SaveAs, 'w', newline='') as csvfile:
#             # 以空白分隔欄位，建立 CSV 檔寫入器
#             writer = csv.writer(csvfile)
#             for l in resultList:
#                 rows_list.insert(1, l)

#             for l in rows_list:
#                 writer.writerow(l)
#     else:
#         with open(SaveAs, 'w', newline='') as csvfile:
#             # 以空白分隔欄位，建立 CSV 檔寫入器
#             writer = csv.writer(csvfile)
#             writer.writerow(resultList_tw)
#             for l in resultList:
#                 writer.writerow(l)

# 1 import csv
    # with open(SaveAs, 'a', newline='') as csvfile:
    #     # 以空白分隔欄位，建立 CSV 檔寫入器
    #     writer = csv.writer(csvfile)
    #     writer.writerow(resultList_tw)
    #     for l in resultList:
    #         writer.writerow(l)
