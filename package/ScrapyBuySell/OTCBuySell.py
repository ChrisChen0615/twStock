# -*- coding: utf-8 -*-
# 上櫃外資 投信買賣超前三十
import requests
from package.Infrastructure import DateObj
from package.Infrastructure import CommonMethon as _cm
from package.ScrapyBuySell import BuySell


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


# class OTC():
#     """證券櫃檯買賣中心物件"""

#     def __init__(self, dateList):
#         """Constructor"""
#         global sheet

#         self.dList = dateList
#         # 是否爬到資料
#         self.IsGetSource = True
#         # 爬到的初始資料
#         self.DataSourceByScrapy = []
#         # 整理過後的預備輸出資料
#         self.ExportData = []

#         # 外資買超前30
#         self.Foreign_Buy = np.array([])
#         # 外資賣超前30
#         self.Foreign_Sell = np.array([])
#         # 投信買超前30
#         self.Local_Buy = np.array([])
#         # 投信賣超前30
#         self.Local_Sell = np.array([])

#     def ScrapyData(self, dateObj):
#         """爬取資料
#         從證券櫃檯買賣中心網站上爬取資料
#         dataObj:日期資料物件
#         """
#         url = 'http://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php'
#         data = {
#             'l': 'zh-tw',
#             'se': 'EW',  # AL:全部、EW:不含權證、牛熊證
#             't': 'D',
#             'd': dateObj.twSlashDate
#         }
#         r = requests.get(url, params=data)
#         data_json = r.json()
#         self.DataSourceByScrapy = data_json["aaData"]
#         if len(self.DataSourceByScrapy) == 0:
#             self.IsGetSource = False

#     def ArrangeData(self):
#         """資料整理
#         將爬取的資料整理並指派給 ExportData
#         取外資、投信買賣超排行前30列
#         """
#         # 從初始資料過濾取得所需欄位
#         adjust_data = []
#         for x in self.DataSourceByScrapy:
#             one_data = [str(x[0]).strip(), x[1].strip(),
#                         _cm.formatNo(x[10]),  # 外陸資買賣超股數
#                         _cm.formatNo(x[13])]  # 投信買賣超股數
#             adjust_data.append(one_data)

#         # 外資買超前三十
#         # dtype=object:為了將張數轉為數字，不然會是字串型態
#         adjust_data.sort(key=lambda x: x[2], reverse=True)
#         foreign_Buy = np.array(copy.deepcopy(adjust_data)[:30], dtype=object)

#         # 外資賣超
#         adjust_data.sort(key=lambda x: x[2], reverse=False)
#         foreign_Sell = np.array(copy.deepcopy(adjust_data)[:30], dtype=object)

#         # 投信買超
#         adjust_data.sort(key=lambda x: x[3], reverse=True)
#         local_Buy = np.array(copy.deepcopy(adjust_data)[:30], dtype=object)

#         # 投信賣超
#         adjust_data.sort(key=lambda x: x[3], reverse=False)
#         local_Sell = np.array(copy.deepcopy(adjust_data)[:30], dtype=object)

#         # axis=None：arr会先按行展开，然后按照obj，删除第obj-1（从0开始）位置的数，返回一个行矩阵。
#         # axis = 0：arr按橫列删除
#         # axis = 1：arr按直行删除
#         # 外資刪除 投信買賣超欄位
#         foreign_Buy = np.delete(foreign_Buy, 3, axis=1)
#         foreign_Sell = np.delete(foreign_Sell, 3, axis=1)
#         # 投信刪除 外資買賣超欄位
#         local_Buy = np.delete(local_Buy, 2, axis=1)
#         local_Sell = np.delete(local_Sell, 2, axis=1)

#         # 準備最後輸出的資料
#         self.ExportData = []
#         empty_list = ['', '', '']
#         for index, elemennt in enumerate(foreign_Buy):
#             row_data = []
#             for npArray in [foreign_Buy, foreign_Sell, local_Buy, local_Sell]:
#                 if npArray[index][2] == 0:
#                     row_data.extend(empty_list)
#                 else:
#                     row_data.extend(npArray[index])
#             self.ExportData.append(row_data)

#         # 去掉買賣超中 張數為0的資料列
#         np4 = [foreign_Buy, foreign_Sell, local_Buy, local_Sell]
#         for npIdx, npElem in enumerate(np4):
#             for idx, elem in enumerate(npElem):
#                 if npElem[idx][2] == 0:
#                     np4[npIdx] = np.delete(np4[npIdx], np.s_[idx:], axis=0)
#                     break
#         self.Foreign_Buy = np4[0]
#         self.Foreign_Sell = np4[1]
#         self.Local_Buy = np4[2]
#         self.Local_Sell = np4[3]

#     def ExportExcel(self, dataObj):
#         """輸出資料至excel
#         excel資料寫入
#         dateObj:日期物件
#         """

#         resultList_tw = [
#             '代號', '公司', '外資買超張數',
#             '代號', '公司', '外資賣超張數',
#             '代號', '公司', '投信買超張數',
#             '代號', '公司', '投信賣超張數']

#         fileObj = FileIO.FileIO('OTCBuySell', dataObj.strdate)
#         row_idx = 1
#         col_idx = 1
#         yellofill = PatternFill(fill_type='solid', fgColor=colors.YELLOW)
#         greenfill = PatternFill(fill_type='solid', fgColor=colors.GREEN)
#         day2Color = colors.Color(rgb='95e1d3')
#         day2fill = PatternFill(fill_type='solid', fgColor=day2Color)
#         day3Color = colors.Color(rgb='eaffd0')
#         day3fill = PatternFill(fill_type='solid', fgColor=day3Color)
#         day4Color = colors.Color(rgb='fce38a')
#         day4fill = PatternFill(fill_type='solid', fgColor=day4Color)
#         day5Color = colors.Color(rgb='f38181')
#         day5fill = PatternFill(fill_type='solid', fgColor=day5Color)

#         wb = fileObj.GetWorkBook()
#         sheetName = dataObj.strdate[4:]

#         # 歷史資料
#         historyList = HistoryFind.GetHistory(fileObj.SaveAs)

#         if fileObj.XlsExist:
#             if sheetName in wb.sheetnames:
#                 print(sheetName + "上櫃" + dataObj.strdate + "買賣超已存在")
#                 sys.exit()
#             else:
#                 sheet = wb.create_sheet(sheetName)
#         else:
#             # 創建一個工作表(注意是一個屬性)
#             sheet = wb.active
#             # excel創建的工作表名默認為sheet1,一下代碼實現了給新創建的工作表創建一個新的名字
#             sheet.title = sheetName

#         sheet.cell(row=row_idx, column=col_idx).value = "同向"
#         sheet.cell(row=row_idx, column=col_idx).fill = yellofill

#         col_idx += 1
#         sheet.cell(row=row_idx, column=col_idx).value = "反向"
#         sheet.cell(row=row_idx, column=col_idx).fill = greenfill

#         col_idx += 1
#         sheet.cell(row=row_idx, column=col_idx).value = "同2"
#         sheet.cell(row=row_idx, column=col_idx).fill = day2fill

#         col_idx += 1
#         sheet.cell(row=row_idx, column=col_idx).value = "同3"
#         sheet.cell(row=row_idx, column=col_idx).fill = day3fill

#         col_idx += 1
#         sheet.cell(row=row_idx, column=col_idx).value = "同4"
#         sheet.cell(row=row_idx, column=col_idx).fill = day4fill

#         col_idx += 1
#         sheet.cell(row=row_idx, column=col_idx).value = "同5"
#         sheet.cell(row=row_idx, column=col_idx).fill = day5fill

#         row_idx += 1
#         col_idx = 1

#         # 向工作表中輸入內容1-標題
#         sheet.append(resultList_tw)

#         row_idx += 1
#         # 內容
#         for row in self.ExportData:
#             col_idx = 1
#             for cell in row:
#                 sheet.cell(row=row_idx, column=col_idx).value = cell
#                 if (col_idx % 3) == 0:
#                     quotient = int(col_idx / 3)
#                     stockNo = sheet.cell(row=row_idx, column=(
#                         (quotient - 1)*3) + 1).value
#                     cnt = HistoryFind.GetItemCount(
#                         historyList, quotient - 1, stockNo)
#                     cnt += 1
#                     sheet.cell(
#                         row=row_idx, column=col_idx).number_format = '#,##0'
#                     if cnt == 2:
#                         sheet.cell(row=row_idx, column=col_idx).fill = day2fill
#                     if cnt == 3:
#                         sheet.cell(row=row_idx, column=col_idx).fill = day3fill
#                     if cnt == 4:
#                         sheet.cell(row=row_idx, column=col_idx).fill = day4fill
#                     if cnt == 5:
#                         sheet.cell(row=row_idx, column=col_idx).fill = day5fill

#                 if col_idx == 2:  # 外資買超
#                     if _cm.FindNumpyIdx(self.Local_Buy, cell, 1) > 0:
#                         sheet.cell(
#                             row=row_idx, column=col_idx).fill = yellofill
#                     elif _cm.FindNumpyIdx(self.Local_Sell, cell, 1) > 0:
#                         sheet.cell(
#                             row=row_idx, column=col_idx).fill = greenfill
#                 if col_idx == 5:  # 外資賣超
#                     if _cm.FindNumpyIdx(self.Local_Sell, cell, 1) > 0:
#                         sheet.cell(
#                             row=row_idx, column=col_idx).fill = yellofill
#                     elif _cm.FindNumpyIdx(self.Local_Buy, cell, 1) > 0:
#                         sheet.cell(
#                             row=row_idx, column=col_idx).fill = greenfill

#                 if col_idx == 8:  # 投信買超
#                     if _cm.FindNumpyIdx(self.Foreign_Buy, cell, 1) > 0:
#                         sheet.cell(
#                             row=row_idx, column=col_idx).fill = yellofill
#                     elif _cm.FindNumpyIdx(self.Foreign_Sell, cell, 1) > 0:
#                         sheet.cell(
#                             row=row_idx, column=col_idx).fill = greenfill
#                 if col_idx == 11:  # 投信賣超
#                     if _cm.FindNumpyIdx(self.Foreign_Sell, cell, 1) > 0:
#                         sheet.cell(
#                             row=row_idx, column=col_idx).fill = yellofill
#                     elif _cm.FindNumpyIdx(self.Foreign_Buy, cell, 1) > 0:
#                         sheet.cell(
#                             row=row_idx, column=col_idx).fill = greenfill
#                 col_idx += 1
#             row_idx += 1
#         # 保存一個文檔
#         wb.save(fileObj.SaveAs)

#     def Execute(self):
#         for d in self.dList:
#             dateObj = DateObj.DateObj(d)
#             # 爬取資料
#             self.ScrapyData(dateObj)
#             # 判斷是否有爬取到資料
#             if not self.IsGetSource:
#                 print("上櫃買賣超讀取失敗")
#                 sys.exit()
#             # 資料整理
#             self.ArrangeData()
#             # 資料輸出
#             self.ExportExcel(dateObj)
