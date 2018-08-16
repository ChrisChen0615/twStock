# -*- coding: utf-8 -*-
# 買賣超前三十資料輸出產生excel sheet物件
import sys
import copy
import numpy as np
from package.Infrastructure import DateObj, FileIO, HistoryFind
from package.Infrastructure import CommonMethon as _cm
# 使用openpyxl内置的格式
from openpyxl.styles import numbers, Font, PatternFill, colors
# enum multiple values
from aenum import MultiValueEnum

"""
使用copy.deepcopy 完整複製一份reference object
因為使用list.copy() 假設新list pop時 會異動到舊list
(只有clear()不會 append()亦會影響)
"""


class ColorEnum(MultiValueEnum):
    """
    顏色代碼aenum
    """
    # 外資、投信同步 買或賣超
    Yello = 0, PatternFill(fill_type='solid', fgColor=colors.YELLOW)
    # 外資、投信不同步 買或賣超
    Green = 1, PatternFill(fill_type='solid', fgColor=colors.GREEN)
    # 5日內累積買或賣超 2天
    Cnt2Days = 2, PatternFill(
        fill_type='solid', fgColor=colors.Color(rgb='95E1D3'))
    # 5日內累積買或賣超 3天
    Cnt3Days = 3, PatternFill(
        fill_type='solid', fgColor=colors.Color(rgb='EAFFD0'))
    # 5日內累積買或賣超 4天
    Cnt4Days = 4, PatternFill(
        fill_type='solid', fgColor=colors.Color(rgb='FCE38A'))
    # 5日內累積買或賣超 5天
    Cnt5Days = 5, PatternFill(
        fill_type='solid', fgColor=colors.Color(rgb='F38181'))


def GetColorRGB(val):
    """
    回傳顏色rgb
    val:enum value[0]
    """
    return ColorEnum(val).values[1]


class BuySell():
    """資料輸出產生excel sheet物件"""

    def __init__(self, dateList):
        """Constructor"""
        self.name = ""
        self.EnName = ""

        global sheet

        self.dList = dateList
        # 是否爬到資料
        self.IsGetSource = True
        # 爬到的初始資料
        self.DataSourceByScrapy = []
        # 從初始資料過濾取得所需欄位
        self.FilterData = []
        # 整理過濾後的預備輸出資料
        self.ExportData = []

        # 外資買超前30
        self.Foreign_Buy = np.array([])
        # 外資賣超前30
        self.Foreign_Sell = np.array([])
        # 投信買超前30
        self.Local_Buy = np.array([])
        # 投信賣超前30
        self.Local_Sell = np.array([])

    def ScrapyData(self, dateObj):
        """爬取資料
        dataObj:日期資料物件
        """

    def FilterScrapyData(self):
        """
        資料整理
        將爬取的資料整理並指派給 FilterData
        """

    def ArrangeData(self):
        """資料整理
        將爬取的資料整理並指派給 ExportData
        取外資、投信買賣超排行前30列
        """
        # 外資買超前三十
        # dtype=object:為了將張數轉為數字，不然會是字串型態
        self.FilterData.sort(key=lambda x: x[2], reverse=True)
        foreign_Buy = np.array(copy.deepcopy(
            self.FilterData)[:30], dtype=object)

        # 外資賣超
        self.FilterData.sort(key=lambda x: x[2], reverse=False)
        foreign_Sell = np.array(copy.deepcopy(
            self.FilterData)[:30], dtype=object)

        # 投信買超
        self.FilterData.sort(key=lambda x: x[3], reverse=True)
        local_Buy = np.array(copy.deepcopy(self.FilterData)[:30], dtype=object)

        # 投信賣超
        self.FilterData.sort(key=lambda x: x[3], reverse=False)
        local_Sell = np.array(copy.deepcopy(
            self.FilterData)[:30], dtype=object)

        # axis=None：arr会先按行展开，然后按照obj，删除第obj-1（从0开始）位置的数，返回一个行矩阵。
        # axis = 0：arr按橫列删除
        # axis = 1：arr按直行删除
        # 外資刪除 投信買賣超欄位
        foreign_Buy = np.delete(foreign_Buy, 3, axis=1)
        foreign_Sell = np.delete(foreign_Sell, 3, axis=1)
        # 投信刪除 外資買賣超欄位
        local_Buy = np.delete(local_Buy, 2, axis=1)
        local_Sell = np.delete(local_Sell, 2, axis=1)

        # 準備最後輸出的資料
        self.ExportData = []
        empty_list = ['', '', '']
        for index, elemennt in enumerate(foreign_Buy):
            row_data = []
            for npArray in [foreign_Buy, foreign_Sell, local_Buy, local_Sell]:
                if npArray[index][2] == 0:
                    row_data.extend(empty_list)
                else:
                    row_data.extend(npArray[index])
            self.ExportData.append(row_data)

        # 去掉買賣超中 張數為0的資料列
        np4 = [foreign_Buy, foreign_Sell, local_Buy, local_Sell]
        for npIdx, npElem in enumerate(np4):
            for idx, elem in enumerate(npElem):
                if npElem[idx][2] == 0:
                    np4[npIdx] = np.delete(np4[npIdx], np.s_[idx:], axis=0)
                    break
        self.Foreign_Buy = np4[0]
        self.Foreign_Sell = np4[1]
        self.Local_Buy = np4[2]
        self.Local_Sell = np4[3]

    def ExportExcel(self, dataObj):
        """輸出資料至excel
        excel資料寫入
        dateObj:日期物件
        """

        resultList_tw = [
            '代號', '公司', '外資買超張數',
            '代號', '公司', '外資賣超張數',
            '代號', '公司', '投信買超張數',
            '代號', '公司', '投信賣超張數']

        fileObj = FileIO.FileIO(self.EnName, dataObj.strdate)
        row_idx = 1
        col_idx = 1

        wb = fileObj.GetWorkBook()
        sheetName = dataObj.strdate[4:]

        # 歷史資料
        historyList = HistoryFind.GetHistory(fileObj.SaveAs)

        if fileObj.XlsExist:
            if sheetName in wb.sheetnames:
                print(sheetName + self.name + dataObj.strdate + "買賣超已存在")
                # sys.exit()
                return
            else:
                sheet = wb.create_sheet(sheetName)
        else:
            # 創建一個工作表(注意是一個屬性)
            sheet = wb.active
            # excel創建的工作表名默認為sheet1,一下代碼實現了給新創建的工作表創建一個新的名字
            sheet.title = sheetName

        sheet.cell(row=row_idx, column=col_idx).value = "同向"
        sheet.cell(
            row=row_idx, column=col_idx).fill = ColorEnum.Yello.values[1]

        col_idx += 1
        sheet.cell(row=row_idx, column=col_idx).value = "反向"
        sheet.cell(
            row=row_idx, column=col_idx).fill = ColorEnum.Green.values[1]

        col_idx += 1
        sheet.cell(row=row_idx, column=col_idx).value = "同2"
        sheet.cell(
            row=row_idx, column=col_idx).fill = ColorEnum.Cnt2Days.values[1]

        col_idx += 1
        sheet.cell(row=row_idx, column=col_idx).value = "同3"
        sheet.cell(
            row=row_idx, column=col_idx).fill = ColorEnum.Cnt3Days.values[1]

        col_idx += 1
        sheet.cell(row=row_idx, column=col_idx).value = "同4"
        sheet.cell(
            row=row_idx, column=col_idx).fill = ColorEnum.Cnt4Days.values[1]

        col_idx += 1
        sheet.cell(row=row_idx, column=col_idx).value = "同5"
        sheet.cell(
            row=row_idx, column=col_idx).fill = ColorEnum.Cnt5Days.values[1]

        row_idx += 1
        col_idx = 1

        # 向工作表中輸入內容1-標題
        sheet.append(resultList_tw)

        row_idx += 1
        # 內容
        for row in self.ExportData:
            col_idx = 1
            for cell in row:
                sheet.cell(row=row_idx, column=col_idx).value = cell
                # 買賣超張數欄位
                if (col_idx % 3) == 0:
                    quotient = int(col_idx / 3)  # 取商數
                    stockNo = sheet.cell(row=row_idx, column=(
                        col_idx - 2)).value  # 取股票代號欄位值
                    cnt = HistoryFind.GetItemCount(
                        historyList, quotient - 1, stockNo)
                    cnt += 1
                    sheet.cell(
                        row=row_idx, column=col_idx).number_format = '#,##0'
                    if cnt > 1:
                        # cnt 2~5 剛好等於 colorenum values[0]
                        sheet.cell(row=row_idx, column=col_idx).fill = GetColorRGB(
                            cnt)
                # buysell = np.array([
                #     [2, self.Local_Buy, self.Local_Sell],
                #     [5, self.Local_Sell, self.Local_Buy],
                #     [8, self.Foreign_Buy, self.Foreign_Sell],
                #     [11, self.Foreign_Sell, self.Foreign_Buy],
                # ])
                # idx _cm.FindNumpyIdx(buysell, col_idx, 0)
                # if idx > 0:
                # 股票中文名稱欄位
                if col_idx == 2:  # 外資買超
                    if _cm.FindNumpyIdx(self.Local_Buy, cell, 1) > 0:
                        sheet.cell(
                            row=row_idx, column=col_idx).fill=ColorEnum.Yello.values[1]
                    elif _cm.FindNumpyIdx(self.Local_Sell, cell, 1) > 0:
                        sheet.cell(
                            row=row_idx, column=col_idx).fill=ColorEnum.Green.values[1]
                if col_idx == 5:  # 外資賣超
                    if _cm.FindNumpyIdx(self.Local_Sell, cell, 1) > 0:
                        sheet.cell(
                            row=row_idx, column=col_idx).fill=ColorEnum.Yello.values[1]
                    elif _cm.FindNumpyIdx(self.Local_Buy, cell, 1) > 0:
                        sheet.cell(
                            row=row_idx, column=col_idx).fill=ColorEnum.Green.values[1]

                if col_idx == 8:  # 投信買超
                    if _cm.FindNumpyIdx(self.Foreign_Buy, cell, 1) > 0:
                        sheet.cell(
                            row=row_idx, column=col_idx).fill=ColorEnum.Yello.values[1]
                    elif _cm.FindNumpyIdx(self.Foreign_Sell, cell, 1) > 0:
                        sheet.cell(
                            row=row_idx, column=col_idx).fill=ColorEnum.Green.values[1]
                if col_idx == 11:  # 投信賣超
                    if _cm.FindNumpyIdx(self.Foreign_Sell, cell, 1) > 0:
                        sheet.cell(
                            row=row_idx, column=col_idx).fill=ColorEnum.Yello.values[1]
                    elif _cm.FindNumpyIdx(self.Foreign_Buy, cell, 1) > 0:
                        sheet.cell(
                            row=row_idx, column=col_idx).fill=ColorEnum.Green.values[1]
                col_idx += 1
            row_idx += 1
        # 保存一個文檔
        wb.save(fileObj.SaveAs)

    def Execute(self):
        for d in self.dList:
            dateObj = DateObj.DateObj(d)
            # 爬取資料
            self.ScrapyData(dateObj)
            # 判斷是否有爬取到資料
            if not self.IsGetSource:
                print(self.name + "買賣超讀取失敗")
                return
                # sys.exit()
            # 資料整理
            self.FilterScrapyData()
            self.ArrangeData()
            # 資料輸出
            self.ExportExcel(dateObj)
