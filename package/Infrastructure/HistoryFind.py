# 取得歷史買賣超紀錄、計算前四日出現次數
from openpyxl import Workbook
from openpyxl import load_workbook
from package.Infrastructure import FileIO


def GetHistory(filePath):
    """
    取得前四天外資、投信買賣超紀錄
    filePath:檔案路徑
    """
    wb = load_workbook(filename=filePath)
    wbList = wb.sheetnames
    # 最多從後面取4個sheet
    overFive = len(wbList) >= 5 and wbList[-4:] or []
    sheet4List = []  # 最多4個sheet 買賣超list
    for sheetName in overFive:
        one_sheet = wb[sheetName]
        foreign_buy = []  # 單sheet 外資買超
        foreign_sell = []  # 單sheet 外資賣超
        local_buy = []  # 單sheet 投信買超
        local_sell = []  # 單sheet 投信賣超
        sheetList = []  # 單sheet外資、投信買賣超
        for rowIdx in range(3, 33):
            for colIdx in [1, 4, 7, 10]:
                val = str(one_sheet.cell(
                    row=rowIdx, column=colIdx).value).strip()
                if val != "None":
                    if colIdx == 1:
                        foreign_buy.append(val)
                    if colIdx == 4:
                        foreign_sell.append(val)
                    if colIdx == 7:
                        local_buy.append(val)
                    if colIdx == 10:
                        local_sell.append(val)
        sheetList.append(foreign_buy)
        sheetList.append(foreign_sell)
        sheetList.append(local_buy)
        sheetList.append(local_sell)
        sheet4List.append(sheetList)

    return sheet4List


def GetItemCount(listObj, typeObj, val):
    """
    搜尋出現次數
    listObj:sheet list
    typeObj:外資買超0、外資賣超1、投信買超2、投信賣超3
    val:股票代號
    """
    cnt = 0
    for l in listObj:
        cnt += l[typeObj].count(val)
    return cnt
