# -*- coding: utf-8 -*-
# 共用方法
from package.Infrastructure import DateObj, FileIO
from openpyxl import Workbook
from openpyxl import load_workbook
import numpy as np


def formatNo(noStr):
    """    
    格式化數字
    noStr:股票股數
    return 股票張數
    """
    noOrg = noStr.replace(',', '')
    noInt = int(noOrg)
    return int(noInt / 1000)

def FindNumpyIdx(npObj, elem, idx):
    """
    從numpy object指定欄位idx取得elem的ndarray
    npObj:Numpy object
    elem:比對value值
    idx:npObj的指定比對欄位
    ref:https://ppt.cc/fHOUmx
    """
    return npObj[np.where(np.isin(npObj[:,idx],elem))]


def FindListIdx(listObj, elem):
    """
    搜尋item在list的索引值
    listObj:搜尋範圍
    elem:搜尋目標
    list.index(ele) 假設沒找到會丟value error
    """
    try:
        idx = listObj.index(elem)
        return 1
    except ValueError:
        return -1

    # if elem.strip() == "":
    #     return -1

    # for row, i in enumerate(listObj):
    #     try:
    #         column = i.index(elem)
    #     except ValueError:
    #         continue
    #     return 1
    # return -1
