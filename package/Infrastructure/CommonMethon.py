# -*- coding: utf-8 -*-
# 共用方法
from package.Infrastructure import DateObj, FileIO
import numpy as np
from decimal import Decimal


def formatNo(noStr):
    """    
    格式化數字
    noStr:股票股數
    return 股票張數
    """
    noOrg = noStr.replace(',', '')
    noInt = int(noOrg)
    return int(noInt / 1000)


def formatNumType(no, t):
    """格式化數字，加千分位"""
    if t == "int":
        return int(no)
    elif t == "float":
        return float(no)
    elif t == "decimal":
        return Decimal(no).quantize(Decimal('0.00'))


def FindNumpyIdx(npObj, elem, idx):
    """
    從numpy object指定欄位idx取得elem的ndarray
    npObj:Numpy object
    elem:比對value值
    idx:npObj的指定比對欄位
    ref:https://ppt.cc/fHOUmx
    """
    return npObj[np.where(np.isin(npObj[:, idx], elem))]


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
