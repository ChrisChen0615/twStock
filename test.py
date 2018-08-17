from enum import Enum
# enum multiple values
from aenum import MultiValueEnum


class ColorEnum(MultiValueEnum):
    """
    顏色代碼aenum
    """
    # 外資、投信同步 買或賣超
    Yello = 0, 'Yello'
    # 外資、投信不同步 買或賣超
    Green = 1, 'Green'
    # 5日內累積買或賣超 2天
    Cnt2Days = 2, 'Cnt2Days'
    # 5日內累積買或賣超 3天
    Cnt3Days = 3, 'Cnt3Days'
    # 5日內累積買或賣超 4天
    Cnt4Days = 4, 'Cnt4Days'
    # 5日內累積買或賣超 5天
    Cnt5Days = 5, 'Cnt5Days'


c = ColorEnum.Yello.values[1]
