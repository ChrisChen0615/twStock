import numpy as np


def GetNumpyIdx(npObj, elem, idx):
    """
    從numpy object指定欄位idx取得elem的index    
    """
    return npObj[np.where(np.isin(npObj[:,idx],elem))]


data = np.array([
    [2, ['2', '22', '222'], ['2', '22', '222']],
    [5, ['5', '55', '555'], ['5', '55', '555']],
    [8, ['8', '88', '888'], ['8', '88', '888']],
    [11, ['11', '1111', '111111'], ['11', '1111', '111111']]
], dtype=object)

r = GetNumpyIdx(data, 5, 0)
print(np.size(r))

#print(data[:,0])