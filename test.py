from enum import Enum
# enum multiple values
from aenum import MultiValueEnum
import numpy as np
import copy

ary = np.array([['1', 11], ['3', 33], ['5', 55], ['2', 22], ['4', 44]])
#ary  = np.sort(ary,axis=0)
ary = ary[ary[:, 1].argsort()]
npObj = np.array(copy.deepcopy()
                 [:2], dtype=object)
print(npObj)
