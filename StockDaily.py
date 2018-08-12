# 主程式
from package import StockMarketExport
from package.ScrapyBuySell import TWSEBuySell, OTCBuySell
import encodings

#dateList = ['20180716','20180717','20180718','20180719','20180720','20180723','20180724']
dateList = ['20180809']

# 大盤excel
StockMarketExport.main(dateList)

# 上市外資(含陸資)、投信買賣超前三十
TWSEBuySell.main(dateList)

# 上櫃外資(含陸資)、投信買賣超前三十
OTCBuySell.main(dateList)
