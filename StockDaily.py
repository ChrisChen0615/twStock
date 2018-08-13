# 主程式
from package import StockMarketExport
from package.ScrapyBuySell import TWSEBuySell, OTCBuySell
import encodings

#dateList = ['20180806','20180807','20180808','20180809']
dateList = ['20180810']

# 大盤excel
StockMarketExport.main(dateList)

# 上市外資(含陸資)、投信買賣超前三十
TWSEBuySell.main(dateList)

# 上櫃外資(含陸資)、投信買賣超前三十
OTCBuySell.main(dateList)
