# 主程式
from package import StockMarketExport
from package.ScrapyBuySell import TWSEBuySell, OTCBuySell
import encodings

#dateList = ['20180813','20180814','20180815']
dateList = ['20180823']

# 大盤excel
StockMarketExport.main(dateList)

# 上市外資(含陸資)、投信買賣超前三十
twse = TWSEBuySell.TWSE(dateList)
twse.Execute()

# 上櫃外資(含陸資)、投信買賣超前三十
otc = OTCBuySell.OTC(dateList)
otc.Execute()
