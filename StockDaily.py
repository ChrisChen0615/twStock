# 主程式
from package import StockMarketExport
from package.ScrapyBuySell import TWSEBuySell, OTCBuySell, BuySell
import encodings

#dateList = ['20180806','20180807','20180808','20180809']
#dateList = ['20180813','20180814','20180815','20180809']
dateList = ['20180813']

# 大盤excel
#StockMarketExport.main(dateList)

# 上市外資(含陸資)、投信買賣超前三十
twse = TWSEBuySell.TWSE(dateList)
twse.Execute()

# 上櫃外資(含陸資)、投信買賣超前三十
otc = OTCBuySell.OTC(dateList)
otc.Execute()
