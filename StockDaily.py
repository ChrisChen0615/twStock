# 主程式
from package import StockMarketExport
from package.ScrapyBuySell import TWSEBuySell, OTCBuySell
import encodings

#dateList = ['20190102','20190103']
dateList = ['20190319']

# 大盤excel
#五大 十大留倉數值有問題
StockMarketExport.main(dateList)

# 上市外資(含陸資)、投信買賣超前三十
twse = TWSEBuySell.TWSE(dateList)
twse.Execute()

# 上櫃外資(含陸資)、投信買賣超前三十
otc = OTCBuySell.OTC(dateList)
otc.Execute()
