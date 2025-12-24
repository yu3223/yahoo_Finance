import yfinance as yf

# 設定股票代碼 (台股要加 .TW)
stock_id = "2408.TW"

# 抓取歷史資料
# period: 時間範圍 (1mo, 1y, max 等)
# interval: 間隔 (1d 代表日 K)
data = yf.download(stock_id, start="2025-01-01", end="2025-12-24")

# 顯示前五筆
print(data.tail())

# 存成 Excel 或 CSV
data.to_csv("2408_history.csv")