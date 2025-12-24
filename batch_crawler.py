import yfinance as yf
import pandas as pd
import twstock
import os

def save_with_style(df, filename):
    print("   -> 正在計算漲跌顏色樣式...")
    # 建立一個與 df 同樣大小的樣式表
    color_df = pd.DataFrame('color: black', index=df.index, columns=df.columns)
    
    # 找出所有的日期欄位 (排除文字描述欄位)
    meta_cols = ['股票代號', '股票名稱', '產業分類']
    date_cols = [c for c in df.columns if c not in meta_cols]
    
    for i in range(len(df)):
        if (i + 1) % 500 == 0:
            print(f"      已處理 {i + 1} 檔股票顏色...")
            
        for j in range(len(date_cols) - 1):
            curr_day = date_cols[j]     # 較新日期 (左)
            prev_day = date_cols[j+1]   # 較舊日期 (右)
            
            price_curr = df.loc[df.index[i], curr_day]
            price_prev = df.loc[df.index[i], prev_day]
            
            if pd.isna(price_curr) or pd.isna(price_prev):
                continue
                
            # 漲紅跌綠判斷
            if price_curr > price_prev:
                color_df.loc[df.index[i], curr_day] = 'color: red'
            elif price_curr < price_prev:
                color_df.loc[df.index[i], curr_day] = 'color: green'

    # 輸出至 Excel
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    df.style.apply(lambda x: color_df, axis=None).to_excel(writer, sheet_name='台股報表', index=False)
    
    worksheet = writer.sheets['台股報表']
    # 調整文字欄位寬度
    worksheet.column_dimensions['A'].width = 10  # 代號
    worksheet.column_dimensions['B'].width = 15  # 名稱
    worksheet.column_dimensions['C'].width = 15  # 產業分類
    
    # 調整日期欄位寬度
    for col in range(4, worksheet.max_column + 1):
        col_letter = worksheet.cell(row=1, column=col).column_letter
        worksheet.column_dimensions[col_letter].width = 12
        
    writer.close()

def run_stock_master():
    print("1. 正在取得完整台股清單 (加入產業分類)...")
    all_stocks = []
    stock_info_map = {} 
    
    for code, info in twstock.codes.items():
        # 排除 00 開頭的 ETF 並限制為 4 碼個股
        if len(code) == 4 and not code.startswith('00'):
            suffix = ".TW" if info.market == '上市' else ".TWO"
            full_code = code + suffix
            all_stocks.append(full_code)
            # 儲存代號、名稱與產業群組
            stock_info_map[full_code] = {
                'code': code, 
                'name': info.name,
                'group': info.group
            }
    
    print(f"   總計共 {len(all_stocks)} 檔個股。")
    
    print(f"2. 正在下載數據 (最近 8 個月以提取 120 天交易日)...")
    data = yf.download(all_stocks, period="8mo", interval="1d", actions=False, threads=True)
    
    if data.empty:
        print("下載失敗！")
        return

    close_prices = data['Close']
    
    print("3. 正在整理數據 (包含產業分類)...")
    final_rows = []
    
    for stock in close_prices.columns:
        s_price = close_prices[stock].dropna()
        if s_price.empty:
            continue
        
        # 取得最近 120 個交易日數據
        last_120 = s_price.tail(120).copy()
        last_120.index = last_120.index.strftime('%Y-%m-%d')
        
        # 建立列資料：代號 + 名稱 + 產業 + 價格
        row_data = {
            '股票代號': stock_info_map[stock]['code'],
            '股票名稱': stock_info_map[stock]['name'],
            '產業分類': stock_info_map[stock]['group']
        }
        
        # 合併價格數據
        prices_dict = last_120.round(1).to_dict()
        row_data.update(prices_dict)
        final_rows.append(row_data)
    
    # 建立 DataFrame
    temp_df = pd.DataFrame(final_rows)
    
    # 確保日期由新到舊排序，並將文字欄位放在最左側
    meta_cols = ['股票代號', '股票名稱', '產業分類']
    date_cols = sorted([c for c in temp_df.columns if c not in meta_cols], reverse=True)
    result_df = temp_df[meta_cols + date_cols]
    
    print(f"4. 正在產生報表 (檔案名稱: tw_stock_with_industry.xlsx)...")
    filename = "tw_stock_with_industry.xlsx"
    save_with_style(result_df, filename)
    
    print(f"\n 完成！")
    print(f"   - 欄位 A: 代號, B: 名稱, C: 產業分類。")
    print(f"   - 資料包含最近 120 個交易日收盤價。")

if __name__ == "__main__":
    run_stock_master()