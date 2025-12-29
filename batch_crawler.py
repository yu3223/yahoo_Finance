import yfinance as yf
import pandas as pd
import twstock
import os

def save_with_style(df, filename):
    print("   -> 正在計算漲跌顏色樣式...")
    color_df = pd.DataFrame('color: black', index=df.index, columns=df.columns)
    
    meta_cols = ['股票代號', '股票名稱', '產業分類']
    date_cols = [c for c in df.columns if c not in meta_cols]
    
    for i in range(len(df)):
        if (i + 1) % 500 == 0:
            print(f"      已處理 {i + 1} 檔標的顏色...")
            
        for j in range(len(date_cols) - 1):
            curr_day = date_cols[j]     # 較新日期 (左)
            prev_day = date_cols[j+1]   # 較舊日期 (右)
            
            price_curr = df.loc[df.index[i], curr_day]
            price_prev = df.loc[df.index[i], prev_day]
            
            if pd.isna(price_curr) or pd.isna(price_prev):
                continue
                
            if price_curr > price_prev:
                color_df.loc[df.index[i], curr_day] = 'color: red'
            elif price_curr < price_prev:
                color_df.loc[df.index[i], curr_day] = 'color: green'

    writer = pd.ExcelWriter(filename, engine='openpyxl')
    df.style.apply(lambda x: color_df, axis=None).to_excel(writer, sheet_name='台股報表', index=False)
    
    worksheet = writer.sheets['台股報表']
    worksheet.column_dimensions['A'].width = 10
    worksheet.column_dimensions['B'].width = 15
    worksheet.column_dimensions['C'].width = 15
    
    for col in range(4, worksheet.max_column + 1):
        col_letter = worksheet.cell(row=1, column=col).column_letter
        worksheet.column_dimensions[col_letter].width = 12
        
    writer.close()

def run_stock_master():
    print("1. 正在取得台股清單 (個股 + 00開頭ETF + 加權指數)...")
    all_stocks = []
    stock_info_map = {} 
    
    # --- A. 加入加權指數 ---
    twii_code = "^TWII"
    all_stocks.append(twii_code)
    stock_info_map[twii_code] = {'code': '指數', 'name': '加權指數', 'group': '大盤'}

    # --- B. 精確篩選清單 ---
    for code, info in twstock.codes.items():
        # 邏輯：4碼個股 或 00開頭的ETF (排除權證)
        is_normal_stock = (len(code) == 4)
        is_etf = (code.startswith('00') and len(code) >= 4)
        
        if is_normal_stock or is_etf:
            suffix = ".TW" if info.market == '上市' else ".TWO"
            full_code = code + suffix
            all_stocks.append(full_code)
            stock_info_map[full_code] = {
                'code': code, 
                'name': info.name,
                'group': info.group
            }
    
    print(f"   總計共 {len(all_stocks)} 檔標的。")
    
    # --- C. 下載數據 (關閉多執行緒以避免 RuntimeError) ---
    print(f"2. 正在下載數據 (最近 8 個月)...")
    data = yf.download(all_stocks, period="8mo", interval="1d", actions=False, threads=False)
    
    if data.empty:
        print("下載失敗！")
        return

    close_prices = data['Close']
    
    print("3. 正在整理數據 (排除無數據個股並對齊最近 120 天)...")
    
    latest_120_dates = close_prices.index.sort_values(ascending=False)[:121]
    date_labels = latest_120_dates.strftime('%Y-%m-%d').tolist()
    
    final_rows = []
    
    for stock in close_prices.columns:
        s_price_aligned = close_prices[stock].reindex(latest_120_dates)
        
        if s_price_aligned.dropna().empty:
            continue 
        
        row_data = {
            '股票代號': stock_info_map[stock]['code'],
            '股票名稱': stock_info_map[stock]['name'],
            '產業分類': stock_info_map[stock]['group']
        }
        
        s_price_aligned.index = s_price_aligned.index.strftime('%Y-%m-%d')
        # 四捨五入並確保數值型態，避免 Excel 計算錯誤
        prices_dict = s_price_aligned.round(2).to_dict()
        row_data.update(prices_dict)
        
        final_rows.append(row_data)
    
    temp_df = pd.DataFrame(final_rows)
    
    # 強制將所有日期欄位轉為數值 (numeric)
    meta_cols = ['股票代號', '股票名稱', '產業分類']
    for col in date_labels:
        temp_df[col] = pd.to_numeric(temp_df[col], errors='coerce')
        
    result_df = temp_df[meta_cols + date_labels]
    
    # 讓加權指數排在第一列
    is_twii = result_df['股票代號'] == '指數'
    result_df = pd.concat([result_df[is_twii], result_df[~is_twii]])
    
    print(f"4. 正在產生報表 (tw_stock_report.xlsx)...")
    filename = "tw_stock_report.xlsx"
    save_with_style(result_df, filename)
    
    print(f"\n 完成！已加入大盤指數與 ETF，共處理 {len(result_df)} 檔有效標的。")

if __name__ == "__main__":
    run_stock_master()