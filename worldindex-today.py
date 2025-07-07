import pandas as pd
from datetime import datetime, timedelta
import os
import yfinance as yf
import time
import random
import csv

def get_ticker_data(ticker, start_date, end_date, retry_count=3, delay=10):
    """
    獲取單一股票代碼數據，含重試機制
    """
    for attempt in range(retry_count):
        try:
            # 加入超時設定
            stock = yf.Ticker(ticker)
            data = stock.history(start=start_date, end=end_date + timedelta(days=1), 
                                interval="1d")
            
            if not data.empty:
                return data
            elif attempt < retry_count - 1:  # 如果數據為空且還有重試次數
                print(f"  獲取 {ticker} 返回空數據，{attempt+1}/{retry_count} 次嘗試，等待 {delay} 秒...")
                time.sleep(delay)
            else:
                return None
        except Exception as e:
            if attempt < retry_count - 1:  # 如果失敗且還有重試次數
                print(f"  獲取 {ticker} 失敗: {str(e)}，{attempt+1}/{retry_count} 次嘗試，等待 {delay} 秒...")
                time.sleep(delay)
            else:
                print(f"  最終無法獲取 {ticker} 數據: {str(e)}")
                return None
    
    return None  # 所有重試都失敗

def get_financial_data():
    # 設定日期範圍（只取當天）
    end_date = datetime.now()
    start_date = end_date - timedelta(days=5)  # 修改為獲取最近5天的數據，增加獲取成功機率
    
    # 使用日期格式建立資料夾名稱 (格式如 20250501)
    date_folder = datetime.now().strftime("%Y%m%d")
    
    # 使用相對路徑建立資料夾 (在程式根目錄下)
    output_dir = os.path.join(date_folder)
    
    print(f"正在獲取 {start_date.strftime('%Y-%m-%d')} 到 {end_date.strftime('%Y-%m-%d')} 的金融數據...")
    
    # 創建日期目錄 (相對位置)
    try:
        os.makedirs(output_dir, exist_ok=True)
        print(f"輸出目錄: {os.path.abspath(output_dir)}")
    except Exception as e:
        print(f"創建目錄失敗: {str(e)}")
        # 備選方案: 使用當前目錄
        output_dir = "."
        print(f"改用當前目錄: {os.path.abspath(output_dir)}")
    
    # 創建DataFrame來存儲數據
    all_results_df = pd.DataFrame(columns=['指標', '開盤', '最高', '最低', '收盤', '更新日期'])
    errors = []
    
    # 測試連接 - 先試一個常見的股票來測試API連接是否正常
    print("正在測試與Yahoo Finance的連接...")
    test_data = get_ticker_data("AAPL", start_date, end_date)
    if test_data is None or test_data.empty:
        print("警告: 無法連接到Yahoo Finance API，請檢查您的網絡連接")
        errors.append("API連接測試失敗")
    else:
        print("API連接測試成功!")
    
    # 1. 獲取公債殖利率 (使用Yahoo Finance替代FRED)
    bonds = {
        "10db": "^TNX",
        "01db": "2YY=F",
        "1001": "USDTWD=X",
        "1003": "HKD=X",
        "1004": "CNY=X",
        "1006": "KRW=F",
        "1005": "JPY=X",
        "100d": "^DJI",
        "100n": "^IXIC",
        "100h": "^HSI",
        "100j": "^N225",
        "100a": "000001.SS",
        "100k": "^KS11",
        "100g": "GC=F",
        "100o": "CL=F",
        "10sb": "ZS=F"
    }
    
    print("\n正在獲取數據...")
    for name, ticker in bonds.items():
        print(f"  處理 {name} ({ticker})...")
        # 隨機延遲 1-3 秒，避免請求過於頻繁
        time.sleep(random.uniform(1, 3))
        
        data = get_ticker_data(ticker, start_date, end_date)
        
        if data is not None and not data.empty:
            # 處理殖利率數據 (需要除以10)
            for col in ['Open', 'High', 'Low', 'Close']:
                data[col] = data[col] / 10  # 轉換為實際百分比
            
            latest_data = data.iloc[-1]
            latest_close = latest_data['Close']*10
            latest_open = latest_data['Open']*10
            latest_high = latest_data['High']*10
            latest_low = latest_data['Low']*10
            latest_date = data.index[-1].strftime('%Y-%m-%d')
            print(f"  {name}: 開:{latest_open:.2f} 高:{latest_high:.2f} 低:{latest_low:.2f} 收:{latest_close:.2f}")
            
            # 將數據添加到結果DataFrame
            all_results_df = pd.concat([all_results_df, pd.DataFrame({
                '指標': [name],
                '開盤': [latest_open],
                '最高': [latest_high],
                '最低': [latest_low],
                '收盤': [latest_close],
                '更新日期': [latest_date]
            })], ignore_index=True)
        else:
            errors.append(f"獲取 {name} 時返回空數據或失敗")
    
    # 保存所有數據到CSV文件 (在日期資料夾內)
    combined_file = os.path.join(output_dir, "worldindex.csv")
    
    try:
        all_results_df.to_csv(combined_file, index=False, encoding='utf-8-sig')
        print(f"\n所有數據已保存到: {os.path.abspath(combined_file)}")
    except Exception as e:
        print(f"保存CSV失敗: {str(e)}")
        # 保存到當前目錄作為備選
        try:
            backup_file = "financial_data_backup.csv"
            all_results_df.to_csv(backup_file, index=False, encoding='utf-8-sig')
            print(f"已將數據保存到: {os.path.abspath(backup_file)}")
        except Exception as e2:
            print(f"保存備份也失敗: {str(e2)}")
            print("\n無法保存數據到文件。以下是收集到的數據:")
            print(all_results_df.head(10))
            print("... 共 %d 條記錄" % len(all_results_df))
    
    # 如果有錯誤，保存錯誤日誌
    if errors:
        try:
            error_file = os.path.join(output_dir, "errors.log")
            with open(error_file, 'w', encoding='utf-8') as f:
                f.write("\n".join(errors))
            print(f"\n有 {len(errors)} 個錯誤發生，詳情請查看: {os.path.abspath(error_file)}")
        except Exception as e:
            print(f"保存錯誤日誌失敗: {str(e)}")
            print("\n錯誤摘要:")
            for i, error in enumerate(errors[:10]):
                print(f"{i+1}: {error}")
            if len(errors) > 10:
                print(f"... 以及另外 {len(errors) - 10} 個錯誤")
    
    print(f"\n成功獲取了 {len(all_results_df)} 條數據記錄，失敗 {len(errors)} 次")
    return {"data": all_results_df, "errors": errors}  # 返回字典

#資料切割
def process_csv_file(csv_file_path, output_directory):
    # 確保輸出目錄存在
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    with open(csv_file_path, 'r', encoding='utf-8') as file:
        # 使用csv模塊讀取檔案
        csv_reader = csv.reader(file)
        
        # 跳過標題行
        next(csv_reader)
        
        # 處理每一行數據
        for row in csv_reader:
            if len(row) >= 6:  # 確保有足夠的列數
                index_name = row[0]  # 指標名稱
                open_price = row[1]  # 開盤價
                high_price = row[2]  # 最高價
                low_price = row[3]   # 最低價
                close_price = row[4] # 收盤價
                update_date = row[5].replace('-', '')  # 更新日期，移除破折號
                
                # 輸出文件路徑
                output_file_path = os.path.join(output_directory, f"{index_name}.txt")
                
                # 創建輸出行，格式: "日期","開盤","最高","最低","收盤","成交量(固定為1000)"
                output_line = f'"{update_date}","{open_price}","{high_price}","{low_price}","{close_price}","1000"\n'
                
                # 寫入到對應的文件
                with open(output_file_path, 'a', encoding='utf-8') as output_file:
                    output_file.write(output_line)
                
                print(f"已處理 {index_name}，寫入到 {output_file_path}")

def start_data_collection():
    """啟動數據收集主函數，含錯誤處理"""
    try:
        print(f"開始執行財務數據收集 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("===============================================================")
        print("注意: 此程序將進行多次API請求，可能需要數分鐘完成。")
        print("程式將在當前目錄下建立以日期命名的資料夾存放數據。")
        print("===============================================================")
        result = get_financial_data()
        result_df = result["data"]
        errors = result["errors"]
        print(f"數據收集完成 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 判斷是否所有數據都成功下載
        if len(errors) == 0:
            print("所有數據下載成功！準備進行數據寫入...")
            
            # 獲取當前日期作為目錄名稱
            date_folder = datetime.now().strftime("%Y%m%d")
            csv_file_path = os.path.join(date_folder, "worldindex.csv")
            output_directory = "D:/stock/txt"  # 你的輸出目錄
            
            # 調用 process_csv_file 處理數據並寫入文件
            try:
                process_csv_file(csv_file_path, output_directory)
                print("所有數據已成功寫入到文件中！")
            except Exception as e:
                print(f"寫入數據到文件時發生錯誤: {str(e)}")
        else:
            print(f"警告：有 {len(errors)} 個數據下載失敗，請檢查錯誤日誌後再決定是否繼續處理。")
            # 你可以在這裡添加代碼，讓用戶選擇是否仍然繼續處理現有數據
            
        return result_df, errors  # 同時返回數據和錯誤信息
    except Exception as e:
        print(f"程序執行過程中出現未捕獲的異常: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return None, [f"未捕獲的異常: {str(e)}"]  # 發生異常時也返回兩個值

if __name__ == "__main__":
    # 執行數據收集並接收返回值
    data_df, data_errors = start_data_collection()
    
    # 在主程序中也可以基於 data_errors 進行進一步的處理
    if data_df is not None and len(data_errors) == 0:
        print("數據處理完成，所有任務已成功執行！")
    elif data_df is not None:
        print(f"數據處理部分完成，但有 {len(data_errors)} 個錯誤。")
    else:
        print("數據處理失敗！")