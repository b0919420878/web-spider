import requests
import urllib3
import os
import urllib.parse
from datetime import datetime
import csv
import io

def download(date_str=None):
    """下載台灣股市資料（上市、上櫃、大盤五秒）
    
    參數:
        date_str: 日期字串 (YYYYMMDD格式)，若為None則使用當天日期
    """
    # 禁用SSL驗證警告
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    # 如果未提供日期，使用當天日期
    if date_str is None:
        today = datetime.now()
        date_str = today.strftime('%Y%m%d')
    
    print(f"下載日期: {date_str}")

    # 獲取櫃買中心所需的格式化日期
    formatted_date = f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:8]}"
    encoded_date = urllib.parse.quote(formatted_date)  # URL編碼
    print(f"格式化日期: {formatted_date}, 編碼: {encoded_date}")

    # 設定保存目錄和檔名 - 使用相對路徑
    # 獲取當前腳本所在目錄
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 創建以日期命名的資料夾
    date_folder = os.path.join(current_dir, date_str)
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"創建日期資料夾: {date_folder}")

    # 1. 下載上櫃資料
    print("開始下載上櫃資料...")
    otc_save_path = os.path.join(date_folder, f'櫃買_{date_str}.csv')
    otc_url = f"https://www.tpex.org.tw/www/zh-tw/afterTrading/dailyQuotes?date={encoded_date}&id=&response=csv"
    download_file(otc_url, otc_save_path, "上櫃")

    # 2. 下載大盤五秒資料
    print("開始下載大盤五秒資料...")
    index_save_path = os.path.join(date_folder, f'大盤5秒_{date_str}.csv')
    index_url = f"https://www.twse.com.tw/rwd/zh/TAIEX/MI_5MINS_INDEX?date={date_str}&response=csv"
    download_file(index_url, index_save_path, "大盤五秒")

    # 3. 下載上市資料
    print("開始下載上市資料...")
    twse_save_path = os.path.join(date_folder, f'上市_{date_str}.csv')
    twse_url = f"https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?date={date_str}&type=ALLBUT0999&response=csv"
    download_file(twse_url, twse_save_path, "上市")

    print("所有資料下載完成")
    return date_folder


def download_file(url, save_path, file_type):
    """下載檔案並儲存
    
    參數:
        url: 下載連結
        save_path: 儲存路徑
        file_type: 檔案類型描述（用於日誌顯示）
    """
    try:
        print(f"下載{file_type}資料，URL: {url}")
        # 設定請求標頭和超時
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
        }
    
        response = requests.get(url, headers=headers, verify=False, timeout=20)
        print(f"{file_type}資料請求完成，狀態碼: {response.status_code}")

        # 檢查是否成功
        if response.status_code == 200:
            # 保存檔案
            with open(save_path, 'wb') as f:
                f.write(response.content)
            print(f"已下載{file_type}資料到: {save_path}")

            # 檢查檔案大小
            file_size = os.path.getsize(save_path)
            print(f"{file_type}檔案大小: {file_size} 字節")
            
            if file_size < 100:
                print(f"警告: {file_type}檔案大小異常小，請檢查內容是否正確")
        else:
            print(f"{file_type}資料下載失敗，狀態碼: {response.status_code}")
            print(f"回應內容: {response.text[:500]}")
    except requests.exceptions.Timeout:
        print(f"{file_type}資料請求超時，服務器沒有在規定時間內響應")
    except requests.exceptions.SSLError as e:
        print(f"{file_type}資料下載出現SSL錯誤: {e}")
    except Exception as e:
        print(f"{file_type}資料下載時發生錯誤: {str(e)}")


def process_stock_data(csv_file_path, date_str, is_otc=False):
    """處理台灣股市資料函數
    
    參數:
        csv_file_path: CSV檔案路徑
        date_str: 日期字串 (YYYYMMDD格式)
        is_otc: 是否為上櫃資料 (預設False為上市資料)
    """
    print(f"開始處理{'上櫃' if is_otc else '上市'}公司資料，檔案: {csv_file_path}")
    
    # 確保目標目錄存在
    target_dir = "D:/stock/txt"
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
        print(f"創建目標目錄: {target_dir}")
    
    # 嘗試不同的編碼讀取CSV文件內容
    encodings_to_try = ['big5', 'cp950', 'utf-8-sig', 'utf-8', 'gbk']
    lines = None
    
    for encoding in encodings_to_try:
        try:
            print(f"嘗試使用 {encoding} 編碼讀取檔案...")
            with open(csv_file_path, 'r', encoding=encoding) as file:
                lines = file.readlines()
                print(f"成功使用 {encoding} 編碼讀取檔案，共 {len(lines)} 行")
                break  # 如果成功讀取則跳出迴圈
        except UnicodeDecodeError:
            print(f"{encoding} 編碼無法讀取檔案")
        except Exception as e:
            print(f"使用 {encoding} 編碼讀取時發生錯誤: {str(e)}")
    
    if lines is None:
        print("所有讀取方式均失敗，無法處理檔案")
        return
    
    # 尋找標題行 (上市和上櫃的標題行格式不同)
    header_idx = -1
    for i, line in enumerate(lines):
        if is_otc:  # 上櫃資料
            if "代號" in line and "名稱" in line and "收盤" in line and "開盤" in line and "最高" in line and "最低" in line:
                header_idx = i
                break
        else:  # 上市資料
            if "證券代號" in line:
                header_idx = i
                break
    
    if header_idx == -1:
        print(f"無法在CSV文件中找到{'上櫃' if is_otc else '上市'}資料的標題行")
        return
    
    print(f"找到標題行，行號: {header_idx}")
    print(f"標題行內容: {lines[header_idx].strip()}")
    
    # 解析標題行
    header_parts = lines[header_idx].strip().split(',')
    header_parts = [part.strip('"') for part in header_parts]
    
    # 尋找需要的列索引
    try:
        code_idx = -1   # 證券代號/代號
        open_idx = -1   # 開盤價/開盤
        high_idx = -1   # 最高價/最高
        low_idx = -1    # 最低價/最低
        close_idx = -1  # 收盤價/收盤
        volume_idx = -1 # 成交股數
        
        for i, part in enumerate(header_parts):
            # 上市和上櫃的欄位名稱可能不同
            if "證券代號" in part or "代號" in part:
                code_idx = i
            elif "開盤價" in part or "開盤" in part:
                open_idx = i
            elif "最高價" in part or "最高" in part:
                high_idx = i
            elif "最低價" in part or "最低" in part:
                low_idx = i
            elif "收盤價" in part or "收盤" in part:
                close_idx = i
            elif "成交股數" in part or "成交量" in part:
                volume_idx = i
    
        print(f"欄位索引: 代號({code_idx}) 開盤({open_idx}) 最高({high_idx}) 最低({low_idx}) 收盤({close_idx}) 成交量({volume_idx})")
        
        # 確認所有需要的列都找到了
        required_indices = [code_idx, open_idx, high_idx, low_idx, close_idx, volume_idx]
        if -1 in required_indices:
            missing_columns = []
            if code_idx == -1: missing_columns.append("證券代號/代號")
            if open_idx == -1: missing_columns.append("開盤價/開盤")
            if high_idx == -1: missing_columns.append("最高價/最高")
            if low_idx == -1: missing_columns.append("最低價/最低")
            if close_idx == -1: missing_columns.append("收盤價/收盤")
            if volume_idx == -1: missing_columns.append("成交股數/成交量")
            
            print(f"無法找到所有需要的列: {', '.join(missing_columns)}")
            return
            
    except Exception as e:
        print(f"處理標題行時出錯: {str(e)}")
        return
    
    # 從標題行後開始處理數據
    processed_count = 0
    skipped_count = 0
    
    # 從標題行下一行開始處理每一行數據
    for i in range(header_idx + 1, len(lines)):
        line = lines[i].strip()
        if not line:  # 跳過空行
            continue
        
        # 使用CSV解析器來正確處理引號和逗號
        try:
            csv_reader = csv.reader(io.StringIO(line))
            row = next(csv_reader)
            
            # 如果行的元素數量不足，則跳過
            max_idx = max(required_indices)
            if len(row) <= max_idx:
                skipped_count += 1
                continue
            
            # 獲取公司代碼
            company_code = row[code_idx].strip().replace('"', '')
            
            # 忽略非股票代碼格式的行
            if not (company_code.isdigit() or (len(company_code) > 0 and company_code[0].isdigit())):
                skipped_count += 1
                continue
                
            # 獲取價格和成交量數據，移除引號和千分位逗號
            try:
                open_price = row[open_idx].strip().replace('"', '').replace(',', '')
                high_price = row[high_idx].strip().replace('"', '').replace(',', '')
                low_price = row[low_idx].strip().replace('"', '').replace(',', '')
                close_price = row[close_idx].strip().replace('"', '').replace(',', '')
                volume = row[volume_idx].strip().replace('"', '').replace(',', '')
                
                # 跳過無效數據
                if '--' in [open_price, high_price, low_price, close_price] or not volume:
                    skipped_count += 1
                    continue
                
                # 轉換成浮點數和整數以驗證（僅驗證，不改變原始字符串）
                try:
                    float(open_price)
                    float(high_price)
                    float(low_price)
                    float(close_price)
                    volume = round(int(volume)/1000)
                except ValueError:
                    skipped_count += 1
                    continue
                
                # 準備寫入的數據行
                taiwan_date = str(int(date_str[:4]) - 1911) + date_str[4:]
                data_line = f'"{taiwan_date}","{open_price}","{high_price}","{low_price}","{close_price}","{volume}"\n'
                
                # 直接寫入對應公司代碼的檔案
                file_path = os.path.join(target_dir, f"{company_code}.txt")
                
                # 檢查檔案是否存在，如果存在則附加，否則創建新檔案
                with open(file_path, 'a', encoding='utf-8') as stock_file:
                    stock_file.write(data_line)
                
                processed_count += 1
                if processed_count % 50 == 0:
                    print(f"已處理 {processed_count} 支{'上櫃' if is_otc else '上市'}股票")
                    
            except Exception as e:
                print(f"處理代碼 {company_code} 時出錯: {str(e)}")
                skipped_count += 1
                continue
            
        except Exception as e:
            print(f"處理行 {i} 時出錯: {str(e)}")
            skipped_count += 1
    
    print(f"處理完成! 成功處理 {processed_count} 支{'上櫃' if is_otc else '上市'}股票，跳過 {skipped_count} 行")

def process_index_5sec_data(index_csv_file_path, date_str):
    """處理台灣證券交易所的每日大盤5秒資料，從9:03:00到13:30:00的資料"""
    print(f"開始處理大盤5秒資料，檔案: {index_csv_file_path}")
    
    # 嘗試不同的編碼讀取CSV文件內容
    encodings_to_try = ['big5', 'cp950', 'utf-8-sig', 'utf-8', 'gbk']
    lines = None
    
    
    for encoding in encodings_to_try:
        try:
            print(f"嘗試使用 {encoding} 編碼讀取檔案...")
            with open(index_csv_file_path, 'r', encoding=encoding) as file:
                lines = file.readlines()
                print(f"成功使用 {encoding} 編碼讀取檔案，共 {len(lines)} 行")
                break  # 如果成功讀取則跳出迴圈
        except UnicodeDecodeError:
            print(f"{encoding} 編碼無法讀取檔案")
        except Exception as e:
            print(f"使用 {encoding} 編碼讀取時發生錯誤: {str(e)}")
    
    if lines is None:
        print("所有讀取方式均失敗，無法處理檔案")
        return
    
    # 找到標題行
    header_idx = -1
    for i, line in enumerate(lines):
        if "時間" in line and "發行量加權股價指數" in line:
            header_idx = i
            break
    
    if header_idx == -1:
        print("無法在CSV文件中找到標題行")
        return
    
    print(f"找到標題行，行號: {header_idx}")
    print(f"標題行內容: {lines[header_idx].strip()}")
    
    # 解析標題行找出成交量欄位位置
    header_parts = lines[header_idx].strip().split(',')
    header_parts = [part.strip('"').replace('="', '') for part in header_parts]
    
    volume_idx = -1
    for i, part in enumerate(header_parts):
        if "成交量" in part or "成交金額" in part:
            volume_idx = i
            break
    
    if volume_idx == -1:
        print("警告: 找不到成交量欄位，將使用預設值")
    else:
        print(f"找到成交量欄位在第 {volume_idx} 欄")
    
    # 尋找開盤時間9:03:00的記錄
    start_time = "09:03:00"
    end_time = "13:30:00"
    open_index = None
    high_index = None
    low_index = None
    close_index = None
    final_volume = None 
    found_start = False
    
    print(f"開始尋找時間範圍 {start_time} 到 {end_time} 的記錄")
    
    # 從標題行後開始處理每一行數據
    for i in range(header_idx + 1, len(lines)):
        line = lines[i].strip()
        if not line:  # 跳過空行
            continue
        
        # 使用CSV解析器來正確處理引號和逗號
        try:
            csv_reader = csv.reader(io.StringIO(line))
            row = next(csv_reader)
            
            # 確保行中有足夠的元素
            if len(row) < 2:  # 至少需要時間和指數值
                continue
            
            # 去除等號前綴和引號
            time_value = row[0].replace('="', '').replace('"', '')
            
            # 提取指數值並轉換為浮點數
            try:
                index_value = float(row[1].replace('"', '').replace(',', ''))
            except (ValueError, IndexError):
                continue
            
            # 提取成交量（如果有的話）
            if volume_idx != -1 and len(row) > volume_idx:
                try:
                    volume_value = row[volume_idx].replace('"', '').replace(',', '').replace('="', '')
                    if volume_value and volume_value.replace('.', '').isdigit():
                        # 轉換成千為單位
                        final_volume = str(round(float(volume_value) / 1000))
                except:
                    pass  # 如果轉換失敗就使用前一個值
            
            # 找到開盤時間
            if start_time in time_value:
                open_index = index_value
                high_index = index_value
                low_index = index_value
                found_start = True
                print(f"找到開盤時間 {start_time}，指數: {open_index}")
                continue
            
            # 更新最高價和最低價
            if found_start:
                if index_value > high_index:
                    high_index = index_value
                if index_value < low_index:
                    low_index = index_value
                
                # 找到收盤時間
                if end_time in time_value:
                    close_index = index_value
                    print(f"找到收盤時間 {end_time}，指數: {close_index}")
                    print(f"最後成交量: {final_volume}")
                    break
                    
        except Exception as e:
            print(f"處理行 {i} 時出錯: {str(e)}")
            continue
    
    if open_index and close_index:
        print(f"大盤5秒資料處理完成!")
        print(f"開盤: {open_index}")
        print(f"最高: {high_index}")
        print(f"最低: {low_index}")
        print(f"收盤: {close_index}")
        print(f"成交量: {final_volume}")
        
        # 儲存到檔案
        output_dir = "D:/stock/txt"
        import os
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"已創建目錄: {output_dir}")
        
        # 格式化數據，去掉小數點
        open_str = str(int(open_index))
        high_str = str(int(high_index))
        low_str = str(int(low_index))
        close_str = str(int(close_index))
        volume_str = final_volume  # 使用抓取到的成交量
        
        # 使用append模式寫入到1000.txt檔案
        output_file = os.path.join(output_dir, "1000.txt")
        
        # 檢查檔案是否存在且不為空
        file_exists = os.path.exists(output_file) and os.path.getsize(output_file) > 0
        
        with open(output_file, 'a', encoding='utf-8') as f:  # 使用'a'模式進行追加
            # 如果檔案不是空的且最後沒有換行符，則先添加一個換行符
            if file_exists:
                # 讀取檔案最後一個字符
                with open(output_file, 'r', encoding='utf-8') as check_file:
                    check_file.seek(0, os.SEEK_END)
                    pos = check_file.tell() - 1
                    if pos > 0:
                        check_file.seek(pos, os.SEEK_SET)
                        last_char = check_file.read(1)
                        if last_char != '\n':
                            f.write('\n')
            # 寫入新記錄
            taiwan_date = str(int(date_str[:4]) - 1911) + date_str[4:]
            f.write(f'"{taiwan_date}","{open_str}","{high_str}","{low_str}","{close_str}","{volume_str}"')
        
        print(f"已追加大盤指數到文件: {output_file}")
    else:
        print(f"未能找到完整的開盤至收盤資料")

def main():
    """主程式"""
    # 獲取當前日期作為預設值
    today = datetime.now()
    default_date_str = today.strftime('%Y%m%d')
    
    # 顯示日期並讓使用者確認或修改
    print("=" * 50)
    print(f"今天日期為: {today.strftime('%Y-%m-%d')} ({default_date_str})")
    print("請確認要處理的日期，或輸入新的日期 (格式: YYYYMMDD)")
    print("直接按 Enter 使用今天日期")
    user_input = input(f"請輸入日期 [{default_date_str}]: ")
    
    # 處理使用者輸入
    if user_input.strip() == "":
        # 使用者直接按 Enter，使用預設日期
        date_str = default_date_str
        print(f"使用預設日期: {date_str}")
    else:
        # 檢查使用者輸入的日期格式是否正確
        if len(user_input) != 8 or not user_input.isdigit():
            print(f"警告: 輸入的日期格式不正確，應為8位數字(YYYYMMDD)。使用預設日期: {default_date_str}")
            date_str = default_date_str
        else:
            try:
                # 嘗試解析日期以驗證有效性
                input_date = datetime.strptime(user_input, '%Y%m%d')
                date_str = user_input
                print(f"使用日期: {date_str} ({input_date.strftime('%Y-%m-%d')})")
            except ValueError:
                print(f"警告: 無效的日期。使用預設日期: {default_date_str}")
                date_str = default_date_str
    
    print("=" * 50)
    
    # 下載資料並取得資料夾路徑
    print("開始下載資料...")
    date_folder = download(date_str)
    
    # 檢查資料夾是否存在
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if not os.path.exists(date_folder):
        print(f"警告: 日期資料夾不存在: {date_folder}")
        print("將在當前目錄尋找檔案")
        date_folder = current_dir
    
    # 處理上市股票資料
    twse_csv_path = os.path.join(date_folder, f'上市_{date_str}.csv')
    if os.path.exists(twse_csv_path):
        process_stock_data(twse_csv_path, date_str, is_otc=False)
    else:
        print(f"找不到上市公司檔案: {twse_csv_path}")

    # 處理上櫃股票資料
    tpex_csv_path = os.path.join(date_folder, f'櫃買_{date_str}.csv')
    if os.path.exists(tpex_csv_path):
        process_stock_data(tpex_csv_path, date_str, is_otc=True)
    else:
        print(f"找不到上櫃公司檔案: {tpex_csv_path}")

    # 處理大盤資料
    index_csv_path = os.path.join(date_folder, f'大盤5秒_{date_str}.csv')
    if os.path.exists(index_csv_path):
        process_index_5sec_data(index_csv_path, date_str)
    else:
        print(f"找不到大盤5秒檔案: {index_csv_path}")
    
    print("資料處理完成!")


if __name__ == "__main__":
    while True:
        main()
    