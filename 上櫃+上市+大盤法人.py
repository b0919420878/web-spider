import requests
import urllib3
import os
import urllib.parse
from datetime import datetime
import csv
import io

def download(date_str=None):
    """下載法人資料（上市、上櫃、大盤）
    
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

    # 獲取上櫃法人所需的格式化日期
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

    # 1. 下載上櫃法人資料
    print("開始下載上櫃法人資料...")
    otc_save_path = os.path.join(date_folder, f'櫃買法人_{date_str}.csv')
    otc_url = f"https://www.tpex.org.tw/www/zh-tw/insti/dailyTrade?type=Daily&sect=EW&date={encoded_date}&id=&response=csv"
    download_file(otc_url, otc_save_path, "上櫃法人")

    # 2. 下載大盤法人資料
    print("開始下載大盤法人資料...")
    index_save_path = os.path.join(date_folder, f'大盤法人_{date_str}.csv')
    index_url = f"https://www.twse.com.tw/rwd/zh/fund/BFI82U?type=day&dayDate={date_str}&response=csv"
    download_file(index_url, index_save_path, "大盤法人")

    # 3. 下載上市資料
    print("開始下載上市法人資料...")
    twse_save_path = os.path.join(date_folder, f'上市法人_{date_str}.csv')
    twse_url = f"https://www.twse.com.tw/rwd/zh/fund/T86?date={date_str}&selectType=ALLBUT0999&response=csv"
    download_file(twse_url, twse_save_path, "上市法人")

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
    print(f"開始處理{'上櫃法人' if is_otc else '上市法人'}公司資料，檔案: {csv_file_path}")
    
    # 確保目標目錄存在
    target_dir = "D:/stock/law"
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
            if "代號" in line and "名稱" in line:
                header_idx = i
                break
        else:  # 上市資料
            if "代號" in line and "名稱" in line:
                header_idx = i
                break
    
    if header_idx == -1:
        print(f"無法在CSV文件中找到{'上櫃法人' if is_otc else '上市法人'}資料的標題行")
        return
    
    print(f"找到標題行，行號: {header_idx}")
    print(f"標題行內容: {lines[header_idx].strip()}")
    
    # 解析標題行
    header_parts = lines[header_idx].strip().split(',')
    header_parts = [part.strip('"') for part in header_parts]
    
    # 尋找需要的列索引
    try:
        code_idx = -1   # 證券代號/代號
        fbuy_idx = -1   # 外資買進
        fsell_idx = -1   # 外資賣出
        itbuy_idx = -1   # 投信買進
        itsell_idx = -1   # 投信賣出
        prbuy_idx = -1   # 自營商賣出
        prsell_idx = -1   # 自營商賣出
        
        for i, part in enumerate(header_parts):
            # 上市和上櫃的欄位名稱可能不同
            if "證券代號" in part or "代號" in part:
                code_idx = i
            elif "外陸資買進股數" in part or "外資及陸資-買進股數" in part:
                fbuy_idx = i
            elif "外陸資賣出股數" in part or "外資及陸資-賣出股數" in part:
                fsell_idx = i
            elif "投信買進" in part or "投信-買進股數" in part:
                itbuy_idx = i
            elif "投信賣出" in part or "投信-賣出股數" in part:
                itsell_idx = i
            elif "自營商買進股數(自行買賣)" in part or "自營商(自行買賣)-買進股數" in part:
                prbuy_idx = i
            elif "自營商賣出股數(自行買賣)" in part or "自營商(自行買賣)-賣出股數" in part:
                prsell_idx = i

    
        print(f"欄位索引: 代號({code_idx}) 外資買進({fbuy_idx}) 外資賣出({fsell_idx}) 投信買進({itbuy_idx}) 投信賣出({itsell_idx}) 自營商買進({prbuy_idx}) 自營商賣出({prsell_idx}))")
        
        # 確認所有需要的列都找到了
        required_indices = [code_idx, fbuy_idx, fsell_idx,itbuy_idx, itsell_idx,prbuy_idx, prsell_idx]
        if -1 in required_indices:
            missing_columns = []
            if code_idx == -1: missing_columns.append("證券代號/代號")
            if fbuy_idx == -1: missing_columns.append("外資買進")
            if fsell_idx == -1: missing_columns.append("外資賣出")
            if itbuy_idx == -1: missing_columns.append("投信買進")
            if itsell_idx == -1: missing_columns.append("投信賣出")
            if prbuy_idx == -1: missing_columns.append("自營商買進")
            if prsell_idx == -1: missing_columns.append("自營商賣出")
            
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
                fbuy_price = row[fbuy_idx].strip().replace('"', '').replace(',', '')
                fsell_price = row[fsell_idx].strip().replace('"', '').replace(',', '')
                itbuy_price = row[itbuy_idx].strip().replace('"', '').replace(',', '')
                itsell_price = row[itsell_idx].strip().replace('"', '').replace(',', '')
                prbuy_price = row[prbuy_idx].strip().replace('"', '').replace(',', '')
                prsell_price = row[prsell_idx].strip().replace('"', '').replace(',', '')
                
                # 跳過無效數據
                if '--' in [fbuy_price,fsell_price,itbuy_price,itsell_price,prbuy_price,prsell_price]:
                    skipped_count += 1
                    continue
                
                # 轉換成浮點數和整數以驗證（僅驗證，不改變原始字符串）
                try:
                    fbuy_price=round(float(fbuy_price))/1000
                    fsell_price=round(float(fsell_price))/1000
                    itbuy_price=round(float(itbuy_price))/1000
                    itsell_price=round(float(itsell_price))/1000
                    prbuy_price=round(float(prbuy_price))/1000
                    prsell_price=round(float(prsell_price))/1000
                except ValueError:
                    skipped_count += 1
                    continue

                # 準備寫入的數據行
                taiwan_date = str(int(date_str[:4]) - 1911) + date_str[4:]
                fdata_line = f'"{taiwan_date}","{fbuy_price}","{fsell_price}","1"\n'
                itdata_line = f'"{taiwan_date}","{itbuy_price}","{itsell_price}","2"\n'
                prdata_line = f'"{taiwan_date}","{prbuy_price}","{prsell_price}","3"\n'
                
                # 直接寫入對應公司代碼的檔案
                file_path = os.path.join(target_dir, f"{company_code}.law")
                
                # 檢查檔案是否存在，如果存在則附加，否則創建新檔案
                with open(file_path, 'a', encoding='utf-8') as stock_file:
                    if fbuy_price!=0 or fsell_price!=0:
                        stock_file.write(fdata_line)
                    if itbuy_price!=0 or itsell_price!=0:
                        stock_file.write(itdata_line)
                    if prbuy_price!=0 or prsell_price!=0:
                        stock_file.write(prdata_line)
                
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

def process_index_data(index_csv_file_path, date_str):
    """處理大盤法人資料"""
    print(f"開始處理大盤法人資料，檔案: {index_csv_file_path}")
    
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
    
    # 檢查檔案是否有足夠的行數
    if len(lines) < 6:
        print(f"檔案行數不足({len(lines)}行)，可能是放假日或資料格式異常，跳過處理")
        return
    
    try:
        # 讀取需要的行進來
        line2 = lines[2].strip()
        line4 = lines[4].strip()
        line5 = lines[5].strip()
        
        # 檢查行是否為空
        if not line2 or not line4 or not line5:
            print("發現空行，可能是放假日或資料異常，跳過處理")
            return
        
        csv_reader2 = csv.reader(io.StringIO(line2))
        csv_reader4 = csv.reader(io.StringIO(line4))
        csv_reader5 = csv.reader(io.StringIO(line5))
        row2 = next(csv_reader2)
        row4 = next(csv_reader4)
        row5 = next(csv_reader5)
        
        # 檢查每行是否有足夠的欄位
        if len(row2) < 4 or len(row4) < 4 or len(row5) < 4:
            print("資料欄位不足，可能是放假日或資料格式異常，跳過處理")
            return
        
        fbuy_price = row5[2].strip().replace('"', '').replace(',', '')
        fsell_price = row5[3].strip().replace('"', '').replace(',', '')
        itbuy_price = row4[2].strip().replace('"', '').replace(',', '')
        itsell_price = row4[3].strip().replace('"', '').replace(',', '')
        prbuy_price = row2[2].strip().replace('"', '').replace(',', '')
        prsell_price = row2[3].strip().replace('"', '').replace(',', '')
        
        # 檢查是否為有效數字
        try:
            fbuy_price = round(float(fbuy_price)/1000)
            fsell_price = round(float(fsell_price)/1000)
            itbuy_price = round(float(itbuy_price)/1000)
            itsell_price = round(float(itsell_price)/1000)
            prbuy_price = round(float(prbuy_price)/1000)
            prsell_price = round(float(prsell_price)/1000)
        except ValueError as e:
            print(f"數據轉換失敗，可能是放假日或資料格式異常: {e}")
            return

        # 準備寫入的數據行
        taiwan_date = str(int(date_str[:4]) - 1911) + date_str[4:]
        fdata_line = f'"{taiwan_date}","{fbuy_price}","{fsell_price}","1"\n'
        itdata_line = f'"{taiwan_date}","{itbuy_price}","{itsell_price}","2"\n'
        prdata_line = f'"{taiwan_date}","{prbuy_price}","{prsell_price}","3"\n'
                
        # 確保目標目錄存在
        target_dir = "D:/stock/law"
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)
        
        # 直接寫入對應公司代碼的檔案
        file_path = os.path.join(target_dir, "1000.law")
                
        # 檢查檔案是否存在，如果存在則附加，否則創建新檔案
        with open(file_path, 'a', encoding='utf-8') as stock_file:
            if fbuy_price != 0 or fsell_price != 0:
                stock_file.write(fdata_line)
            if itbuy_price != 0 or itsell_price != 0:
                stock_file.write(itdata_line)
            if prbuy_price != 0 or prsell_price != 0:
                stock_file.write(prdata_line)
        
        print("大盤法人資料處理完成")
        
    except IndexError as e:
        print(f"索引錯誤，可能是放假日或資料格式異常: {e}")
        return
    except Exception as e:
        print(f"處理大盤法人資料時發生錯誤: {e}")
        return

    
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
    
    # 處理上市法人資料
    twse_csv_path = os.path.join(date_folder, f'上市法人_{date_str}.csv')
    if os.path.exists(twse_csv_path):
        process_stock_data(twse_csv_path, date_str, is_otc=False)
    else:
        print(f"找不到上市公司檔案: {twse_csv_path}")

    # 處理上櫃法人資料
    tpex_csv_path = os.path.join(date_folder, f'櫃買法人_{date_str}.csv')
    if os.path.exists(tpex_csv_path):
        process_stock_data(tpex_csv_path, date_str, is_otc=True)
    else:
        print(f"找不到上櫃公司檔案: {tpex_csv_path}")

    # 處理大盤法人資料
    index_csv_file_path = os.path.join(date_folder, f'大盤法人_{date_str}.csv')
    if os.path.exists(index_csv_file_path):
        process_index_data(index_csv_file_path, date_str)
    else:
        print(f"找不到上市公司檔案: {index_csv_file_path}")

if __name__ == "__main__":
    while True:
        main()