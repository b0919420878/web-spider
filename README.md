# web-spider

A comprehensive Python-based financial data collection system that automatically downloads and processes Taiwan stock market data and international financial indicators.

##  Overview

This project consists of four main scripts that collect different types of financial data:

1. **International Financial Data** - Global indices, bonds, currencies, and commodities
2. **Taiwan Stock Market Data** - Listed and OTC stocks with 5-second interval market data
3. **Institutional Trading Data** - Foreign investors, investment trusts, and proprietary trading data
4. **Margin Trading Data** - Margin financing and securities lending information

##  Features

###  International Data Collection (`worldindex-today.py`)
- **Bond Yields**: 10-year and 1-year treasury yields
- **Currency Exchange Rates**: USD/TWD, USD/HKD, USD/CNY, USD/KRW, USD/JPY
- **Stock Indices**: Dow Jones, NASDAQ, Hang Seng, Nikkei, Shanghai Composite, KOSPI
- **Commodities**: Gold and Oil futures, Soybean futures
- **Data Source**: Yahoo Finance API
- **Output**: CSV and TXT format files

###  Taiwan Stock Market Data (`上櫃+上市+5秒.py`)
- **Listed Companies**: Taiwan Stock Exchange (TWSE) daily quotes
- **OTC Companies**: Taipei Exchange (TPEx) daily quotes  
- **Market Index**: 5-second interval TAIEX data (9:03:00 - 13:30:00)
- **Data Points**: Open, High, Low, Close, Volume
- **Output**: Individual TXT files for each stock (by stock code)

###  Institutional Trading Data (`上櫃+上市+大盤法人.py`)
- **Foreign Investors**: Buy/sell volumes
- **Investment Trusts**: Buy/sell volumes
- **Proprietary Trading**: Buy/sell volumes
- **Coverage**: Individual stocks and market-wide data
- **Output**: LAW files with institutional trading records

###  Margin Trading Data (`上櫃+上市融資.py`)
- **Margin Financing**: Buy, sell, and outstanding balances
- **Securities Lending**: Buy, sell, and outstanding balances
- **Coverage**: Individual stocks and market aggregates
- **Output**: INV files with margin trading records

##  Requirements

### Python Dependencies
```bash
pip install pandas yfinance requests urllib3
```

### System Requirements
- Python 3.6+
- Stable internet connection
- Sufficient storage space for daily data files

##  Project Structure

```
financial-data-scraper/
├── worldindex-today.py           # International financial data
├── 上櫃+上市+5秒.py                # Taiwan stock market data  
├── 上櫃+上市+大盤法人.py            # Institutional trading data
├── 上櫃+上市融資.py                # Margin trading data
├── YYYYMMDD/                     # Daily data folders
│   ├── worldindex.csv            # International data
│   ├── 上市_YYYYMMDD.csv         # Listed stocks data
│   ├── 櫃買_YYYYMMDD.csv         # OTC stocks data
│   ├── 大盤5秒_YYYYMMDD.csv      # 5-second market data
│   ├── 上市法人_YYYYMMDD.csv     # Listed institutional data
│   ├── 櫃買法人_YYYYMMDD.csv     # OTC institutional data
│   ├── 大盤法人_YYYYMMDD.csv     # Market institutional data
│   ├── 上市融資_YYYYMMDD.csv     # Listed margin data
│   └── 櫃買融資_YYYYMMDD.csv     # OTC margin data
└── D:/stock/                     # Output directory
    ├── txt/                      # Stock price data
    │   ├── 1000.txt             # TAIEX index
    │   ├── 2330.txt             # TSMC
    │   └── [stock_code].txt     # Individual stocks
    ├── law/                      # Institutional data
    │   ├── 1000.law             # Market institutional data
    │   └── [stock_code].law     # Individual stocks
    └── inv/                      # Margin trading data
        ├── 1000.inv             # Market margin data
        └── [stock_code].inv     # Individual stocks
```

##  Usage

### 1. International Financial Data
```bash
python worldindex-today.py
```
- Automatically collects current day's international financial data
- Creates date-named folder with CSV output
- Processes data into TXT format for further analysis

### 2. Taiwan Stock Market Data
```bash
python 上櫃+上市+5秒.py
```
- Interactive date selection (default: current date)
- Downloads data from TWSE and TPEx
- Processes into individual stock files

### 3. Institutional Trading Data
```bash
python 上櫃+上市+大盤法人.py
```
- Collects foreign investor, investment trust, and proprietary trading data
- Supports both listed and OTC markets
- Outputs to LAW format files

### 4. Margin Trading Data
```bash
python 上櫃+上市融資.py
```
- Downloads margin financing and securities lending data
- Includes market-wide aggregates
- Outputs to INV format files

##  Data Format

### Stock Price Data (TXT files)
```
"YYYMMDD","Open","High","Low","Close","Volume"
"1140707","150.50","152.00","149.80","151.25","1000"
```

### Institutional Data (LAW files)
```
"YYYMMDD","Buy","Sell","Type"
"1140707","1000","800","1"    # Type 1: Foreign investors
"1140707","500","300","2"     # Type 2: Investment trusts  
"1140707","200","100","3"     # Type 3: Proprietary trading
```

### Margin Trading Data (INV files)
```
"YYYMMDD","MarginBuy","MarginSell","MarginBalance","ShortBuy","ShortSell","ShortBalance"
"1140707","1000","800","5000","200","150","300"
```

##  Configuration

### Output Directories
- **Stock Data**: `D:/stock/txt/`
- **Institutional Data**: `D:/stock/law/`  
- **Margin Data**: `D:/stock/inv/`

### Date Format
- **Input**: YYYYMMDD (e.g., 20250707)
- **Taiwan Date**: YYYMMDD (ROC calendar, e.g., 1140707)

### International Data Mapping
| Code | Description |
|------|-------------|
| 10db | 10-Year Treasury Yield |
| 01db | 1-Year Treasury Yield |
| 1001 | USD/TWD Exchange Rate |
| 1003 | USD/HKD Exchange Rate |
| 1004 | USD/CNY Exchange Rate |
| 1005 | USD/JPY Exchange Rate |
| 1006 | USD/KRW Exchange Rate |
| 100d | Dow Jones Industrial Average |
| 100n | NASDAQ Composite |
| 100h | Hang Seng Index |
| 100j | Nikkei 225 |
| 100a | Shanghai Composite |
| 100k | KOSPI |
| 100g | Gold Futures |
| 100o | Crude Oil Futures |
| 10sb | Soybean Futures |

##  Important Notes

### Data Sources
- **International Data**: Yahoo Finance (rate-limited)
- **Taiwan Data**: Official exchanges (TWSE, TPEx)

### Rate Limiting
- International data collection includes random delays (1-3 seconds)
- Retry mechanism with exponential backoff
- Respectful API usage to avoid blocking

### Market Hours
- **Taiwan Market**: Monday-Friday, 9:00 AM - 1:30 PM (Taiwan Time)
- **Data Availability**: Usually available after market close

### Error Handling
- Comprehensive error logging
- Graceful handling of missing data
- Automatic retry mechanisms
- Fallback encoding support for CSV files

##  Troubleshooting

### Common Issues

1. **SSL Certificate Errors**
   - SSL verification is disabled for Taiwan exchange APIs
   - Ensure stable internet connection

2. **Encoding Issues**  
   - Scripts support multiple encodings (Big5, UTF-8, CP950)
   - Automatic encoding detection and fallback

3. **Missing Data**
   - Check if market is open (no data on weekends/holidays)
   - Verify internet connection
   - Check error logs in date folders

4. **Rate Limiting**
   - Yahoo Finance may temporarily block requests
   - Wait and retry after some time
   - Consider using VPN if persistent issues

### File Permissions
Ensure write permissions for:
- Current directory (for date folders)
- `D:/stock/` directory and subdirectories

##  Data Analysis

The collected data can be used for:
- Technical analysis and backtesting
- Institutional behavior analysis  
- Margin trading trend analysis
- Cross-market correlation studies
- Automated trading system development

##  Contributing

1. Fork the repository
2. Create a feature branch
3. Add comprehensive error handling
4. Test with different market conditions
5. Submit a pull request

##  License

This project is provided as-is for educational and research purposes. Please ensure compliance with data provider terms of service.

##  Quick Start

1. Install dependencies: `pip install pandas yfinance requests urllib3`
2. Create output directories: `mkdir -p D:/stock/{txt,law,inv}`
3. Run scripts in order:
   ```bash
   python worldindex-today.py
   python 上櫃+上市+5秒.py  
   python 上櫃+上市+大盤法人.py
   python 上櫃+上市融資.py
   ```
4. Check `D:/stock/` for processed data files

---

**Note**: This tool is designed for research and educational purposes. Always verify data accuracy and comply with relevant financial data usage policies.
