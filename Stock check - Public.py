# ================================================
#  Charbotte Stock Checker v1.0
#  🌻 Built by Chazou | 🦊 Protected by Bizzou
# ================================================
#
#  A clean and professional tool for Excel-based analysis
#  • Features: News, Insider Trades, Full Financial Stats
#  • Note: This tool is customized to me. Not licensed advice.
#  • A deeper version with alert triggers is available privately.
#
#  🌐 charbotte.com
# ================================================

import yfinance as yf
import pandas as pd
import xlwings as xw
import requests
import time
import sys
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta

os.chdir(os.path.dirname(os.path.abspath(__file__)))

# === Logging Setup ===
os.chdir(os.path.dirname(os.path.abspath(__file__)))
os.makedirs("logs", exist_ok=True)
log_filename = datetime.now().strftime("logs/log_%Y-%m-%d_%H-%M-%S.txt")
sys.stdout = open(log_filename, "w", encoding="utf-8")

# === Constants ===
load_dotenv()
api_key = os.getenv("FINNHUB_API_KEY")
DAYS_BACK_NEWS = 3
DAYS_BACK_TRADES = 365

# === Fetch News from Finnhub ===
def fetch_finnhub_news(ticker_list, api_key, days=DAYS_BACK_NEWS):
    base_url = "https://finnhub.io/api/v1/company-news"
    to_date = datetime.now().date()
    from_date = to_date - timedelta(days=days)
    news_rows = []

    for ticker in ticker_list:
        if ticker.endswith(".TO"):
            print(f"Finnhub - Skipping blocked ticker: {ticker}")
            continue

        try:
            url = f"{base_url}?symbol={ticker}&from={from_date}&to={to_date}&token={api_key}"
            response = requests.get(url)
            data = response.json()

            if isinstance(data, dict) and "error" in data:
                print(f"Bad API response for {ticker}: {data['error']}")
                if "access" in data["error"] or "resource" in data["error"]:
                    continue
                elif "limit" in data["error"]:
                    time.sleep(60)
                    continue

            if isinstance(data, list):
                for item in data:
                    try:
                        news_date = datetime.fromtimestamp(item['datetime'])
                        if news_date.date() >= from_date:
                            news_rows.append([
                                ticker,
                                news_date.strftime('%Y-%m-%d %H:%M'),
                                item.get('headline'),
                                item.get('source'),
                                'finnhub',
                                item.get('url')
                            ])
                    except:
                        pass
            time.sleep(1)

        except Exception as e:
            print(f"News error for {ticker}: {e}")
    return news_rows

# === Fetch News from Yahoo ===
def fetch_yahoo_news(ticker_list, days=DAYS_BACK_NEWS):
    yahoo_rows = []
    cutoff_date = datetime.now() - timedelta(days=days)

    for ticker in ticker_list:
        try:
            stock = yf.Ticker(ticker)
            news = stock.news
            for item in news:
                if 'providerPublishTime' not in item:
                    continue
                pub_date = datetime.fromtimestamp(item['providerPublishTime'])
                if pub_date >= cutoff_date:
                    yahoo_rows.append([
                        ticker,
                        pub_date.strftime('%Y-%m-%d %H:%M'),
                        item.get('title'),
                        item.get('publisher'),
                        'yahoo',
                        item.get('link')
                    ])
        except:
            pass
    return yahoo_rows

# === Fetch Insider Trades ===
def fetch_insider_trades(ticker_list, api_key, days=DAYS_BACK_TRADES):
    base_url = "https://finnhub.io/api/v1/stock/insider-transactions"
    to_date = datetime.now().date()
    from_date = to_date - timedelta(days=days)
    trades = []

    for ticker in ticker_list:
        if ticker.endswith(".TO"):
            continue
        try:
            url = f"{base_url}?symbol={ticker}&from={from_date}&token={api_key}"
            response = requests.get(url)
            data = response.json()
            if isinstance(data, dict) and "data" in data:
                for item in data["data"]:
                    trades.append([
                        ticker,
                        item.get("transactionDate", ""),
                        item.get("name"),
                        item.get("transactionCode"),
                        item.get("share"),
                        item.get("transactionPrice")
                    ])
        except:
            pass
    return trades

# === Excel Integration ===
try:
    wb = xw.Book.caller()
except:
    wb = xw.Book("Stock check - Public.xlsm")

file_path = wb.fullname
print("🐍 Connected to:", file_path)

ws = wb.sheets["Watchlist check"]

df_check = ws.range("A1").expand("table").options(pd.DataFrame, index=False).value

ticker_list = df_check['Ticker'].dropna().tolist()
errored = []

# === Collect Stock Stats ===
stock_data = []
for ticker in ticker_list:
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period="2y")
        info = stock.info

        if hist.empty:
            continue

        # Price, Volume
        current_price = hist['Close'].iloc[-1]
        price_at_open = hist['Open'].iloc[-1]
        percent_change = ((current_price - price_at_open) / price_at_open) * 100
        high_52w = hist['High'].rolling(252).max().iloc[-1]
        low_52w = hist['Low'].rolling(252).min().iloc[-1]
        vol_today = hist['Volume'].iloc[-1]
        vol_avg_30 = hist['Volume'].rolling(30).mean().iloc[-1]

        # RSI
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0).rolling(14).mean()
        loss = -delta.clip(upper=0).rolling(14).mean()
        rs = gain / loss
        rsi = 100 - (100 / (1 + rs)).iloc[-1]

        # MA/EMA
        sma50 = hist['Close'].rolling(50).mean().iloc[-1]
        sma200 = hist['Close'].rolling(200).mean().iloc[-1]
        sma_crossover = 'Yes' if sma50 > sma200 else 'No'
        ema9 = hist['Close'].ewm(span=9).mean().iloc[-1]
        ema21 = hist['Close'].ewm(span=21).mean().iloc[-1]
        ema_crossover = 'Yes' if ema9 > ema21 else 'No'

        # Financials
        stock_data.append([
            ticker, current_price, price_at_open, percent_change, high_52w, low_52w,
            vol_today, vol_avg_30, rsi, sma50, sma200, sma_crossover,
            ema9, ema21, ema_crossover,
            info.get("trailingPE"), info.get("priceToSalesTrailing12Months"),
            info.get("marketCap"), info.get("totalRevenue"), info.get("revenueGrowth"),
            info.get("grossMargins"), info.get("profitMargins"), info.get("debtToEquity"),
            info.get("earningsQuarterlyGrowth"), info.get("returnOnEquity"), info.get("returnOnAssets")
        ])

    except:
        errored.append(ticker)

columns = [
    "Ticker", "Current Price", "Price at open", "% Change", "52W High", "52W Low",
    "Vol today", "Vol Avg (30)", "RSI", "SMA50", "SMA200", "SMA crossover",
    "EMA9", "EMA21", "EMA crossover", "P/E", "P/S", "Market Cap",
    "Revenue", "Rev (YoY)", "Gross Margin", "Profit Margin", "Debt-to-Eq",
    "EPS Growth", "ROE", "ROA"
]

# === Write to Excel ===
df_out = pd.DataFrame(stock_data, columns=columns).set_index("Ticker")
df_check.set_index("Ticker", inplace=True)

# Add missing columns from df_out to df_check
for col in df_out.columns:
    if col not in df_check.columns:
        df_check[col] = None

# Update values from df_out into df_check (no column slice)
for ticker in df_out.index:
    for col in df_out.columns:
        if ticker in df_check.index:
            df_check.at[ticker, col] = df_out.at[ticker, col]



for ticker in df_check.index:
    if ticker in df_out.index:
        for col in df_out.columns:
            df_check.loc[ticker, col] = df_out.loc[ticker, col]

ws.range("A1").value = [df_check.reset_index().columns.tolist()] + df_check.reset_index().values.tolist()

# === Clear Old News/Insiders ===
wb.sheets["News"].range("A2:F10000").clear_contents()
wb.sheets["Insider trade"].range("A2:F10000").clear_contents()

# === Insert News ===
combined_news = fetch_finnhub_news(ticker_list, api_key) + fetch_yahoo_news(ticker_list)
if combined_news:
    df_news = pd.DataFrame(combined_news, columns=["Ticker", "Date", "Headline", "Source", "API", "Link"])
    df_news["Link"] = df_news["Link"].apply(lambda x: f'=HYPERLINK("{x}", "{x}")')
    wb.sheets["News"].range("A1").value = [df_news.columns.tolist()] + df_news.values.tolist()

# === Insert Insider Trades ===
insider_data = fetch_insider_trades(ticker_list, api_key)
if insider_data:
    df_insiders = pd.DataFrame(insider_data, columns=["Ticker", "Date", "Name", "Transaction Type", "Shares", "Price"])
    wb.sheets["Insider trade"].range("A1").value = [df_insiders.columns.tolist()] + df_insiders.values.tolist()

ws_news = wb.sheets["News"]
ws_insider = wb.sheets["Insider trade"]

# === Write News to Excel ===
if not df_news.empty:
    ws_news.range("A1").value = [df_news.columns.tolist()] + df_news.values.tolist()
else:
    print("📰 No news data to write.")

# === Write Insider Trades to Excel ===
if not df_insiders.empty:
    ws_insider.range("A1").value = [df_insiders.columns.tolist()] + df_insiders.values.tolist()
else:
    print("🕵️ No insider trade data to write.")


# === Save and Close ===
#wb.save()
#wb.close()

if errored:
    print("\nTickers with errors:", errored)

sys.stdout.close()

if errored:
    print("\nTickers with errors:", errored)