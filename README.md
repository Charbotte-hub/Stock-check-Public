# 📈 Charbotte Stock Checker – Public Version

A clean, Excel-powered Python tool for stock analysis.  
Built by Chazou 💛 Protected by Bizzou 🦊

---

## ✅ Features

- Pulls the latest **news** (last 3 days) (FMP + Yahoo)
- Fetches **insider trades** (last365 days) (U.S. tickers only)
- Loads full **financial statistics** for each stock
- Writes data directly to Excel with VBA buttons
- All activity is logged to a local `logs/` folder

---

## 📦 Requirements

Ensure you have **Python 3.9+** installed.  
Install the following libraries:

```bash
pip install requests openpyxl xlwings python-dotenv
```

Optional:
```bash
pip install pandas
```

---

## 🔐 API Key Setup

1. Sign up for a **free API key** at [finnhub.io](https://finnhub.io/)
2. In your project folder, create a file named `.env`
3. Add the following line:

```env
FINNHUB_API_KEY=your_finnhub_key_here
```

✅ Your `.env` file is ignored by Git and stays private.

---

## 🚀 How to Run

Open the provided Excel file:

- In **watchlist** tab, work on your tickers, add, analysis, keep, remove etc
- **Button 1: End of Session** Save all your note, analysis, grade into history columns and extract tickers that action not "delete" to **watchlist check**
- **Button 2:** Find the path to Python script
- **Button 3:** Clear old data and run Python script

All results are displayed in:
- `News` tab
- `Insider trades` tab
- `Watchlist check` tab

---

## 🧠 Notes

- ⚠️ **Canadian insider trade data is not available** (no free API)
- Yahoo sometimes returns partial or missing news — fallback logic handles this
- The public version removes debugging prints for performance
- API errors and skipped tickers are logged for traceability
- `.env`, `.pyc`, and log files are ignored via `.gitignore`

---

## 🌐 License

MIT License.  
Use freely with credit to **Chazou & Bizzou** ✨

---

## 🧩 Future Features (Private Builds Only)

- Sentiment scoring
- Alert triggers
- Earnings predictions
- Multi-source AI-powered reports

---

## 📫 Contact

✨ charbotte.com – where Botte&Bizzou bloom 🌻
