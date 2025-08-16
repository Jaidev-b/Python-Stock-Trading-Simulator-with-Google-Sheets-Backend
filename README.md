# 📈 Mock Stock Simulator

A Python-based stock market simulation that runs entirely on **Google Sheets**.  
Participants trade with each other through a **Broker Terminal**, with trades verified, executed, and updated in real time.  
Prices fluctuate using **VWAP (Volume Weighted Average Price)**, random market movements, and **circuit breakers** — mimicking real stock market behavior.

---

## 🚀 Features
- **Broker Terminal**: Centralized order book where trades are logged.
- **VWAP Pricing**: Prices adjust dynamically based on the last 3 trades.
- **Manual Overrides**: Admins can set prices manually via Google Sheets.
- **Circuit Breakers**: Automatic 20% upper and lower limits on stock prices.
- **Auto-Updates**: Live price chart updates every cycle with conditional formatting.
- **Participant Portfolios**: Each team/participant gets a sheet tracking cash, holdings, and trade history.

---

## 📂 Repository Structure
- `main.py` → Core simulation script  
- `credentials.json.example` → Template for Google Service Account credentials  
- `requirements.txt` → Required Python packages  
- `README.md` → Documentation  

---

## ⚙️ Setup

### 1. Clone the Repository
```bash
git clone https://github.com/your-username/mock-stock-simulator.git
cd mock-stock-simulator
