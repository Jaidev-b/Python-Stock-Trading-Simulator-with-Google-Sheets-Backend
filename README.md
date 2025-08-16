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
git clone https://github.com/your-username/mock-stock-simulator.git
cd mock-stock-simulator

### 2. Install Dependencies
pip install -r requirements.txt

### 3. Setup Google Sheets API
- Create a **Google Cloud Service Account**.
- Enable **Google Sheets API**.
- Download `credentials.json` and place it in the repo root.
- Share all simulation sheets (BrokerTerminal, Price_Chart, etc.) with the service account email.

### 4. Configure Sheets
- Update `SHEET_MAPPING` in `main.py` with the correct Google Sheet URLs.
- Ensure each participant sheet has:
  - **B2** → Cash balance  
  - **A6:B14** → Holdings table  

### 5. Run the Simulator
python main.py

The script will:
- Authenticate with Google Sheets.
- Start a simulation loop.
- Process trades from the BrokerTerminal.
- Update all sheets every **10 seconds**.

---

## 🏦 Simulation Rules
- **Starting Cash**: ₹1,00,000 per participant.  
- **Minimum Order Value**: ₹7,500.  
- **Trade Pricing**: VWAP of last 3 trades + random fluctuation.  
- **Circuit Limits**: ±20% from last price.  

---

## 📊 Example Flow
1. Buyer and Seller log a trade in **BrokerTerminal**.  
2. Script verifies:
   - Buyer has enough cash  
   - Seller has enough stock  
   - Price is within circuit limits  
3. If valid → Trade executed ✅  
   - Buyer’s holdings and cash updated  
   - Seller’s holdings and cash updated  
   - Price Chart reflects new LTP, VWAP, % change, and volume  

---

## 🛠️ Future Improvements
- Web dashboard for live price display.  
- Order matching engine (instead of manual trade entry).  
- Historical trade analysis & reporting.  
