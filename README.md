# Python-Stock-Trading-Simulator-with-Google-Sheets-Backend
This project is a Python-based mock stock market simulation script that uses Google Sheets as its database and user interface. It creates a dynamic trading environment where users can buy and sell stocks, with prices that fluctuate based on trading activity and simulated market volatility. 
Of course! Here is a well-structured GitHub post for your Python script, formatted as a README.md file. This includes a title, description, features, a detailed setup guide, usage instructions, and the code itself.

Python Stock Trading Simulator with Google Sheets Backend
This project is a Python-based mock stock market simulation script that uses Google Sheets as its database and user interface. It creates a dynamic trading environment where users can buy and sell stocks, with prices that fluctuate based on trading activity and simulated market volatility.

The script reads trade orders from a central "Broker Terminal" sheet, validates them, executes them by updating individual participant sheets, and continuously updates a live "Price Chart" sheet.

Key Features
Google Sheets Integration: Uses Google Sheets as a "database" for storing participant holdings, cash balances, transaction histories, and trade orders. This makes the data easily viewable and manually editable.

Real-Time Trade Processing: The script runs in a continuous loop, scanning for new trade orders and processing them in near real-time.

Batch Processing: To optimize Google Sheets API usage, the script gathers all new trades in a cycle and processes them as a single batch, updating all relevant sheets efficiently.

Robust Trade Validation: Before executing a trade, the script validates:

Sufficient Funds: Checks if the buyer has enough cash.

Sufficient Holdings: Verifies the seller owns the required quantity of stock.

Minimum Order Value: Rejects trades below a configurable threshold (e.g., ₹7,500).

Circuit Limits: Ensures the trade price is within the daily upper (20%) and lower (-20%) circuit limits.

Dynamic Price Simulation:

Prices are influenced by the Volume Weighted Average Price (VWAP) of the last three trades.

A randomized fluctuation is added each cycle to simulate market volatility.

Prices are automatically updated on a central Price_Chart sheet.

Admin Controls: A dedicated Admin_Controls sheet allows a moderator to manually override the price of any stock, which is useful for correcting errors or simulating market-moving news.

Detailed Logging: Provides comprehensive console output for monitoring script activity, including successful trades, errors, and price updates.

How It Works
The simulation is driven by a main loop that orchestrates two key functions:

process_trades():

Scans the BrokerTerminal sheet for new rows where the Process checkbox is marked TRUE.

Pre-fetches the cash and holdings data for all participants involved in the current batch of trades.

For each trade, it performs the full suite of validation checks.

If a trade is valid, it updates the local data for the buyer and seller (subtracting/adding cash and stock).

It then performs a batch update to Google Sheets, writing the new cash/holding values and appending transaction records to the respective participant sheets.

The trade's status is updated to "✅" (success) or "❌" (failure) in the BrokerTerminal.

update_price_chart():

First, it checks the Admin_Controls sheet for any manual price overrides.

If no override is present, it calculates a new price for each stock based on its VWAP and a small random fluctuation.

It calculates the new 20% upper and lower circuit limits based on this new price.

It performs a batch update to the Price_Chart sheet, displaying the new Live Price, Last Traded Price (LTP), Volume, Change %, and circuit limits.

This cycle repeats every 10 seconds, creating a continuously running market.

Setup and Installation
1. Prerequisites
Python 3.x

The following Python libraries. You can install them using pip:

Bash

pip install gspread oauth2client
2. Google Cloud Platform Setup
This script requires a Google Service Account to interact with your Google Sheets.

Create a Google Cloud Project: Go to the Google Cloud Console and create a new project.

Enable APIs: In your new project, go to the "APIs & Services" > "Library" and enable the Google Drive API and the Google Sheets API.

Create a Service Account:

Go to "APIs & Services" > "Credentials".

Click "Create Credentials" and select "Service account".

Give it a name (e.g., "Stock Sim Bot") and grant it the "Editor" role for this project.

Click "Done".

Generate a JSON Key:

In the credentials screen, find the service account you just created and click on it.

Go to the "KEYS" tab, click "ADD KEY", and select "Create new key".

Choose JSON as the key type and click "CREATE". A .json file will be downloaded.

Rename the Key: Rename the downloaded file to credentials.json and place it in the same directory as your Python script.

3. Google Sheets Setup
Create Your Spreadsheets: Create all the necessary Google Sheets. You will need:

A sheet for each participant (e.g., MasterAccount, Silhouette, Quanta).

A BrokerTerminal sheet to place orders.

A Price_Chart sheet to display live prices.

An Admin_Controls sheet for manual overrides.

Share Sheets with the Service Account:

Open your credentials.json file and find the client_email value (it will look something like your-bot-name@your-project.iam.gserviceaccount.com).

For every single spreadsheet you created, click the "Share" button and share it with this email address, giving it "Editor" permissions. This is a critical step.

4. Configure the Script
Update SHEET_MAPPING: In the main.py script, replace the placeholder URLs in the SHEET_MAPPING dictionary with the URLs of your own Google Sheets.

Python

SHEET_MAPPING = {
    "BrokerTerminal": "https://docs.google.com/spreadsheets/d/YOUR_BROKER_URL/...",
    "MasterAccount": "https://docs.google.com/spreadsheets/d/YOUR_MASTER_ACCOUNT_URL/...",
    # ... and so on for all your sheets
}
Set Initial Prices: Modify the INITIAL_COMPANY_PRICES dictionary to include the stocks you want to trade and their starting prices.

Usage
Run the script from your terminal:

Bash

python main.py
Placing a Trade:

Open your BrokerTerminal Google Sheet.

Add a new row with the Buyer, Seller, Company, Qty, and Price.

In the Process column for that row, check the box (or type TRUE).

The script will pick up the order on its next cycle, process it, and update the Status and Reason columns.

Manual Price Override:

Open the Admin_Controls sheet.

Enter the company name and the new desired price.

Check the Apply Override checkbox.

The script will apply this price on the next update_price_chart cycle and uncheck the box.

Monitor: Watch the script's console output to see a log of all actions.
