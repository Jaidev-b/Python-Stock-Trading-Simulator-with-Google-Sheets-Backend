import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import time
import random
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# === CONFIGURATION ===
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

try:
    CREDS = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", SCOPE)
    client = gspread.authorize(CREDS)
    logging.info("Successfully authenticated with Google Sheets API.")
except Exception as e:
    logging.error(
        f"Authentication Error: Could not load credentials.json or authorize gspread. Please ensure 'credentials.json' is correct and accessible. Error: {e}")
    exit()

SHEET_MAPPING = {
    "BrokerTerminal": "https://docs.google.com/spreadsheets/d/your_file_link",
    "MasterAccount": "https://docs.google.com/spreadsheets/d/your_file_link",
    "Silhouette": "https://docs.google.com/spreadsheets/d/your_file_link",
}

MIN_ORDER_VALUE = 7500  # ₹7,500

# All dictionary keys are now uppercase for consistency
INITIAL_COMPANY_PRICES = {
    "RELIANCE": 1500.00,
    "TATA STEEL": 160.00,
    "BPCL": 330.00,
    "BAJAJ FINANCE": 940.00,
    "IRFC": 140.00,
    "VI": 7.50,
    "GTL INFRA": 1.80,
    "ULTRATECH": 12100.00,
}

# Dictionary to store previous prices for Change % calculation: {company: previous_price}
PREVIOUS_PRICES = {company: price for company, price in INITIAL_COMPANY_PRICES.items()}

# Dictionary to store actual LTP from trades: {company: last_traded_price}
LAST_TRADED_PRICES = {company: price for company, price in
                      INITIAL_COMPANY_PRICES.items()}

# Dictionary to store current VWAP price for each company
CURRENT_VWAP_PRICES = {company: price for company, price in
                       INITIAL_COMPANY_PRICES.items()}

# Dictionary to store recent trades for VWAP calculation: {company: [(price, quantity), (price, quantity), ...]}
RECENT_TRADES_HISTORY = {company: [] for company in INITIAL_COMPANY_PRICES.keys()}

# Dictionary to store current volume for each company
COMPANY_VOLUME = {company: 0 for company in INITIAL_COMPANY_PRICES.keys()}

# --- Global variable to store current circuit limits (fetched from price chart) ---
CURRENT_CIRCUITS = {company: {"upper": 0.0, "lower": 0.0} for company in INITIAL_COMPANY_PRICES.keys()}

# === SHEET HELPER FUNCTIONS ===

# Global cache for worksheets to avoid opening them repeatedly in a single cycle
_WORKSHEET_CACHE = {}


def get_worksheet(sheet_name_or_url):
    """
    Opens a Google Sheet by its URL and returns the first worksheet (sheet1).
    This function caches worksheets to avoid reopening them within the same run.
    """
    if sheet_name_or_url not in _WORKSHEET_CACHE:
        try:
            if sheet_name_or_url.startswith("http"):
                spreadsheet = client.open_by_url(sheet_name_or_url)
            else:
                spreadsheet = client.open(sheet_name_or_url)

            _WORKSHEET_CACHE[sheet_name_or_url] = spreadsheet.sheet1
            logging.debug(
                f"DEBUG: Successfully opened and cached new spreadsheet '{spreadsheet.title}' from URL/name: {sheet_name_or_url} (API call).")
        except gspread.exceptions.SpreadsheetNotFound:
            logging.error(f"Error: Spreadsheet not found at URL/Name: {sheet_name_or_url}")
            raise
        except Exception as e:
            logging.error(f"Error opening sheet at URL/Name {sheet_name_or_url}: {e}")
            raise
    else:
        logging.debug(
            f"DEBUG: Retrieving spreadsheet '{_WORKSHEET_CACHE[sheet_name_or_url].title}' from cache (no API call).")
    return _WORKSHEET_CACHE[sheet_name_or_url]


def get_cash_balance(ws):
    """
    Retrieves the cash balance from a specific cell (B2) of the given worksheet.
    """
    try:
        cash_str = ws.acell("B2").value
        if cash_str is None or cash_str.strip() == "":
            logging.warning(f"Cash balance cell B2 is empty for worksheet '{ws.title}'. Returning 0.0.")
            return 0.0
        return float(cash_str)
    except (ValueError, TypeError) as e:
        logging.warning(
            f"Could not convert cash balance in B2 ('{cash_str}') to float for worksheet '{ws.title}'. Returning 0.0. Error: {e}")
        return 0.0
    except Exception as e:
        logging.error(f"Error retrieving cash balance from worksheet '{ws.title}': {e}")
        return 0.0


def get_holdings(ws):
    """
    Retrieves company holdings from a specified range (A6:B14) of the worksheet.
    It parses the data into a dictionary: {company_name: quantity}.
    """
    holdings = {}
    try:
        data = ws.get("A6:B14")  # CHECK THIS RANGE IN YOUR ACTUAL SHEETS
    except Exception as e:
        logging.error(f"Error fetching holdings data from range A6:B14 for worksheet '{ws.title}': {e}")
        return {}
    for row in data:
        if len(row) < 2 or row[0].strip().lower() == "company":
            continue
        try:
            company_name = row[0].strip().upper()
            quantity_str = row[1].strip() if len(row) > 1 else "0"
            quantity = int(quantity_str) if quantity_str else 0
            holdings[company_name] = quantity
        except (ValueError, IndexError) as e:
            logging.warning(
                f"Could not parse holding row '{row}' for worksheet '{ws.title}'. Setting quantity to 0. Error: {e}")
            holdings[row[0].strip().upper()] = 0
    return holdings


def prepare_transaction_row(action, company, qty, price, total):
    """
    Prepares a transaction record as a new row for later batch appending.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return [timestamp, action, company, qty, price, total]


# === PRICE CHART AND ADMIN CONTROL LOGIC ===

def calculate_vwap(company):
    """
    Calculates the Volume Weighted Average Price (VWAP) for a given company
    based on the last 3 trades stored in RECENT_TRADES_HISTORY.
    """
    trades = RECENT_TRADES_HISTORY.get(company, [])
    if not trades:
        return None

    total_price_volume = 0
    total_volume = 0
    for price, qty in trades:
        total_price_volume += price * qty
        total_volume += qty

    return total_price_volume / total_volume if total_volume > 0 else None


def get_manual_overrides(admin_ws):
    """
    Reads manual override requests from the Admin_Controls sheet.
    Returns a dictionary: {company_name: override_price} for active overrides.
    """
    overrides = {}
    try:
        admin_data = admin_ws.get("A4:C11")
        updates_to_clear_checkboxes = []

        for i, row in enumerate(admin_data):
            if len(row) >= 3:
                company = row[0].strip().upper()
                override_price_str = row[1].strip()
                apply_override_checkbox = row[2].strip()

                if apply_override_checkbox.upper() == 'TRUE' and company and override_price_str:
                    try:
                        override_price = float(override_price_str)
                        overrides[company] = override_price
                        updates_to_clear_checkboxes.append({'range': f'C{i + 4}', 'values': [['FALSE']]})
                    except ValueError:
                        logging.warning(f"Invalid override price '{override_price_str}' for company '{company}'.")

        if updates_to_clear_checkboxes:
            admin_ws.batch_update(updates_to_clear_checkboxes)
            logging.info(f"Cleared {len(updates_to_clear_checkboxes)} manual override checkboxes.")
    except Exception as e:
        logging.error(f"Error reading manual overrides from Admin_Controls sheet: {e}")
    return overrides


def update_price_chart():
    """
    Updates the 'Price_Chart' sheet with current prices, LTP, volume, change, and circuits.
    Handles manual overrides.
    Populates CURRENT_VWAP_PRICES and CURRENT_CIRCUITS global variables.
    """
    logging.info("\n--- Updating Price Chart ---")
    try:
        price_chart_ws = get_worksheet(SHEET_MAPPING["Price_Chart"])
        admin_ws = get_worksheet(SHEET_MAPPING["Admin_Controls"])

        price_data_from_sheet = price_chart_ws.get_all_values()
        manual_overrides = get_manual_overrides(admin_ws)

        updates = []
        for i, row in enumerate(price_data_from_sheet):
            if i == 0:
                continue

            company_name = row[0].strip().upper()
            if not company_name:
                break

            if company_name not in INITIAL_COMPANY_PRICES:
                logging.warning(
                    f"Company '{company_name}' from Price_Chart not found in INITIAL_COMPANY_PRICES. Skipping price update for this company.")
                continue

            # Robustly get previous_price_for_change
            previous_price_for_change = INITIAL_COMPANY_PRICES.get(company_name, 0.0)
            try:
                if len(row) > 1 and row[1].strip():
                    sheet_live_price_str = row[1].strip().replace('%', '')
                    sheet_price = float(sheet_live_price_str)
                    if sheet_price > 0:
                        previous_price_for_change = sheet_price
                    else:
                        logging.warning(
                            f"Previous price for {company_name} from sheet ('{row[1]}') is zero or negative. Using initial price.")
                else:
                    logging.warning(f"Previous price cell for {company_name} is empty. Using initial price.")
            except (ValueError, TypeError) as e:
                logging.warning(
                    f"Could not parse previous price for {company_name} from sheet ('{row[1]}'). Using initial price. Error: {e}")

            PREVIOUS_PRICES[company_name] = previous_price_for_change

            current_ltp = LAST_TRADED_PRICES.get(company_name, INITIAL_COMPANY_PRICES.get(company_name, 0.0))
            current_volume = COMPANY_VOLUME.get(company_name, 0)
            current_vwap = CURRENT_VWAP_PRICES.get(company_name, INITIAL_COMPANY_PRICES.get(company_name, 0.0))

            new_current_price = current_vwap
            new_ltp_display = current_ltp

            if company_name in manual_overrides:
                new_current_price = manual_overrides[company_name]
                new_ltp_display = manual_overrides[company_name]
                logging.info(f"  Manual override applied for {company_name}: setting price to {new_current_price:.2f}")

                # === FIX: Clear the recent trade history to make the override the new baseline price. ===
                RECENT_TRADES_HISTORY[company_name].clear()
                logging.info(f"  Trade history for {company_name} cleared due to manual override.")

            else:
                base_price = CURRENT_VWAP_PRICES.get(company_name, INITIAL_COMPANY_PRICES.get(company_name, 0.0))
                vwap_price_from_trades = calculate_vwap(company_name)

                if vwap_price_from_trades is not None:
                    base_price = vwap_price_from_trades

                percentage_fluctuation = base_price * (random.uniform(-0.015, 0.015))
                fixed_fluctuation = random.uniform(-0.02, 0.02)
                fluctuation = percentage_fluctuation + fixed_fluctuation
                new_current_price = base_price + fluctuation

                if new_current_price < 0:
                    new_current_price = 0.01

            upper_circuit = new_current_price * 1.20
            lower_circuit = new_current_price * 0.80

            if new_current_price > upper_circuit:
                new_current_price = upper_circuit
                logging.info(f"  {company_name} hit Upper Circuit at {upper_circuit:.2f}")
            elif new_current_price < lower_circuit:
                new_current_price = lower_circuit
                logging.info(f"  {company_name} hit Lower Circuit at {lower_circuit:.2f}")

            CURRENT_VWAP_PRICES[company_name] = new_current_price
            CURRENT_CIRCUITS[company_name]["upper"] = upper_circuit
            CURRENT_CIRCUITS[company_name]["lower"] = lower_circuit

            change_percent = 0.0
            if PREVIOUS_PRICES.get(company_name, 0) != 0:
                change_percent = (new_current_price - PREVIOUS_PRICES[company_name]) / PREVIOUS_PRICES[
                    company_name]

            updates.append({
                'range': f'B{i + 1}:G{i + 1}',
                'values': [[
                    round(new_current_price, 2),
                    round(new_ltp_display, 2),
                    current_volume,
                    change_percent,
                    round(upper_circuit, 2),
                    round(lower_circuit, 2)
                ]]
            })

        if updates:
            price_chart_ws.batch_update(updates)
            logging.info("Price chart updated successfully.")
        else:
            logging.info("No companies found to update in Price Chart.")

    except Exception as e:
        logging.error(f"Error updating price chart: '{e}'")


def apply_price_chart_conditional_formatting():
    """
    Applies conditional formatting rules to the 'Price_Chart' sheet.
    This function should be called once on script startup as rules persist on the sheet.
    """
    logging.info("Applying conditional formatting to Price Chart sheet.")
    try:
        price_chart_ws = get_worksheet(SHEET_MAPPING["Price_Chart"])
        sheet_id = price_chart_ws.worksheet_id

        requests = []

        range_for_columns = {
            "sheetId": sheet_id,
            "startRowIndex": 1,
            "endRowIndex": 1000
        }

        light_green_fill = {"red": 220 / 255, "green": 255 / 255, "blue": 220 / 255}
        dark_green_text = {"red": 33 / 255, "green": 136 / 255, "blue": 56 / 255}
        light_red_fill = {"red": 255 / 255, "green": 220 / 255, "blue": 220 / 255}
        dark_red_text = {"red": 204 / 255, "green": 51 / 255, "blue": 51 / 255}

        rule_live_price_up = {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{**range_for_columns, "startColumnIndex": 1, "endColumnIndex": 2}],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": "=$E2>0"}]},
                        "format": {"backgroundColor": light_green_fill, "foregroundColor": dark_green_text}
                    }
                }, "index": 0
            }
        }
        requests.append(rule_live_price_up)

        rule_live_price_down = {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{**range_for_columns, "startColumnIndex": 1, "endColumnIndex": 2}],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": "=$E2<0"}]},
                        "format": {"backgroundColor": light_red_fill, "foregroundColor": dark_red_text}
                    }
                }, "index": 1
            }
        }
        requests.append(rule_live_price_down)

        rule_change_percent_up = {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{**range_for_columns, "startColumnIndex": 4, "endColumnIndex": 5}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER_THAN", "values": [{"userEnteredValue": "0"}]},
                        "format": {"backgroundColor": light_green_fill, "foregroundColor": dark_green_text}
                    }
                }, "index": 2
            }
        }
        requests.append(rule_change_percent_up)

        rule_change_percent_down = {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{**range_for_columns, "startColumnIndex": 4, "endColumnIndex": 5}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_LESS_THAN", "values": [{"userEnteredValue": "0"}]},
                        "format": {"backgroundColor": light_red_fill, "foregroundColor": dark_red_text}
                    }
                }, "index": 3
            }
        }
        requests.append(rule_change_percent_down)

        if requests:
            price_chart_ws.batch_update({'requests': requests})
            logging.info("Conditional formatting rules applied successfully.")
        else:
            logging.warning("No conditional formatting rules to apply.")

    except gspread.exceptions.SpreadsheetNotFound:
        logging.error(f"Price Chart spreadsheet not found for conditional formatting.")
    except Exception as e:
        logging.error(f"Error applying conditional formatting: {e}")


# === MAIN TRADE PROCESSING LOGIC ===

def process_trades():
    """
    This is the core function that orchestrates the trade simulation.
    """
    logging.info(f"\n--- Starting trade processing cycle at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")

    _WORKSHEET_CACHE.clear()

    try:
        broker_ws = get_worksheet(SHEET_MAPPING["BrokerTerminal"])
        orders = broker_ws.get_all_values()[1:]
    except Exception as e:
        logging.critical(f"Could not access BrokerTerminal sheet. Skipping this cycle. Error: {e}")
        return False

    new_transactions_to_process = []

    for i, row in enumerate(orders, start=2):
        if len(row) < 10:
            if not any(cell.strip() for cell in row): continue
            logging.warning(
                f"  Incomplete row data at row {i}. Expected at least 10 columns, got {len(row)}. Skipping.")
            continue

        status = row[7].strip()
        process_checkbox = row[9].strip().upper()

        if status or process_checkbox != 'TRUE':
            continue

        new_transactions_to_process.append((i, row))

    if not new_transactions_to_process:
        logging.info("No new fresh transactions marked for processing in this cycle.")
        logging.info("--- Trade processing cycle completed ---")
        return False

    logging.info(f"Found {len(new_transactions_to_process)} new transactions marked for processing as a batch.")

    found_any_processed_in_batch = False
    updates_to_broker_terminal = []

    participant_data_for_batch_update = {}

    all_participants_in_batch = set()
    for _, row_data in new_transactions_to_process:
        _, buyer, seller, _, _, _, _, _, _, _ = row_data
        all_participants_in_batch.add(buyer.strip())
        all_participants_in_batch.add(seller.strip())

    for participant_name in all_participants_in_batch:
        if participant_name not in SHEET_MAPPING:
            logging.error(f"Invalid participant '{participant_name}' encountered. Skipping data pre-fetch.")
            continue

        try:
            participant_ws = get_worksheet(SHEET_MAPPING[participant_name])
            participant_data_for_batch_update[participant_name] = {
                'cash': get_cash_balance(participant_ws),
                'holdings': get_holdings(participant_ws),
                'transactions_to_append_data': [],
                'holding_cell_ranges': {}
            }
            holdings_data_for_ranges = participant_ws.get("A6:B14")
            for r_idx, h_row in enumerate(holdings_data_for_ranges):
                if len(h_row) >= 1 and h_row[0].strip():
                    company_name = h_row[0].strip().upper()
                    participant_data_for_batch_update[participant_name]['holding_cell_ranges'][
                        company_name] = f'B{r_idx + 6}'

        except Exception as e:
            logging.error(f"Error pre-fetching data for participant '{participant_name}': {e}")
            participant_data_for_batch_update[participant_name] = None

    for original_row_index, row_data in new_transactions_to_process:
        order_id, buyer, seller, company_raw, qty_str, price_str, total_str, _, _, _ = row_data
        company = company_raw.strip().upper()

        logging.info(f"Processing order in row {original_row_index} (Order ID: {order_id})...")

        try:
            qty = int(qty_str)
            price = float(price_str)
            calculated_total = round(qty * price, 2)
            if not total_str or abs(float(total_str) - calculated_total) > 0.01:
                updates_to_broker_terminal.append({'range': f'G{original_row_index}', 'values': [[calculated_total]]})
            total = calculated_total
        except (ValueError, TypeError) as e:
            updates_to_broker_terminal.append({'range': f'H{original_row_index}', 'values': [["❌"]]})
            updates_to_broker_terminal.append(
                {'range': f'I{original_row_index}', 'values': [[f"Invalid qty or price: {e}"]]})
            updates_to_broker_terminal.append({'range': f'J{original_row_index}', 'values': [['FALSE']]})
            logging.error(
                f"  Error in row {original_row_index}: Invalid quantity ('{qty_str}') or price ('{price_str}'). Error: {e}")
            continue

        if total < MIN_ORDER_VALUE:
            updates_to_broker_terminal.append({'range': f'H{original_row_index}', 'values': [["❌"]]})
            updates_to_broker_terminal.append(
                {'range': f'I{original_row_index}', 'values': [[f"Order < ₹{MIN_ORDER_VALUE} (Total: {total:.2f})"]]})
            updates_to_broker_terminal.append({'range': f'J{original_row_index}', 'values': [['FALSE']]})
            logging.warning(
                f"  Error in row {original_row_index}: Order value ({total:.2f}) is below minimum allowed ({MIN_ORDER_VALUE}).")
            continue

        if participant_data_for_batch_update.get(buyer) is None or participant_data_for_batch_update.get(
                seller) is None:
            updates_to_broker_terminal.append({'range': f'H{original_row_index}', 'values': [["❌"]]})
            updates_to_broker_terminal.append(
                {'range': f'I{original_row_index}', 'values': [[f"Participant sheet access error."]]})
            updates_to_broker_terminal.append({'range': f'J{original_row_index}', 'values': [['FALSE']]})
            logging.error(
                f"  Error in row {original_row_index}: Could not access cached data for buyer '{buyer}' or seller '{seller}'.")
            continue

        company_circuits = CURRENT_CIRCUITS.get(company)
        if company_circuits and company_circuits["upper"] != 0.0 and company_circuits["lower"] != 0.0:
            upper_limit = company_circuits["upper"]
            lower_limit = company_circuits["lower"]
            if not (lower_limit <= price <= upper_limit):
                updates_to_broker_terminal.append({'range': f'H{original_row_index}', 'values': [["❌"]]})
                updates_to_broker_terminal.append({'range': f'I{original_row_index}', 'values': [
                    [f"Price ₹{price:.2f} outside circuit ({lower_limit:.2f}-{upper_limit:.2f})"]]})
                updates_to_broker_terminal.append({'range': f'J{original_row_index}', 'values': [['FALSE']]})
                logging.warning(
                    f"  Error in row {original_row_index}: Order price '{price:.2f}' for '{company}' is outside circuit limits.")
                continue

        buyer_current_cash = participant_data_for_batch_update[buyer]['cash']
        seller_current_holdings = participant_data_for_batch_update[seller]['holdings']

        if buyer_current_cash < total:
            updates_to_broker_terminal.append({'range': f'H{original_row_index}', 'values': [["❌"]]})
            updates_to_broker_terminal.append({'range': f'I{original_row_index}', 'values': [
                [f"Insufficient cash (Buyer has {buyer_current_cash:.2f}, needs {total:.2f})"]]})
            updates_to_broker_terminal.append({'range': f'J{original_row_index}', 'values': [['FALSE']]})
            logging.warning(f"  Error in row {original_row_index}: Buyer '{buyer}' has insufficient cash.")
            continue

        if company not in seller_current_holdings or seller_current_holdings[company] < qty:
            updates_to_broker_terminal.append({'range': f'H{original_row_index}', 'values': [["❌"]]})
            updates_to_broker_terminal.append({'range': f'I{original_row_index}', 'values': [
                [f"Insufficient stock (Seller has {seller_current_holdings.get(company, 0)}, needs {qty})"]]})
            updates_to_broker_terminal.append({'range': f'J{original_row_index}', 'values': [['FALSE']]})
            logging.warning(
                f"  Error in row {original_row_index}: Seller '{seller}' has insufficient stock of '{company}'.")
            continue

        # === Execute Trade (Update local cache first) ===
        try:
            participant_data_for_batch_update[buyer]['cash'] -= total
            participant_data_for_batch_update[seller]['cash'] += total

            participant_data_for_batch_update[buyer]['holdings'][company] = participant_data_for_batch_update[buyer][
                                                                                'holdings'].get(company, 0) + qty
            participant_data_for_batch_update[seller]['holdings'][company] -= qty

            participant_data_for_batch_update[buyer]['transactions_to_append_data'].append(
                prepare_transaction_row("BUY", company, qty, price, total)
            )
            participant_data_for_batch_update[seller]['transactions_to_append_data'].append(
                prepare_transaction_row("SELL", company, qty, price, total)
            )

            RECENT_TRADES_HISTORY[company].append((price, qty))
            if len(RECENT_TRADES_HISTORY[company]) > 3:
                RECENT_TRADES_HISTORY[company].pop(0)
            COMPANY_VOLUME[company] += qty
            LAST_TRADED_PRICES[company] = price

            updates_to_broker_terminal.append({'range': f'H{original_row_index}', 'values': [["✅"]]})
            updates_to_broker_terminal.append({'range': f'I{original_row_index}', 'values': [["Trade completed"]]})
            updates_to_broker_terminal.append({'range': f'J{original_row_index}', 'values': [['FALSE']]})
            logging.info(f"  Trade {order_id} completed successfully.")
            found_any_processed_in_batch = True

        except Exception as e:
            updates_to_broker_terminal.append({'range': f'H{original_row_index}', 'values': [["❌"]]})
            updates_to_broker_terminal.append(
                {'range': f'I{original_row_index}', 'values': [[f"Trade execution failed: {e}"]]})
            updates_to_broker_terminal.append({'range': f'J{original_row_index}', 'values': [['FALSE']]})
            logging.error(f"  Error executing trade {order_id}: {e}")
            continue

    if updates_to_broker_terminal:
        try:
            broker_ws.batch_update(updates_to_broker_terminal)
            logging.info(f"Updated {len(updates_to_broker_terminal)} cells in BrokerTerminal.")
        except Exception as e:
            logging.error(f"Failed to perform batch update on BrokerTerminal: {e}")

    for participant_name, data in participant_data_for_batch_update.items():
        if data is None:
            continue

        try:
            participant_ws = get_worksheet(SHEET_MAPPING[participant_name])
            participant_sheet_updates = []

            participant_sheet_updates.append({'range': 'B2', 'values': [[round(data['cash'], 2)]]})

            for company, qty in data['holdings'].items():
                range_name = data['holding_cell_ranges'].get(company)
                if range_name:
                    participant_sheet_updates.append({'range': range_name, 'values': [[str(qty)]]})
                else:
                    logging.warning(
                        f"Could not find range for company '{company}' in '{participant_name}' holdings for batch update.")

            if participant_sheet_updates:
                participant_ws.batch_update(participant_sheet_updates)
                logging.info(f"Updated cash and holdings for {participant_name}.")

            if data['transactions_to_append_data']:
                participant_ws.append_rows(data['transactions_to_append_data'])
                logging.info(
                    f"Appended {len(data['transactions_to_append_data'])} transactions for {participant_name}.")

        except Exception as e:
            logging.error(f"Error performing batch updates for participant '{participant_name}': {e}")

    if not found_any_processed_in_batch:
        logging.info("No new transactions were processed in this batch.")
    else:
        logging.info("Finished processing all new transactions in this batch.")
    logging.info("--- Trade processing cycle completed ---")
    return found_any_processed_in_batch


# === SCRIPT ENTRY POINT ===
if __name__ == "__main__":
    logging.info("Mock Stock Simulation Script Started. Press Ctrl+C to stop.")

    for company, price in INITIAL_COMPANY_PRICES.items():
        PREVIOUS_PRICES[company] = price
        LAST_TRADED_PRICES[company] = price
        CURRENT_VWAP_PRICES[company] = price
        CURRENT_CIRCUITS[company]["upper"] = price * 1.20
        CURRENT_CIRCUITS[company]["lower"] = price * 0.80

    try:
        update_price_chart()
        apply_price_chart_conditional_formatting()
    except Exception as e:
        logging.error(f"Initial setup (price chart update or conditional formatting) failed: {e}")

    CYCLE_INTERVAL_SECONDS = 10

    while True:
        try:
            process_trades()
        except Exception as e:
            logging.critical(f"An unexpected error occurred during a trade processing cycle: {e}")

        update_price_chart()

        logging.info(f"Waiting {CYCLE_INTERVAL_SECONDS} seconds before next full cycle...")

        time.sleep(CYCLE_INTERVAL_SECONDS)
