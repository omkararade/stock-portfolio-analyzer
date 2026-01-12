import pandas as pd                                                     # Import pandas for data handling and DataFrame manipulation
import yfinance as yf                                                   # Import yfinance to fetch stock market data
import os                                                               # Import os for file and path handling
import sys                                                              # Import sys to modify Python path for module imports


# ----------------------------------------------- Add the scripts directory to Python path ---------------------------------------------

scripts_dir = os.path.dirname(os.path.abspath(__file__))                # Get the directory where this script is located
if scripts_dir not in sys.path:                                         # Check if that directory is not already in Python's search path
    sys.path.append(scripts_dir)                                        # Add it to Python’s module search path


# ------------------------------------------------- Import indicators with fallback ----------------------------------------------------

try:
    from indicators import calculate_rsi, calculate_macd, calculate_sma # Try importing custom RSI and MACD functions
    print(" Imported indicators from local module")
except ImportError:
    try:
        from .indicators import calculate_rsi, calculate_macd, calculate_sma # Alternative import style (relative import if using packages)
        print(" Imported indicators from relative module")
    except ImportError:
        print("❌ Could not import indicators module")
        # Define fallback functions if module not found
        def calculate_rsi(close_prices, period=14):
            delta = close_prices.diff()                                 # Price change from previous day
            gain = delta.where(delta > 0, 0)                            # Positive changes only
            loss = -delta.where(delta < 0, 0)                           # Negative changes only
            avg_gain = gain.rolling(window=period).mean()               # Average gain over 'period'
            avg_loss = loss.rolling(window=period).mean()               # Average loss over 'period'
            rs = avg_gain / avg_loss                                    # Relative strength (RS)
            rsi = 100 - (100 / (1 + rs))                                # Final RSI formula
            return rsi

        def calculate_macd(close_prices, short_window=12, long_window=26, signal_window=9):
            ema_short = close_prices.ewm(span=short_window, adjust=False).mean()    # 12-day EMA
            ema_long = close_prices.ewm(span=long_window, adjust=False).mean()      # 26-day EMA
            macd = ema_short - ema_long                                             # MACD = short EMA - long EMA
            signal = macd.ewm(span=signal_window, adjust=False).mean()              # Signal line (9-day EMA of MACD)
            histogram = macd - signal                                               # Difference = histogram
            return macd, signal, histogram

        def calculate_sma(close_prices, window=20):
            return close_prices.rolling(window=window).mean()

# -------------------------------------------------- Read tickers from Excel Subsheet(Sheet1) ----------------------------------------------------

def get_tickers_from_excel(excel_path=None, sheet_name="Sheet1"):
    if excel_path is None:                                                          # If no path is provided, locate Stock_data.xlsm automatically
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))      # Go one folder up
        excel_path = os.path.join(base_dir, "Stock_data.xlsm")                       # Default Excel file name

    print(f"📁 Looking for Excel file at: {excel_path}")

    if not os.path.exists(excel_path):                                              # If the file does not exist
        print(f"❌ Excel file not found at: {excel_path}")
        return ["AAPL", "MSFT", "GOOGL", "TSLA", "AMZN"]                            # Return fallback tickers

    try:
        excel_file = pd.ExcelFile(excel_path)                                       # Load Excel file to check sheet names
        print(f" Available sheets: {excel_file.sheet_names}")

        df = pd.read_excel(excel_path, sheet_name=sheet_name)                       # Read the specified sheet
        print(f" Successfully read sheet: {sheet_name}")
        print(f"Columns found: {df.columns.tolist()}")                              # Show columns found in Excel

        tickers = df.iloc[:, 0].dropna().tolist()                                   # Read tickers from first column
        print(f" Tickers extracted: {tickers}")
        return tickers

    except Exception as e:                                                          # If any error occurs
        print(f"❌ Error reading Excel file: {e}")
        return ["AAPL", "MSFT", "GOOGL", "TSLA", "AMZN"]                            # Return fallback tickers


# -------------------------------------------------- Read ESG from from Excel Subsheet(Manual_ESG) ----------------------------------------------------
def read_manual_esg(excel_path):
    try:
        df = pd.read_excel(excel_path, sheet_name="Manual_ESG")

        # Normalize ticker for safe joining
        df["Ticker"] = (
            df["Ticker"]
            .astype(str)
            .str.upper()
            .str.strip()
        )

        print(" Manual_ESG loaded successfully")
        return df

    except Exception as e:
        print(" Manual_ESG sheet not found or empty")
        return pd.DataFrame(columns=[
            "Ticker",
            "ESG Theme",
            "Manual ESG Score",
            "Confidence Level",
            "Assessment Criteria",
            "Review Date",
            "Analyst Notes"
        ])

# --------------------------------------- Fetch data and calculate indicators for each ticker ------------------------------------------

def calculate_quarterly_growth(stock):
    """Calculate quarterly revenue and earnings growth from financial statements"""
    
    revenue_q_growth = None
    earnings_q_growth = None
    
    try:
        # Get QUARTERLY financial statements
        quarterly_income = stock.quarterly_financials  # This is quarterly!
        
        if not quarterly_income.empty:
            # Check available revenue metrics
            revenue_metric = None
            for metric in ['Total Revenue', 'Revenue', 'Operating Revenue', 'Sales Revenue']:
                if metric in quarterly_income.index:
                    revenue_metric = metric
                    break
            
            # Calculate revenue growth
            if revenue_metric:
                revenues = quarterly_income.loc[revenue_metric]
                if len(revenues) >= 2:
                    revenue_current = revenues.iloc[0]  # Most recent quarter
                    revenue_previous = revenues.iloc[1]  # Previous quarter
                    if revenue_previous != 0:
                        revenue_q_growth = (revenue_current - revenue_previous) / revenue_previous
            
            # Calculate earnings growth
            earnings_metric = None
            for metric in ['Net Income', 'Net Income Common Stockholders', 'Net Income Continuous Operations']:
                if metric in quarterly_income.index:
                    earnings_metric = metric
                    break
            
            if earnings_metric:
                earnings = quarterly_income.loc[earnings_metric]
                if len(earnings) >= 2:
                    earnings_current = earnings.iloc[0]
                    earnings_previous = earnings.iloc[1]
                    if earnings_previous != 0:
                        earnings_q_growth = (earnings_current - earnings_previous) / earnings_previous
    
    except Exception as e:
        print(f"⚠️  Could not calculate quarterly growth: {e}")
    
    return revenue_q_growth, earnings_q_growth


def fetch_stock_data_with_indicators(tickers):
    all_data = []                                                                  

    for ticker in tickers:                                                          
        try:
            print(f"📊 Fetching data for {ticker}...")
            stock = yf.Ticker(ticker)                                              

            info = stock.info if hasattr(stock, 'info') else {}                                                    
            hist = stock.history(period="5y")                                      

            # ---------------- Fundamental Metrics ----------------
            current_price = info.get("currentPrice") or info.get("regularMarketPrice")
            pe_ratio = info.get("trailingPE") or info.get("forwardPE")
            market_cap = info.get("marketCap")
            dividend_yield = info.get("dividendYield")
            
            # ---------------- Get Financial Statements ----------------
            grossProfit = operatingIncome = netIncome = None
            totalCash = totalDebt = totalDebtToEquity = None
            freeCashflow = operatingCashflow = None
            
            try:
                # Get annual financial statements
                income_stmt = stock.financials
                balance_sheet = stock.balance_sheet
                cash_flow = stock.cashflow
                
                if not income_stmt.empty:
                    # Get latest annual values
                    if 'Gross Profit' in income_stmt.index:
                        grossProfit = income_stmt.loc['Gross Profit'].iloc[0]
                    if 'Operating Income' in income_stmt.index:
                        operatingIncome = income_stmt.loc['Operating Income'].iloc[0]
                    elif 'EBIT' in income_stmt.index:
                        operatingIncome = income_stmt.loc['EBIT'].iloc[0]
                    if 'Net Income' in income_stmt.index:
                        netIncome = income_stmt.loc['Net Income'].iloc[0]
                
                if not balance_sheet.empty:
                    # Get balance sheet values
                    if 'Total Cash' in balance_sheet.index:
                        totalCash = balance_sheet.loc['Total Cash'].iloc[0]
                    if 'Total Debt' in balance_sheet.index:
                        totalDebt = balance_sheet.loc['Total Debt'].iloc[0]
                    
                    # Calculate Debt to Equity
                    totalEquity = None
                    for equity_metric in ['Total Equity', 'Total Stockholder Equity', 'Stockholders Equity']:
                        if equity_metric in balance_sheet.index:
                            totalEquity = balance_sheet.loc[equity_metric].iloc[0]
                            break
                    
                    if totalDebt and totalEquity and totalEquity != 0:
                        totalDebtToEquity = totalDebt / totalEquity
                
                if not cash_flow.empty:
                    if 'Free Cash Flow' in cash_flow.index:
                        freeCashflow = cash_flow.loc['Free Cash Flow'].iloc[0]
                    if 'Operating Cash Flow' in cash_flow.index:
                        operatingCashflow = cash_flow.loc['Operating Cash Flow'].iloc[0]
                        
            except Exception as e:
                print(f"⚠️  Financial statements error for {ticker}: {e}")
            
            # ---------------- Calculate Growth Metrics ----------------
            # 1. Annual Growth (from info or calculate)
            earnings_growth = info.get("earningsGrowth")
            revenue_growth = info.get("revenueGrowth")
            
            if not earnings_growth or not revenue_growth:
                try:
                    if not income_stmt.empty:
                        # Annual revenue growth
                        if 'Total Revenue' in income_stmt.index:
                            revenues = income_stmt.loc['Total Revenue']
                            if len(revenues) >= 2:
                                rev_current = revenues.iloc[0]
                                rev_previous = revenues.iloc[1]
                                if rev_previous != 0:
                                    revenue_growth = (rev_current - rev_previous) / rev_previous
                        
                        # Annual earnings growth
                        if 'Net Income' in income_stmt.index:
                            earnings = income_stmt.loc['Net Income']
                            if len(earnings) >= 2:
                                earn_current = earnings.iloc[0]
                                earn_previous = earnings.iloc[1]
                                if earn_previous != 0:
                                    earnings_growth = (earn_current - earn_previous) / earn_previous
                except:
                    pass
            
            # 2. QUARTERLY Growth (calculate from quarterly financials)
            revenue_q_growth, earnings_q_growth = calculate_quarterly_growth(stock)
            
            # If still None, try from info as fallback
            if earnings_q_growth is None:
                earnings_q_growth = info.get("earningsQuarterlyGrowth")
            if revenue_q_growth is None:
                revenue_q_growth = info.get("revenueQuarterlyGrowth")
            
            # ---------------- Analyst Indicators ----------------
            recommendations = stock.recommendations
            strong_buy = buy = hold = sell = strong_sell = None
            total_analysts = None

            if recommendations is not None and not recommendations.empty:
                if 'period' in recommendations.columns:
                    latest = recommendations.loc[recommendations["period"] == "0m"]
                else:
                    latest = recommendations.tail(1)
                
                if not latest.empty:
                    strong_buy = int(latest.get("strongBuy", [0]).iloc[0]) if "strongBuy" in latest.columns else 0
                    buy = int(latest.get("buy", [0]).iloc[0]) if "buy" in latest.columns else 0
                    hold = int(latest.get("hold", [0]).iloc[0]) if "hold" in latest.columns else 0
                    sell = int(latest.get("sell", [0]).iloc[0]) if "sell" in latest.columns else 0
                    strong_sell = int(latest.get("strongSell", [0]).iloc[0]) if "strongSell" in latest.columns else 0
                    
                    total_analysts = strong_buy + buy + hold + sell + strong_sell

            target_mean = info.get("targetMeanPrice")
            target_high = info.get("targetHighPrice")
            target_low = info.get("targetLowPrice")

            # ---------------- Upside / Downside % ----------------
            upside_pct = None
            if current_price and target_mean:
                try:
                    upside_pct = ((target_mean - current_price) / current_price) * 100
                except (TypeError, ZeroDivisionError):
                    upside_pct = None

            if upside_pct is not None:
                if upside_pct >= 15:
                    upside_label = "High Upside"
                elif upside_pct >= 5:
                    upside_label = "Moderate Upside"
                else:
                    upside_label = "Limited / Downside"
            else:
                upside_label = "N/A"

            # -------------- Technical Indicators ---------------
            if not hist.empty and len(hist) > 200:
                close_prices = hist["Close"]
                
                rsi = calculate_rsi(close_prices).iloc[-1] if len(close_prices) >= 14 else None
                
                macd, signal, _ = calculate_macd(close_prices)
                macd_value = macd.iloc[-1] if not macd.empty else None
                signal_value = signal.iloc[-1] if not signal.empty else None
                
                sma_20 = calculate_sma(close_prices, window=20).iloc[-1] if len(close_prices) >= 20 else None
                sma_50 = calculate_sma(close_prices, window=50).iloc[-1] if len(close_prices) >= 50 else None
                sma_200 = calculate_sma(close_prices, window=200).iloc[-1] if len(close_prices) >= 200 else None
            else:
                rsi = macd_value = signal_value = sma_20 = sma_50 = sma_200 = None 

            # ---------------- ESG DATA ----------------
            esg_total = esg_env = esg_social = esg_gov = esg_percentile = None

            try:
                esg = stock.sustainability

                if isinstance(esg, pd.DataFrame) and not esg.empty:
                    if "totalEsg" in esg.index:
                        esg_total = esg.loc["totalEsg"].values[0] if len(esg.loc["totalEsg"].values) > 0 else None
                    if "environmentScore" in esg.index:
                        esg_env = esg.loc["environmentScore"].values[0] if len(esg.loc["environmentScore"].values) > 0 else None
                    if "socialScore" in esg.index:
                        esg_social = esg.loc["socialScore"].values[0] if len(esg.loc["socialScore"].values) > 0 else None
                    if "governanceScore" in esg.index:
                        esg_gov = esg.loc["governanceScore"].values[0] if len(esg.loc["governanceScore"].values) > 0 else None
                    if "percentile" in esg.index:
                        esg_percentile = esg.loc["percentile"].values[0] if len(esg.loc["percentile"].values) > 0 else None

            except Exception as e:
                print(f"ESG unavailable for {ticker}: {str(e)[:100]}")

            data = {
                "Ticker": ticker,
                "Current Price": round(current_price, 2) if current_price else "N/A",

                # --- Valuation ---
                "PE Ratio": round(pe_ratio, 2) if pe_ratio else "N/A",
                "Market Cap": f"{round(market_cap / 1e9, 2)}B" if market_cap else "N/A",
                "Dividend Yield": round(dividend_yield, 4) if dividend_yield else "N/A",

                # --- Financial Performance ---
                "Gross Profit": f"{round(grossProfit / 1e9, 2)}B" if grossProfit else "N/A",
                "Operating Income": f"{round(operatingIncome / 1e9, 2)}B" if operatingIncome else "N/A",
                "Net Income": f"{round(netIncome / 1e9, 2)}B" if netIncome else "N/A",

                # --- Balance Sheet ---
                "Total Cash": f"{round(totalCash / 1e9, 2)}B" if totalCash else "N/A",
                "Total Debt": f"{round(totalDebt / 1e9, 2)}B" if totalDebt else "N/A",
                "Debt to Equity": round(totalDebtToEquity, 2) if totalDebtToEquity else "N/A",

                # --- Cash Flow ---
                "Free Cash Flow": f"{round(freeCashflow / 1e9, 2)}B" if freeCashflow else "N/A",
                "Operating Cash Flow": f"{round(operatingCashflow / 1e9, 2)}B" if operatingCashflow else "N/A",

                # --- Growth Metrics ---
                "Earnings Growth YoY": f"{round(earnings_growth*100, 2)}%" if earnings_growth else "N/A",
                "Revenue Growth YoY": f"{round(revenue_growth*100, 2)}%" if revenue_growth else "N/A",
                "Earnings QoQ Growth": f"{round(earnings_q_growth*100, 2)}%" if earnings_q_growth else "N/A",
                "Revenue QoQ Growth": f"{round(revenue_q_growth*100, 2)}%" if revenue_q_growth else "N/A",

                # --- Technicals ---
                "RSI (14)": round(rsi, 2) if rsi is not None else "N/A",
                "SMA 20": round(sma_20, 2) if sma_20 is not None else "N/A",
                "SMA 50": round(sma_50, 2) if sma_50 is not None else "N/A",
                "SMA 200": round(sma_200, 2) if sma_200 is not None else "N/A",
                "MACD": round(macd_value, 2) if macd_value is not None else "N/A",
                "Signal Line": round(signal_value, 2) if signal_value is not None else "N/A",

                # --- Analyst Estimates ---
                "Strong Buy": strong_buy if strong_buy is not None else "N/A",
                "Buy": buy if buy is not None else "N/A",
                "Hold": hold if hold is not None else "N/A",
                "Sell": sell if sell is not None else "N/A",
                "Strong Sell": strong_sell if strong_sell is not None else "N/A",
                "Total Analysts (Breakdown)": total_analysts if total_analysts is not None else "N/A",
                "Target Mean": round(target_mean, 2) if target_mean else "N/A",
                "Target High": round(target_high, 2) if target_high else "N/A",
                "Target Low": round(target_low, 2) if target_low else "N/A",
                "Upside %": f"{round(upside_pct, 2)}%" if upside_pct is not None else "N/A",
                "Upside View": upside_label,

                # --- ESG Scores ---
                "ESG Total Score": round(esg_total, 2) if esg_total is not None else "N/A",
                "ESG Environment": round(esg_env, 2) if esg_env is not None else "N/A",
                "ESG Social": round(esg_social, 2) if esg_social is not None else "N/A",
                "ESG Governance": round(esg_gov, 2) if esg_gov is not None else "N/A",
                "ESG Percentile": round(esg_percentile, 2) if esg_percentile is not None else "N/A",
            }

            all_data.append(data)                                                  
            print(f"✅ Successfully processed {ticker}")

        except Exception as e:                                                      # If stock data fetching fails
            print(f"❌ Error fetching {ticker}: {e}")
            all_data.append({
                "Ticker": ticker,
                "Current Price": "Error",
                "Error Message": str(e)[:100]
            })

    return pd.DataFrame(all_data)
# --------------------------------------- Main Funtion  ------------------------------------------
def main():

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    excel_path = os.path.join(base_dir, "Stock_data.xlsm")

    # Step 1: Read stock tickers
    tickers = get_tickers_from_excel(excel_path)                                              
    print(" Tickers found:", tickers)

    # Step 2: Fetch stock data & indicators 
    df = fetch_stock_data_with_indicators(tickers)                                  
    print(df)

    # 3. Read Manual ESG
    manual_esg_df = read_manual_esg(excel_path)

    # 4. LEFT JOIN (no row loss, no overwrite)
    final_df = df.merge(
        manual_esg_df,
        on="Ticker",
        how="left"
    )

    print(final_df)

    #Save to Excel for verification
    # output_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "stock_data_output.xlsx")
    # final_df.to_excel(output_path, index=False)
    # print(f"📁 Saved to {output_path}")

if __name__ == "__main__":                                                          # Run only if file is executed directly
    main()