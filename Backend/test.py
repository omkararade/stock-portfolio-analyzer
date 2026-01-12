import yfinance as yf
import pandas as pd

# Define the ticker symbol
ticker_symbol = "AAPL"

# Create a Ticker object
ticker = yf.Ticker(ticker_symbol)

# Get the analyst recommendations
recommendations = ticker.recommendations
# Or use the recommendations_trend attribute for a summary
recommendation_trend = ticker.recommendations

# Display the data
print(f"--- Analyst Recommendations for {ticker_symbol} ---")
print(recommendations)

print(f"\n--- Recommendation Trend for {ticker_symbol} ---")
print(recommendation_trend)

import yesg

try:
    esg_table = yesg.get_esg_full(ticker)
    total_score = esg_table["Total-Score"].iloc[-1]
    e_score = esg_table["E-Score"].iloc[-1]
    s_score = esg_table["S-Score"].iloc[-1]
    g_score = esg_table["G-Score"].iloc[-1]
except Exception:
    total_score = e_score = s_score = g_score = None


print(total_score)