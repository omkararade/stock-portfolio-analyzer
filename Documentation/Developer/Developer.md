# Stock Data Fetcher - Developer Documentation

## 📋 Table of Contents

1. [Project Overview](#project-overview)
2. [Project Structure](#project-structure)
3. [Module Documentation](#module-documentation)
4. [Data Flow Architecture](#data-flow-architecture)
5. [API Integration](#api-integration)
6. [Technical Implementation](#technical-implementation)
7. [Error Handling](#error-handling)
8. [Extension Points](#extension-points)
9. [Testing & Debugging](#testing--debugging)
10. [Performance Considerations](#performance-considerations)

---

## 🎯 Project Overview

### Purpose
The Stock Data Fetcher is a Python-based financial data aggregation system that collects, processes, and formats stock market data from multiple sources into a unified Excel dashboard.

### Technology Stack
- **Python**: 3.11+ (portable distribution included)
- **Core Libraries**:
  - `pandas`: Data manipulation and DataFrame operations
  - `yfinance`: Yahoo Finance API wrapper for market data
  - `xlwings`: Excel automation via COM interface (Windows)
  - `openpyxl`: Excel file reading/writing
- **Dependencies**: See `requirements.txt`

### Key Features
- Automated data fetching from Yahoo Finance API
- Technical indicator calculations (RSI, MACD, SMA)
- Financial statement parsing and analysis
- ESG score integration (automated + manual)
- Excel formatting and presentation
- Error handling and fallback mechanisms

---

## 📁 Project Structure

```
DrArthur_Project/
├── Backend/                    # Core Python modules
│   ├── fetch_data.py          # Main data fetching logic
│   ├── indicators.py          # Technical indicator calculations
│   ├── update_excel.py        # Excel update and formatting
│   └── test.py                # Testing utilities
├── Documentation/
│   ├── Client/                # Client-facing documentation
│   └── Developer/             # Developer documentation
├── python/                    # Portable Python distribution (excluded from analysis)
├── Stock_data.xlsm            # Excel input/output file
├── requirements.txt           # Python dependencies
└── run_update.bat             # Batch execution script
```

---

## 📚 Module Documentation

### 1. `fetch_data.py`

#### Purpose
Main data fetching and processing module. Handles reading tickers from Excel, fetching stock data via yfinance, calculating metrics, and merging datasets.

#### Key Functions

##### `get_tickers_from_excel(excel_path=None, sheet_name="Sheet1")`
**Purpose**: Reads stock ticker symbols from Excel file.

**Parameters**:
- `excel_path` (str, optional): Path to Excel file. If None, auto-detects `Stock_data.xlsm` in project root.
- `sheet_name` (str): Sheet name containing tickers (default: "Sheet1")

**Returns**: `list[str]` - List of ticker symbols

**Implementation Details**:
```python
# Path resolution logic
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
excel_path = os.path.join(base_dir, "Stock_data.xlsm")

# Excel reading with error handling
excel_file = pd.ExcelFile(excel_path)
df = pd.read_excel(excel_path, sheet_name=sheet_name)
tickers = df.iloc[:, 0].dropna().tolist()  # First column, remove NaN
```

**Error Handling**:
- File not found → Returns default tickers: `["AAPL", "MSFT", "GOOGL", "TSLA", "AMZN"]`
- Sheet not found → Returns default tickers
- Empty sheet → Returns empty list

##### `read_manual_esg(excel_path)`
**Purpose**: Reads manual ESG assessments from Excel.

**Parameters**:
- `excel_path` (str): Path to Excel file

**Returns**: `pd.DataFrame` - DataFrame with ESG columns or empty DataFrame with column structure

**Column Structure**:
```python
columns = [
    "Ticker",
    "ESG Theme",
    "Manual ESG Score",
    "Confidence Level",
    "Assessment Criteria",
    "Review Date",
    "Analyst Notes"
]
```

**Data Normalization**:
- Tickers converted to uppercase
- Whitespace stripped
- String type conversion for safe joining

**Error Handling**:
- Sheet not found → Returns empty DataFrame with column structure
- Empty sheet → Returns empty DataFrame

##### `calculate_quarterly_growth(stock)`
**Purpose**: Calculates quarter-over-quarter growth rates from financial statements.

**Parameters**:
- `stock` (yf.Ticker): yfinance Ticker object

**Returns**: `tuple[float|None, float|None]` - (revenue_q_growth, earnings_q_growth)

**Implementation Logic**:
```python
# Get quarterly financial statements
quarterly_income = stock.quarterly_financials

# Find revenue metric (handles multiple naming conventions)
for metric in ['Total Revenue', 'Revenue', 'Operating Revenue', 'Sales Revenue']:
    if metric in quarterly_income.index:
        revenue_metric = metric
        break

# Calculate growth: (current - previous) / previous
revenue_q_growth = (revenue_current - revenue_previous) / revenue_previous
```

**Error Handling**:
- Missing data → Returns `(None, None)`
- Division by zero → Returns `None`
- Exception caught → Prints warning, returns `(None, None)`

##### `fetch_stock_data_with_indicators(tickers)`
**Purpose**: Main data fetching function. Processes each ticker and collects all metrics.

**Parameters**:
- `tickers` (list[str]): List of stock ticker symbols

**Returns**: `pd.DataFrame` - DataFrame with all stock data

**Data Collection Process**:

1. **Market Data** (`stock.info`):
   ```python
   current_price = info.get("currentPrice") or info.get("regularMarketPrice")
   pe_ratio = info.get("trailingPE") or info.get("forwardPE")
   market_cap = info.get("marketCap")
   dividend_yield = info.get("dividendYield")
   ```

2. **Financial Statements**:
   ```python
   income_stmt = stock.financials      # Annual income statement
   balance_sheet = stock.balance_sheet  # Annual balance sheet
   cash_flow = stock.cashflow           # Annual cash flow
   ```
   - Extracts latest annual values (`.iloc[0]`)
   - Handles multiple metric name variations

3. **Growth Metrics**:
   - **YoY**: From `info` dict or calculated from annual statements
   - **QoQ**: Calculated via `calculate_quarterly_growth()`

4. **Technical Indicators**:
   ```python
   hist = stock.history(period="5y")  # 5 years of daily prices
   close_prices = hist["Close"]
   rsi = calculate_rsi(close_prices).iloc[-1]
   macd, signal, _ = calculate_macd(close_prices)
   sma_20 = calculate_sma(close_prices, window=20).iloc[-1]
   ```

5. **Analyst Data**:
   ```python
   recommendations = stock.recommendations
   # Extract latest period (0m) recommendations
   # Count Strong Buy, Buy, Hold, Sell, Strong Sell
   ```

6. **ESG Scores**:
   ```python
   esg = stock.sustainability
   # Extract totalEsg, environmentScore, socialScore, governanceScore, percentile
   ```

**Data Dictionary Structure**:
```python
data = {
    "Ticker": ticker,
    "Current Price": round(current_price, 2) if current_price else "N/A",
    "PE Ratio": round(pe_ratio, 2) if pe_ratio else "N/A",
    # ... 40+ more fields
}
```

**Error Handling**:
- Individual ticker failure → Appends error row, continues processing
- Missing data → Uses "N/A" placeholder
- API errors → Catches exception, logs error, continues

##### `main()`
**Purpose**: Entry point for standalone execution.

**Execution Flow**:
1. Locate Excel file
2. Read tickers
3. Fetch stock data
4. Read manual ESG
5. Merge datasets (LEFT JOIN)
6. Print results

**Note**: Does not write to Excel (use `update_excel.py` for full pipeline)

---

### 2. `indicators.py`

#### Purpose
Technical indicator calculation module. Provides RSI, MACD, and SMA functions.

#### Functions

##### `calculate_rsi(close_prices, period=14)`
**Purpose**: Calculates Relative Strength Index.

**Formula**:
```
RS = Average Gain / Average Loss (over period)
RSI = 100 - (100 / (1 + RS))
```

**Implementation**:
```python
delta = close_prices.diff()                    # Price changes
gain = delta.where(delta > 0, 0)              # Positive changes only
loss = -delta.where(delta < 0, 0)              # Negative changes only
avg_gain = gain.rolling(window=period).mean() # Rolling average gain
avg_loss = loss.rolling(window=period).mean()  # Rolling average loss
rs = avg_gain / avg_loss                        # Relative strength
rsi = 100 - (100 / (1 + rs))                   # Final RSI
```

**Returns**: `pd.Series` - RSI values (same length as input)

**Edge Cases**:
- Insufficient data (< period) → Returns NaN for initial values
- Division by zero (avg_loss = 0) → Returns NaN

##### `calculate_macd(close_prices, short_window=12, long_window=26, signal_window=9)`
**Purpose**: Calculates MACD (Moving Average Convergence Divergence).

**Components**:
- **MACD Line**: `12-day EMA - 26-day EMA`
- **Signal Line**: `9-day EMA of MACD`
- **Histogram**: `MACD - Signal`

**Implementation**:
```python
ema_short = close_prices.ewm(span=short_window, adjust=False).mean()
ema_long = close_prices.ewm(span=long_window, adjust=False).mean()
macd = ema_short - ema_long
signal = macd.ewm(span=signal_window, adjust=False).mean()
histogram = macd - signal
```

**Returns**: `tuple[pd.Series, pd.Series, pd.Series]` - (macd, signal, histogram)

**Note**: Uses Exponential Moving Average (EMA), not Simple Moving Average (SMA)

##### `calculate_sma(close_prices, window=20)`
**Purpose**: Calculates Simple Moving Average.

**Formula**: `SMA = Sum of closing prices over window / window`

**Implementation**:
```python
return close_prices.rolling(window=window).mean()
```

**Returns**: `pd.Series` - SMA values

**Edge Cases**:
- Insufficient data (< window) → Returns NaN for initial values

---

### 3. `update_excel.py`

#### Purpose
Excel update and formatting module. Orchestrates data fetching, merging, writing, and formatting.

#### Key Functions

##### `format_excel(sheet)`
**Purpose**: Applies comprehensive visual formatting to Excel sheet.

**Formatting Categories**:

1. **Base Formatting** (All cells):
   ```python
   used_range.api.Interior.Color = 0x000000      # Black background
   used_range.api.Font.Color = 0xFFFFFF           # White text
   used_range.api.Font.Size = 10                  # Font size
   used_range.api.HorizontalAlignment = -4108     # Center align (xlCenter)
   used_range.api.VerticalAlignment = -4108       # Center align
   ```

2. **Borders**:
   ```python
   for border_id in range(7, 13):
       used_range.api.Borders(border_id).LineStyle = 1    # Continuous
       used_range.api.Borders(border_id).Weight = 1        # Thin
       used_range.api.Borders(border_id).Color = 0x404040  # Dark gray
   ```

3. **Header Row**:
   ```python
   header.api.Font.Bold = True
   header.api.Font.Size = 11
   header.api.Font.Color = 0xFFFFFF
   header.api.Interior.Color = 0x2E75B5  # Dark blue
   ```

4. **Frozen Panes**:
   ```python
   window.SplitRow = 1        # Freeze row 1
   window.SplitColumn = 2     # Freeze columns A-B
   window.FreezePanes = True
   ```

5. **Conditional Formatting**:
   - **Positive Values** (Dividend Yield, Upside %): Green (#4CAF50), bold
   - **Negative Values**: Red (#F44336), bold
   - **RSI < 30**: Blue (#2196F3)
   - **RSI > 70**: Orange (#FF9800)

**Helper Functions**:
- `col_range(col_name)`: Returns data range for column (rows 2 to last)
- `col_letter(col_name)`: Converts column name to Excel letter (A, B, C)

**Implementation Notes**:
- Uses `openpyxl.utils.get_column_letter()` for column conversion
- Uses xlwings COM API for formatting (Windows-specific)
- Handles missing columns gracefully

##### `collapse_duplicate_ticker_rows(sheet)`
**Purpose**: Creates visual grouping for duplicate ticker rows (multiple ESG themes).

**Logic**:
```python
# For each row, if ticker matches previous row:
if current_ticker == prev_ticker:
    # Clear all non-ESG columns
    for col_name in headers:
        if col_name not in ESG_ONLY_COLUMNS and col_name != "Ticker":
            sheet.cells(row, col_idx).value = ""
    # Also clear ticker cell
    sheet.cells(row, ticker_col_idx).value = ""
```

**ESG-Only Columns** (preserved in duplicate rows):
```python
ESG_ONLY_COLUMNS = {
    "ESG Theme", "Manual ESG Score", "Confidence Level",
    "Assessment Criteria", "Review Date", "Analyst Notes",
    "Upside Bucket", "ESG Category", "RSI Status"
}
```

**Result**: Creates stacked/grouped appearance where only first row shows full data.

##### `add_upside_bucket(df)`
**Purpose**: Categorizes Upside % into buckets.

**Classification**:
```python
if num >= 0.10:
    return "High (>10%)"
elif num >= 0:
    return "Medium (0–10%)"
else:
    return "Negative"
```

**Input Handling**:
- Removes "%" sign
- Converts to decimal
- Handles "N/A" and NaN

**Returns**: Modified DataFrame with new "Upside Bucket" column

##### `add_esg_category(df)`
**Purpose**: Categorizes Manual ESG Score.

**Classification**:
```python
if score >= 60:
    return "Good (≥60)"
elif score >= 40:
    return "Average (40–59)"
else:
    return "Poor (<40)"
```

**Returns**: Modified DataFrame with new "ESG Category" column

##### `add_rsi_status(df)`
**Purpose**: Categorizes RSI values.

**Classification**:
```python
if rsi > 70:
    return "Overbought (>70)"
elif rsi < 30:
    return "Oversold (<30)"
else:
    return "Neutral"
```

**Returns**: Modified DataFrame with new "RSI Status" column

##### `update_excel()`
**Purpose**: Main orchestration function.

**Execution Steps**:

1. **Locate Excel File**:
   ```python
   base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
   excel_path = os.path.join(base_dir, "Stock_data.xlsm")
   ```

2. **Connect to Excel**:
   ```python
   app = xw.apps.active  # Get active Excel instance
   workbook = next((wb for wb in app.books if "Stock_data.xlsm" in wb.name), None)
   if workbook is None:
       workbook = app.books.open(excel_path)
   ```

3. **Fetch Data**:
   ```python
   tickers = get_tickers_from_excel(excel_path, sheet_name="Sheet1")
   df = fetch_stock_data_with_indicators(tickers)
   manual_esg_df = read_manual_esg(excel_path)
   ```

4. **Merge & Enrich**:
   ```python
   final_df = df.merge(manual_esg_df, on="Ticker", how="left")
   final_df = add_upside_bucket(final_df)
   final_df = add_esg_category(final_df)
   final_df = add_rsi_status(final_df)
   ```

5. **Write to Excel**:
   ```python
   sheet = workbook.sheets["RawData"]
   sheet.clear()
   sheet.range("A1").value = final_df
   ```

6. **Apply Formatting**:
   ```python
   format_excel(sheet)
   collapse_duplicate_ticker_rows(sheet)
   ```

**Error Handling**:
- Returns `False` on any exception
- Prints error messages (commented out traceback for production)

---

## 🔄 Data Flow Architecture

### Complete Pipeline

```
┌─────────────────────────────────────────────────────────────┐
│                    EXECUTION START                           │
│                  (run_update.bat)                            │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│              update_excel.py::update_excel()                 │
│  Step 1: Locate Stock_data.xlsm                              │
│  Step 2: Connect to Excel application (xlwings)             │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│              fetch_data.py::get_tickers_from_excel()         │
│  - Read Sheet1, Column A                                     │
│  - Extract ticker symbols                                    │
│  - Return: list[str]                                         │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│        fetch_data.py::fetch_stock_data_with_indicators()     │
│  For each ticker:                                            │
│  ├─ Create yf.Ticker object                                 │
│  ├─ Fetch stock.info (market data)                          │
│  ├─ Fetch stock.history(period="5y") (price data)           │
│  ├─ Fetch stock.financials (income statement)               │
│  ├─ Fetch stock.balance_sheet                               │
│  ├─ Fetch stock.cashflow                                    │
│  ├─ Fetch stock.quarterly_financials                        │
│  ├─ Fetch stock.recommendations (analyst data)             │
│  ├─ Fetch stock.sustainability (ESG)                         │
│  ├─ Calculate technical indicators (RSI, MACD, SMA)         │
│  ├─ Calculate growth metrics (YoY, QoQ)                     │
│  └─ Build data dictionary                                    │
│  Return: pd.DataFrame                                        │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│              fetch_data.py::read_manual_esg()                 │
│  - Read Manual_ESG sheet                                     │
│  - Normalize ticker column                                   │
│  - Return: pd.DataFrame (or empty with structure)           │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│                    DATA MERGING                              │
│  final_df = df.merge(manual_esg_df, on="Ticker", how="left")│
│  - LEFT JOIN preserves all stock data                        │
│  - Adds ESG columns where available                         │
│  - Creates multiple rows for multiple ESG themes             │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│                  DATA ENRICHMENT                              │
│  - add_upside_bucket(final_df)                               │
│  - add_esg_category(final_df)                                │
│  - add_rsi_status(final_df)                                  │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│                  EXCEL WRITING                                │
│  sheet = workbook.sheets["RawData"]                          │
│  sheet.clear()                                               │
│  sheet.range("A1").value = final_df                          │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│                  EXCEL FORMATTING                            │
│  format_excel(sheet)                                         │
│  - Apply base formatting (colors, borders, alignment)       │
│  - Format header row                                         │
│  - Apply conditional formatting                              │
│  - Freeze panes                                              │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│              PRESENTATION LOGIC                               │
│  collapse_duplicate_ticker_rows(sheet)                       │
│  - Clear non-ESG columns in duplicate rows                  │
│  - Create visual grouping                                   │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│                    COMPLETION                                 │
│  Return: True (success) or False (failure)                 │
└─────────────────────────────────────────────────────────────┘
```

### Data Transformation Stages

#### Stage 1: Raw API Data
```python
{
    "currentPrice": 185.64,
    "trailingPE": 30.15,
    "marketCap": 2850000000000,
    # ... raw API response
}
```

#### Stage 2: Processed Dictionary
```python
{
    "Ticker": "AAPL",
    "Current Price": 185.64,
    "PE Ratio": 30.15,
    "Market Cap": "2.85B",
    # ... formatted values
}
```

#### Stage 3: DataFrame
```python
pd.DataFrame([
    {"Ticker": "AAPL", "Current Price": 185.64, ...},
    {"Ticker": "MSFT", "Current Price": 378.85, ...},
    # ... one row per ticker
])
```

#### Stage 4: Merged DataFrame
```python
# After LEFT JOIN with Manual_ESG
# Multiple rows possible per ticker (if multiple ESG themes)
```

#### Stage 5: Enriched DataFrame
```python
# Added columns:
# - Upside Bucket
# - ESG Category
# - RSI Status
```

#### Stage 6: Excel Output
```python
# Written to RawData sheet
# Formatted with colors, borders, conditional formatting
# Duplicate rows collapsed for presentation
```

---

## 🔌 API Integration

### Yahoo Finance API (via yfinance)

#### Ticker Object Creation
```python
import yfinance as yf
stock = yf.Ticker("AAPL")
```

#### Available Data Sources

1. **Market Info** (`stock.info`):
   - Dictionary with 100+ fields
   - Includes: price, ratios, market cap, dividend yield
   - **Update Frequency**: Real-time during market hours

2. **Historical Prices** (`stock.history(period="5y")`):
   - Returns DataFrame with OHLCV data
   - Columns: Open, High, Low, Close, Volume
   - **Period Options**: "1d", "5d", "1mo", "3mo", "6mo", "1y", "2y", "5y", "10y", "ytd", "max"
   - **Interval Options**: "1m", "2m", "5m", "15m", "30m", "60m", "90m", "1h", "1d", "5d", "1wk", "1mo", "3mo"

3. **Financial Statements**:
   - `stock.financials`: Annual income statement
   - `stock.quarterly_financials`: Quarterly income statement
   - `stock.balance_sheet`: Annual balance sheet
   - `stock.quarterly_balance_sheet`: Quarterly balance sheet
   - `stock.cashflow`: Annual cash flow statement
   - `stock.quarterly_cashflow`: Quarterly cash flow statement
   - **Format**: DataFrame with dates as columns, metrics as rows

4. **Analyst Data** (`stock.recommendations`):
   - Returns DataFrame with columns: Date, Firm, To Grade, From Grade, Action
   - **Alternative**: `stock.recommendations_summary` (aggregated view)

5. **ESG Data** (`stock.sustainability`):
   - Returns DataFrame with ESG scores
   - **Availability**: Not available for all stocks
   - **Fields**: totalEsg, environmentScore, socialScore, governanceScore, percentile

#### API Limitations

1. **Rate Limiting**:
   - No official rate limit documented
   - Practical limit: ~100 requests/minute
   - **Mitigation**: Sequential processing (not parallel)

2. **Data Availability**:
   - Some metrics missing for small-cap stocks
   - ESG scores not available for all companies
   - Historical data limited by stock listing date

3. **Data Freshness**:
   - Real-time during market hours
   - Delayed 15-20 minutes for free tier
   - Financial statements updated quarterly

#### Error Handling

```python
try:
    stock = yf.Ticker(ticker)
    info = stock.info if hasattr(stock, 'info') else {}
except Exception as e:
    print(f"❌ Error fetching {ticker}: {e}")
    # Continue with next ticker
```

---

## 🛠️ Technical Implementation

### Path Resolution

#### Project Root Detection
```python
# In fetch_data.py
scripts_dir = os.path.dirname(os.path.abspath(__file__))  # Backend/
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # Project root

# In update_excel.py
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
scripts_dir = os.path.join(project_root, "Backend")
```

#### Excel File Location
```python
excel_path = os.path.join(base_dir, "Stock_data.xlsm")
```

### Module Import Strategy

#### Dynamic Path Addition
```python
# Add project directories to Python path
if scripts_dir not in sys.path:
    sys.path.append(scripts_dir)
if project_root not in sys.path:
    sys.path.append(project_root)
```

#### Import with Fallback
```python
try:
    from indicators import calculate_rsi, calculate_macd, calculate_sma
except ImportError:
    try:
        from .indicators import calculate_rsi, calculate_macd, calculate_sma
    except ImportError:
        # Define fallback functions inline
        def calculate_rsi(close_prices, period=14):
            # ... implementation
```

### DataFrame Operations

#### Merging Strategy
```python
# LEFT JOIN preserves all stock data rows
final_df = df.merge(
    manual_esg_df,
    on="Ticker",
    how="left"  # LEFT JOIN
)
```

**Result**:
- All tickers from `df` preserved
- ESG columns added where ticker matches
- Multiple ESG themes create multiple rows per ticker

#### Column Normalization
```python
# Normalize ticker column for safe joining
df["Ticker"] = (
    df["Ticker"]
    .astype(str)
    .str.upper()
    .str.strip()
)
```

### Excel Integration (xlwings)

#### COM Interface Usage
```python
import xlwings as xw

# Get active Excel application
app = xw.apps.active

# Find or open workbook
workbook = next(
    (wb for wb in app.books if "Stock_data.xlsm" in wb.name),
    None
)
if workbook is None:
    workbook = app.books.open(excel_path)

# Access sheet
sheet = workbook.sheets["RawData"]
```

#### Direct API Access
```python
# xlwings provides COM API access via .api property
sheet.api.Interior.Color = 0x000000  # Black background
sheet.api.Font.Color = 0xFFFFFF      # White text
sheet.api.Font.Bold = True           # Bold text
```

**Note**: Requires Excel to be installed and running (Windows-specific)

### Data Formatting

#### Numeric Formatting
```python
# Rounding
"Current Price": round(current_price, 2) if current_price else "N/A"

# Billions conversion
"Market Cap": f"{round(market_cap / 1e9, 2)}B" if market_cap else "N/A"

# Percentage conversion
"Earnings Growth YoY": f"{round(earnings_growth*100, 2)}%" if earnings_growth else "N/A"
```

#### Missing Data Handling
```python
# Consistent "N/A" for missing values
value if value is not None else "N/A"

# Check for NaN
if pd.isna(val) or val == "N/A":
    return "N/A"
```

---

## ⚠️ Error Handling

### Error Handling Strategy

#### 1. Graceful Degradation
- Missing data → Use "N/A" placeholder
- Failed ticker → Continue with next ticker
- Missing sheet → Use empty DataFrame with structure

#### 2. Fallback Values
```python
# Excel file not found
if not os.path.exists(excel_path):
    return ["AAPL", "MSFT", "GOOGL", "TSLA", "AMZN"]  # Default tickers

# Price not available
current_price = info.get("currentPrice") or info.get("regularMarketPrice")
```

#### 3. Exception Catching
```python
try:
    # Risky operation
    esg = stock.sustainability
except Exception as e:
    print(f"ESG unavailable for {ticker}: {str(e)[:100]}")
    # Continue with None values
```

#### 4. Data Validation
```python
# Check for empty DataFrames
if not income_stmt.empty:
    # Process data
else:
    # Use None values

# Check for sufficient data
if len(close_prices) >= 14:
    rsi = calculate_rsi(close_prices).iloc[-1]
else:
    rsi = None
```

### Common Error Scenarios

#### 1. Excel File Not Found
- **Detection**: `os.path.exists(excel_path)` returns False
- **Handling**: Use default tickers
- **User Impact**: System continues with sample data

#### 2. Excel Not Running
- **Detection**: `xw.apps.active` returns None
- **Handling**: Print error, return False
- **User Impact**: Script fails gracefully

#### 3. API Timeout
- **Detection**: Exception during `stock.history()` or `stock.info`
- **Handling**: Catch exception, log error, continue with next ticker
- **User Impact**: Missing data for affected ticker

#### 4. Missing Financial Data
- **Detection**: Empty DataFrame or missing index
- **Handling**: Check for empty, use None, display "N/A"
- **User Impact**: Some columns show "N/A"

#### 5. Division by Zero
- **Detection**: `revenue_previous == 0` in growth calculation
- **Handling**: Check before division, return None
- **User Impact**: Growth metric shows "N/A"

---

## 🔧 Extension Points

### Adding New Metrics

#### Step 1: Fetch Data
```python
# In fetch_stock_data_with_indicators()
new_metric = info.get("newMetricName") or calculate_new_metric(stock)
```

#### Step 2: Add to Data Dictionary
```python
data = {
    # ... existing fields
    "New Metric": format_value(new_metric),
}
```

#### Step 3: (Optional) Add Formatting
```python
# In format_excel()
if "New Metric" in headers:
    rng = col_range("New Metric")
    # Apply conditional formatting
```

### Adding New Technical Indicators

#### Step 1: Implement Function
```python
# In indicators.py
def calculate_new_indicator(close_prices, param1=10, param2=20):
    # Implementation
    return indicator_values
```

#### Step 2: Import and Use
```python
# In fetch_data.py
from indicators import calculate_new_indicator

# In fetch_stock_data_with_indicators()
new_indicator = calculate_new_indicator(close_prices, param1=10, param2=20)
data["New Indicator"] = new_indicator.iloc[-1] if not new_indicator.empty else "N/A"
```

### Adding New Data Sources

#### Example: Alternative ESG Source
```python
def fetch_alternative_esg(ticker):
    try:
        # Call alternative API
        esg_data = alternative_api.get_esg(ticker)
        return esg_data
    except Exception as e:
        return None

# In fetch_stock_data_with_indicators()
alternative_esg = fetch_alternative_esg(ticker)
if alternative_esg:
    data["Alternative ESG"] = alternative_esg
```

### Custom Formatting Rules

#### Add Conditional Formatting
```python
# In format_excel()
new_indicator_rng = col_range("New Indicator")
if new_indicator_rng:
    new_indicator_rng.api.FormatConditions.Delete()
    
    # Add rule for high values
    cond = new_indicator_rng.api.FormatConditions.Add(
        Type=2, Operator=5, Formula1="100"
    )
    cond.Font.Color = 0x4CAF50  # Green
```

### Parallel Processing

#### Current: Sequential Processing
```python
for ticker in tickers:
    data = fetch_stock_data_with_indicators([ticker])
    all_data.append(data)
```

#### Potential: Parallel Processing
```python
from concurrent.futures import ThreadPoolExecutor

def fetch_single_ticker(ticker):
    return fetch_stock_data_with_indicators([ticker])

with ThreadPoolExecutor(max_workers=5) as executor:
    results = executor.map(fetch_single_ticker, tickers)
    all_data = list(results)
```

**Note**: Yahoo Finance API may rate limit parallel requests

---

## 🧪 Testing & Debugging

### Testing Utilities

#### `test.py` Module
```python
# Test individual components
import yfinance as yf
ticker = yf.Ticker("AAPL")
recommendations = ticker.recommendations
print(recommendations)
```

### Debugging Strategies

#### 1. Print Statements
```python
print(f"📊 Fetching data for {ticker}...")
print(f"✅ Successfully processed {ticker}")
print(f"❌ Error fetching {ticker}: {e}")
```

#### 2. DataFrame Inspection
```python
# In main() or update_excel()
print(df.head())
print(df.columns.tolist())
print(df.info())
```

#### 3. Excel Verification
```python
# Save intermediate DataFrame
df.to_excel("debug_output.xlsx", index=False)
```

#### 4. Error Logging
```python
import traceback

try:
    # Risky operation
except Exception as e:
    traceback.print_exc()  # Full stack trace
    print(f"Error: {e}")
```

### Common Debugging Scenarios

#### Issue: Missing Data for Specific Ticker
**Debug Steps**:
1. Test ticker individually: `yf.Ticker("TICKER").info`
2. Check if ticker symbol is correct
3. Verify data availability on Yahoo Finance website
4. Check for API errors in console output

#### Issue: Excel Formatting Not Applied
**Debug Steps**:
1. Verify Excel is running: `xw.apps.active`
2. Check sheet name: `workbook.sheets["RawData"]`
3. Verify data written: Check cell values
4. Test formatting manually in Excel

#### Issue: Merge Not Working
**Debug Steps**:
1. Check ticker normalization: Print both DataFrames' Ticker columns
2. Verify column names match: `df.columns.tolist()`
3. Check for duplicates: `df.duplicated(subset=["Ticker"])`

---

## ⚡ Performance Considerations

### Current Performance

#### Processing Time
- **Per Ticker**: ~30-60 seconds
- **10 Tickers**: ~5-10 minutes
- **50 Tickers**: ~25-50 minutes

#### Bottlenecks
1. **API Calls**: Sequential fetching (one ticker at a time)
2. **Historical Data**: 5 years of daily data (~1,250 rows per ticker)
3. **Excel Writing**: Large DataFrames take time to write
4. **Formatting**: Conditional formatting applied cell-by-cell

### Optimization Opportunities

#### 1. Parallel Processing
```python
# Use ThreadPoolExecutor for concurrent API calls
# Risk: API rate limiting
# Benefit: 3-5x speedup for large portfolios
```

#### 2. Caching
```python
# Cache historical data (doesn't change frequently)
import pickle

def get_cached_history(ticker, cache_dir="cache"):
    cache_file = os.path.join(cache_dir, f"{ticker}_history.pkl")
    if os.path.exists(cache_file):
        # Check if cache is fresh (< 1 hour old)
        if os.path.getmtime(cache_file) > time.time() - 3600:
            return pd.read_pickle(cache_file)
    # Fetch and cache
    hist = stock.history(period="5y")
    hist.to_pickle(cache_file)
    return hist
```

#### 3. Batch API Calls
```python
# yfinance supports multiple tickers
stocks = yf.Tickers("AAPL MSFT GOOGL")
# Fetch all at once (if API supports)
```

#### 4. Incremental Updates
```python
# Only update changed tickers
existing_tickers = set(df_existing["Ticker"])
new_tickers = [t for t in tickers if t not in existing_tickers]
# Only fetch new tickers
```

### Memory Considerations

#### DataFrame Size
- **Per Ticker Row**: ~50 columns × ~100 bytes = ~5 KB
- **100 Tickers**: ~500 KB (negligible)
- **Historical Data**: ~1,250 rows × 6 columns × 8 bytes = ~60 KB per ticker

#### Excel File Size
- **Raw Data**: ~1-5 MB (depending on number of rows)
- **Formatting**: Minimal overhead
- **Total**: Typically < 10 MB

---

## 📝 Code Style & Conventions

### Naming Conventions
- **Functions**: `snake_case` (e.g., `fetch_stock_data_with_indicators`)
- **Variables**: `snake_case` (e.g., `current_price`, `pe_ratio`)
- **Constants**: `UPPER_SNAKE_CASE` (e.g., `ESG_ONLY_COLUMNS`)
- **Classes**: `PascalCase` (not used in current codebase)

### Documentation
- **Docstrings**: Not consistently used (opportunity for improvement)
- **Comments**: Extensive inline comments explaining logic
- **Print Statements**: Used for user feedback (✅, ❌, ⚠️ emojis)

### Code Organization
- **Imports**: Grouped by type (standard library, third-party, local)
- **Functions**: Logical grouping within modules
- **Error Handling**: Try-except blocks around risky operations

---

## 🔐 Security Considerations

### Input Validation
- **Ticker Symbols**: No validation (relies on API to handle invalid tickers)
- **Excel File**: Path resolution prevents directory traversal
- **User Data**: No sensitive data stored or transmitted

### API Security
- **No Authentication**: Yahoo Finance API is public (no API keys)
- **Rate Limiting**: Sequential processing prevents excessive requests
- **Error Handling**: Prevents information leakage in error messages

### File System
- **Excel File**: Read/write operations limited to project directory
- **No External Files**: All operations within project structure

---

## 🚀 Deployment Considerations

### Portable Python Distribution
- **Location**: `python/` folder
- **Usage**: `python\python.exe Backend\update_excel.py`
- **Benefits**: No system-wide Python installation required
- **Limitations**: Windows-specific paths

### Batch File Execution
```batch
@echo off
pushd %~dp0
python\python.exe Backend\update_excel.py
pause
popd
```

### Dependencies
- **requirements.txt**: Lists all Python packages
- **Installation**: Not required (portable Python includes packages)
- **Updates**: Manual update of portable Python packages

### Excel Dependency
- **Required**: Microsoft Excel installed
- **Version**: Excel 2016 or later
- **Platform**: Windows (xlwings COM interface)

---

## 📚 Additional Resources

### Library Documentation
- **yfinance**: https://github.com/ranaroussi/yfinance
- **pandas**: https://pandas.pydata.org/docs/
- **xlwings**: https://docs.xlwings.org/
- **openpyxl**: https://openpyxl.readthedocs.io/

### Financial Data Sources
- **Yahoo Finance**: https://finance.yahoo.com/
- **Alternative APIs**: Alpha Vantage, IEX Cloud, Polygon.io

### Technical Indicators
- **RSI**: https://www.investopedia.com/terms/r/rsi.asp
- **MACD**: https://www.investopedia.com/terms/m/macd.asp
- **Moving Averages**: https://www.investopedia.com/terms/m/movingaverage.asp

---

*Last Updated: 2025*

