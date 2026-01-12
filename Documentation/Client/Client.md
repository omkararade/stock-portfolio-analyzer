# Stock Data Fetcher - Client Documentation

## 📋 Overview

The **Stock Data Fetcher** is an automated financial data aggregation system that collects comprehensive stock market information from multiple sources and presents it in a formatted Excel dashboard. The system integrates fundamental analysis, technical indicators, ESG ratings, and analyst recommendations into a unified dataset for investment decision-making.

---

## 🎯 What This System Does

### Core Functionality

The system automatically:

1. **Reads your stock portfolio** from an Excel file (`Stock_data.xlsm`)
2. **Fetches real-time market data** for each stock ticker
3. **Calculates technical indicators** (RSI, MACD, Moving Averages)
4. **Retrieves financial statements** (Income Statement, Balance Sheet, Cash Flow)
5. **Gathers analyst recommendations** and price targets
6. **Collects ESG scores** from public sources
7. **Merges manual ESG assessments** (if provided)
8. **Formats and displays results** in a professional Excel dashboard

### Data Sources

- **Yahoo Finance API** (via `yfinance` library): Primary data source for market data, financials, and ESG scores
- **Excel File**: Input (tickers) and output (formatted dashboard)
- **Manual ESG Sheet**: Optional custom ESG assessments

---

## 📊 Parameters & Inputs

### Required Inputs

#### 1. Excel File Structure (`Stock_data.xlsm`)

**Sheet1** - Stock Tickers (Required)
- **Location**: Column A
- **Format**: One ticker symbol per row (e.g., AAPL, MSFT, GOOGL)
- **Case**: Case-insensitive (automatically converted to uppercase)
- **Example**:
  ```
  Column A: Tickers
  AAPL
  MSFT
  GOOGL
  TSLA
  AMZN
  ```

**Manual_ESG** - Custom ESG Assessments (Optional)
- **Purpose**: Add your own ESG evaluations that override or supplement automated scores
- **Required Columns**:
  - `Ticker`: Stock symbol (must match Sheet1)
  - `ESG Theme`: Category or focus area (e.g., "Climate Change", "Labor Practices")
  - `Manual ESG Score`: Numeric score (0-100)
  - `Confidence Level`: Your confidence in the assessment
  - `Assessment Criteria`: Description of evaluation method
  - `Review Date`: Date of assessment
  - `Analyst Notes`: Additional comments

**RawData** - Output Sheet (Auto-generated)
- **Purpose**: Contains all fetched and calculated data
- **Format**: Automatically formatted with black background theme
- **Note**: This sheet is overwritten each time the system runs

### System Parameters

#### Data Fetching Period
- **Historical Data**: 5 years (`period="5y"`)
- **Purpose**: Used for calculating technical indicators (RSI, MACD, SMA)
- **Modification**: Can be changed in code (requires developer assistance)

#### Technical Indicator Parameters

| Indicator | Period/Window | Purpose |
|-----------|---------------|---------|
| **RSI** | 14 days | Momentum indicator (0-100 scale) |
| **MACD** | Short: 12, Long: 26, Signal: 9 | Trend-following momentum indicator |
| **SMA 20** | 20 days | Short-term moving average |
| **SMA 50** | 50 days | Medium-term moving average |
| **SMA 200** | 200 days | Long-term moving average |

---

## 🔢 Calculations & Metrics Explained

### 1. Valuation Metrics

#### Price-to-Earnings (PE) Ratio
- **Formula**: `Market Price per Share / Earnings per Share`
- **Interpretation**: 
  - Lower PE = Potentially undervalued
  - Higher PE = Potentially overvalued or high growth expectations
- **Source**: `trailingPE` or `forwardPE` from Yahoo Finance

#### Market Capitalization
- **Formula**: `Current Price × Total Shares Outstanding`
- **Display**: Converted to billions (B) or trillions (T)
- **Purpose**: Company size indicator

#### Dividend Yield
- **Formula**: `Annual Dividend per Share / Current Price × 100`
- **Display**: Percentage (e.g., 2.5%)
- **Purpose**: Income return on investment

### 2. Financial Performance Metrics

#### Gross Profit
- **Source**: Income Statement
- **Calculation**: `Revenue - Cost of Goods Sold`
- **Display**: Billions (B)

#### Operating Income
- **Source**: Income Statement
- **Calculation**: `Gross Profit - Operating Expenses`
- **Also Known As**: EBIT (Earnings Before Interest and Taxes)
- **Display**: Billions (B)

#### Net Income
- **Source**: Income Statement
- **Calculation**: `Operating Income - Interest - Taxes`
- **Display**: Billions (B)
- **Purpose**: Bottom-line profitability

#### Debt-to-Equity Ratio
- **Formula**: `Total Debt / Total Equity`
- **Interpretation**:
  - < 1.0 = Low debt, conservative
  - > 2.0 = High debt, higher risk
- **Purpose**: Financial leverage indicator

#### Free Cash Flow
- **Source**: Cash Flow Statement
- **Calculation**: `Operating Cash Flow - Capital Expenditures`
- **Display**: Billions (B)
- **Purpose**: Cash available for dividends, buybacks, or growth

### 3. Growth Metrics

#### Year-over-Year (YoY) Growth
- **Revenue Growth YoY**: `(Current Year Revenue - Previous Year Revenue) / Previous Year Revenue × 100`
- **Earnings Growth YoY**: `(Current Year Earnings - Previous Year Earnings) / Previous Year Earnings × 100`
- **Source**: Annual financial statements
- **Display**: Percentage (e.g., 15.5%)

#### Quarter-over-Quarter (QoQ) Growth
- **Revenue Growth QoQ**: `(Current Quarter Revenue - Previous Quarter Revenue) / Previous Quarter Revenue × 100`
- **Earnings Growth QoQ**: `(Current Quarter Earnings - Previous Quarter Earnings) / Previous Quarter Earnings × 100`
- **Source**: Quarterly financial statements
- **Display**: Percentage (e.g., 3.2%)
- **Purpose**: Short-term trend indicator

### 4. Technical Indicators

#### RSI (Relative Strength Index) - 14 Day
- **Formula**: `100 - (100 / (1 + RS))` where `RS = Average Gain / Average Loss`
- **Range**: 0-100
- **Interpretation**:
  - **< 30**: Oversold (potential buy signal)
  - **30-70**: Neutral
  - **> 70**: Overbought (potential sell signal)
- **Purpose**: Momentum oscillator

#### MACD (Moving Average Convergence Divergence)
- **Components**:
  - **MACD Line**: `12-day EMA - 26-day EMA`
  - **Signal Line**: `9-day EMA of MACD`
  - **Histogram**: `MACD - Signal`
- **Interpretation**:
  - MACD > Signal = Bullish
  - MACD < Signal = Bearish
- **Purpose**: Trend-following momentum indicator

#### Simple Moving Averages (SMA)
- **SMA 20**: Average closing price over 20 days
- **SMA 50**: Average closing price over 50 days
- **SMA 200**: Average closing price over 200 days
- **Interpretation**:
  - Price > SMA = Uptrend
  - Price < SMA = Downtrend
  - Golden Cross: SMA 50 crosses above SMA 200 (bullish)
  - Death Cross: SMA 50 crosses below SMA 200 (bearish)

### 5. Analyst Data

#### Recommendation Breakdown
- **Strong Buy**: Number of analysts recommending strong buy
- **Buy**: Number of analysts recommending buy
- **Hold**: Number of analysts recommending hold
- **Sell**: Number of analysts recommending sell
- **Strong Sell**: Number of analysts recommending strong sell
- **Total Analysts**: Sum of all recommendations
- **Source**: Latest analyst recommendations from Yahoo Finance

#### Price Targets
- **Target Mean**: Average of all analyst price targets
- **Target High**: Highest analyst price target
- **Target Low**: Lowest analyst price target
- **Upside %**: `((Target Mean - Current Price) / Current Price) × 100`
- **Upside View**:
  - **High Upside**: ≥ 15%
  - **Moderate Upside**: 5-15%
  - **Limited/Downside**: < 5%

### 6. ESG Scores

#### ESG Components
- **ESG Total Score**: Overall ESG rating (0-100)
- **ESG Environment**: Environmental score (0-100)
- **ESG Social**: Social responsibility score (0-100)
- **ESG Governance**: Corporate governance score (0-100)
- **ESG Percentile**: Ranking compared to industry peers (0-100)
- **Source**: Yahoo Finance sustainability data

#### Manual ESG Override
- **Priority**: Manual ESG scores take precedence when provided
- **Multiple Themes**: One ticker can have multiple ESG theme rows
- **Display**: Stacked rows in Excel (duplicate ticker rows show only ESG columns)

### 7. Calculated Categories

#### Upside Bucket
- **High (>10%)**: Upside potential ≥ 10%
- **Medium (0-10%)**: Upside potential 0-10%
- **Negative**: Downside risk (negative upside)

#### ESG Category
- **Good (≥60)**: ESG score ≥ 60
- **Average (40-59)**: ESG score 40-59
- **Poor (<40)**: ESG score < 40

#### RSI Status
- **Overbought (>70)**: RSI > 70
- **Oversold (<30)**: RSI < 30
- **Neutral**: RSI between 30-70

---

## 🏗️ Architecture & Pipeline

### System Architecture

```
┌─────────────────┐
│  Excel File     │
│  Stock_data.xlsm│
│  - Sheet1       │
│  - Manual_ESG   │
└────────┬────────┘
         │
         ▼
┌─────────────────────────────────────┐
│  Data Fetching Module               │
│  (fetch_data.py)                    │
│  - Read tickers from Excel          │
│  - Fetch data via yfinance API      │
│  - Calculate technical indicators   │
└────────┬────────────────────────────┘
         │  
         ▼
┌─────────────────────────────────────┐
│  Data Processing                    │
│  - Merge automated + manual ESG     │
│  - Calculate growth metrics         │
│  - Add categorical buckets          │
└────────┬────────────────────────────┘
         │
         ▼
┌─────────────────────────────────────┐
│  Excel Update Module                 │
│  (update_excel.py)                   │
│  - Write to RawData sheet            │
│  - Apply formatting                  │
│  - Collapse duplicate rows           │
└────────┬─────────────────────────────┘
         │
         ▼
┌─────────────────┐
│  Formatted      │
│  Excel Dashboard│
│  (RawData sheet)│
└─────────────────┘
```

### Data Pipeline Flow

#### Step 1: Input Reading
1. System locates `Stock_data.xlsm` in project root
2. Reads ticker symbols from `Sheet1`, Column A
3. Reads manual ESG data from `Manual_ESG` sheet (if exists)
4. Falls back to default tickers if Excel file not found

#### Step 2: Data Fetching
For each ticker:
1. **Market Data**: Current price, PE ratio, market cap, dividend yield
2. **Financial Statements**: Income statement, balance sheet, cash flow statement
3. **Historical Prices**: 5 years of daily closing prices
4. **Analyst Data**: Recommendations and price targets
5. **ESG Data**: Sustainability scores from Yahoo Finance

#### Step 3: Calculations
1. **Technical Indicators**: RSI, MACD, SMA (20, 50, 200)
2. **Growth Metrics**: YoY and QoQ growth rates
3. **Debt Ratios**: Debt-to-equity calculation
4. **Upside Potential**: Percentage difference from analyst targets

#### Step 4: Data Merging
1. Combine automated stock data with manual ESG assessments
2. Use LEFT JOIN (keeps all stock data, adds ESG where available)
3. Handle multiple ESG themes per ticker (creates multiple rows)

#### Step 5: Enrichment
1. Add categorical buckets (Upside Bucket, ESG Category, RSI Status)
2. Format numeric values (rounding, percentage conversion)
3. Handle missing data (display as "N/A")

#### Step 6: Excel Output
1. Clear existing `RawData` sheet
2. Write DataFrame to Excel starting at A1
3. Apply black theme formatting
4. Collapse duplicate ticker rows (for ESG stacking)

---

## 📈 Output Format

### Excel Dashboard Structure

The `RawData` sheet contains the following columns (in order):

#### Identification
- **Ticker**: Stock symbol

#### Valuation
- **Current Price**: Latest market price
- **PE Ratio**: Price-to-earnings ratio
- **Market Cap**: Market capitalization (B or T)
- **Dividend Yield**: Annual dividend yield (%)

#### Financial Performance
- **Gross Profit**: Annual gross profit (B)
- **Operating Income**: Annual operating income (B)
- **Net Income**: Annual net income (B)

#### Balance Sheet
- **Total Cash**: Cash and cash equivalents (B)
- **Total Debt**: Total debt outstanding (B)
- **Debt to Equity**: Debt-to-equity ratio

#### Cash Flow
- **Free Cash Flow**: Annual free cash flow (B)
- **Operating Cash Flow**: Annual operating cash flow (B)

#### Growth Metrics
- **Earnings Growth YoY**: Year-over-year earnings growth (%)
- **Revenue Growth YoY**: Year-over-year revenue growth (%)
- **Earnings QoQ Growth**: Quarter-over-quarter earnings growth (%)
- **Revenue QoQ Growth**: Quarter-over-quarter revenue growth (%)

#### Technical Indicators
- **RSI (14)**: Relative Strength Index
- **SMA 20**: 20-day Simple Moving Average
- **SMA 50**: 50-day Simple Moving Average
- **SMA 200**: 200-day Simple Moving Average
- **MACD**: MACD line value
- **Signal Line**: MACD signal line value

#### Analyst Data
- **Strong Buy**: Count of strong buy recommendations
- **Buy**: Count of buy recommendations
- **Hold**: Count of hold recommendations
- **Sell**: Count of sell recommendations
- **Strong Sell**: Count of strong sell recommendations
- **Total Analysts (Breakdown)**: Total analyst count
- **Target Mean**: Average price target
- **Target High**: Highest price target
- **Target Low**: Lowest price target
- **Upside %**: Percentage upside potential
- **Upside View**: Categorical view (High/Moderate/Limited)

#### ESG Scores
- **ESG Total Score**: Overall ESG rating
- **ESG Environment**: Environmental score
- **ESG Social**: Social score
- **ESG Governance**: Governance score
- **ESG Percentile**: Industry percentile ranking

#### Manual ESG (if provided)
- **ESG Theme**: Custom ESG category
- **Manual ESG Score**: Custom score (0-100)
- **Confidence Level**: Assessment confidence
- **Assessment Criteria**: Evaluation method
- **Review Date**: Assessment date
- **Analyst Notes**: Additional comments

#### Calculated Categories
- **Upside Bucket**: Categorical upside classification
- **ESG Category**: Categorical ESG classification
- **RSI Status**: Categorical RSI classification

### Visual Formatting

#### Color Scheme
- **Background**: Black (#000000)
- **Text**: White (#FFFFFF)
- **Header**: Dark blue background (#2E75B5) with white bold text
- **Ticker Column**: Bright blue text (#4FC3F7), bold

#### Conditional Formatting
- **Positive Values** (Dividend Yield, Upside %): Green text (#4CAF50), bold
- **Negative Values**: Red text (#F44336), bold
- **RSI < 30** (Oversold): Blue text (#2196F3)
- **RSI > 70** (Overbought): Orange text (#FF9800)

#### Layout Features
- **Frozen Panes**: Header row and first 2 columns frozen
- **Auto-fit Columns**: Column widths adjust to content
- **Grid Borders**: Dark gray borders (#404040) on all cells
- **Centered Alignment**: All cells center-aligned

---

## 🚀 How to Use the System

### Prerequisites

1. **Excel File**: `Stock_data.xlsm` must exist in project root
2. **Excel Application**: Microsoft Excel must be installed and running
3. **Internet Connection**: Required for data fetching
4. **Python Environment**: Portable Python included in `python/` folder

### Execution Methods

#### Method 1: Batch File (Recommended)
Click on Button 1 in `Stock_data.xlsm` (Sheet1)

#### Method 2: Batch File
1. Double-click `run_update.bat`
2. System will automatically:
   - Navigate to project directory
   - Run Python script
   - Update Excel dashboard
   - Display completion message

#### Method 3: Direct Python Execution
1. Open Command Prompt or PowerShell
2. Navigate to project directory
3. Run: `python\python.exe Backend\update_excel.py`
4. Ensure Excel is open with `Stock_data.xlsm`

### Step-by-Step Workflow

1. **Prepare Input**:
   - Open `Stock_data.xlsm`
   - Add/update tickers in `Sheet1`, Column A
   - (Optional) Add manual ESG assessments in `Manual_ESG` sheet

2. **Run Update**:
   - Open `Stock_data.xlsm` (Sheet1)
   - Click `Button 1`
   - Wait for data fetching (may take 1-2 minutes for 10 ticker)
   - Monitor console output for progress

3. **Review Results**:
   - Check `RawData` sheet for updated data
   - Verify all tickers processed successfully
   - Review formatted dashboard

4. **Analyze Data**:
   - Use conditional formatting colors for quick insights
   - Compare metrics across tickers
   - Review ESG scores and analyst recommendations

### Expected Processing Time

- **10 Tickers**: ~1-2 minutes
- **50 Tickers**: ~5-10 minutes
- **100 Tickers**: ~10-15 minutes

*Note: Processing time depends on internet speed and API response times*

---

## ⚠️ Error Handling

### Common Issues & Solutions

#### Excel File Not Found
- **Symptom**: System uses default tickers (AAPL, MSFT, GOOGL, TSLA, AMZN)
- **Solution**: Ensure `Stock_data.xlsm` exists in project root directory

#### Excel Not Running
- **Symptom**: Error message "Please open Excel first"
- **Solution**: Open Microsoft Excel before running the script

#### Missing Data Fields
- **Symptom**: Some columns show "N/A"
- **Explanation**: Data may not be available for certain stocks (e.g., ESG scores for small caps)
- **Solution**: This is normal - system handles missing data gracefully

#### Internet Connection Issues
- **Symptom**: Timeout errors or failed fetches
- **Solution**: Check internet connection and retry

#### Manual_ESG Sheet Not Found
- **Symptom**: Warning message, but processing continues
- **Solution**: Create `Manual_ESG` sheet if you want to add custom ESG assessments

---

## 🔧 Customization Options

### Adding More Tickers
- **Method**: Edit `Sheet1`, Column A in Excel
- **No Code Changes**: System automatically reads all tickers

### Custom ESG Scoring
- **Method**: Use `Manual_ESG` sheet
- **Format**: Follow column structure described above
- **Multiple Themes**: Add multiple rows per ticker for different ESG themes

### Adjusting Time Periods
- **Current**: 5 years historical data
- **Modification**: Requires code change (contact developer)
- **Impact**: Affects technical indicator calculations

### Adding New Metrics
- **Method**: Requires code modification
- **Contact**: Developer for custom metric additions

---

## 📋 Data Quality & Limitations

### Data Accuracy
- **Source**: Yahoo Finance (public data)
- **Update Frequency**: Real-time during execution
- **Reliability**: High for major stocks, variable for small caps

### Limitations
- **API Rate Limits**: Yahoo Finance may throttle requests for many tickers
- **Data Availability**: Some metrics may not be available for all stocks
- **Historical Data**: Limited to what Yahoo Finance provides
- **ESG Scores**: Not available for all companies

### Best Practices
1. **Verify Critical Data**: Cross-check important metrics with official sources
2. **Regular Updates**: Run updates daily or weekly for current data
3. **Portfolio Size**: Optimal for 10-50 tickers (larger portfolios may take longer)
4. **Manual Review**: Always review ESG scores and analyst recommendations

---

## 📞 Support & Maintenance

### System Requirements
- **Operating System**: Windows (primary), macOS/Linux (with modifications)
- **Excel Version**: Microsoft Excel 2016 or later
- **Python Version**: 3.11+ (included in portable Python)
- **Storage**: Minimal (Excel file + Python scripts)

### Regular Maintenance
- **No Maintenance Required**: System is self-contained
- **Updates**: Check for yfinance library updates periodically
- **Data Refresh**: Run update script as needed

### Troubleshooting
- **Check Console Output**: Error messages provide diagnostic information
- **Verify Excel File**: Ensure `Stock_data.xlsm` structure is correct
- **Internet Connection**: Ensure stable connection during execution
- **Excel Permissions**: Ensure Excel file is not locked or read-only

---

## 📚 Glossary

- **Ticker**: Stock symbol (e.g., AAPL for Apple Inc.)
- **PE Ratio**: Price-to-Earnings ratio
- **Market Cap**: Market capitalization
- **RSI**: Relative Strength Index
- **MACD**: Moving Average Convergence Divergence
- **SMA**: Simple Moving Average
- **YoY**: Year-over-Year
- **QoQ**: Quarter-over-Quarter
- **ESG**: Environmental, Social, and Governance
- **EMA**: Exponential Moving Average
- **EBIT**: Earnings Before Interest and Taxes

---

*Last Updated: 12/01/2025*

