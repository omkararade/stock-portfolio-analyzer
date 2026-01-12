<div align="center">

# 📊 Stock Portfolio Analyzer

**Automated stock portfolio analysis tool that aggregates financial data, calculates technical indicators, and generates formatted Excel dashboards**

[![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
[![Excel](https://img.shields.io/badge/Excel-2016+-brightgreen.svg)](https://www.microsoft.com/microsoft-365/excel)
[![Status](https://img.shields.io/badge/Status-Active-success.svg)](https://github.com)

![Stock Analysis](https://img.shields.io/badge/Stock-Analysis-orange)
![Financial Data](https://img.shields.io/badge/Financial-Data-blue)
![Excel Dashboard](https://img.shields.io/badge/Excel-Dashboard-green)

</div>

---

## 🎯 Overview

**Stock Portfolio Analyzer** is a comprehensive Python-based financial data aggregation system that automatically collects, analyzes, and visualizes stock market data for investment portfolios. The tool integrates fundamental analysis, technical indicators, ESG ratings, and analyst recommendations into a unified, professionally formatted Excel dashboard.

### ✨ Key Features

- 📈 **Real-time Market Data** - Current prices, PE ratios, market cap, dividend yields
- 💰 **Financial Statements** - Income statements, balance sheets, cash flow analysis
- 📊 **Technical Indicators** - RSI, MACD, Moving Averages (20, 50, 200-day)
- 🌱 **ESG Integration** - Automated ESG scores + manual assessment override
- 👥 **Analyst Data** - Recommendations, price targets, upside potential
- 📉 **Growth Metrics** - Year-over-year and quarter-over-quarter analysis
- 🎨 **Excel Dashboard** - Professional formatting with conditional styling
- ⚡ **Automated Workflow** - One-click execution via batch file

---

## 🖼️ Screenshots

<div align="center">

### Dashboard Preview

![Dashboard](https://via.placeholder.com/800x400/2E75B6/FFFFFF?text=Excel+Dashboard+Preview)

*Professional Excel dashboard with black theme and conditional formatting*

### Data Categories

![Data](https://via.placeholder.com/600x300/4CAF50/FFFFFF?text=Comprehensive+Stock+Data)

*40+ data points per stock including fundamentals, technicals, and ESG scores*

</div>

---

## 🚀 Quick Start

### Prerequisites

- **Windows OS** (Excel automation requires Windows)
- **Microsoft Excel 2016+** installed and running
- **Internet Connection** for data fetching

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/stock-portfolio-analyzer.git
   cd stock-portfolio-analyzer
   ```

2. **No Python installation required!**
   - Portable Python distribution included in `python/` folder
   - All dependencies pre-installed

3. **Prepare your Excel file**
   - Ensure `Stock_data.xlsm` exists in project root
   - Add ticker symbols to `Sheet1`, Column A (one per row)

### Usage

#### Method 1: Batch File (Recommended)
```bash
# Double-click run_update.bat
# Or run from command line:
run_update.bat
```

#### Method 2: Direct Python Execution
```bash
python\python.exe Backend\update_excel.py
```

#### Method 3: Standalone Data Fetching
```bash
python\python.exe Backend\fetch_data.py
```

### Excel File Structure

**Sheet1** - Stock Tickers (Required)
```
Column A: Tickers
AAPL
MSFT
GOOGL
TSLA
```

**Manual_ESG** - Custom ESG Assessments (Optional)
```
Ticker | ESG Theme | Manual ESG Score | Confidence Level | ...
```

**RawData** - Output Sheet (Auto-generated)
- Contains all fetched and calculated data
- Automatically formatted with professional styling

---

## 📋 Features in Detail

### 📊 Data Collection

| Category | Metrics |
|----------|---------|
| **Valuation** | Current Price, PE Ratio, Market Cap, Dividend Yield |
| **Financials** | Gross Profit, Operating Income, Net Income, Cash Flow |
| **Balance Sheet** | Total Cash, Total Debt, Debt-to-Equity Ratio |
| **Growth** | Revenue/Earnings Growth (YoY & QoQ) |
| **Technical** | RSI, MACD, SMA (20/50/200), Signal Line |
| **Analyst** | Strong Buy/Buy/Hold/Sell counts, Price Targets |
| **ESG** | Total Score, Environment, Social, Governance, Percentile |

### 🔧 Technical Indicators

- **RSI (14-day)** - Momentum oscillator (0-100 scale)
  - < 30: Oversold (potential buy signal)
  - > 70: Overbought (potential sell signal)

- **MACD** - Trend-following momentum indicator
  - MACD Line: 12-day EMA - 26-day EMA
  - Signal Line: 9-day EMA of MACD
  - Histogram: MACD - Signal

- **Moving Averages** - Trend identification
  - SMA 20: Short-term trend
  - SMA 50: Medium-term trend
  - SMA 200: Long-term trend

### 🎨 Excel Formatting

- **Black Theme** - Professional dark background with white text
- **Conditional Formatting** - Color-coded values (green/red for positive/negative)
- **Frozen Panes** - Header row and first 2 columns always visible
- **Auto-fit Columns** - Optimal column widths
- **Grid Borders** - Clean, organized appearance

---

## 📚 Documentation

Comprehensive documentation is available in the `Documentation/` folder:

- **[Client Documentation](Documentation/Client/Client.md)** - User guide with parameters, calculations, and usage instructions
- **[Developer Documentation](Documentation/Developer/Developer.md)** - Technical implementation details, API integration, and extension points

### Quick Links

- [Installation Guide](Documentation/Client/Client.md#-how-to-use-the-system)
- [Data Categories](Documentation/Client/Client.md#-data-categories-collected)
- [Calculations Explained](Documentation/Client/Client.md#-calculations--metrics-explained)
- [API Integration](Documentation/Developer/Developer.md#-api-integration)
- [Extension Points](Documentation/Developer/Developer.md#-extension-points)

---

## 🛠️ Technology Stack

<div align="center">

| Technology | Purpose |
|------------|---------|
| ![Python](https://img.shields.io/badge/Python-3.11+-3776AB?logo=python&logoColor=white) | Core programming language |
| ![Pandas](https://img.shields.io/badge/Pandas-2.3+-150458?logo=pandas&logoColor=white) | Data manipulation and analysis |
| ![yfinance](https://img.shields.io/badge/yfinance-0.2.66-FF6B6B?logo=yahoo&logoColor=white) | Yahoo Finance API integration |
| ![xlwings](https://img.shields.io/badge/xlwings-0.33.16-217346?logo=excel&logoColor=white) | Excel automation (Windows) |
| ![openpyxl](https://img.shields.io/badge/openpyxl-3.1.5-2E7D32?logo=excel&logoColor=white) | Excel file operations |

</div>

### Dependencies

```
pandas>=2.3.3
yfinance>=0.2.66
openpyxl>=3.1.5
xlwings>=0.33.16
```

---

## 📖 Usage Examples

### Basic Usage

```python
# Add tickers to Stock_data.xlsm, Sheet1, Column A
# Run: run_update.bat
# View results in RawData sheet
```

### Custom ESG Assessment

1. Create `Manual_ESG` sheet in Excel
2. Add columns: Ticker, ESG Theme, Manual ESG Score, Confidence Level, etc.
3. Run update script
4. ESG data will be merged with automated data

### Multiple ESG Themes

Add multiple rows per ticker for different ESG themes:
```
Ticker | ESG Theme          | Manual ESG Score
AAPL   | Climate Change     | 85
AAPL   | Labor Practices    | 78
AAPL   | Data Privacy       | 82
```

---

## 🔍 Data Sources

- **Yahoo Finance API** (via `yfinance`) - Primary data source
  - Market data, financial statements, analyst recommendations
  - ESG scores and sustainability metrics
  - Historical price data (5 years)

- **Excel File** - Input/output interface
  - Ticker list input
  - Manual ESG assessments
  - Formatted dashboard output

---

## ⚙️ Configuration

### Adjusting Time Period

Modify in `Backend/fetch_data.py`:
```python
hist = stock.history(period="5y")  # Change to "1y", "2y", "10y", etc.
```

### Adding Custom Metrics

1. Fetch data in `fetch_stock_data_with_indicators()`
2. Add to data dictionary
3. (Optional) Add formatting in `update_excel.py`

See [Developer Documentation](Documentation/Developer/Developer.md#-extension-points) for details.

---

## 🐛 Troubleshooting

### Common Issues

**Excel file not found**
- Ensure `Stock_data.xlsm` exists in project root
- System will use default tickers if file missing

**Excel not running**
- Open Microsoft Excel before running script
- Error: "Please open Excel first"

**Missing data fields**
- Some metrics may show "N/A" for certain stocks
- Normal behavior - data not always available

**Internet connection issues**
- Check internet connection
- Yahoo Finance API may be temporarily unavailable

See [Error Handling](Documentation/Client/Client.md#-error-handling) for more details.

---

## 📊 Performance

- **Processing Time**: ~30-60 seconds per ticker
- **10 Tickers**: ~5-10 minutes
- **50 Tickers**: ~25-50 minutes

*Processing time depends on internet speed and API response times*

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Setup

```bash
# Install dependencies (if using system Python)
pip install -r requirements.txt

# Run tests
python Backend/test.py
```

---

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

- [yfinance](https://github.com/ranaroussi/yfinance) - Yahoo Finance API wrapper
- [pandas](https://pandas.pydata.org/) - Data analysis library
- [xlwings](https://www.xlwings.org/) - Excel automation library
- [Yahoo Finance](https://finance.yahoo.com/) - Financial data source

---

## 📞 Support

- 📖 [Documentation](Documentation/)
- 🐛 [Report Issues](https://github.com/yourusername/stock-portfolio-analyzer/issues)
- 💬 [Discussions](https://github.com/yourusername/stock-portfolio-analyzer/discussions)

---

## ⭐ Star History

[![Star History Chart](https://api.star-history.com/svg?repos=yourusername/stock-portfolio-analyzer&type=Date)](https://star-history.com/#yourusername/stock-portfolio-analyzer&Date)

---

<div align="center">

**Made with ❤️ for investors and financial analysts**

[⬆ Back to Top](#-stock-portfolio-analyzer)

</div>

