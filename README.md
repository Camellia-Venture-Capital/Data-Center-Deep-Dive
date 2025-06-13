# 🚀 Financial Data Extractor

A professional-grade financial data extraction tool that combines SEC Edgar and Yahoo Finance data sources with an intuitive GUI interface. Features automatic Excel sheet processing, organized folder structures, and comprehensive preview capabilities.

## 📊 Overview

This tool provides seamless access to financial data from two major sources:

- **SEC Edgar**: Official regulatory filings (10-Q, 10-K, 8-K) with Excel and HTML formats
- **Yahoo Finance**: Historical financial statements and market data in CSV format

## ✨ Key Features

### 🏢 Dual Data Sources

- **SEC Edgar**: Official SEC documents, Excel reports, HTML filings
- **Yahoo Finance**: Historical financial data with easy CSV export

### 📊 Advanced Excel Processing

- **Automatic Sheet Selection**: Consolidated financial statements are auto-selected
- **CSV Export**: Excel sheets automatically exported as individual CSV files
- **Smart Fallback**: Selects main financial statements if no consolidated sheets found

### 📁 Organized File Structure

```
data/
├── TICKER/
│   ├── FORM-TYPE/          # For SEC data
│   │   └── YEAR/
│   │       ├── Financial_Report.xlsx
│   │       └── *_Consolidated_*.csv
│   └── PERIOD-TYPE/        # For Yahoo Finance data
│       └── YEAR/
│           ├── Income_Statement.csv
│           ├── Balance_Sheet.csv
│           └── Cash_Flow.csv
```

## 🛠️ Installation & Setup

### Prerequisites
- Python 3.7+
- Required packages (install via requirements.txt)

### Quick Start
1. Clone the repository
2. Install dependencies: `pip install -r requirements.txt`
3. Run the application: `python data_extractor.py`
4. Use the GUI to select data sources and configure extraction parameters

## 📈 Usage Guide

### SEC Edgar Data Extraction
1. Enter company ticker symbol
2. Select filing type (10-Q, 10-K, 8-K)
3. Choose date range
4. Click "Extract SEC Data"
5. Review preview and export options

### Yahoo Finance Data Extraction
1. Enter company ticker symbol
2. Select statement type (Income, Balance Sheet, Cash Flow)
3. Choose time period (Annual/Quarterly)
4. Click "Extract Yahoo Data"
5. Download CSV files
