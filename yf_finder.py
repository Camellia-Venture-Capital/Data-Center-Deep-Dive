#!/usr/bin/env python3
"""
Interactive Data Center Financial Data Extractor
Allows users to select companies, time periods, and filing types for data extraction
"""

import yfinance as yf
import pandas as pd
import os
from datetime import datetime, timedelta
import logging
from typing import Dict, List, Optional
import time

# Set up logging
logging.basicConfig(
    filename='data_extraction.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class DataCenterExtractor:
    def __init__(self):
        """Initialize the Data Center Extractor"""
        # Create data directory
        self.data_dir = './financial_data'
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
        
        # Define filing types
        self.filing_types = {
            '10-K': 'Annual Report',
            '10-Q': 'Quarterly Report',
            '8-K': 'Current Report',
            'DEF 14A': 'Proxy Statement',
            'S-1': 'Registration Statement'
        }
    
    def search_company(self, query: str) -> Optional[Dict]:
        """Search for a company by name or ticker with enhanced information"""
        try:
            # Try to get company info directly
            company = yf.Ticker(query)
            info = company.info
            
            if info and 'symbol' in info:
                # Extract comprehensive company information
                company_data = {
                    'ticker': info.get('symbol', query.upper()),
                    'name': info.get('longName') or info.get('shortName') or query.upper(),
                    'sector': info.get('sector', 'N/A'),
                    'industry': info.get('industry', 'N/A'),
                    'country': info.get('country', 'N/A'),
                    'city': info.get('city', 'N/A'),
                    'website': info.get('website', 'N/A'),
                    'business_summary': info.get('longBusinessSummary', 'N/A')[:200] + '...' if info.get('longBusinessSummary') else 'N/A',
                    'market_cap': info.get('marketCap', 'N/A'),
                    'employees': info.get('fullTimeEmployees', 'N/A'),
                    'exchange': info.get('exchange', 'N/A')
                }
                
                # Format market cap for better display
                if isinstance(company_data['market_cap'], (int, float)):
                    if company_data['market_cap'] > 1e12:
                        company_data['market_cap_formatted'] = f"${company_data['market_cap']/1e12:.2f}T"
                    elif company_data['market_cap'] > 1e9:
                        company_data['market_cap_formatted'] = f"${company_data['market_cap']/1e9:.2f}B"
                    elif company_data['market_cap'] > 1e6:
                        company_data['market_cap_formatted'] = f"${company_data['market_cap']/1e6:.2f}M"
                    else:
                        company_data['market_cap_formatted'] = f"${company_data['market_cap']:,.0f}"
                else:
                    company_data['market_cap_formatted'] = 'N/A'
                
                # Format employee count
                if isinstance(company_data['employees'], (int, float)):
                    company_data['employees_formatted'] = f"{company_data['employees']:,}"
                else:
                    company_data['employees_formatted'] = 'N/A'
                
                return company_data
            
            return None
            
        except Exception as e:
            logging.error(f"Error searching for company {query}: {str(e)}")
            # Fallback to basic search
            try:
                company = yf.Ticker(query)
                info = company.info
                if info and info.get('symbol'):
                    return {
                        'ticker': info.get('symbol', query.upper()),
                        'name': info.get('shortName', query.upper()),
                        'sector': 'N/A',
                        'industry': 'N/A'
                    }
            except:
                pass
            return None
    
    def get_available_periods(self, ticker: str) -> Dict:
        """Get available time periods for a company with enhanced period detection"""
        try:
            company = yf.Ticker(ticker)
            
            # Get historical data to determine available periods
            hist = company.history(period="max")
            
            if hist.empty:
                return {}
            
            # Get unique years and quarters with more comprehensive analysis
            years = sorted(hist.index.year.unique(), reverse=True)
            quarters = sorted(hist.index.quarter.unique(), reverse=True)
            
            # Try to get actual financial statement dates
            try:
                # Get quarterly financials to see actual reporting periods
                quarterly_financials = company.quarterly_financials
                annual_financials = company.financials
                
                actual_quarters = []
                actual_years = []
                
                if not quarterly_financials.empty:
                    actual_quarters = sorted([col.year for col in quarterly_financials.columns], reverse=True)
                
                if not annual_financials.empty:
                    actual_years = sorted([col.year for col in annual_financials.columns], reverse=True)
                
                # Use actual financial periods if available
                if actual_years:
                    years = actual_years[:10]  # Last 10 years
                if actual_quarters:
                    years_with_quarters = sorted(list(set(actual_quarters)), reverse=True)[:5]  # Last 5 years with quarterly data
                else:
                    years_with_quarters = years[:5]
                    
            except Exception:
                # Fallback to historical price data years
                years_with_quarters = years[:5]
            
            return {
                'years': years[:10],  # Last 10 years for annual
                'years_with_quarters': years_with_quarters,  # Years with quarterly data
                'quarters': [1, 2, 3, 4],  # Standard quarters
                'earliest_date': hist.index.min(),
                'latest_date': hist.index.max(),
                'data_points': len(hist)
            }
            
        except Exception as e:
            logging.error(f"Error getting available periods for {ticker}: {str(e)}")
            return {}
    
    def get_company_data(self, ticker: str, start_date: str = None, end_date: str = None, filing_type: str = None) -> Optional[Dict]:
        """Get financial data for a company within specified period with enhanced error handling"""
        try:
            company = yf.Ticker(ticker)
            
            # Get basic info
            info = company.info
            
            # Initialize result dictionary
            result = {
                'info': info,
                'income_stmt': None,
                'balance_sheet': None,
                'cash_flow': None,
                'historical': None
            }
            
            # Get financial statements based on filing type
            try:
                if filing_type == 'quarterly':
                    result['income_stmt'] = company.quarterly_income_stmt
                    result['balance_sheet'] = company.quarterly_balance_sheet
                    result['cash_flow'] = company.quarterly_cashflow
                else:
                    # Default to annual
                    result['income_stmt'] = company.income_stmt
                    result['balance_sheet'] = company.balance_sheet
                    result['cash_flow'] = company.cashflow
                    
            except Exception as e:
                logging.warning(f"Error fetching financial statements for {ticker}: {str(e)}")
                # Continue with other data even if financials fail
            
            # Get historical data for the period
            try:
                if start_date and end_date:
                    result['historical'] = company.history(start=start_date, end=end_date)
                else:
                    result['historical'] = company.history(period="1y")
            except Exception as e:
                logging.warning(f"Error fetching historical data for {ticker}: {str(e)}")
            
            # Validate that we got at least some data
            has_data = any([
                result['income_stmt'] is not None and not result['income_stmt'].empty,
                result['balance_sheet'] is not None and not result['balance_sheet'].empty,
                result['cash_flow'] is not None and not result['cash_flow'].empty,
                result['historical'] is not None and not result['historical'].empty,
                result['info'] and len(result['info']) > 0
            ])
            
            if has_data:
                return result
            else:
                logging.warning(f"No data found for {ticker}")
                return None
                
        except Exception as e:
            logging.error(f"Error fetching data for {ticker}: {str(e)}")
            return None
    
    def save_company_data(self, ticker: str, data: Dict, filing_type: str = None):
        """Save company data to CSV files with enhanced file naming"""
        try:
            company_dir = os.path.join(self.data_dir, ticker)
            if not os.path.exists(company_dir):
                os.makedirs(company_dir)
            
            # Create timestamp for file naming
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filing_suffix = f"_{filing_type}" if filing_type else ""
            
            # Save financial statements with better error handling
            files_saved = []
            
            if data.get('income_stmt') is not None and not data['income_stmt'].empty:
                filename = f'income_statement{filing_suffix}_{timestamp}.csv'
                filepath = os.path.join(company_dir, filename)
                data['income_stmt'].to_csv(filepath)
                files_saved.append(filename)
            
            if data.get('balance_sheet') is not None and not data['balance_sheet'].empty:
                filename = f'balance_sheet{filing_suffix}_{timestamp}.csv'
                filepath = os.path.join(company_dir, filename)
                data['balance_sheet'].to_csv(filepath)
                files_saved.append(filename)
            
            if data.get('cash_flow') is not None and not data['cash_flow'].empty:
                filename = f'cash_flow{filing_suffix}_{timestamp}.csv'
                filepath = os.path.join(company_dir, filename)
                data['cash_flow'].to_csv(filepath)
                files_saved.append(filename)
            
            if data.get('historical') is not None and not data['historical'].empty:
                filename = f'historical_data_{timestamp}.csv'
                filepath = os.path.join(company_dir, filename)
                data['historical'].to_csv(filepath)
                files_saved.append(filename)
            
            # Save company info with enhanced formatting
            if data.get('info') and len(data['info']) > 0:
                filename = f'company_info_{timestamp}.csv'
                filepath = os.path.join(company_dir, filename)
                
                # Create a more readable company info file
                info_series = pd.Series(data['info'])
                # Filter out None values and convert problematic types
                clean_info = {}
                for key, value in info_series.items():
                    if value is not None:
                        try:
                            # Convert to string if it's a complex type
                            if isinstance(value, (list, dict)):
                                clean_info[key] = str(value)
                            else:
                                clean_info[key] = value
                        except:
                            clean_info[key] = str(value)
                
                pd.Series(clean_info).to_csv(filepath)
                files_saved.append(filename)
            
            if files_saved:
                logging.info(f"Successfully saved {len(files_saved)} files for {ticker}: {', '.join(files_saved)}")
                return files_saved
            else:
                logging.warning(f"No files saved for {ticker} - no valid data found")
                return []
            
        except Exception as e:
            logging.error(f"Error saving data for {ticker}: {str(e)}")
            return []
    
    def interactive_extraction(self):
        """Interactive data extraction process with enhanced user experience"""
        print("\n=== Enhanced Data Center Financial Data Extractor ===")
        print("🔍 Example inputs:")
        print("  Company: AAPL, MSFT, EQIX, DLR, GOOGL, TSLA")
        print("  Filing type: annual, quarterly") 
        print("  Date format: YYYY-MM-DD (e.g., 2023-01-01)")
        print("="*60)
        
        while True:
            # Get company input
            company_input = input("\n📊 Enter company name or ticker (or 'quit' to exit): ").strip()
            if company_input.lower() == 'quit':
                break
            
            # Search for company with enhanced feedback
            print(f"🔍 Searching for: {company_input}")
            company_info = self.search_company(company_input)
            
            if not company_info:
                print(f"❌ Company not found: {company_input}")
                print("💡 Try using the stock ticker symbol (e.g., AAPL for Apple)")
                continue
            
            # Display enhanced company information
            print(f"\n✅ Found company:")
            print(f"   📈 Name: {company_info['name']} ({company_info['ticker']})")
            print(f"   🏢 Sector: {company_info.get('sector', 'N/A')}")
            print(f"   🏭 Industry: {company_info.get('industry', 'N/A')}")
            if company_info.get('market_cap_formatted', 'N/A') != 'N/A':
                print(f"   💰 Market Cap: {company_info['market_cap_formatted']}")
            if company_info.get('employees_formatted', 'N/A') != 'N/A':
                print(f"   👥 Employees: {company_info['employees_formatted']}")
            
            # Get available periods with enhanced display
            periods = self.get_available_periods(company_info['ticker'])
            if not periods:
                print("❌ No historical data available")
                continue
            
            print(f"\n📅 Available time periods:")
            print(f"   📊 Data points: {periods.get('data_points', 'Unknown')}")
            print(f"   📈 Date range: {periods.get('earliest_date', 'Unknown').strftime('%Y-%m-%d') if periods.get('earliest_date') else 'Unknown'} to {periods.get('latest_date', 'Unknown').strftime('%Y-%m-%d') if periods.get('latest_date') else 'Unknown'}")
            
            if periods.get('years'):
                print(f"   🗓️  Annual data: {', '.join(map(str, periods['years'][:5]))}{'...' if len(periods['years']) > 5 else ''}")
            if periods.get('years_with_quarters'):
                print(f"   📈 Quarterly data: {', '.join(map(str, periods['years_with_quarters'][:3]))}{'...' if len(periods['years_with_quarters']) > 3 else ''}")
            
            # Get filing type with enhanced options
            print(f"\n📋 Select data type:")
            print("   1. Annual (yearly financial statements)")
            print("   2. Quarterly (quarterly financial statements)")
            
            while True:
                filing_choice = input("Enter choice (1-2): ").strip()
                if filing_choice == "1":
                    filing_type = "annual"
                    break
                elif filing_choice == "2":
                    filing_type = "quarterly"
                    break
                else:
                    print("❌ Invalid choice. Please enter 1 or 2.")
            
            # Get date range with smart defaults
            print(f"\n📅 Enter date range (YYYY-MM-DD format)")
            
            # Smart default dates based on filing type
            if filing_type == "quarterly":
                default_start = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")  # 1 year ago
                print(f"   💡 Suggestion: Last year for quarterly data")
            else:
                default_start = (datetime.now() - timedelta(days=365*3)).strftime("%Y-%m-%d")  # 3 years ago
                print(f"   💡 Suggestion: Last 3 years for annual data")
            
            while True:
                start_date = input(f"📅 Start date (press Enter for {default_start}): ").strip()
                if not start_date:
                    start_date = default_start
                    break
                try:
                    datetime.strptime(start_date, "%Y-%m-%d")
                    break
                except ValueError:
                    print("❌ Invalid date format. Please use YYYY-MM-DD (e.g., 2023-01-01)")
            
            while True:
                end_date = input("📅 End date (press Enter for today): ").strip()
                if not end_date:
                    end_date = datetime.now().strftime("%Y-%m-%d")
                    break
                try:
                    datetime.strptime(end_date, "%Y-%m-%d")
                    break
                except ValueError:
                    print("❌ Invalid date format. Please use YYYY-MM-DD (e.g., 2023-12-31)")
            
            # Confirm selection with enhanced summary
            print(f"\n📋 Extraction Summary:")
            print(f"   🏢 Company: {company_info['name']} ({company_info['ticker']})")
            print(f"   📊 Data type: {filing_type.title()}")
            print(f"   📅 Period: {start_date} to {end_date}")
            print(f"   🎯 Expected files: Income Statement, Balance Sheet, Cash Flow, Company Info, Historical Data")
            
            confirm = input("\n✅ Proceed with data extraction? (y/n): ").lower()
            if confirm != 'y':
                print("❌ Extraction cancelled.")
                continue
            
            # Get and save data with progress feedback
            print(f"\n🔄 Extracting data for {company_info['ticker']}...")
            print("   📥 Fetching financial statements...")
            
            data = self.get_company_data(
                company_info['ticker'],
                start_date,
                end_date,
                filing_type
            )
            
            if data:
                print("   💾 Saving files...")
                saved_files = self.save_company_data(company_info['ticker'], data, filing_type)
                
                if saved_files:
                    print(f"\n✅ Extraction completed successfully!")
                    print(f"   📁 Location: {os.path.join(self.data_dir, company_info['ticker'])}")
                    print(f"   📄 Files created: {len(saved_files)}")
                    for file in saved_files:
                        print(f"      • {file}")
                else:
                    print(f"\n⚠️  No files were saved (no valid data found)")
            else:
                print(f"\n❌ Failed to fetch data for {company_info['ticker']}")
                print("💡 This might be due to:")
                print("   • Invalid ticker symbol")
                print("   • No data available for the selected period")
                print("   • Temporary network issues")
            
            # Ask if user wants to continue
            if input(f"\n🔄 Extract data for another company? (y/n): ").lower() != 'y':
                break
        
        print(f"\n👋 Thank you for using Enhanced Financial Data Extractor!")


def main():
    """Main function with enhanced startup"""
    print("🚀 Enhanced Financial Data Extractor")
    print("=" * 50)
    print("📊 Extract comprehensive financial data from Yahoo Finance")
    print("🔍 Enhanced with detailed company information and smart defaults")
    print("=" * 50)
    
    extractor = DataCenterExtractor()
    extractor.interactive_extraction()


if __name__ == "__main__":
    main()