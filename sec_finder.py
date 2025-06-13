#!/usr/bin/env python3
"""
Financial Report Finder - Interactive CLI Tool with Download Functionality
Run with: python finder.py
"""

import requests
import json
import re
from datetime import datetime
import sys
import os
from pathlib import Path
import time
import io
import pandas as pd
import openpyxl
import io
from sec_api import RenderApi

class FinancialReportFinder:
    def __init__(self, sec_api_key=None):
        """Initialize the Financial Report Finder with optional SEC API integration"""
        self.headers = {
            'User-Agent': 'Financial Report Finder zhengdingnan@gmail.com',
            'Accept-Encoding': 'gzip, deflate',
            'Accept': 'application/json, text/html, */*'
        }
        self.pandas_available = True
        self.sec_api_key = sec_api_key
        self.render_api = None
        
        
        
        # Try to initialize SEC API if key provided
        if sec_api_key:
            try:
                from sec_api import RenderApi
                self.render_api = RenderApi(api_key=sec_api_key)
                print("✅ SEC API initialized for enhanced downloading")
            except ImportError:
                print("⚠️  sec_api not installed. Install with: pip install sec-api")
                print("   Using basic download method instead")
            except Exception as e:
                print(f"⚠️  SEC API initialization failed: {e}")
                print("   Using basic download method instead")
    
    def get_company_cik(self, ticker_or_name):
        """
        Get company CIK from ticker symbol or company name using multiple methods
        
        Args:
            ticker_or_name (str): Stock ticker or company name
            
        Returns:
            str: CIK number (padded to 10 digits)
        """
        print(f"🔍 Searching for: {ticker_or_name}")
        
        # Method 1: Try the company tickers JSON (primary method)
        try:
            url = "https://www.sec.gov/files/company_tickers.json"
            response = requests.get(url, headers=self.headers, timeout=10)
            response.raise_for_status()
            
            companies = response.json()
            search_term = ticker_or_name.upper().strip()
            
            for company_data in companies.values():
                ticker = company_data.get('ticker', '').upper()
                title = company_data.get('title', '').upper()
                
                if search_term == ticker or search_term in title:
                    cik = str(company_data['cik_str']).zfill(10)
                    print(f"✅ Found: {company_data['title']} ({company_data['ticker']}) - CIK: {cik}")
                    return cik
            
        except Exception as e:
            print(f"⚠️  Primary search method failed, trying alternatives...")
        
        # Method 2: Try the company tickers exchange JSON (alternative endpoint)
        try:
            url = "https://www.sec.gov/files/company_tickers_exchange.json"
            response = requests.get(url, headers=self.headers, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            search_term = ticker_or_name.upper().strip()
            
            for field_name, companies in data['fields'].items():
                for company in companies:
                    if len(company) >= 3:
                        ticker = str(company[0]).upper() if company[0] else ""
                        title = str(company[1]).upper() if company[1] else ""
                        cik_val = company[2] if len(company) > 2 else None
                        
                        if search_term == ticker or search_term in title:
                            cik = str(cik_val).zfill(10)
                            print(f"✅ Found: {company[1]} ({company[0]}) - CIK: {cik}")
                            return cik
            
        except Exception:
            pass
        
        # Method 3: Hardcoded popular tickers (fallback)
        popular_tickers = {
            'AAPL': '0000320193',    # Apple Inc
            'MSFT': '0000789019',    # Microsoft Corporation
            'GOOGL': '0001652044',   # Alphabet Inc
            'GOOG': '0001652044',    # Alphabet Inc
            'AMZN': '0001018724',    # Amazon.com Inc
            'TSLA': '0001318605',    # Tesla Inc
            'META': '0001326801',    # Meta Platforms Inc
            'NVDA': '0001045810',    # NVIDIA Corporation
            'NFLX': '0001065280',    # Netflix Inc
            'EQIX': '0001101239',    # Equinix Inc
            'CRM': '0001108524',     # Salesforce Inc
            'ORCL': '0001341439',    # Oracle Corporation
            'IBM': '0000051143',     # International Business Machines
            'INTC': '0000050863',    # Intel Corporation
            'AMD': '0000002488',     # Advanced Micro Devices
            'UBER': '0001543151',    # Uber Technologies Inc
            'SPOT': '0001639920',    # Spotify Technology SA
            'PYPL': '0001633917',    # PayPal Holdings Inc
            'DIS': '0001001039',     # Walt Disney Company
            'KO': '0000021344',      # Coca-Cola Company
            'PEP': '0000077476',     # PepsiCo Inc
            'WMT': '0000104169',     # Walmart Inc
            'JPM': '0000019617',     # JPMorgan Chase & Co
            'BAC': '0000070858',     # Bank of America Corporation
            'V': '0001403161',       # Visa Inc
            'MA': '0001141391',      # Mastercard Incorporated
            'IRM': '0001020569',     # Iron Mountain Inc
            'DLR': '0001558370',     # Digital Realty Trust Inc
            'AMT': '0001065280',     # American Tower Corporation
            'VRT': '0001065280',     # Vertiv Holdings Inc
            'LUMN': '0001065280',     # Luminex Corporation
            'SWCH': '0001065280',     # Switch Inc
            'MSFT': '0000789019',     # Microsoft Corporation
            'NVDA': '0001045810',     # NVIDIA Corporation
            'AVGO': '0001065280',     # Broadcom Inc
            'ORCL': '0001341439',     # Oracle Corporation
            'GDS': '0001065280',     # Global Data Centers Inc  
            'EQIX': '0001101239',     # Equinix Inc
            'CRM': '0001108524',     # Salesforce Inc
            'IBM': '0000051143',     # International Business Machines
            'INTC': '0000050863',    # Intel Corporation
            'AMD': '0000002488',     # Advanced Micro Devices
            'UBER': '0001543151',    # Uber Technologies Inc
            'SPOT': '0001639920',    # Spotify Technology SA
            'PYPL': '0001633917',    # PayPal Holdings Inc
        }
        
        ticker_upper = ticker_or_name.upper().strip()
        if ticker_upper in popular_tickers:
            cik = popular_tickers[ticker_upper]
            print(f"✅ Found via database: {ticker_upper} - CIK: {cik}")
            return cik
        
        print(f"❌ Company not found: {ticker_or_name}")
        return None
    
    def get_company_info_from_cik(self, cik):
        try:
            # Convert CIK to proper format (10-digit string with leading zeros)
            if isinstance(cik, str):
                cik = cik.strip()
            cik_formatted = f"{int(cik):010d}"
            
            # SEC company tickers endpoint
            url = "https://www.sec.gov/files/company_tickers.json"
            headers = {'User-Agent': 'zhengdingnan@gmail.com'}
            
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            companies = response.json()
            
            # Search for the CIK in the data
            for company_data in companies.values():
                if f"{company_data['cik_str']:010d}" == cik_formatted:
                    return {
                        'cik': company_data['cik_str'],
                        'ticker': company_data['ticker'],
                        'title': company_data['title']
                    }
            
        except Exception as e:
            print(f"Error fetching company info for CIK {cik}: {e}")
            return None
        
    def get_recent_filings(self, cik, form_type="10-Q", count=6):
        """Get recent filings for a company"""
        try:
            print(f"📋 Getting recent {form_type} filings...")
            url = f"https://data.sec.gov/submissions/CIK{cik}.json"
            response = requests.get(url, headers=self.headers, timeout=15)
            response.raise_for_status()
            
            data = response.json()
            filings = data['filings']['recent']
            
            # Filter by form type
            filtered_filings = []
            for i, form in enumerate(filings['form']):
                if form == form_type and len(filtered_filings) < count:
                    filing_info = {
                        'form': form,
                        'filingDate': filings['filingDate'][i],
                        'accessionNumber': filings['accessionNumber'][i],
                        'reportDate': filings['reportDate'][i],
                        'primaryDocument': filings['primaryDocument'][i]
                    }
                    filtered_filings.append(filing_info)
            
            return filtered_filings
            
        except Exception as e:
            print(f"❌ Error getting filings: {e}")
            return []
    
    def generate_financial_report_urls(self, cik, accession_number):
        """Generate URLs for financial reports"""
        clean_accession = accession_number.replace('-', '')
        base_url = f"https://www.sec.gov/Archives/edgar/data/{cik}/{clean_accession}"
        
        urls = {
            'Excel Financial Report': f"{base_url}/Financial_Report.xlsx",
            'Income Statement (HTML)': f"{base_url}/R4.htm",
            'Balance Sheet (HTML)': f"{base_url}/R2.htm",
            'Balance Sheet Parenthetical (HTML)': f"{base_url}/R3.htm",
            'Cash Flow Statement (HTML)': f"{base_url}/R7.htm",
            'Stockholder Equity (HTML)': f"{base_url}/R6.htm"
        }
        
        return urls
    
    def create_download_directory(self, ticker, form_type):
        """Create directory structure for downloads"""
        base_dir = Path("sec-data")
        download_dir = base_dir / ticker.upper() / form_type.upper()
        download_dir.mkdir(parents=True, exist_ok=True)
        return download_dir
    
    def get_safe_filename(self, url, filing_date, report_date, report_type, ticker):
        """Generate safe filename for downloaded files with ticker and end date"""
        # Extract file extension
        if 'Financial_Report.xlsx' in url:
            ext = '.xlsx'
            base_name = 'Financial_Report'
        elif url.endswith('.htm'):
            ext = '.htm'
            # Map report types to readable names
            type_mapping = {
                'Income Statement (HTML)': 'Income_Statement',
                'Balance Sheet (HTML)': 'Balance_Sheet',
                'Balance Sheet Parenthetical (HTML)': 'Balance_Sheet_Parenthetical',
                'Cash Flow Statement (HTML)': 'Cash_Flow_Statement',
                'Stockholder Equity (HTML)': 'Stockholder_Equity'
            }
            base_name = type_mapping.get(report_type, 'Financial_Report')
        else:
            ext = '.html'
            base_name = 'Financial_Report'
        
        # Use report date (end date) instead of filing date for filename
        # Format: TICKER-report_type-YYYYMMDD.ext
        safe_end_date = report_date.replace('-', '')
        filename = f"{ticker.upper()}-{base_name}-{safe_end_date}{ext}"
        return filename
    
    def download_file_basic(self, url, filepath):
        """Download file using basic requests (fallback method)"""
        try:
            response = requests.get(url, headers=self.headers, timeout=30)
            response.raise_for_status()
            
            with open(filepath, 'wb') as f:
                f.write(response.content)
            return True
        except Exception as e:
            print(f"❌ Basic download failed: {e}")
            return False
    
    def download_file_sec_api(self, url, filepath):
        """Download file using SEC API (enhanced method)"""
        try:
            if not self.render_api:
                return False
            
            # Use SEC API for better reliability
            file_content = self.render_api.get_file(url, return_binary=True)
            
            with open(filepath, 'wb') as f:
                f.write(file_content)
            return True
        except Exception as e:
            print(f"❌ SEC API download failed: {e}")
            return False
    
    def analyze_excel_file(self, file_content):
        """Analyze Excel file and return sheet information"""
        if not self.pandas_available:
            return None
        
        try:
            import pandas as pd
            import openpyxl
            
            # Read Excel file from bytes
            excel_file = io.BytesIO(file_content)
            
            # Get all sheet names
            workbook = openpyxl.load_workbook(excel_file, read_only=True)
            sheet_names = workbook.sheetnames
            workbook.close()
            
            # Analyze each sheet
            excel_file.seek(0)  # Reset file pointer
            sheet_info = {}
            
            for sheet_name in sheet_names:
                try:
                    # Read just the first few rows to get an idea of content
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=5)
                    
                    sheet_info[sheet_name] = {
                        'columns': list(df.columns),
                        'shape': f"{len(df.index)}+ rows x {len(df.columns)} columns",
                        'sample_data': df.head(2).to_dict('records') if not df.empty else []
                    }
                except Exception as e:
                    sheet_info[sheet_name] = {
                        'columns': [],
                        'shape': 'Unable to read',
                        'sample_data': [],
                        'error': str(e)
                    }
            
            return sheet_info
            
        except Exception as e:
            print(f"❌ Error analyzing Excel file: {e}")
            return None
    
    def select_excel_sheets_to_export(self, sheet_info):
        """Let user select which Excel sheets to export as CSV"""
        if not sheet_info:
            return []
        
        print(f"\n📊 EXCEL FILE ANALYSIS:")
        print("="*80)
        
        sheet_list = list(sheet_info.keys())
        auto_selected = []
        
        for i, (sheet_name, info) in enumerate(sheet_info.items(), 1):
            # Check if sheet starts with "consolidated" (case insensitive)
            is_consolidated = 'consolidated' in sheet_name.lower()

            status_icon = "🟢 AUTO-SELECTED" if is_consolidated else ""
            
            print(f"\n{i}. Sheet: '{sheet_name}' {status_icon}")
            print(f"   Size: {info['shape']}")
            
            if info.get('columns'):
                print(f"   Columns: {', '.join(info['columns'][:5])}")
                if len(info['columns']) > 5:
                    print(f"            ... and {len(info['columns']) - 5} more")
            
            if info.get('error'):
                print(f"   ⚠️  Error: {info['error']}")
            
            # Auto-select consolidated sheets
            if is_consolidated:
                auto_selected.append(sheet_name)
        
        # Show auto-selected sheets
        if auto_selected:
            print(f"\n🟢 AUTO-SELECTED CONSOLIDATED SHEETS:")
            for sheet in auto_selected:
                print(f"   └─ '{sheet}'")
        
        print(f"\n📋 SELECT ADDITIONAL SHEETS TO EXPORT AS CSV:")
        print("Enter numbers separated by commas (e.g., 1,3,5)")
        print("Enter 'auto' to only download financial statements")
        print("Enter 'all' for all sheets")
        print("Enter 'skip' to skip all CSV exports (including auto-selected)")
        
        while True:
            selection = input("Selection: ").strip().lower()
            
            if selection == 'skip':
                return []
            elif selection == 'auto':
                return auto_selected
            elif selection == 'all':
                return sheet_list
            
            try:
                # Parse comma-separated numbers
                indices = [int(x.strip()) - 1 for x in selection.split(',')]
                selected_sheets = auto_selected.copy()  # Start with auto-selected
                
                for idx in indices:
                    if 0 <= idx < len(sheet_list):
                        sheet_name = sheet_list[idx]
                        if sheet_name not in selected_sheets:  # Avoid duplicates
                            selected_sheets.append(sheet_name)
                    else:
                        print(f"❌ Invalid selection: {idx + 1}")
                        continue
                
                if selected_sheets:
                    return selected_sheets
                else:
                    print("❌ No valid selections made")
                    
            except ValueError:
                print("❌ Invalid input. Use numbers separated by commas, 'all', 'auto', or 'skip'")
    
    def export_excel_sheets_to_csv(self, file_content, selected_sheets, download_dir, ticker, report_date):
        """Export selected Excel sheets to CSV files"""
        if not self.pandas_available or not selected_sheets:
            return 0
        
        try:
            import pandas as pd
            
            excel_file = io.BytesIO(file_content)
            exported_count = 0
            
            print(f"\n📤 EXPORTING SHEETS TO CSV:")
            
            for sheet_name in selected_sheets:
                try:
                    # Read the entire sheet
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    
                    # Generate CSV filename
                    safe_sheet_name = re.sub(r'[<>:"/\\|?*]', '_', sheet_name)  # Replace invalid chars
                    safe_end_date = report_date.replace('-', '')
                    csv_filename = f"{ticker.upper()}-{safe_sheet_name}-{safe_end_date}.csv"
                    csv_filepath = download_dir / csv_filename
                    
                    # Export to CSV
                    df.to_csv(csv_filepath, index=False)
                    
                    file_size = os.path.getsize(csv_filepath) / 1024  # Size in KB
                    print(f"   ✅ Exported '{sheet_name}' to {csv_filename} ({file_size:.1f} KB)")
                    exported_count += 1
                    
                except Exception as e:
                    print(f"   ❌ Failed to export '{sheet_name}': {e}")
            
            print(f"\n   📊 CSV Export Summary: {exported_count}/{len(selected_sheets)} sheets exported")
            return exported_count
            
        except Exception as e:
            print(f"❌ Error during CSV export: {e}")
            return 0
    
    def download_and_process_excel(self, url, download_dir, filing_date, report_date, report_type, ticker):
        """Download Excel file and optionally export sheets to CSV"""
        filename = self.get_safe_filename(url, filing_date, report_date, report_type, ticker)
        filepath = download_dir / filename
        
        print(f"   📥 Downloading {report_type}...")
        print(f"      └─ File: {filename}")
        
        # Download the Excel file first
        file_content = None
        success = False
        
        if self.render_api:
            print(f"      └─ Using SEC API...")
            try:
                file_content = self.render_api.get_file(url, return_binary=True)
                success = True
            except Exception as e:
                print(f"❌ SEC API download failed: {e}")
        
        if not success:
            print(f"      └─ Using basic download...")
            try:
                response = requests.get(url, headers=self.headers, timeout=30)
                response.raise_for_status()
                file_content = response.content
                success = True
            except Exception as e:
                print(f"❌ Basic download failed: {e}")
        
        if not success or not file_content:
            print(f"      ❌ Download failed")
            return False
        
        # Save the Excel file
        with open(filepath, 'wb') as f:
            f.write(file_content)
        
        file_size = os.path.getsize(filepath) / 1024  # Size in KB
        print(f"      ✅ Downloaded successfully ({file_size:.1f} KB)")
        
        # Analyze Excel file and offer CSV export
        if 'Excel' in report_type and self.pandas_available:
            print(f"      🔍 Analyzing Excel file structure...")
            sheet_info = self.analyze_excel_file(file_content)
            
            if sheet_info:
                selected_sheets = self.select_excel_sheets_to_export(sheet_info)
                if selected_sheets:
                    self.export_excel_sheets_to_csv(file_content, selected_sheets, download_dir, ticker, report_date)
            else:
                print(f"      ⚠️  Could not analyze Excel file structure")
        
        return True
    
    def download_financial_report(self, url, download_dir, filing_date, report_date, report_type, ticker):
        """Download a single financial report with multiple methods"""
        
        # Handle Excel files specially (with sheet analysis and CSV export)
        if 'Excel' in report_type:
            return self.download_and_process_excel(url, download_dir, filing_date, report_date, report_type, ticker)
        
        # Handle regular HTML files
        filename = self.get_safe_filename(url, filing_date, report_date, report_type, ticker)
        filepath = download_dir / filename
        
        print(f"   📥 Downloading {report_type}...")
        print(f"      └─ File: {filename}")
        
        # Try SEC API first (if available), then fallback to basic method
        success = False
        
        if self.render_api:
            print(f"      └─ Using SEC API...")
            success = self.download_file_sec_api(url, filepath)
        
        if not success:
            print(f"      └─ Using basic download...")
            success = self.download_file_basic(url, filepath)
        
        if success:
            file_size = os.path.getsize(filepath) / 1024  # Size in KB
            print(f"      ✅ Downloaded successfully ({file_size:.1f} KB)")
            return True
        else:
            print(f"      ❌ Download failed")
            return False
    
    def download_filing_reports(self, filing, urls, ticker, form_type):
        """Download all reports for a specific filing"""
        filing_date = filing['filingDate']
        report_date = filing['reportDate']
        accession = filing['accessionNumber']
        
        print(f"\n📥 DOWNLOADING REPORTS")
        print(f"   Filing Date: {filing_date}")
        print(f"   Report Date (End Date): {report_date}")
        print(f"   Accession: {accession}")
        
        # Create download directory
        download_dir = self.create_download_directory(ticker, form_type)
        print(f"   Directory: {download_dir}")
        
        # Track download results
        downloaded = 0
        total = len(urls)
        
        # Download each report
        for report_type, url in urls.items():
            try:
                if self.download_financial_report(url, download_dir, filing_date, report_date, report_type, ticker):
                    downloaded += 1
                time.sleep(1)  # Be respectful with requests
            except Exception as e:
                print(f"      ❌ Error downloading {report_type}: {e}")
        
        print(f"\n   📊 Download Summary: {downloaded}/{total} files downloaded")
        print(f"   📁 Files saved to: {download_dir}")
        
        return downloaded > 0
    
    def ask_download_preference(self):
        """Ask user about download preferences"""
        print(f"\n💾 DOWNLOAD OPTIONS:")
        print("1. Download all reports for selected filing")
        print("2. Download specific report types")
        print("3. View links only (no download)")
        
        while True:
            choice = input("\nEnter choice (1-3): ").strip()
            if choice in ['1', '2', '3']:
                return choice
            else:
                print("❌ Invalid choice. Please enter 1, 2, or 3.")
    
    def select_specific_reports(self, urls):
        """Let user select specific reports to download"""
        print(f"\n📋 SELECT REPORTS TO DOWNLOAD:")
        url_list = list(urls.items())
        
        for i, (report_type, url) in enumerate(url_list, 1):
            print(f"{i}. {report_type}")
        
        print("Enter numbers separated by commas (e.g., 1,3,5) or 'all' for all reports:")
        
        while True:
            selection = input("Selection: ").strip().lower()
            
            if selection == 'all':
                return urls
            
            try:
                # Parse comma-separated numbers
                indices = [int(x.strip()) - 1 for x in selection.split(',')]
                selected_urls = {}
                
                for idx in indices:
                    if 0 <= idx < len(url_list):
                        report_type, url = url_list[idx]
                        selected_urls[report_type] = url
                    else:
                        print(f"❌ Invalid selection: {idx + 1}")
                        continue
                
                if selected_urls:
                    return selected_urls
                else:
                    print("❌ No valid selections made")
                    
            except ValueError:
                print("❌ Invalid input. Use numbers separated by commas or 'all'")
    
    def select_filing_to_download(self, filings):
        """Let user select which filing to download"""
        if len(filings) == 1:
            return 0
        
        print(f"\n📋 SELECT FILING TO DOWNLOAD:")
        for i, filing in enumerate(filings, 1):
            print(f"{i}. Filing Date: {filing['filingDate']} | Report Date: {filing['reportDate']}")
        
        while True:
            try:
                choice = input(f"\nEnter filing number (1-{len(filings)}): ").strip()
                idx = int(choice) - 1
                if 0 <= idx < len(filings):
                    return idx
                else:
                    print(f"❌ Invalid choice. Please enter 1-{len(filings)}")
            except ValueError:
                print("❌ Invalid input. Please enter a number.")
    
    def search_company(self, company_input):
        """Main search function"""
        print(f"\n{'='*80}")
        print(f"🏢 SEARCHING FOR: {company_input.upper()}")
        print(f"{'='*80}")
        
        # Get CIK
        cik = self.get_company_cik(company_input)
        if not cik:
            return False
        
        # Ask user for filing type
        print(f"\n📊 Select filing type:")
        print("1. 10-Q (Quarterly Reports)")
        print("2. 10-K (Annual Reports)")
        print("3. 8-K (Current Reports)")
        
        while True:
            choice = input("\nEnter choice (1-3) or press Enter for 10-Q: ").strip()
            if choice == "" or choice == "1":
                form_type = "10-Q"
                break
            elif choice == "2":
                form_type = "10-K"
                break
            elif choice == "3":
                form_type = "8-K"
                break
            else:
                print("❌ Invalid choice. Please enter 1, 2, or 3.")
        
        # Get recent filings
        filings = self.get_recent_filings(cik, form_type, count=3)
        
        if not filings:
            print(f"❌ No {form_type} filings found")
            return False
        
        print(f"\n📈 Found {len(filings)} recent {form_type} filing(s):")
        print("-" * 80)
        
        # Display each filing with its URLs
        for i, filing in enumerate(filings, 1):
            print(f"\n📋 FILING #{i}")
            print(f"   Form Type: {filing['form']}")
            print(f"   Filing Date: {filing['filingDate']}")
            print(f"   Report Date: {filing['reportDate']}")
            print(f"   Accession Number: {filing['accessionNumber']}")
            
            # Generate URLs
            urls = self.generate_financial_report_urls(cik, filing['accessionNumber'])
            
            print(f"\n   📎 FINANCIAL REPORT LINKS:")
            for report_type, url in urls.items():
                print(f"   └─ {report_type}: {url}")
        
        # Ask user about downloading
        download_choice = self.ask_download_preference()
        
        if download_choice == '3':  # View links only
            print("\n✅ Links displayed above. No downloads performed.")
            return True
        
        # Select filing to download (if multiple)
        filing_idx = self.select_filing_to_download(filings)
        selected_filing = filings[filing_idx]
        selected_urls = self.generate_financial_report_urls(cik, selected_filing['accessionNumber'])
        
        # Select specific reports if requested
        if download_choice == '2':
            selected_urls = self.select_specific_reports(selected_urls)
        
        # Get company ticker for folder structure
        ticker = company_input if len(company_input) <= 5 else company_input[:5]
        
        # Perform downloads
        success = self.download_filing_reports(selected_filing, selected_urls, ticker, form_type)
        
        return success


def main():
    """Main interactive function"""
    print("💼 Financial Report Finder & Downloader")
    print("=" * 55)
    print("Find and download SEC financial reports for any public company")
    print("Excel files can be analyzed and exported as individual CSV sheets")
    print("Enter company ticker (e.g., AAPL) or company name")
    print("Type 'quit' or 'exit' to stop")
    
    # Ask for SEC API key (optional)
    print("\n🔑 SEC API Integration (Optional):")
    print("   For better download reliability, you can provide a SEC API key")
    print("   Get one at: https://sec-api.io")
    sec_api_key = input("   Enter SEC API key (or press Enter to skip): ").strip()
    
    if not sec_api_key:
        sec_api_key = None
        print("   ℹ️  Using basic download method")
    
    finder = FinancialReportFinder(sec_api_key=sec_api_key)
    
    print("\n" + "="*55)
    
    while True:
        try:
            # Get user input
            company_input = input("\n🔍 Enter company ticker or name: ").strip()
            
            # Check for exit commands
            if company_input.lower() in ['quit', 'exit', 'q']:
                print("👋 Goodbye!")
                break
            
            if not company_input:
                print("❌ Please enter a company ticker or name")
                continue
            
            # Search for company
            success = finder.search_company(company_input)
            
            if success:
                print("\n" + "="*80)
                print("✅ Search completed successfully!")
            else:
                print("\n❌ Search failed. Try a different ticker or company name.")
            
            # Ask if user wants to search another company
            print("\n" + "-"*50)
            continue_search = input("🔄 Search another company? (y/n): ").strip().lower()
            if continue_search in ['n', 'no']:
                print("👋 Goodbye!")
                break
            
        except KeyboardInterrupt:
            print("\n\n👋 Goodbye!")
            break
        except Exception as e:
            print(f"\n❌ An error occurred: {e}")
            print("Please try again.")


if __name__ == "__main__":
    main()