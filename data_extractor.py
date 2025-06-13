#!/usr/bin/env python3
"""
Enhanced Financial Data Extractor
Combines SEC Edgar and Yahoo Finance data sources with improved UI
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import zipfile
import tempfile
import os
import sys
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta
import json
import webbrowser
import requests

# Import the existing modules
from sec_finder import FinancialReportFinder
from yf_finder import DataCenterExtractor


class ScrollableFrame(ttk.Frame):
    """A scrollable frame widget for tkinter"""
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        
        # Create canvas and scrollbar
        self.canvas = tk.Canvas(self)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        # Configure scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Create window in canvas
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel
        self.bind_mousewheel()
    
    def bind_mousewheel(self):
        """Bind mousewheel events for scrolling"""
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(event):
            self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            self.canvas.unbind_all("<MouseWheel>")
        
        self.canvas.bind('<Enter>', _bind_to_mousewheel)
        self.canvas.bind('<Leave>', _unbind_from_mousewheel)


class EnhancedFinancialExtractor:
    def __init__(self, root):
        self.outer_root = root
        self.outer_root.title("Enhanced Financial Data Extractor")
        self.outer_root.geometry("1200x800")
        self.outer_root.configure(bg='#f0f0f0')

        # Create canvas and scrollbar
        canvas = tk.Canvas(self.outer_root, borderwidth=0, background="#f0f0f0")
        scrollbar = tk.Scrollbar(self.outer_root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Create internal frame and bind it to the canvas
        internal_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=internal_frame, anchor="nw")

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        internal_frame.bind("<Configure>", on_frame_configure)

        # Optional: enable mousewheel scrolling
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        # Use internal frame for layout from now on
        self.root = internal_frame
        
        # Initialize data source modules
        self.sec_finder = FinancialReportFinder()
        self.yf_extractor = DataCenterExtractor()
        
        # Data storage
        self.selected_company = None
        self.available_periods = []
        self.selected_periods = []
        self.preview_data = {}
        self.selected_files = set()
        

        self.setup_ui()
        
    def setup_ui(self):
        """Setup the main user interface"""
        # Create main container with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Enhanced Financial Data Extractor", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Step 1: Data Source Selection
        self.create_data_source_section(main_frame, 1)
        
        # Step 2: Company Search
        self.create_company_search_section(main_frame, 2)
        
        # Step 3: Form Type Selection
        self.create_form_type_section(main_frame, 3)
        
        # Step 4: Time Period Selection
        self.create_time_period_section(main_frame, 4)
        
        # Step 5: File Preview and Selection
        self.create_file_preview_section(main_frame, 5)
        
        # Step 6: Download Options
        self.create_download_section(main_frame, 6)
        
        # Progress bar
        self.progress_var = tk.StringVar(value="Ready")
        self.progress_label = ttk.Label(main_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=7, column=0, columnspan=2, pady=(20, 0))
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        
    def create_data_source_section(self, parent, row):
        """Create data source selection section"""
        frame = ttk.LabelFrame(parent, text="Step 1: Select Data Source", padding="10")
        frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.data_source_var = tk.StringVar(value="sec")
        
        ttk.Radiobutton(frame, text="SEC Edgar (Official SEC Filings)", 
                       variable=self.data_source_var, value="sec",
                       command=self.on_data_source_change).grid(row=0, column=0, sticky=tk.W)
        
        ttk.Radiobutton(frame, text="Yahoo Finance (Financial Statements)", 
                       variable=self.data_source_var, value="yahoo",
                       command=self.on_data_source_change).grid(row=1, column=0, sticky=tk.W)
        
        # Info labels
        self.sec_info = ttk.Label(frame, text="✓ Official SEC documents, Excel reports, HTML filings", 
                                 foreground="green")
        self.sec_info.grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        
        self.yahoo_info = ttk.Label(frame, text="✓ Historical financial data, easy CSV export", 
                                   foreground="blue")
        self.yahoo_info.grid(row=1, column=1, sticky=tk.W, padx=(20, 0))
        
    def create_company_search_section(self, parent, row):
        """Create company search section"""
        frame = ttk.LabelFrame(parent, text="Step 2: Search Company", padding="10")
        frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(frame, text="Company Ticker or Name:").grid(row=0, column=0, sticky=tk.W)
        
        self.company_entry = ttk.Entry(frame, width=30)
        self.company_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        self.company_entry.bind('<Return>', lambda e: self.search_company())
        
        self.search_button = ttk.Button(frame, text="Search", command=self.search_company)
        self.search_button.grid(row=0, column=2, padx=(10, 0))
        
        # Company info display
        self.company_info_label = ttk.Label(frame, text="", foreground="green")
        self.company_info_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        
        frame.columnconfigure(1, weight=1)
        
    def create_form_type_section(self, parent, row):
        """Create form type selection section"""
        self.form_frame = ttk.LabelFrame(parent, text="Step 3: Select Form Type", padding="10")
        self.form_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.form_frame.grid_remove()  # Hide initially
        
        ttk.Label(self.form_frame, text="Filing Type:").grid(row=0, column=0, sticky=tk.W)
        
        self.form_type_var = tk.StringVar()
        self.form_combo = ttk.Combobox(self.form_frame, textvariable=self.form_type_var, 
                                      state="readonly", width=40)
        self.form_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        self.form_combo.bind('<<ComboboxSelected>>', self.on_form_type_change)
        
        self.form_frame.columnconfigure(1, weight=1)
        
    def create_time_period_section(self, parent, row):
        """Create time period selection section"""
        self.period_frame = ttk.LabelFrame(parent, text="Step 4: Select Time Periods", padding="10")
        self.period_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.period_frame.grid_remove()  # Hide initially
        
        # Create frame for period list with scrollbar
        period_container = ttk.Frame(self.period_frame)
        period_container.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollable listbox for periods
        self.period_listbox = tk.Listbox(period_container, selectmode=tk.MULTIPLE, height=6)
        scrollbar = ttk.Scrollbar(period_container, orient=tk.VERTICAL, command=self.period_listbox.yview)
        self.period_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.period_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        period_container.columnconfigure(0, weight=1)
        period_container.rowconfigure(0, weight=1)
        
        # Buttons
        button_frame = ttk.Frame(self.period_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(button_frame, text="Select All", 
                  command=self.select_all_periods).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(button_frame, text="Clear All", 
                  command=self.clear_all_periods).grid(row=0, column=1, padx=(5, 0))
        ttk.Button(button_frame, text="Load Periods", 
                  command=self.load_time_periods).grid(row=0, column=2, padx=(10, 0))
        
        ttk.Button(button_frame, text="Confirm Periods / Show Files",command=self.load_available_files).grid(row=1, column=1, padx=(10, 0))  

        self.period_frame.columnconfigure(0, weight=1)
        self.period_frame.rowconfigure(0, weight=1)

        
        
    def create_file_preview_section(self, parent, row):
        """Create file preview and selection section"""
        self.preview_frame = ttk.LabelFrame(parent, text="Step 5: Preview and Select Files", padding="10")
        self.preview_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.preview_frame.grid_remove()  # Hide initially
        
        # File list with checkboxes
        files_container = ttk.Frame(self.preview_frame)
        files_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.file_tree = ttk.Treeview(files_container, columns=('Type', 'Description'), 
                                     show='tree headings', height=8)
        self.file_tree.heading('#0', text='File Name')
        self.file_tree.heading('Type', text='Type')
        self.file_tree.heading('Description', text='Description')
        
        file_scrollbar = ttk.Scrollbar(files_container, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=file_scrollbar.set)
        
        self.file_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        file_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        files_container.columnconfigure(0, weight=1)
        files_container.rowconfigure(0, weight=1)
        
        # Preview buttons
        preview_button_frame = ttk.Frame(self.preview_frame)
        preview_button_frame.grid(row=1, column=0, pady=(10, 0))
        
        ttk.Button(preview_button_frame, text="Preview Selected", 
                  command=self.preview_selected_file).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(preview_button_frame, text="Select All Files", 
                  command=self.select_all_files).grid(row=0, column=1, padx=(5, 0))
        ttk.Button(preview_button_frame, text="Load Available Files", 
                  command=self.load_available_files).grid(row=0, column=2, padx=(10, 0))
        
        self.preview_frame.columnconfigure(0, weight=1)
        self.preview_frame.rowconfigure(0, weight=1)
    
       
    def create_download_section(self, parent, row):
        """Create download options section"""
        self.download_frame = ttk.LabelFrame(parent, text="Step 6: Download Options", padding="10")
        self.download_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.download_frame.grid_remove()  # Hide initially
        
        # Storage options
        storage_frame = ttk.Frame(self.download_frame)
        storage_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        self.storage_var = tk.StringVar(value="local")
        
        ttk.Radiobutton(storage_frame, text="Save locally", 
                       variable=self.storage_var, value="local",
                       command=self.on_storage_change).grid(row=0, column=0, sticky=tk.W)
        
        ttk.Radiobutton(storage_frame, text="Download as ZIP", 
                       variable=self.storage_var, value="zip",
                       command=self.on_storage_change).grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        
        # Local path selection
        self.local_frame = ttk.Frame(self.download_frame)
        self.local_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(self.local_frame, text="Local Path:").grid(row=0, column=0, sticky=tk.W)
        
        self.local_path_var = tk.StringVar(value=str(Path.home() / "Downloads" / "FinancialData"))
        self.local_path_entry = ttk.Entry(self.local_frame, textvariable=self.local_path_var, width=50)
        self.local_path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        
        ttk.Button(self.local_frame, text="Browse", 
                  command=self.browse_local_path).grid(row=0, column=2, padx=(10, 0))
        
        self.local_frame.columnconfigure(1, weight=1)
        
        # Extract button
        self.extract_button = ttk.Button(self.download_frame, text="Extract Data", 
                                        command=self.extract_data)
        self.extract_button.grid(row=2, column=0, columnspan=2, pady=(20, 0))
        
    def on_data_source_change(self):
        """Handle data source selection change"""
        # Reset form when data source changes
        self.reset_form()
        
        # Update form type options based on data source
        if self.data_source_var.get() == "sec":
            form_types = [
                "10-Q (Quarterly Reports)",
                "10-K (Annual Reports)", 
                "8-K (Current Reports)"
            ]
        else:  # yahoo
            form_types = [
                "Annual (Yearly Financial Statements)",
                "Quarterly (Quarterly Financial Statements)"
            ]
        
        self.form_combo['values'] = form_types
        
    def reset_form(self):
        """Reset form to initial state"""
        self.selected_company = None
        self.available_periods = []
        self.selected_periods = []
        self.preview_data = {}
        self.selected_files = set()
        
        # Hide sections
        self.form_frame.grid_remove()
        self.period_frame.grid_remove()
        self.preview_frame.grid_remove()
        self.download_frame.grid_remove()
        
        # Clear widgets
        self.company_info_label.config(text="")
        self.period_listbox.delete(0, tk.END)
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
    def search_company(self):
        """Search for company using selected data source"""
        company_input = self.company_entry.get().strip()
        if not company_input:
            messagebox.showerror("Error", "Please enter a company ticker or name")
            return
        
        self.progress_var.set("Searching for company...")
        self.progress_bar.start()
        
        def search_thread():
            try:
                if self.data_source_var.get() == "sec":
                    # Use SEC finder
                    cik = self.sec_finder.get_company_cik(company_input)
                    if cik:
                        # For SEC, we'll store the CIK and company input
                        
                        company_info = self.sec_finder.get_company_info_from_cik(cik)
                        self.selected_company = {
                            'cik': cik,
                            'ticker': company_info['ticker'],
                            'name': company_info['title'],
                            'source': 'sec'
                        }
                        success_msg = f"Found company: {company_info['title']} ({company_info['ticker']}) (CIK: {cik})"
                    else:
                        self.selected_company = None
                        success_msg = None
                else:
                    # Use Yahoo Finance
                    company_info = self.yf_extractor.search_company(company_input)
                    if company_info:
                        company_info['source'] = 'yahoo'
                        self.selected_company = company_info
                        success_msg = f"Found: {company_info['name']} ({company_info['ticker']})\nSector: {company_info['sector']}"
                    else:
                        self.selected_company = None
                        success_msg = None
                
                # Update UI in main thread
                self.root.after(0, self.update_company_search_result, success_msg)
                
            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("Error", f"Search failed: {str(e)}"))

            finally:
                self.root.after(0, self.stop_progress)
        
        threading.Thread(target=search_thread, daemon=True).start()
        
    def update_company_search_result(self, success_msg):
        """Update UI with company search results"""
        if success_msg:
            self.company_info_label.config(text=success_msg)
            self.form_frame.grid()
        else:
            messagebox.showerror("Error", "Company not found")
        
    def on_form_type_change(self, event=None):
        """Handle form type selection change"""
        if self.selected_company:
            self.period_frame.grid()
            
    def load_time_periods(self):
        """Load available time periods for selected company and form type"""
        if not self.selected_company or not self.form_type_var.get():
            messagebox.showerror("Error", "Please select company and form type first")
            return
        
        self.progress_var.set("Loading available periods...")
        self.progress_bar.start()
        
        def load_periods_thread():
            try:
                periods = []
                
                if self.selected_company['source'] == 'sec':
                    # Get SEC filings
                    form_type = self.form_type_var.get().split()[0]  # Extract form type (10-Q, 10-K, etc.)
                    filings = self.sec_finder.get_recent_filings(self.selected_company['cik'], form_type, count=10)
                    
                    for filing in filings:
                        periods.append({
                            'description': f"{filing['form']} - Filing: {filing['filingDate']} | Report: {filing['reportDate']}",
                            'filing_date': filing['filingDate'],
                            'report_date': filing['reportDate'],
                            'accession': filing['accessionNumber'],
                            'form': filing['form']
                        })
                        
                else:
                    # Get Yahoo Finance periods
                    ticker = self.selected_company['ticker']
                    yf_periods = self.yf_extractor.get_available_periods(ticker)
                    
                    if 'Quarterly' in self.form_type_var.get():
                        # Generate quarterly periods
                        for year in yf_periods.get('years', [])[:5]:  # Last 5 years
                            for quarter in [1, 2, 3, 4]:
                                periods.append({
                                    'description': f"Q{quarter} {year}",
                                    'year': year,
                                    'quarter': quarter,
                                    'period_type': 'quarterly'
                                })
                    else:
                        # Generate annual periods
                        for year in yf_periods.get('years', [])[:10]:  # Last 10 years
                            periods.append({
                                'description': f"Annual {year}",
                                'year': year,
                                'period_type': 'annual'
                            })
                
                self.available_periods = periods
                self.root.after(0, self.update_period_list)
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to load periods: {str(e)}"))
            finally:
                self.root.after(0, self.stop_progress)
        
        threading.Thread(target=load_periods_thread, daemon=True).start()
        
    def update_period_list(self):
        """Update the period listbox with available periods"""
        self.period_listbox.delete(0, tk.END)
        for period in self.available_periods:
            self.period_listbox.insert(tk.END, period['description'])
        
    def select_all_periods(self):
        """Select all periods in the list"""
        self.period_listbox.select_set(0, tk.END)
        
    def clear_all_periods(self):
        """Clear all period selections"""
        self.period_listbox.selection_clear(0, tk.END)
        
    def load_available_files(self):
        """Load available files based on selected periods"""
        selected_indices = self.period_listbox.curselection()
        
        if not selected_indices:
            messagebox.showerror("Error", "Please select at least one time period")
            return
        
        self.selected_periods = [self.available_periods[i] for i in selected_indices]
        
        # Clear file tree
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # Add available file types based on data source
        if self.selected_company['source'] == 'sec':
            file_types = [
                ('Excel Financial Report', 'Financial_Report.xlsx', 'Complete financial statements in Excel format'),
                ('Income Statement', 'Income_Statement.htm', 'Income statement in HTML format'),
                ('Balance Sheet', 'Balance_Sheet.htm', 'Balance sheet in HTML format'),
                ('Cash Flow Statement', 'Cash_Flow_Statement.htm', 'Cash flow statement in HTML format'),
                ('Stockholder Equity', 'Stockholder_Equity.htm', 'Statement of stockholder equity')
            ]
        else:
            file_types = [
                ('Income Statement', 'income_statement.csv', 'Income statement data'),
                ('Balance Sheet', 'balance_sheet.csv', 'Balance sheet data'),
                ('Cash Flow', 'cash_flow.csv', 'Cash flow statement data'),
                ('Company Info', 'company_info.csv', 'Basic company information'),
                ('Historical Data', 'historical_data.csv', 'Historical stock price data')
            ]
        
        for file_type, filename, description in file_types:
            item_id = self.file_tree.insert('', tk.END, text=filename, 
                                          values=(file_type, description),
                                          tags=('file',))
        
        self.preview_frame.grid()
        self.download_frame.grid()
        
    def select_all_files(self):
        """Select all files in the tree"""
        for item in self.file_tree.get_children():
            self.file_tree.selection_add(item)
        
    def preview_selected_file(self):
        """Enhanced preview with Excel sheet analysis for SEC files"""
        selected_items = self.file_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a file to preview")
            return
        
        # For now, show a simple preview window
        item = selected_items[0]
        file_name = self.file_tree.item(item, 'text')
        file_type = self.file_tree.item(item, 'values')[0]
        
        # Create enhanced preview window
        preview_window = tk.Toplevel(self.root)
        preview_window.title(f"Preview: {file_name}")
        preview_window.geometry("900x700")
        
        # Create main frame with scrollable content
        main_frame = ttk.Frame(preview_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add notebook for different preview types
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Basic info tab
        info_frame = ttk.Frame(notebook)
        notebook.add(info_frame, text="📄 File Info")
        
        info_text = tk.Text(info_frame, wrap=tk.WORD, height=15)
        info_scrollbar = ttk.Scrollbar(info_frame, orient=tk.VERTICAL, command=info_text.yview)
        info_text.configure(yscrollcommand=info_scrollbar.set)
        
        info_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        info_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Enhanced preview content
        preview_text = f"📊 File Type: {file_type}\n"
        preview_text += f"📁 File Name: {file_name}\n"
        preview_text += f"📅 Selected Periods: {len(self.selected_periods)}\n"
        preview_text += f"🏢 Company: {self.selected_company['name']} ({self.selected_company['ticker']})\n\n"
        
        if self.selected_company['source'] == 'sec':
            preview_text += "🏛️ SEC Edgar Data Format:\n\n"
            
            if 'Excel' in file_type:
                preview_text += "📊 Excel Financial Report Features:\n"
                preview_text += "• Multiple sheets with comprehensive financial data\n"
                preview_text += "• Consolidated financial statements (auto-selected)\n"
                preview_text += "• Income statement, balance sheet, cash flow\n"
                preview_text += "• Notes and footnotes\n"
                preview_text += "• Automatic CSV export of consolidated sheets\n\n"
                
                preview_text += "🔄 Automatic Processing:\n"
                preview_text += "• Consolidated sheets are automatically selected\n"
                preview_text += "• Main financial statements are exported as CSV\n"
                preview_text += "• Original Excel file is preserved\n"
                preview_text += "• Files organized by year in folder structure\n\n"
                
                # Add Excel-specific preview tab if we have sample data
                if self.selected_periods:
                    excel_frame = ttk.Frame(notebook)
                    notebook.add(excel_frame, text="📊 Excel Sheets")
                    
                    excel_text = tk.Text(excel_frame, wrap=tk.WORD)
                    excel_scrollbar = ttk.Scrollbar(excel_frame, orient=tk.VERTICAL, command=excel_text.yview)
                    excel_text.configure(yscrollcommand=excel_scrollbar.set)
                    
                    excel_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                    excel_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                    
                    excel_preview = "📊 Expected Excel Sheet Structure:\n\n"
                    excel_preview += "🟢 Consolidated Statements (Auto-Selected):\n"
                    excel_preview += "   • Consolidated Income Statement\n"
                    excel_preview += "   • Consolidated Balance Sheet\n"
                    excel_preview += "   • Consolidated Cash Flow\n\n"
                    
                    excel_preview += "📋 Other Typical Sheets:\n"
                    excel_preview += "   • Cover Page\n"
                    excel_preview += "   • Notes to Financial Statements\n"
                    excel_preview += "   • Supplementary Information\n\n"
                    
                    excel_preview += "⚙️ Processing Notes:\n"
                    excel_preview += "   • Consolidated sheets are automatically exported as CSV\n"
                    excel_preview += "   • Each sheet becomes a separate CSV file\n"
                    excel_preview += "   • Files are organized in year-based folders\n"
                    excel_preview += "   • CSV files are named with ticker, sheet name, and date\n"
                    
                    excel_text.insert(tk.END, excel_preview)
                    excel_text.configure(state=tk.DISABLED)
                
            else:
                preview_text += f"📄 {file_type} Features:\n"
                preview_text += "• Formatted financial statements in HTML\n"
                preview_text += "• Professional SEC formatting\n"
                preview_text += "• Direct browser preview available\n"
                preview_text += "• Copy data directly from tables\n\n"
                
        else:
            preview_text += "📈 Yahoo Finance Data Format:\n\n"
            preview_text += f"📊 {file_type} Features:\n"
            preview_text += "• CSV files with time series data\n"
            preview_text += "• Columns for different financial metrics\n"
            preview_text += "• Quarterly/annual frequency based on selection\n"
            preview_text += "• Easy import into Excel, Python, R\n"
            preview_text += "• Organized by period type and year\n\n"
        
        # Add folder structure preview
        preview_text += "📁 Folder Structure:\n"
        if self.selected_company['source'] == 'sec':
            preview_text += f"   📂 {self.selected_company['ticker']}/\n"
            for period in self.selected_periods[:3]:  # Show first 3 periods
                year = period['report_date'][:4]
                preview_text += f"   └── 📂 {period['form']}/\n"
                preview_text += f"       └── 📂 {year}/\n"
                preview_text += f"           ├── 📄 Financial_Report.xlsx\n"
                preview_text += f"           └── 📄 *_Consolidated_*.csv\n"
        else:
            preview_text += f"   📂 {self.selected_company['ticker']}/\n"
            for period in self.selected_periods[:3]:  # Show first 3 periods
                period_type = period['period_type'].upper()
                year = str(period['year'])
                preview_text += f"   └── 📂 {period_type}/\n"
                preview_text += f"       └── 📂 {year}/\n"
                preview_text += f"           ├── 📄 {self.selected_company['ticker']}-Income_Statement-*.csv\n"
                preview_text += f"           ├── 📄 {self.selected_company['ticker']}-Balance_Sheet-*.csv\n"
                preview_text += f"           └── 📄 {self.selected_company['ticker']}-Cash_Flow-*.csv\n"
        
        if len(self.selected_periods) > 3:
            preview_text += f"   └── ... and {len(self.selected_periods) - 3} more period folders\n"
        
        preview_text += "\n💡 Pro Tips:\n"
        preview_text += "• Files are automatically organized by year for easy analysis\n"
        preview_text += "• Excel files include automatic CSV export of key sheets\n"
        preview_text += "• Use 'Open Folder' button after extraction to view results\n"
        preview_text += "• ZIP downloads preserve the same folder structure\n"
        
        info_text.insert(tk.END, preview_text)
        info_text.configure(state=tk.DISABLED)
        
        # Add advanced preview tab for specific data sources
        if self.selected_company['source'] == 'sec' and len(self.selected_periods) > 0:
            advanced_frame = ttk.Frame(notebook)
            notebook.add(advanced_frame, text="🔍 Advanced Preview")
            
            advanced_text = tk.Text(advanced_frame, wrap=tk.WORD)
            advanced_scrollbar = ttk.Scrollbar(advanced_frame, orient=tk.VERTICAL, command=advanced_text.yview)
            advanced_text.configure(yscrollcommand=advanced_scrollbar.set)
            
            advanced_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            advanced_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Generate URLs for preview
            period = self.selected_periods[0]  # Use first period for preview
            urls = self.sec_finder.generate_financial_report_urls(
                self.selected_company['cik'], 
                period['accession']
            )
            
            advanced_preview = f"🔗 SEC Edgar URLs for {period['description']}:\n\n"
            
            url_mapping = {
                'Excel Financial Report': '📊 Excel File',
                'Income Statement (HTML)': '📄 Income Statement',
                'Balance Sheet (HTML)': '📄 Balance Sheet',
                'Cash Flow Statement (HTML)': '📄 Cash Flow',
                'Stockholder Equity (HTML)': '📄 Stockholder Equity'
            }
            
            for url_key, url in urls.items():
                icon = url_mapping.get(url_key, '📄')
                advanced_preview += f"{icon}: {url_key}\n"
                advanced_preview += f"   🔗 {url}\n\n"
            
            advanced_preview += "⚙️ Automatic Processing Details:\n\n"
            advanced_preview += "🔄 Excel File Processing:\n"
            advanced_preview += "1. Download Excel file from SEC EDGAR\n"
            advanced_preview += "2. Analyze sheet structure automatically\n"
            advanced_preview += "3. Identify consolidated financial statements\n"
            advanced_preview += "4. Export consolidated sheets as individual CSV files\n"
            advanced_preview += "5. Preserve original Excel file for reference\n\n"
            
            advanced_preview += "📊 Expected Sheet Categories:\n"
            advanced_preview += "• Cover Page: Filing information and summary\n"
            advanced_preview += "• Consolidated Income Statement: Revenue, expenses, profit\n"
            advanced_preview += "• Consolidated Balance Sheet: Assets, liabilities, equity\n"
            advanced_preview += "• Consolidated Cash Flow: Operating, investing, financing flows\n"
            advanced_preview += "• Notes: Footnotes and additional disclosures\n\n"
            
            advanced_preview += "🎯 Auto-Selection Logic:\n"
            advanced_preview += "• Sheets with 'consolidated' in name are automatically selected\n"
            advanced_preview += "• If no consolidated sheets found, main financial statements are selected\n"
            advanced_preview += "• Cover pages and notes are excluded from auto-selection\n"
            advanced_preview += "• Maximum of 3-5 sheets are auto-exported to prevent clutter\n"
            
            advanced_text.insert(tk.END, advanced_preview)
            advanced_text.configure(state=tk.DISABLED)
        
        # Close button
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))
        
        ttk.Button(button_frame, text="✅ Close Preview", 
                command=preview_window.destroy).pack(side=tk.LEFT, padx=(0, 10))
        
        # Add direct link button for SEC HTML files
        if self.selected_company['source'] == 'sec' and 'Statement' in file_type and len(self.selected_periods) > 0:
            period = self.selected_periods[0]
            urls = self.sec_finder.generate_financial_report_urls(
                self.selected_company['cik'], 
                period['accession']
            )
            
            # Map file type to URL
            url_mapping = {
                'Income Statement': 'Income Statement (HTML)',
                'Balance Sheet': 'Balance Sheet (HTML)',
                'Cash Flow Statement': 'Cash Flow Statement (HTML)',
                'Stockholder Equity': 'Stockholder Equity (HTML)'
            }
            
            url_key = url_mapping.get(file_type)
            if url_key and url_key in urls:
                def open_html_preview():
                    webbrowser.open(urls[url_key])
                
                ttk.Button(button_frame, text="🌐 Open HTML Preview", 
                        command=open_html_preview).pack(side=tk.LEFT)

    # =====================================
    # NEW METHODS - COMPLETELY NEW
    # =====================================

    def download_financial_report_enhanced(self, url, download_dir, filing_date, report_date, report_type, ticker):
        """Enhanced download method with auto Excel sheet selection"""
        
        # Handle Excel files specially (with automatic sheet selection)
        if 'Excel' in report_type:
            return self.download_and_process_excel_auto(url, download_dir, filing_date, report_date, report_type, ticker)
        
        # Handle regular HTML files using existing method
        return self.sec_finder.download_financial_report(url, download_dir, filing_date, report_date, report_type, ticker)

    def download_and_process_excel_auto(self, url, download_dir, filing_date, report_date, report_type, ticker):
        """Download Excel file and automatically export consolidated sheets to CSV"""
        filename = self.sec_finder.get_safe_filename(url, filing_date, report_date, report_type, ticker)
        filepath = download_dir / filename
        
        print(f"   📥 Downloading {report_type}...")
        print(f"      └─ File: {filename}")
        
        # Download the Excel file first
        file_content = None
        success = False
        
        if self.sec_finder.render_api:
            print(f"      └─ Using SEC API...")
            try:
                file_content = self.sec_finder.render_api.get_file(url, return_binary=True)
                success = True
            except Exception as e:
                print(f"❌ SEC API download failed: {e}")
        
        if not success:
            print(f"      └─ Using basic download...")
            try:
                response = requests.get(url, headers=self.sec_finder.headers, timeout=30)
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
        
        # Analyze Excel file and automatically export consolidated sheets
        if self.sec_finder.pandas_available:
            print(f"      🔍 Analyzing Excel file structure...")
            sheet_info = self.sec_finder.analyze_excel_file(file_content)
            
            if sheet_info:
                # Auto-select consolidated sheets (default behavior)
                auto_selected = []
                for sheet_name in sheet_info.keys():
                    if 'consolidated' in sheet_name.lower():
                        auto_selected.append(sheet_name)
                
                if auto_selected:
                    print(f"      🟢 AUTO-SELECTING CONSOLIDATED SHEETS:")
                    for sheet in auto_selected:
                        print(f"         └─ '{sheet}'")
                    
                    # Export auto-selected sheets to CSV
                    exported = self.sec_finder.export_excel_sheets_to_csv(
                        file_content, auto_selected, download_dir, ticker, report_date
                    )
                    print(f"      📊 Auto-exported {exported} consolidated sheet(s) to CSV")
                else:
                    print(f"      ⚠️  No consolidated sheets found for auto-selection")
                    # If no consolidated sheets, try to find main financial statement sheets
                    financial_sheets = []
                    for sheet_name in sheet_info.keys():
                        if any(term in sheet_name.lower() for term in ['income', 'balance', 'cash', 'statement']):
                            financial_sheets.append(sheet_name)
                    
                    if financial_sheets:
                        print(f"      🔵 AUTO-SELECTING FINANCIAL STATEMENT SHEETS:")
                        for sheet in financial_sheets[:3]:  # Limit to first 3
                            print(f"         └─ '{sheet}'")
                        
                        exported = self.sec_finder.export_excel_sheets_to_csv(
                            file_content, financial_sheets[:3], download_dir, ticker, report_date
                        )
                        print(f"      📊 Auto-exported {exported} financial statement sheet(s) to CSV")
            else:
                print(f"      ⚠️  Could not analyze Excel file structure")
        
        return True

    def add_excel_preview_functionality(self):
        """Add Excel preview functionality inspired by flask app"""
        # This method would be called when Excel files are selected for preview
        # Currently a placeholder for future enhancements
        pass
        
    def on_storage_change(self):
        """Handle storage option change"""
        if self.storage_var.get() == "local":
            self.local_frame.grid()
        else:
            self.local_frame.grid_remove()
            
    def browse_local_path(self):
        """Browse for local storage path"""
        path = filedialog.askdirectory(initialdir=self.local_path_var.get())
        if path:
            self.local_path_var.set(path)
            
    def extract_data(self):
        """Extract the selected data"""
        # Validate selections
        if not self.selected_company:
            messagebox.showerror("Error", "Please select a company")
            return
            
        if not self.selected_periods:
            messagebox.showerror("Error", "Please select time periods")
            return
            
        selected_files = self.file_tree.selection()
        if not selected_files:
            messagebox.showerror("Error", "Please select files to extract")
            return
        
        if self.storage_var.get() == "local" and not self.local_path_var.get():
            messagebox.showerror("Error", "Please specify local storage path")
            return
        
        self.progress_var.set("Extracting data...")
        self.progress_bar.start()
        
        def extract_thread():
            try:
                extracted_files = []
                
                if self.selected_company['source'] == 'sec':
                    # Extract SEC data
                    extracted_files = self.extract_sec_data(selected_files)
                else:
                    # Extract Yahoo Finance data
                    extracted_files = self.extract_yahoo_data(selected_files)
                
                # Handle storage
                if self.storage_var.get() == "zip":
                    zip_path = self.create_zip_file(extracted_files)
                    self.root.after(0, lambda: self.show_extraction_success(f"ZIP file created: {zip_path}"))
                else:
                    self.root.after(0, lambda: self.show_extraction_success(f"Files saved to: {self.local_path_var.get()}"))
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Extraction failed: {str(e)}"))
            finally:
                self.root.after(0, self.stop_progress)
        
        threading.Thread(target=extract_thread, daemon=True).start()
        
    def extract_sec_data(self, selected_files):
        """Extract SEC data for selected files and periods with enhanced folder structure"""
        extracted_files = []
        
        # Create local directory if needed
        if self.storage_var.get() == "local":
            local_dir = Path(self.local_path_var.get())
            local_dir.mkdir(parents=True, exist_ok=True)
        
        for period in self.selected_periods:
            # Extract year from report date for folder structure
            try:
                period_year = period['report_date'][:4]  # Extract YYYY from YYYY-MM-DD
            except:
                period_year = "unknown"
            
            # Create organized folder structure: data/TICKER/FORM-TYPE/YEAR/
            if self.storage_var.get() == "local":
                download_dir = local_dir / self.selected_company['ticker'] / period['form'] / period_year
                download_dir.mkdir(parents=True, exist_ok=True)
            
            # Generate URLs for this period
            urls = self.sec_finder.generate_financial_report_urls(
                self.selected_company['cik'], 
                period['accession']
            )
            
            for file_item in selected_files:
                file_name = self.file_tree.item(file_item, 'text')
                file_type = self.file_tree.item(file_item, 'values')[0]
                
                # Map file type to URL
                url_key = None
                for key in urls:
                    if file_type.lower() in key.lower():
                        url_key = key
                        break
                
                if url_key and url_key in urls:
                    try:
                        # Download file
                        if self.storage_var.get() == "local":
                            # Use enhanced download method with auto Excel sheet selection
                            success = self.download_financial_report_enhanced(
                                urls[url_key], download_dir, 
                                period['filing_date'], period['report_date'], 
                                file_type, self.selected_company['ticker']
                            )
                            
                            if success:
                                # Find the downloaded file(s) - could be multiple if Excel sheets were exported
                                safe_filename = self.sec_finder.get_safe_filename(
                                    urls[url_key], period['filing_date'], 
                                    period['report_date'], file_type, 
                                    self.selected_company['ticker']
                                )
                                file_path = download_dir / safe_filename
                                extracted_files.append(str(file_path))
                                
                                # Also add any CSV files that were exported from Excel
                                if 'Excel' in file_type:
                                    csv_files = list(download_dir.glob(f"{self.selected_company['ticker']}-*-{period['report_date'].replace('-', '')}.csv"))
                                    extracted_files.extend([str(f) for f in csv_files])
                        else:
                            # For ZIP download, we'll handle this differently
                            # Store file info for later processing
                            extracted_files.append({
                                'url': urls[url_key],
                                'filename': f"{self.selected_company['ticker']}-{file_type}-{period['report_date']}.{file_name.split('.')[-1]}",
                                'period': period,
                                'type': file_type,
                                'year': period_year
                            })
                    except Exception as e:
                        print(f"Failed to download {file_type} for {period['description']}: {e}")
        
        return extracted_files

    
    def extract_yahoo_data(self, selected_files):
        """Extract Yahoo Finance data for selected files and periods with enhanced folder structure"""
        extracted_files = []
        
        # Create local directory if needed
        if self.storage_var.get() == "local":
            local_dir = Path(self.local_path_var.get())
            local_dir.mkdir(parents=True, exist_ok=True)
        
        for period in self.selected_periods:
            # Extract year for folder structure
            period_year = str(period['year'])
            period_type = period['period_type'].upper()  # QUARTERLY or ANNUAL
            
            # Create organized folder structure: data/TICKER/PERIOD_TYPE/YEAR/
            if self.storage_var.get() == "local":
                download_dir = local_dir / self.selected_company['ticker'] / period_type / period_year
                download_dir.mkdir(parents=True, exist_ok=True)
            
            # Get data for this period
            if period['period_type'] == 'quarterly':
                # Calculate start and end dates for quarter
                year = period['year']
                quarter = period['quarter']
                start_month = (quarter - 1) * 3 + 1
                end_month = quarter * 3
                start_date = f"{year}-{start_month:02d}-01"
                end_date = f"{year}-{end_month:02d}-{30 if end_month in [6, 9] else 31}"
                period_suffix = f"Q{quarter}_{year}"
            else:
                # Annual data
                start_date = f"{period['year']}-01-01"
                end_date = f"{period['year']}-12-31"
                period_suffix = f"Annual_{year}"
            
            # Get company data
            company_data = self.yf_extractor.get_company_data(
                self.selected_company['ticker'], 
                start_date, end_date,
                'quarterly' if period['period_type'] == 'quarterly' else 'annual'
            )
            
            if company_data:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                for file_item in selected_files:
                    file_name = self.file_tree.item(file_item, 'text')
                    file_type = self.file_tree.item(file_item, 'values')[0]
                    
                    # Map file type to data
                    data_to_save = None
                    if 'Income Statement' in file_type:
                        data_to_save = company_data.get('income_stmt')
                    elif 'Balance Sheet' in file_type:
                        data_to_save = company_data.get('balance_sheet')
                    elif 'Cash Flow' in file_type:
                        data_to_save = company_data.get('cash_flow')
                    elif 'Company Info' in file_type:
                        import pandas as pd
                        data_to_save = pd.Series(company_data.get('info', {}))
                    elif 'Historical Data' in file_type:
                        data_to_save = company_data.get('historical')
                    
                    if data_to_save is not None and not data_to_save.empty:
                        # Enhanced filename with period information
                        filename = f"{self.selected_company['ticker']}-{file_type.replace(' ', '_')}-{period_suffix}-{timestamp}.csv"
                        
                        if self.storage_var.get() == "local":
                            file_path = download_dir / filename
                            data_to_save.to_csv(file_path)
                            extracted_files.append(str(file_path))
                            
                            file_size = os.path.getsize(file_path) / 1024  # Size in KB
                            print(f"   ✅ Saved {file_type} for {period_suffix} ({file_size:.1f} KB)")
                        else:
                            # For ZIP download
                            extracted_files.append({
                                'data': data_to_save,
                                'filename': filename,
                                'period': period,
                                'type': file_type,
                                'year': period_year,
                                'period_type': period_type
                            })
        
        return extracted_files
    
    def create_zip_file(self, extracted_files):
        """Create ZIP file with extracted data using organized folder structure"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"{self.selected_company['ticker']}_financial_data_{timestamp}.zip"
        zip_path = Path.home() / "Downloads" / zip_filename
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            if self.selected_company['source'] == 'sec':
                # Handle SEC files with organized folder structure
                for file_info in extracted_files:
                    if isinstance(file_info, dict):
                        # Download and add to ZIP with folder structure
                        try:
                            response = requests.get(file_info['url'], 
                                                headers=self.sec_finder.headers, 
                                                timeout=30)
                            response.raise_for_status()
                            
                            # Create organized path in ZIP: TICKER/FORM/YEAR/filename
                            zip_path_in_archive = f"{self.selected_company['ticker']}/{file_info['period']['form']}/{file_info['year']}/{file_info['filename']}"
                            zipf.writestr(zip_path_in_archive, response.content)
                            
                        except Exception as e:
                            print(f"Failed to add {file_info['filename']} to ZIP: {e}")
                    else:
                        # Local file path
                        if os.path.exists(file_info):
                            # Preserve folder structure in ZIP
                            rel_path = os.path.relpath(file_info, Path.home() / "Downloads" / "FinancialData")
                            zipf.write(file_info, rel_path)
            else:
                # Handle Yahoo Finance data with organized folder structure
                for file_info in extracted_files:
                    if isinstance(file_info, dict) and 'data' in file_info:
                        csv_content = file_info['data'].to_csv()
                        
                        # Create organized path in ZIP: TICKER/PERIOD_TYPE/YEAR/filename
                        zip_path_in_archive = f"{self.selected_company['ticker']}/{file_info['period_type']}/{file_info['year']}/{file_info['filename']}"
                        zipf.writestr(zip_path_in_archive, csv_content)
                        
                    elif isinstance(file_info, str) and os.path.exists(file_info):
                        # Preserve folder structure in ZIP
                        rel_path = os.path.relpath(file_info, Path.home() / "Downloads" / "FinancialData")
                        zipf.write(file_info, rel_path)
        
        return str(zip_path)
    
    def show_extraction_success(self, message):
        """Show extraction success message"""
        success_window = tk.Toplevel(self.root)
        success_window.title("Extraction Complete")
        success_window.geometry("400x200")
        success_window.resizable(False, False)
        
        # Center the window
        success_window.transient(self.root)
        success_window.grab_set()
        
        frame = ttk.Frame(success_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Success icon and message
        ttk.Label(frame, text="✅", font=('Arial', 24)).pack(pady=(0, 10))
        ttk.Label(frame, text="Data extraction completed successfully!", 
                 font=('Arial', 12, 'bold')).pack(pady=(0, 10))
        ttk.Label(frame, text=message, font=('Arial', 10)).pack(pady=(0, 20))
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack()
        
        if "ZIP file" in message:
            ttk.Button(button_frame, text="Open Folder", 
                      command=lambda: self.open_folder(os.path.dirname(message.split(": ")[1]))).pack(side=tk.LEFT, padx=(0, 10))
        else:
            ttk.Button(button_frame, text="Open Folder", 
                      command=lambda: self.open_folder(message.split(": ")[1])).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="Close", 
                  command=success_window.destroy).pack(side=tk.LEFT)
    
    def open_folder(self, path):
        """Open folder in system file explorer"""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(path)
            elif os.name == 'posix':  # macOS and Linux
                os.system(f'open "{path}"' if sys.platform == 'darwin' else f'xdg-open "{path}"')
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {e}")
    
    def stop_progress(self):
        """Stop progress bar and reset status"""
        self.progress_bar.stop()
        self.progress_var.set("Ready")

class WelcomeWindow:
    """Enhanced welcome screen with instructions"""
    def __init__(self, root):
        self.root = root
        self.show_welcome()
    
    def show_welcome(self):
        """Show welcome dialog with enhanced styling"""
        welcome = tk.Toplevel(self.root)
        welcome.title("Welcome to Enhanced Financial Data Extractor")
        welcome.geometry("750x650")
        welcome.resizable(False, False)
        welcome.transient(self.root)
        welcome.grab_set()
        
        # Center the window
        welcome.update_idletasks()
        x = (welcome.winfo_screenwidth() // 2) - (750 // 2)
        y = (welcome.winfo_screenheight() // 2) - (650 // 2)
        welcome.geometry(f"750x650+{x}+{y}")
        
        # Create scrollable frame for welcome content
        welcome_scroll = ScrollableFrame(welcome)
        welcome_scroll.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        frame = welcome_scroll.scrollable_frame
        
        # Title section
        title_frame = ttk.Frame(frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(title_frame, text="🚀 Enhanced Financial Data Extractor", 
                 font=('Arial', 18, 'bold')).pack()
        ttk.Label(title_frame, text="Professional Financial Data Acquisition Tool", 
                 font=('Arial', 12, 'italic'), foreground='gray').pack()
        
        # Features section
        features_frame = ttk.LabelFrame(frame, text="✨ Key Features", padding="15")
        features_frame.pack(fill=tk.X, pady=(0, 15))
        
        features_text = """
🏢 Dual Data Sources:
   • SEC Edgar: Official regulatory filings (10-Q, 10-K, 8-K)
   • Yahoo Finance: Historical financial statements and market data

📊 Comprehensive File Types:
   • Excel financial reports with multiple worksheets
   • HTML formatted statements for easy viewing
   • CSV data files for analysis and modeling

🎯 Advanced Selection:
   • Multi-period selection with intuitive interface
   • File preview before downloading
   • Flexible storage options (local folders or ZIP archives)

⚡ Enhanced User Experience:
   • Fully scrollable interface
   • Real-time progress tracking
   • Detailed error handling and validation
   • Professional styling and icons
        """
        
        features_label = ttk.Label(features_frame, text=features_text.strip(), 
                                 font=('Arial', 10), justify=tk.LEFT)
        features_label.pack(anchor=tk.W)
        
        # How to use section
        howto_frame = ttk.LabelFrame(frame, text="📋 How to Use", padding="15")
        howto_frame.pack(fill=tk.X, pady=(0, 15))
        
        howto_text = """
1. 📊 Select Data Source: Choose between SEC Edgar or Yahoo Finance
2. 🔍 Search Company: Enter ticker symbol or company name
3. 📋 Choose Form Type: Select the type of financial report you need
4. 📅 Load & Select Periods: Choose specific time periods for your analysis
5. 📁 Preview Files: Review available files and select what you need
6. 💾 Configure Download: Choose local storage or ZIP download
7. 🚀 Extract Data: Start the download process and track progress
        """
        
        howto_label = ttk.Label(howto_frame, text=howto_text.strip(), 
                               font=('Arial', 10), justify=tk.LEFT)
        howto_label.pack(anchor=tk.W)
        
        # Data sources comparison
        sources_frame = ttk.LabelFrame(frame, text="📊 Data Sources Comparison", padding="15")
        sources_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Create two columns for comparison
        comparison_frame = ttk.Frame(sources_frame)
        comparison_frame.pack(fill=tk.X)
        
        # SEC Edgar column
        sec_frame = ttk.Frame(comparison_frame)
        sec_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        ttk.Label(sec_frame, text="🏛️ SEC Edgar", font=('Arial', 12, 'bold'), 
                 foreground='green').pack()
        
        sec_details = """
✅ Official regulatory filings
✅ Most accurate and comprehensive
✅ Excel and HTML formats
✅ Quarterly and annual reports
✅ Current and historical data
⚠️  Requires internet connection
⚠️  May have download delays
        """
        
        ttk.Label(sec_frame, text=sec_details.strip(), font=('Arial', 9), 
                 justify=tk.LEFT).pack(anchor=tk.W)
        
        # Yahoo Finance column
        yahoo_frame = ttk.Frame(comparison_frame)
        yahoo_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        ttk.Label(yahoo_frame, text="📈 Yahoo Finance", font=('Arial', 12, 'bold'), 
                 foreground='blue').pack()
        
        yahoo_details = """
✅ Fast and reliable access
✅ Easy-to-use CSV format
✅ Historical market data
✅ Company information
✅ Good for analysis/modeling
⚠️  Aggregated data
⚠️  May have slight delays
        """
        
        ttk.Label(yahoo_frame, text=yahoo_details.strip(), font=('Arial', 9), 
                 justify=tk.LEFT).pack(anchor=tk.W)
        
        # Tips section
        tips_frame = ttk.LabelFrame(frame, text="💡 Pro Tips", padding="15")
        tips_frame.pack(fill=tk.X, pady=(0, 15))
        
        tips_text = """
🎯 For Academic Research: Use SEC Edgar for official, audited data
📊 For Quick Analysis: Use Yahoo Finance for rapid data acquisition
🔄 Multi-Period Analysis: Select multiple quarters/years for trend analysis
📁 File Organization: Use descriptive local folder names for easy access
💾 Large Downloads: Use ZIP format for multiple companies/periods
🔍 Preview First: Always preview files to understand data structure
⚡ Scroll Interface: Use scroll buttons or mouse wheel to navigate
        """
        
        tips_label = ttk.Label(tips_frame, text=tips_text.strip(), 
                              font=('Arial', 10), justify=tk.LEFT)
        tips_label.pack(anchor=tk.W)
        
        # Example companies section
        examples_frame = ttk.LabelFrame(frame, text="🏢 Example Companies", padding="15")
        examples_frame.pack(fill=tk.X, pady=(0, 20))
        
        examples_text = """
Technology: AAPL (Apple), MSFT (Microsoft), GOOGL (Google), NVDA (NVIDIA)
Finance: JPM (JPMorgan), BAC (Bank of America), V (Visa), MA (Mastercard)
Healthcare: JNJ (Johnson & Johnson), PFE (Pfizer), UNH (UnitedHealth)
Energy: XOM (ExxonMobil), CVX (Chevron), COP (ConocoPhillips)
Consumer: AMZN (Amazon), TSLA (Tesla), WMT (Walmart), KO (Coca-Cola)
        """
        
        examples_label = ttk.Label(examples_frame, text=examples_text.strip(), 
                                  font=('Arial', 10), justify=tk.LEFT)
        examples_label.pack(anchor=tk.W)
        
        # Button frame
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=(20, 0))
        
        ttk.Button(button_frame, text="🚀 Get Started", 
                  command=welcome.destroy).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(button_frame, text="📚 Documentation", 
                  command=lambda: self.show_documentation(welcome)).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(button_frame, text="❌ Exit", 
                  command=self.root.quit).pack(side=tk.LEFT)
    
    def show_documentation(self, parent_window):
        """Show comprehensive documentation window"""
        doc_window = tk.Toplevel(parent_window)
        doc_window.title("📚 Documentation & Help")
        doc_window.geometry("900x700")
        doc_window.resizable(True, True)
        
        # Create scrollable documentation
        doc_scroll = ScrollableFrame(doc_window)
        doc_scroll.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        frame = doc_scroll.scrollable_frame
        
        # Create notebook for tabbed documentation
        notebook = ttk.Notebook(frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # SEC Edgar Tab
        sec_frame = ttk.Frame(notebook, padding="20")
        notebook.add(sec_frame, text="🏛️ SEC Edgar")
        
        sec_scroll = ScrollableFrame(sec_frame)
        sec_scroll.pack(fill=tk.BOTH, expand=True)
        
        sec_content_frame = sec_scroll.scrollable_frame
        
        ttk.Label(sec_content_frame, text="🏛️ SEC Edgar Data Source", 
                 font=('Arial', 16, 'bold')).pack(pady=(0, 15))
        
        sec_content = """
📋 Available Form Types:
• 10-Q: Quarterly reports with unaudited financial statements
• 10-K: Annual reports with audited financial statements  
• 8-K: Current reports for significant corporate events

📄 File Types Available:
• Excel Financial Report (.xlsx): Complete financial statements with multiple worksheets
• Income Statement (.htm): Revenue, expenses, and profit data in HTML format
• Balance Sheet (.htm): Assets, liabilities, and stockholders' equity
• Cash Flow Statement (.htm): Operating, investing, and financing cash flows
• Stockholder Equity (.htm): Changes in equity over time

🎯 Best Practices:
• Excel files contain the most comprehensive data with multiple worksheets
• HTML files are formatted for easy reading and printing
• 10-K reports are more detailed and comprehensive than 10-Q reports
• Check multiple periods for trend analysis and comparative studies
• Use official SEC data for academic research and regulatory compliance

⚠️ Important Notes:
• Files are downloaded directly from SEC EDGAR database
• Download speeds may vary based on SEC server load
• Some older filings may not have Excel format available
• Always verify data accuracy for critical financial decisions

🔗 Data Sources:
• All data sourced from official SEC EDGAR database
• URLs follow SEC standard naming conventions
• Files are organized by CIK (Central Index Key) and accession numbers
        """
        
        sec_text_label = ttk.Label(sec_content_frame, text=sec_content.strip(), 
                                  font=('Arial', 10), justify=tk.LEFT, wraplength=800)
        sec_text_label.pack(anchor=tk.W)
        
        # Yahoo Finance Tab
        yahoo_frame = ttk.Frame(notebook, padding="20")
        notebook.add(yahoo_frame, text="📈 Yahoo Finance")
        
        yahoo_scroll = ScrollableFrame(yahoo_frame)
        yahoo_scroll.pack(fill=tk.BOTH, expand=True)
        
        yahoo_content_frame = yahoo_scroll.scrollable_frame
        
        ttk.Label(yahoo_content_frame, text="📈 Yahoo Finance Data Source", 
                 font=('Arial', 16, 'bold')).pack(pady=(0, 15))
        
        yahoo_content = """
📊 Available Period Types:
• Annual: Yearly financial statements (up to 10 years of history)
• Quarterly: Quarterly financial statements (up to 5 years of history)

📄 File Types Available:
• Income Statement (.csv): Revenue, expenses, gross profit, net income, EPS
• Balance Sheet (.csv): Assets, liabilities, stockholders' equity, working capital
• Cash Flow (.csv): Operating, investing, financing cash flows, free cash flow
• Company Info (.csv): Market cap, sector, industry, employee count, business summary
• Historical Data (.csv): Stock prices, trading volume, dividends, splits

🎯 Best Practices:
• CSV format is ideal for data analysis and modeling
• Use quarterly data for detailed trend analysis
• Historical data is excellent for market analysis and backtesting
• Combine multiple file types for comprehensive company analysis
• Data is updated regularly and reflects market conditions

⚡ Advantages:
• Fast and reliable data access
• Standardized CSV format for easy importing
• Comprehensive historical coverage
• Regular updates and maintenance
• Works well with analysis tools like Excel, Python, R

⚠️ Limitations:
• Data is aggregated from various sources
• May have slight delays compared to real-time data
• Less detailed than official SEC filings
• Some metrics may be calculated/estimated
        """
        
        yahoo_text_label = ttk.Label(yahoo_content_frame, text=yahoo_content.strip(), 
                                    font=('Arial', 10), justify=tk.LEFT, wraplength=800)
        yahoo_text_label.pack(anchor=tk.W)
        
        # Technical Tab
        tech_frame = ttk.Frame(notebook, padding="20")
        notebook.add(tech_frame, text="🔧 Technical")
        
        tech_scroll = ScrollableFrame(tech_frame)
        tech_scroll.pack(fill=tk.BOTH, expand=True)
        
        tech_content_frame = tech_scroll.scrollable_frame
        
        ttk.Label(tech_content_frame, text="🔧 Technical Information", 
                 font=('Arial', 16, 'bold')).pack(pady=(0, 15))
        
        tech_content = """
💻 System Requirements:
• Python 3.7 or higher
• Required packages: tkinter, pandas, requests, yfinance
• Internet connection for data downloads
• Sufficient disk space for downloaded files

📁 File Organization:
• Local storage creates organized folder structure
• Files are named with ticker, type, and date information
• ZIP downloads create timestamped archive files
• Separate folders for each company and form type

🔧 Advanced Features:
• Multi-threading for responsive interface
• Progress tracking for long downloads
• Automatic retry mechanisms for failed downloads
• Comprehensive error handling and logging
• Full interface scrolling for large screens

⚙️ Configuration Options:
• Customizable download locations
• Flexible file naming conventions
• Multiple storage format options
• Batch processing capabilities
• Preview functionality before downloading

🚀 Performance Tips:
• Use local storage for frequent access
• Use ZIP format for archival purposes
• Select specific periods to reduce download time
• Close other applications during large downloads
• Ensure stable internet connection for best results

🛠️ Troubleshooting:
• Check internet connection if downloads fail
• Verify company ticker symbols are correct
• Ensure sufficient disk space for downloads
• Try different time periods if data is unavailable
• Contact support for persistent issues
        """
        
        tech_text_label = ttk.Label(tech_content_frame, text=tech_content.strip(), 
                                   font=('Arial', 10), justify=tk.LEFT, wraplength=800)
        tech_text_label.pack(anchor=tk.W)
        
        # FAQ Tab
        faq_frame = ttk.Frame(notebook, padding="20")
        notebook.add(faq_frame, text="❓ FAQ")
        
        faq_scroll = ScrollableFrame(faq_frame)
        faq_scroll.pack(fill=tk.BOTH, expand=True)
        
        faq_content_frame = faq_scroll.scrollable_frame
        
        ttk.Label(faq_content_frame, text="❓ Frequently Asked Questions", 
                 font=('Arial', 16, 'bold')).pack(pady=(0, 15))
        
        faq_content = """
Q: Which data source should I use?
A: Use SEC Edgar for official regulatory data and Yahoo Finance for quick analysis. SEC data is more comprehensive but slower to download.

Q: How many periods can I select at once?
A: You can select multiple periods using Ctrl+click (Windows) or Cmd+click (Mac). The system handles batch downloads efficiently.

Q: What file formats are available?
A: SEC Edgar provides Excel (.xlsx) and HTML (.htm) files. Yahoo Finance provides CSV (.csv) files that work well with analysis tools.

Q: Can I download data for multiple companies?
A: Currently, the tool processes one company at a time. You can run multiple sessions or restart the tool for different companies.

Q: How is the data organized when saved locally?
A: Files are organized in folders by company ticker, then by form type, with descriptive filenames including dates.

Q: What if a download fails?
A: The system includes retry mechanisms and error handling. Check your internet connection and try again with a smaller date range.

Q: Can I preview files before downloading?
A: Yes! Use the "Preview Selected" button to see file structure and expected content before downloading.

Q: Is there a limit on download size?
A: No artificial limits, but large downloads may take longer. Consider using ZIP format for multiple periods.

Q: How current is the data?
A: SEC data is updated as companies file reports. Yahoo Finance data is updated regularly but may have slight delays.

Q: Can I use this for commercial purposes?
A: The tool itself can be used commercially, but check the terms of service for SEC Edgar and Yahoo Finance data usage.
        """
        
        faq_text_label = ttk.Label(faq_content_frame, text=faq_content.strip(), 
                                  font=('Arial', 10), justify=tk.LEFT, wraplength=800)
        faq_text_label.pack(anchor=tk.W)
        
        # Close button
        close_frame = ttk.Frame(frame)
        close_frame.pack(pady=(20, 0))
        
        ttk.Button(close_frame, text="✅ Close Documentation", 
                  command=doc_window.destroy).pack()


def main():
    """Main function to run the application"""
    try:
        # Initialize root window with professional styling
        root = tk.Tk()
        root.title("Enhanced Financial Data Extractor")
        
        # Configure professional styling
        style = ttk.Style()
        available_themes = style.theme_names()
        
        # Try to set a professional theme
        if 'clam' in available_themes:
            style.theme_use('clam')
        elif 'vista' in available_themes:
            style.theme_use('vista')
        elif 'default' in available_themes:
            style.theme_use('default')
            
        # Configure common style elements
        style.configure('TLabel', font=('Arial', 10))
        style.configure('TButton', font=('Arial', 10))
        style.configure('TFrame', background='#f0f0f0')
        
        # Initialize the main application
        app = EnhancedFinancialExtractor(root)
        
        # Show welcome screen
        welcome = WelcomeWindow(root)
        
        # Start the application
        root.mainloop()
        
    except Exception as e:
        # Handle any initialization errors
        error_msg = f"Application initialization failed: {str(e)}"
        if 'root' in locals():
            messagebox.showerror("Error", error_msg)
        else:
            print(error_msg)
        sys.exit(1)


if __name__ == "__main__":
    main()