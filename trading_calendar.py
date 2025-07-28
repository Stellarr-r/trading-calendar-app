#!/usr/bin/env python3
"""
Trading Calendar

Version: 1.0.0
Author: Star
Repository: https://github.com/Stellarr-r/trading-calendar-app

"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import json
import shutil
import threading
import time
import tempfile
from datetime import datetime, timedelta
from collections import defaultdict
import calendar

# Check and install dependencies
def check_dependencies():
    """Check and install required packages"""
    required_packages = ['openpyxl', 'requests']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"Installing required packages: {', '.join(missing_packages)}")
        import subprocess
        for package in missing_packages:
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
                print(f"✅ {package} installed successfully")
            except subprocess.CalledProcessError:
                print(f"❌ Failed to install {package}")
                return False
    
    return True

# Install dependencies before importing them
if not check_dependencies():
    print("Failed to install required dependencies")
    input("Press Enter to exit...")
    sys.exit(1)

# Now import the dependencies
import openpyxl
import requests

# ============================================================================
# EMBEDDED CONFIGURATION
# ============================================================================

DEFAULT_CONFIG = {
    "theme": {
        "bg_dark": "#0f0f0f",
        "bg_card": "#1a1a1a",
        "bg_accent": "#262626",
        "text_primary": "#ffffff",
        "text_secondary": "#9ca3af",
        "text_muted": "#6b7280",
        "accent_blue": "#3b82f6",
        "accent_green": "#10b981",
        "accent_red": "#ef4444",
        "accent_yellow": "#f59e0b",
        "border": "#374151",
        "hover": "#1f2937"
    },
    "excel_columns": {
        "datetime": "Date/Time",
        "pnl": "P&L USD",
        "trade_num": "Trade #",
        "type": "Type"
    },
    "sheet_names": [
        "List of trades",
        "list of trades",
        "trades"
    ],
    "ui": {
        "window_width": 1200,
        "window_height": 800,
        "loading_timeout": 2000,
        "progress_update_interval": 10
    },
    "version": "1.0.0"
}

# ============================================================================
# DATA MANAGER CLASS
# ============================================================================

class DataManager:
    """Handles data persistence and history management"""
    
    def __init__(self, data_dir):
        self.data_dir = data_dir
        self.ensure_data_directory()
    
    def ensure_data_directory(self):
        """Create data directory if it doesn't exist"""
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir, exist_ok=True)
    
    def save_trading_data(self, filename, trades, daily_pnl, stats):
        """Save trading data with timestamp for history"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(os.path.basename(filename))[0]
            save_file = os.path.join(self.data_dir, f"{base_name}_{timestamp}.json")
            
            # Convert trades to serializable format
            serializable_trades = []
            for trade in trades:
                trade_data = trade.copy()
                if 'date' in trade_data:
                    trade_data['date'] = trade_data['date'].isoformat()
                serializable_trades.append(trade_data)
            
            data = {
                'timestamp': timestamp,
                'original_filename': os.path.basename(filename),
                'trades': serializable_trades,
                'daily_pnl': dict(daily_pnl),
                'stats': stats,
                'total_trades': len(trades)
            }
            
            with open(save_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
            
            print(f"💾 Trading data saved: {os.path.basename(save_file)}")
            return save_file
            
        except Exception as e:
            print(f"❌ Error saving data: {e}")
            return None
    
    def get_data_history(self):
        """Get list of saved trading data files"""
        try:
            if not os.path.exists(self.data_dir):
                return []
            
            files = []
            for filename in os.listdir(self.data_dir):
                if filename.endswith('.json'):
                    filepath = os.path.join(self.data_dir, filename)
                    files.append({
                        'filename': filename,
                        'path': filepath,
                        'modified': os.path.getmtime(filepath)
                    })
            
            # Sort by modification time (newest first)
            files.sort(key=lambda x: x['modified'], reverse=True)
            return files
            
        except Exception as e:
            print(f"❌ Error getting data history: {e}")
            return []

# ============================================================================
# TRADE PROCESSOR CLASS
# ============================================================================

class TradeProcessor:
    """Handles processing of TradingView Excel exports"""
    
    def __init__(self, config=None):
        """Initialize with configuration"""
        self.config = config or DEFAULT_CONFIG
        self.trades = []
        self.daily_pnl = defaultdict(float)
        self.stats = {}
    
    def process_file(self, file_path, progress_callback=None):
        """
        Process TradingView Excel export
        
        Args:
            file_path (str): Path to Excel file
            progress_callback (callable): Optional callback for progress updates
            
        Returns:
            dict: Processing results with trades, daily_pnl, and stats
        """
        try:
            if progress_callback:
                progress_callback("Opening Excel file...")
            
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            
            if progress_callback:
                progress_callback("Scanning for trading data sheets...")
            
            sheet = self._find_trades_sheet(wb, progress_callback)
            if not sheet:
                raise ValueError("Could not find trading data sheet")
            
            if progress_callback:
                progress_callback("Analyzing column headers...")
            
            headers = self._analyze_headers(sheet, progress_callback)
            
            if progress_callback:
                progress_callback("Processing trade data...")
            
            self._process_trades(sheet, headers, progress_callback)
            
            if progress_callback:
                progress_callback("Calculating statistics...")
            
            self.calculate_stats()
            
            return {
                'trades': self.trades,
                'daily_pnl': dict(self.daily_pnl),
                'stats': self.stats
            }
            
        except Exception as e:
            raise Exception(f"Error processing Excel file: {str(e)}")
    
    def _find_trades_sheet(self, wb, progress_callback=None):
        """Find the trades sheet in the workbook"""
        sheet_names = self.config.get("sheet_names", ["List of trades"])
        
        for sheet_name in wb.sheetnames:
            for target_name in sheet_names:
                if target_name.lower() in sheet_name.lower():
                    if progress_callback:
                        progress_callback(f"Found sheet: '{sheet_name}'")
                    return wb[sheet_name]
        
        return None
    
    def _analyze_headers(self, sheet, progress_callback=None):
        """Analyze column headers in the sheet"""
        headers = {}
        column_config = self.config.get("excel_columns", {})
        first_row = list(sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]
        
        for i, header in enumerate(first_row):
            if header:
                header_str = str(header).strip()
                
                # Check against configured column names
                for key, expected_name in column_config.items():
                    if header_str == expected_name:
                        headers[key] = i
                        if progress_callback:
                            progress_callback(f"Found {expected_name} column")
                        break
        
        # Validate required columns
        required_columns = ['datetime', 'pnl']
        missing_columns = [col for col in required_columns if col not in headers]
        
        if missing_columns:
            column_names = [column_config.get(col, col) for col in missing_columns]
            raise ValueError(f"Missing required columns: {', '.join(column_names)}")
        
        return headers
    
    def _process_trades(self, sheet, headers, progress_callback=None):
        """Process all trades from the sheet"""
        self.trades = []
        self.daily_pnl = defaultdict(float)
        processed_trade_numbers = set()
        
        total_rows = sheet.max_row - 1  # Exclude header
        processed_rows = 0
        progress_interval = self.config.get("ui", {}).get("progress_update_interval", 10)
        
        for row in sheet.iter_rows(min_row=2, values_only=True):
            processed_rows += 1
            
            # Progress update
            if progress_callback and processed_rows % progress_interval == 0:
                progress = int((processed_rows / total_rows) * 100)
                progress_callback(f"Processing trades... {progress}% ({processed_rows}/{total_rows})")
            
            try:
                trade = self._process_single_trade(row, headers, processed_trade_numbers)
                if trade:
                    self.trades.append(trade)
                    
                    # Add to daily P&L
                    date_key = trade['date'].strftime('%Y-%m-%d')
                    self.daily_pnl[date_key] += trade['pnl']
                    
            except Exception as e:
                # Log individual trade processing errors but continue
                continue
        
        if progress_callback:
            progress_callback(f"Successfully processed {len(self.trades)} unique trades")
    
    def _process_single_trade(self, row, headers, processed_trade_numbers):
        """Process a single trade row"""
        trade_num = row[headers['trade_num']] if 'trade_num' in headers else None
        datetime_val = row[headers['datetime']]
        pnl_val = row[headers['pnl']]
        
        # Validate required fields
        if not datetime_val or pnl_val is None:
            return None
        
        # Skip duplicates by trade number
        if trade_num and trade_num in processed_trade_numbers:
            return None
        
        # Parse and validate date
        trade_date = self._parse_date(datetime_val)
        if not trade_date:
            return None
        
        # Parse and validate P&L
        pnl = self._parse_pnl(pnl_val)
        if pnl is None:
            return None
        
        # Mark as processed
        if trade_num:
            processed_trade_numbers.add(trade_num)
        
        return {
            'date': trade_date,
            'pnl': pnl,
            'trade_num': trade_num
        }
    
    def _parse_date(self, datetime_val):
        """Parse date value with validation"""
        try:
            if isinstance(datetime_val, datetime):
                parsed_date = datetime_val.date()
            else:
                # Try to parse string date
                date_str = str(datetime_val).split()[0]
                parsed_date = datetime.strptime(date_str, '%Y-%m-%d').date()
            
            # Validate date is reasonable (not in future, not too old)
            if parsed_date > datetime.now().date():
                return None
            if parsed_date.year < 2000:  # Reasonable minimum year
                return None
                
            return parsed_date
        except (ValueError, TypeError, AttributeError):
            return None
    
    def _parse_pnl(self, pnl_val):
        """Parse P&L value with validation"""
        try:
            if isinstance(pnl_val, (int, float)):
                return float(pnl_val)
            elif isinstance(pnl_val, str):
                # Clean string and convert
                cleaned = pnl_val.replace(',', '').replace('$', '').strip()
                if not cleaned or cleaned == '-':
                    return None
                return float(cleaned)
            else:
                return None
        except (ValueError, TypeError):
            return None
    
    def calculate_stats(self):
        """Calculate trading statistics"""
        if not self.trades:
            self.stats = {
                'total_pnl': 0.0,
                'win_rate': 0.0,
                'avg_daily': 0.0,
                'total_trades': 0
            }
            return
        
        total_pnl = sum(trade['pnl'] for trade in self.trades)
        winning_trades = sum(1 for trade in self.trades if trade['pnl'] > 0)
        total_trades = len(self.trades)
        win_rate = (winning_trades / total_trades * 100) if total_trades > 0 else 0
        
        trading_days = len([pnl for pnl in self.daily_pnl.values() if pnl != 0])
        avg_daily = total_pnl / trading_days if trading_days > 0 else 0
        
        self.stats = {
            'total_pnl': total_pnl,
            'win_rate': win_rate,
            'avg_daily': avg_daily,
            'total_trades': total_trades
        }

# ============================================================================
# MAIN TRADING CALENDAR APPLICATION
# ============================================================================

class ModernTradingCalendar:
    def __init__(self, root):
        self.root = root
        self.root.title("Trading Calendar")
        
        # Load configuration
        self.config = DEFAULT_CONFIG
        ui_config = self.config.get("ui", {})
        
        self.root.geometry(f"{ui_config.get('window_width', 1200)}x{ui_config.get('window_height', 800)}")
        self.root.configure(bg=self.config['theme']['bg_dark'])
        self.root.resizable(True, True)
        
        # Get version from batch file environment variable, fallback to config
        self.CURRENT_VERSION = os.environ.get('TRADING_CALENDAR_VERSION', self.config.get('version', '1.0.0'))
        
        # Validate and display version info
        self.validate_version()
        
        # Setup data persistence
        data_dir = os.environ.get('TRADING_CALENDAR_DATA_DIR', os.path.join(os.getcwd(), 'data'))
        self.data_manager = DataManager(data_dir)
        
        # Setup trade processor
        self.trade_processor = TradeProcessor(self.config)
        
        # Theme from config
        self.theme = self.config['theme']
        
        # Data
        self.trades = []
        self.daily_pnl = defaultdict(float)
        self.stats = {}
        self.current_month = datetime.now().month
        self.current_year = datetime.now().year
        self.current_file = None
        
        # Initialize app immediately
        self.setup_ui()
    
    def validate_version(self):
        """Validate version format and show startup info"""
        import re
        
        # Check if version matches X.Y.Z format
        version_pattern = r'^\d+\.\d+\.\d+$'
        if not re.match(version_pattern, self.CURRENT_VERSION):
            print(f"⚠️  Warning: Invalid version format '{self.CURRENT_VERSION}'. Using fallback.")
            self.CURRENT_VERSION = "1.0.0"
        
        # Show version info in console
        print(f"🚀 Trading Calendar v{self.CURRENT_VERSION} starting...")
        
        # Check version source
        version_source = os.environ.get('TRADING_CALENDAR_VERSION')
        if version_source:
            print(f"📋 Version loaded from TradingCalendar.bat: {version_source}")
        else:
            print(f"📋 Using default version (run via TradingCalendar.bat for version management)")
        
    def setup_ui(self):
        """Create modern UI layout - optimized for speed"""
        # Main container
        main_container = tk.Frame(self.root, bg=self.theme['bg_dark'])
        main_container.pack(fill='both', expand=True, padx=0, pady=0)
        
        # Top bar - create immediately
        self.create_top_bar(main_container)
        
        # Content area
        content_frame = tk.Frame(main_container, bg=self.theme['bg_dark'])
        content_frame.pack(fill='both', expand=True, padx=20, pady=0)
        
        # Create placeholder and defer heavy UI elements
        loading_label = tk.Label(content_frame, 
                               text="Loading...", 
                               font=('Inter', 12),
                               fg=self.theme['text_muted'],
                               bg=self.theme['bg_dark'])
        loading_label.pack(pady=50)
        
        # Create heavy UI elements after a brief delay
        self.root.after(100, lambda: self.create_heavy_ui_elements(content_frame, loading_label))
        
    def create_heavy_ui_elements(self, content_frame, loading_label):
        """Create the heavy UI elements after initial load"""
        # Remove loading label
        loading_label.destroy()
        
        # Stats section
        self.create_stats_section(content_frame)
        
        # Calendar section  
        self.create_calendar_section(content_frame)
        
        # Bottom info
        self.create_bottom_info(content_frame)
        
    def create_top_bar(self, parent):
        """Modern top navigation bar"""
        top_bar = tk.Frame(parent, bg=self.theme['bg_card'], height=70)
        top_bar.pack(fill='x', padx=0, pady=0)
        top_bar.pack_propagate(False)
        
        # Left side - Logo/Title
        left_frame = tk.Frame(top_bar, bg=self.theme['bg_card'])
        left_frame.pack(side='left', fill='y', padx=20)
        
        title = tk.Label(left_frame,
                        text="Trading Calendar",
                        font=('Inter', 20, 'bold'),
                        fg=self.theme['text_primary'],
                        bg=self.theme['bg_card'])
        title.pack(side='left', pady=20)
        
        # Version indicator next to title  
        version_label = tk.Label(left_frame,
                               text=f"v{self.CURRENT_VERSION}",
                               font=('Inter', 10),
                               fg=self.theme['text_muted'],
                               bg=self.theme['bg_card'])
        version_label.pack(side='left', padx=(10, 0), pady=20)
        
        # Right side - Actions
        right_frame = tk.Frame(top_bar, bg=self.theme['bg_card'])
        right_frame.pack(side='right', fill='y', padx=20)
        
        # Load button
        load_btn = tk.Button(right_frame,
                            text="Load File",
                            font=('Inter', 10, 'bold'),
                            bg=self.theme['accent_blue'],
                            fg='white',
                            border=0,
                            padx=20,
                            pady=8,
                            cursor='hand2',
                            activebackground='#2563eb',
                            command=self.load_file)
        load_btn.pack(side='right', pady=20)
        
        # File status
        self.file_status = tk.Label(left_frame,
                                   text="No file loaded",
                                   font=('Inter', 9),
                                   fg=self.theme['text_muted'],
                                   bg=self.theme['bg_card'])
        self.file_status.pack(side='left', padx=(20, 0), pady=20)
        
    def create_stats_section(self, parent):
        """Modern statistics cards"""
        stats_container = tk.Frame(parent, bg=self.theme['bg_dark'])
        stats_container.pack(fill='x', pady=20)
        
        # Stats grid
        stats_grid = tk.Frame(stats_container, bg=self.theme['bg_dark'])
        stats_grid.pack(fill='x')
        
        self.stat_cards = {}
        stats_config = [
            ("Total P&L", "$0.00", "total_pnl"),
            ("Win Rate", "0%", "win_rate"),
            ("Avg Daily", "$0.00", "avg_daily"),
            ("Total Trades", "0", "total_trades")
        ]
        
        for i, (label, value, key) in enumerate(stats_config):
            card = self.create_stat_card(stats_grid, label, value)
            card.grid(row=0, column=i, padx=(0, 15 if i < 3 else 0), sticky='ew')
            self.stat_cards[key] = card
            stats_grid.grid_columnconfigure(i, weight=1)
    
    def create_stat_card(self, parent, label, value):
        """Create a modern stat card"""
        card = tk.Frame(parent, 
                       bg=self.theme['bg_card'],
                       relief='flat',
                       bd=0)
        
        # Content
        content = tk.Frame(card, bg=self.theme['bg_card'])
        content.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Value
        value_label = tk.Label(content,
                              text=value,
                              font=('Inter', 24, 'bold'),
                              fg=self.theme['text_primary'],
                              bg=self.theme['bg_card'])
        value_label.pack(anchor='w')
        
        # Label
        label_widget = tk.Label(content,
                               text=label,
                               font=('Inter', 11),
                               fg=self.theme['text_secondary'],
                               bg=self.theme['bg_card'])
        label_widget.pack(anchor='w', pady=(5, 0))
        
        # Store references
        card.value_label = value_label
        card.label_widget = label_widget
        
        return card
    
    def create_calendar_section(self, parent):
        """Modern calendar display"""
        calendar_container = tk.Frame(parent, bg=self.theme['bg_dark'])
        calendar_container.pack(fill='both', expand=True, pady=(20, 0))
        
        # Calendar header
        cal_header = tk.Frame(calendar_container, bg=self.theme['bg_dark'])
        cal_header.pack(fill='x', pady=(0, 20))
        
        # Navigation
        nav_frame = tk.Frame(cal_header, bg=self.theme['bg_dark'])
        nav_frame.pack()
        
        # Previous button
        prev_btn = tk.Button(nav_frame,
                            text="‹",
                            font=('Inter', 18),
                            bg=self.theme['bg_accent'],
                            fg=self.theme['text_primary'],
                            border=0,
                            width=3,
                            pady=8,
                            cursor='hand2',
                            activebackground=self.theme['hover'],
                            command=self.prev_month)
        prev_btn.pack(side='left')
        
        # Month display
        month_name = calendar.month_name[self.current_month]
        self.month_display = tk.Label(nav_frame,
                                     text=f"{month_name} {self.current_year}",
                                     font=('Inter', 16, 'bold'),
                                     fg=self.theme['text_primary'],
                                     bg=self.theme['bg_dark'],
                                     padx=30)
        self.month_display.pack(side='left')
        
        # Next button
        next_btn = tk.Button(nav_frame,
                            text="›",
                            font=('Inter', 18),
                            bg=self.theme['bg_accent'],
                            fg=self.theme['text_primary'],
                            border=0,
                            width=3,
                            pady=8,
                            cursor='hand2',
                            activebackground=self.theme['hover'],
                            command=self.next_month)
        next_btn.pack(side='left')
        
        # Calendar grid
        self.calendar_frame = tk.Frame(calendar_container, bg=self.theme['bg_dark'])
        self.calendar_frame.pack(fill='both', expand=True)
        
        self.setup_calendar_grid()
        
    def setup_calendar_grid(self):
        """Setup calendar grid layout"""
        # Day headers
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        for i, day in enumerate(days):
            header = tk.Label(self.calendar_frame,
                             text=day,
                             font=('Inter', 11, 'bold'),
                             fg=self.theme['text_secondary'],
                             bg=self.theme['bg_accent'],
                             pady=12)
            header.grid(row=0, column=i, sticky='ew', padx=1, pady=1)
        
        # Configure grid
        for i in range(7):
            self.calendar_frame.grid_columnconfigure(i, weight=1)
        for i in range(7):
            self.calendar_frame.grid_rowconfigure(i, weight=1)
            
        self.update_calendar()
    
    def create_day_cell(self, day, pnl=0, row=0, col=0):
        """Create modern day cell"""
        # Determine styling
        if pnl > 0:
            bg_color = self.theme['accent_green']
            text_color = 'white'
        elif pnl < 0:
            bg_color = self.theme['accent_red']
            text_color = 'white'
        else:
            bg_color = self.theme['bg_card']
            text_color = self.theme['text_secondary']
        
        # Cell container
        cell = tk.Frame(self.calendar_frame,
                       bg=bg_color,
                       relief='flat',
                       bd=0)
        cell.grid(row=row, column=col, sticky='nsew', padx=1, pady=1)
        
        # Cell content
        content = tk.Frame(cell, bg=bg_color)
        content.pack(fill='both', expand=True, padx=12, pady=12)
        
        # Day number
        day_label = tk.Label(content,
                            text=str(day),
                            font=('Inter', 14, 'bold'),
                            fg=text_color,
                            bg=bg_color)
        day_label.pack(anchor='nw')
        
        # P&L if trading day
        if pnl != 0:
            pnl_text = f"${pnl:,.0f}" if abs(pnl) >= 1 else f"${pnl:.2f}"
            pnl_label = tk.Label(content,
                                text=pnl_text,
                                font=('Inter', 10),
                                fg=text_color,
                                bg=bg_color)
            pnl_label.pack(anchor='sw')
    
    def create_bottom_info(self, parent):
        """Bottom information panel"""
        bottom_frame = tk.Frame(parent, bg=self.theme['bg_dark'])
        bottom_frame.pack(fill='x', pady=(20, 10))
        
        # Legend
        legend_frame = tk.Frame(bottom_frame, bg=self.theme['bg_dark'])
        legend_frame.pack(side='left')
        
        tk.Label(legend_frame,
                text="Legend:",
                font=('Inter', 10, 'bold'),
                fg=self.theme['text_primary'],
                bg=self.theme['bg_dark']).pack(side='left')
        
        legend_items = [
            ("● Profit", self.theme['accent_green']),
            ("● Loss", self.theme['accent_red']),
            ("● No Trading", self.theme['text_muted'])
        ]
        
        for text, color in legend_items:
            tk.Label(legend_frame,
                    text=text,
                    font=('Inter', 10),
                    fg=color,
                    bg=self.theme['bg_dark']).pack(side='left', padx=(15, 0))
    
    def load_file(self):
        """Load and process Excel file with loading screen"""
        file_path = filedialog.askopenfilename(
            title="Select TradingView Export",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        # Store current file for data saving
        self.current_file = file_path
            
        # Show loading dialog
        self.show_loading_dialog(file_path)
    
    def show_loading_dialog(self, file_path):
        """Show loading progress dialog"""
        self.loading_dialog = tk.Toplevel(self.root)
        self.loading_dialog.title("Processing Excel File")
        self.loading_dialog.geometry("500x350")
        self.loading_dialog.configure(bg=self.theme['bg_dark'])
        self.loading_dialog.transient(self.root)
        self.loading_dialog.grab_set()
        
        # Center the dialog
        self.loading_dialog.geometry("+{}+{}".format(
            int(self.root.winfo_screenwidth()/2 - 250),
            int(self.root.winfo_screenheight()/2 - 175)
        ))
        
        # Allow closing with Escape key but prevent X button
        self.loading_dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        self.loading_dialog.bind('<Escape>', lambda e: self.continue_to_calendar())
        self.loading_dialog.bind('<Return>', lambda e: self.continue_to_calendar())
        self.loading_dialog.attributes('-topmost', True)
        
        # Main container
        main_frame = tk.Frame(self.loading_dialog, bg=self.theme['bg_dark'])
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Title
        title_label = tk.Label(main_frame,
                              text="Processing Excel File",
                              font=('Inter', 18, 'bold'),
                              fg=self.theme['accent_blue'],
                              bg=self.theme['bg_dark'])
        title_label.pack(pady=(0, 20))
        
        # File name
        filename = os.path.basename(file_path)
        file_label = tk.Label(main_frame,
                             text=f"File: {filename}",
                             font=('Inter', 12),
                             fg=self.theme['text_secondary'],
                             bg=self.theme['bg_dark'])
        file_label.pack(pady=(0, 20))
        
        # Progress log area
        log_frame = tk.Frame(main_frame, bg=self.theme['bg_card'])
        log_frame.pack(fill='both', expand=True, pady=(0, 20))
        
        # Log header
        log_header = tk.Label(log_frame,
                             text="Processing Log:",
                             font=('Inter', 12, 'bold'),
                             fg=self.theme['text_primary'],
                             bg=self.theme['bg_card'])
        log_header.pack(anchor='w', padx=15, pady=(15, 5))
        
        # Log text
        self.loading_log = tk.Text(log_frame,
                                  font=('Consolas', 10),
                                  bg=self.theme['bg_dark'],
                                  fg=self.theme['text_primary'],
                                  border=0,
                                  height=10,
                                  wrap='word')
        self.loading_log.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        
        # Continue button (initially hidden)
        self.continue_btn = tk.Button(main_frame,
                                     text="Continue to Calendar",
                                     font=('Inter', 12, 'bold'),
                                     bg=self.theme['accent_green'],
                                     fg='white',
                                     border=0,
                                     padx=30,
                                     pady=12,
                                     cursor='hand2',
                                     command=self.continue_to_calendar)
        # Don't pack yet - will show after processing
        
        # Start processing in background
        self.process_file_background(file_path)
        
    def process_file_background(self, file_path):
        """Process file in background thread"""
        processing_thread = threading.Thread(
            target=self.process_file_with_logging,
            args=(file_path,),
            daemon=True
        )
        processing_thread.start()
        
    def process_file_with_logging(self, file_path):
        """Process file with detailed logging using TradeProcessor"""
        try:
            # Use the embedded trade processor
            result = self.trade_processor.process_file(file_path, self.log_loading)
            
            # Update application data
            self.trades = result['trades']
            self.daily_pnl = defaultdict(float, result['daily_pnl'])
            self.stats = result['stats']
            
            # Show final statistics
            total_pnl = self.stats.get('total_pnl', 0)
            total_trades = self.stats.get('total_trades', 0)
            win_rate = self.stats.get('win_rate', 0)
            
            self.log_loading(f"Total P&L: ${total_pnl:,.2f}")
            self.log_loading(f"Total Trades: {total_trades}")
            self.log_loading(f"Win Rate: {win_rate:.1f}%")
            
            self.log_loading("Preparing calendar display...")
            
            # Update UI on main thread
            self.loading_dialog.after(0, self.finish_processing)
            
        except Exception as e:
            error_msg = str(e)
            self.log_loading(f"ERROR: {error_msg}")
            self.loading_dialog.after(0, lambda: self.show_processing_error(error_msg))
    
    def log_loading(self, message):
        """Add message to loading log"""
        def update_log():
            self.loading_log.insert('end', f"{message}\n")
            self.loading_log.see('end')
            self.loading_dialog.update()
        
        self.loading_dialog.after(0, update_log)
        time.sleep(0.1)  # Brief pause for visual effect
    
    def finish_processing(self):
        """Finish processing and show continue button"""
        self.log_loading("Processing complete! Ready to view calendar.")
        
        # Show continue button - make sure it's visible
        self.continue_btn.pack(pady=(10, 0))
        self.continue_btn.focus_set()  # Focus on the button
        
        # Update main window displays
        self.update_displays()
        
        # Save trading data for history
        if self.current_file and self.trades and self.stats:
            self.data_manager.save_trading_data(self.current_file, self.trades, self.daily_pnl, self.stats)
        
        # Auto-close the dialog after configured timeout
        timeout = self.config.get('ui', {}).get('loading_timeout', 2000)
        self.loading_dialog.after(timeout, self.continue_to_calendar)
        
        filename = os.path.basename(self.current_file) if self.current_file else "Unknown"
        self.file_status.config(text=f"Loaded: {filename}")
    
    def show_processing_error(self, error_msg):
        """Show processing error"""
        # Show error button instead of continue
        error_btn = tk.Button(self.loading_dialog,
                             text="Close",
                             font=('Inter', 12, 'bold'),
                             bg=self.theme['accent_red'],
                             fg='white',
                             border=0,
                             padx=30,
                             pady=12,
                             cursor='hand2',
                             command=self.close_loading_dialog)
        error_btn.pack(pady=(10, 0))
        
        # Also show error messagebox
        self.loading_dialog.after(100, lambda: messagebox.showerror("Processing Error", f"Failed to process file: {error_msg}"))
    
    def continue_to_calendar(self):
        """Continue to main calendar view"""
        self.close_loading_dialog()
    
    def close_loading_dialog(self):
        """Close the loading dialog"""
        if hasattr(self, 'loading_dialog') and self.loading_dialog:
            self.loading_dialog.destroy()
    
    def update_displays(self):
        """Update all displays with new data"""
        self.update_stats_display()
        self.update_calendar()
    
    def update_stats_display(self):
        """Update statistics cards"""
        if not self.stats:
            return
        
        # Total P&L
        pnl_color = self.theme['accent_green'] if self.stats['total_pnl'] >= 0 else self.theme['accent_red']
        self.stat_cards['total_pnl'].value_label.config(
            text=f"${self.stats['total_pnl']:,.2f}",
            fg=pnl_color
        )
        
        # Win Rate
        self.stat_cards['win_rate'].value_label.config(
            text=f"{self.stats['win_rate']:.1f}%"
        )
        
        # Average Daily
        avg_color = self.theme['accent_green'] if self.stats['avg_daily'] >= 0 else self.theme['accent_red']
        self.stat_cards['avg_daily'].value_label.config(
            text=f"${self.stats['avg_daily']:.2f}",
            fg=avg_color
        )
        
        # Total Trades
        self.stat_cards['total_trades'].value_label.config(
            text=str(self.stats['total_trades'])
        )
    
    def update_calendar(self):
        """Update calendar display"""
        # Clear existing cells (keep headers)
        for widget in self.calendar_frame.winfo_children():
            if int(widget.grid_info()['row']) > 0:
                widget.destroy()
        
        # Generate calendar
        cal = calendar.monthcalendar(self.current_year, self.current_month)
        
        for week_num, week in enumerate(cal):
            for day_num, day in enumerate(week):
                if day == 0:
                    continue
                    
                row = week_num + 1
                col = day_num
                
                # Get P&L for this day
                date_key = f"{self.current_year}-{self.current_month:02d}-{day:02d}"
                pnl = self.daily_pnl.get(date_key, 0)
                
                self.create_day_cell(day, pnl, row, col)
        
        # Update month display
        month_name = calendar.month_name[self.current_month]
        self.month_display.config(text=f"{month_name} {self.current_year}")
    
    def prev_month(self):
        """Navigate to previous month"""
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.update_calendar()
    
    def next_month(self):
        """Navigate to next month"""
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.update_calendar()

# ============================================================================
# MAIN APPLICATION ENTRY POINT
# ============================================================================

def main():
    """Main application entry point"""
    try:
        print("🚀 Trading Calendar - Complete Self-Contained Application")
        print("=" * 60)
        
        # Create and run the application
        root = tk.Tk()
        app = ModernTradingCalendar(root)
        
        print("✅ Application initialized successfully")
        print("📊 Ready to process TradingView Excel exports")
        print("=" * 60)
        
        root.mainloop()
        
    except ImportError as e:
        # Show error if dependencies missing
        print(f"❌ Missing dependency: {e}")
        print("\nThis shouldn't happen as dependencies are auto-installed.")
        print("Please ensure you have internet access and pip is available.")
        input("\nPress Enter to exit...")
        sys.exit(1)
    except Exception as e:
        # Handle any other startup errors
        print(f"❌ Application error: {e}")
        input("\nPress Enter to exit...")
        sys.exit(1)

if __name__ == "__main__":
    main()