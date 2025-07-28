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

def check_dependencies():
    """Check and install required packages if missing."""
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
                print(f"{package} installed successfully")
            except subprocess.CalledProcessError as e:
                print(f"Failed to install {package}: {e}")
                return False
    
    return True

if not check_dependencies():
    print("Failed to install required dependencies")
    input("Press Enter to exit...")
    sys.exit(1)

import openpyxl
import requests

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
        "list of trades"
    ],
    "ui": {
        "window_width": 1200,
        "window_height": 800,
        "loading_timeout": 2000,
        "progress_update_interval": 10
    },
    "version": "1.0.0"
}

class DataManager:
    """Manages data persistence for trading records."""
    
    def __init__(self, data_dir):
        self.data_dir = data_dir
        self.ensure_data_directory()
    
    def ensure_data_directory(self):
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir, exist_ok=True)
    
    def save_trading_data(self, filename, trades, daily_pnl, stats):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(os.path.basename(filename))[0]
            save_file = os.path.join(self.data_dir, f"{base_name}_{timestamp}.json")
            
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
            
            print(f"Trading data saved: {os.path.basename(save_file)}")
            return save_file
            
        except (IOError, OSError, json.JSONEncodeError) as e:
            print(f"Error saving data: {e}")
            return None
    
    def get_data_history(self):
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
            
            files.sort(key=lambda x: x['modified'], reverse=True)
            return files
            
        except (IOError, OSError) as e:
            print(f"Error getting data history: {e}")
            return []

class TradeProcessor:
    """Processes Excel trading data and calculates statistics."""
    
    def __init__(self, config=None):
        self.config = config or DEFAULT_CONFIG
        self.trades = []
        self.daily_pnl = defaultdict(float)
        self.stats = {}
    
    def process_file(self, file_path, progress_callback=None):
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
            raise RuntimeError(f"Error processing Excel file: {str(e)}") from e
    
    def _find_trades_sheet(self, wb, progress_callback=None):
        sheet_names = self.config.get("sheet_names", ["List of trades"])
        
        for sheet_name in wb.sheetnames:
            for target_name in sheet_names:
                if target_name.lower() in sheet_name.lower():
                    if progress_callback:
                        progress_callback(f"Found sheet: '{sheet_name}'")
                    return wb[sheet_name]
        
        return None
    
    def _analyze_headers(self, sheet, progress_callback=None):
        headers = {}
        column_config = self.config.get("excel_columns", {})
        first_row = list(sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]
        
        for i, header in enumerate(first_row):
            if header:
                header_str = str(header).strip()
                
                for key, expected_name in column_config.items():
                    if header_str == expected_name:
                        headers[key] = i
                        if progress_callback:
                            progress_callback(f"Found {expected_name} column")
                        break
        
        required_columns = ['datetime', 'pnl']
        missing_columns = [col for col in required_columns if col not in headers]
        
        if missing_columns:
            column_names = [column_config.get(col, col) for col in missing_columns]
            raise ValueError(f"Missing required columns: {', '.join(column_names)}")
        
        return headers
    
    def _process_trades(self, sheet, headers, progress_callback=None):
        self.trades = []
        self.daily_pnl = defaultdict(float)
        processed_trade_numbers = set()
        
        total_rows = sheet.max_row - 1
        processed_rows = 0
        progress_interval = self.config.get("ui", {}).get("progress_update_interval", 10)
        
        for row in sheet.iter_rows(min_row=2, values_only=True):
            processed_rows += 1
            
            if progress_callback and processed_rows % progress_interval == 0:
                progress = int((processed_rows / total_rows) * 100)
                progress_callback(f"Processing trades... {progress}% ({processed_rows}/{total_rows})")
            
            try:
                trade = self._process_single_trade(row, headers, processed_trade_numbers)
                if trade:
                    self.trades.append(trade)
                    
                    date_key = trade['date'].strftime('%Y-%m-%d')
                    self.daily_pnl[date_key] += trade['pnl']
                    
            except (ValueError, TypeError, AttributeError):
                continue
        
        if progress_callback:
            progress_callback(f"Successfully processed {len(self.trades)} unique trades")
    
    def _process_single_trade(self, row, headers, processed_trade_numbers):
        trade_num = row[headers['trade_num']] if 'trade_num' in headers else None
        datetime_val = row[headers['datetime']]
        pnl_val = row[headers['pnl']]
        
        if not datetime_val or pnl_val is None:
            return None
        
        if trade_num and trade_num in processed_trade_numbers:
            return None
        
        trade_date = self._parse_date(datetime_val)
        if not trade_date:
            return None
        
        pnl = self._parse_pnl(pnl_val)
        if pnl is None:
            return None
        
        if trade_num:
            processed_trade_numbers.add(trade_num)
        
        return {
            'date': trade_date,
            'pnl': pnl,
            'trade_num': trade_num
        }
    
    def _parse_date(self, datetime_val):
        """Parse date from various Excel date formats."""
        try:
            if isinstance(datetime_val, datetime):
                parsed_date = datetime_val.date()
            else:
                date_str = str(datetime_val).split()[0]
                parsed_date = datetime.strptime(date_str, '%Y-%m-%d').date()
            
            if parsed_date > datetime.now().date():
                return None
            if parsed_date.year < 2000:
                return None
                
            return parsed_date
        except (ValueError, TypeError, AttributeError):
            return None
    
    def _parse_pnl(self, pnl_val):
        """Parse P&L value from various Excel formats."""
        try:
            if isinstance(pnl_val, (int, float)):
                return float(pnl_val)
            elif isinstance(pnl_val, str):
                cleaned = pnl_val.replace(',', '').replace('$', '').strip()
                if not cleaned or cleaned == '-':
                    return None
                return float(cleaned)
            else:
                return None
        except (ValueError, TypeError):
            return None
    
    def calculate_stats(self):
        """Calculate trading statistics from processed trades."""
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

class StrategyAnalyzer:
    """Main application class for the Strategy Analyzer GUI."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Strategy Analyzer")
        
        self.config = DEFAULT_CONFIG
        ui_config = self.config.get("ui", {})
        
        self.root.geometry(f"{ui_config.get('window_width', 1200)}x{ui_config.get('window_height', 800)}")
        self.root.configure(bg=self.config['theme']['bg_dark'])
        self.root.resizable(True, True)
        
        self.CURRENT_VERSION = os.environ.get('STRATEGY_ANALYZER_VERSION', self.config.get('version', '1.0.0'))
        
        self.validate_version()
        
        data_dir = os.environ.get('STRATEGY_ANALYZER_DATA_DIR', os.path.join(os.getcwd(), 'data'))
        self.data_manager = DataManager(data_dir)
        
        self.trade_processor = TradeProcessor(self.config)
        
        self.theme = self.config['theme']
        
        self.trades = []
        self.daily_pnl = defaultdict(float)
        self.stats = {}
        self.current_month = datetime.now().month
        self.current_year = datetime.now().year
        self.current_file = None
        
        self.setup_ui()
    
    def validate_version(self):
        import re
        
        version_pattern = r'^(\d+\.\d+\.\d+|DEV)$'
        if not re.match(version_pattern, self.CURRENT_VERSION):
            print(f"Warning: Invalid version format '{self.CURRENT_VERSION}'. Using fallback.")
            self.CURRENT_VERSION = "1.0.0"
        
        print(f"Strategy Analyzer v{self.CURRENT_VERSION} starting...")
        
        version_source = os.environ.get('STRATEGY_ANALYZER_VERSION')
        if version_source:
            print(f"Version loaded from Strategy Analyzer launcher: {version_source}")
        else:
            print(f"Using default version (run via Strategy Analyzer launcher for version management)")
        
    def setup_ui(self):
        main_container = tk.Frame(self.root, bg=self.theme['bg_dark'])
        main_container.pack(fill='both', expand=True, padx=0, pady=0)
        
        self.create_top_bar(main_container)
        
        content_frame = tk.Frame(main_container, bg=self.theme['bg_dark'])
        content_frame.pack(fill='both', expand=True, padx=20, pady=0)
        
        loading_label = tk.Label(content_frame, 
                               text="Loading...", 
                               font=('Inter', 12),
                               fg=self.theme['text_muted'],
                               bg=self.theme['bg_dark'])
        loading_label.pack(pady=50)
        
        self.root.after(100, lambda: self.create_heavy_ui_elements(content_frame, loading_label))
        
    def create_heavy_ui_elements(self, content_frame, loading_label):
        loading_label.destroy()
        
        self.create_stats_section(content_frame)
        
        self.create_calendar_section(content_frame)
        
        self.create_bottom_info(content_frame)
        
    def create_top_bar(self, parent):
        top_bar = tk.Frame(parent, bg=self.theme['bg_card'], height=70)
        top_bar.pack(fill='x', padx=0, pady=0)
        top_bar.pack_propagate(False)
        
        left_frame = tk.Frame(top_bar, bg=self.theme['bg_card'])
        left_frame.pack(side='left', fill='y', padx=20)
        
        title = tk.Label(left_frame,
                        text="Strategy Analyzer",
                        font=('Inter', 20, 'bold'),
                        fg=self.theme['text_primary'],
                        bg=self.theme['bg_card'])
        title.pack(side='left', pady=20)
        
        version_label = tk.Label(left_frame,
                               text=f"v{self.CURRENT_VERSION}",
                               font=('Inter', 10),
                               fg=self.theme['text_muted'],
                               bg=self.theme['bg_card'])
        version_label.pack(side='left', padx=(10, 0), pady=20)
        
        right_frame = tk.Frame(top_bar, bg=self.theme['bg_card'])
        right_frame.pack(side='right', fill='y', padx=20)
        
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
        
        self.file_status = tk.Label(left_frame,
                                   text="No file loaded",
                                   font=('Inter', 9),
                                   fg=self.theme['text_muted'],
                                   bg=self.theme['bg_card'])
        self.file_status.pack(side='left', padx=(20, 0), pady=20)
        
    def create_stats_section(self, parent):
        stats_container = tk.Frame(parent, bg=self.theme['bg_dark'])
        stats_container.pack(fill='x', pady=20)
        
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
        card = tk.Frame(parent, 
                       bg=self.theme['bg_card'],
                       relief='flat',
                       bd=0)
        
        content = tk.Frame(card, bg=self.theme['bg_card'])
        content.pack(fill='both', expand=True, padx=20, pady=20)
        
        value_label = tk.Label(content,
                              text=value,
                              font=('Inter', 24, 'bold'),
                              fg=self.theme['text_primary'],
                              bg=self.theme['bg_card'])
        value_label.pack(anchor='w')
        
        label_widget = tk.Label(content,
                               text=label,
                               font=('Inter', 11),
                               fg=self.theme['text_secondary'],
                               bg=self.theme['bg_card'])
        label_widget.pack(anchor='w', pady=(5, 0))
        
        card.value_label = value_label
        card.label_widget = label_widget
        
        return card
    
    def create_calendar_section(self, parent):
        calendar_container = tk.Frame(parent, bg=self.theme['bg_dark'])
        calendar_container.pack(fill='both', expand=True, pady=(20, 0))
        
        cal_header = tk.Frame(calendar_container, bg=self.theme['bg_dark'])
        cal_header.pack(fill='x', pady=(0, 20))
        
        nav_frame = tk.Frame(cal_header, bg=self.theme['bg_dark'])
        nav_frame.pack()
        
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
        
        month_name = calendar.month_name[self.current_month]
        self.month_display = tk.Label(nav_frame,
                                     text=f"{month_name} {self.current_year}",
                                     font=('Inter', 16, 'bold'),
                                     fg=self.theme['text_primary'],
                                     bg=self.theme['bg_dark'],
                                     padx=30)
        self.month_display.pack(side='left')
        
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
        
        self.calendar_frame = tk.Frame(calendar_container, bg=self.theme['bg_dark'])
        self.calendar_frame.pack(fill='both', expand=True)
        
        self.setup_calendar_grid()
        
    def setup_calendar_grid(self):
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        for i, day in enumerate(days):
            header = tk.Label(self.calendar_frame,
                             text=day,
                             font=('Inter', 11, 'bold'),
                             fg=self.theme['text_secondary'],
                             bg=self.theme['bg_accent'],
                             pady=12)
            header.grid(row=0, column=i, sticky='ew', padx=1, pady=1)
        
        for i in range(7):
            self.calendar_frame.grid_columnconfigure(i, weight=1)
        for i in range(7):
            self.calendar_frame.grid_rowconfigure(i, weight=1)
            
        self.update_calendar()
    
    def create_day_cell(self, day, pnl=0, row=0, col=0):
        if pnl > 0:
            bg_color = self.theme['accent_green']
            text_color = 'white'
        elif pnl < 0:
            bg_color = self.theme['accent_red']
            text_color = 'white'
        else:
            bg_color = self.theme['bg_card']
            text_color = self.theme['text_secondary']
        
        cell = tk.Frame(self.calendar_frame,
                       bg=bg_color,
                       relief='flat',
                       bd=0)
        cell.grid(row=row, column=col, sticky='nsew', padx=1, pady=1)
        
        content = tk.Frame(cell, bg=bg_color)
        content.pack(fill='both', expand=True, padx=12, pady=12)
        
        day_label = tk.Label(content,
                            text=str(day),
                            font=('Inter', 14, 'bold'),
                            fg=text_color,
                            bg=bg_color)
        day_label.pack(anchor='nw')
        
        if pnl != 0:
            pnl_text = f"${pnl:,.0f}" if abs(pnl) >= 1 else f"${pnl:.2f}"
            pnl_label = tk.Label(content,
                                text=pnl_text,
                                font=('Inter', 10),
                                fg=text_color,
                                bg=bg_color)
            pnl_label.pack(anchor='sw')
    
    def create_bottom_info(self, parent):
        bottom_frame = tk.Frame(parent, bg=self.theme['bg_dark'])
        bottom_frame.pack(fill='x', pady=(20, 10))
        
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
        file_path = filedialog.askopenfilename(
            title="Select TradingView Export",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        self.current_file = file_path
            
        self.show_loading_dialog(file_path)
    
    def show_loading_dialog(self, file_path):
        self.loading_dialog = tk.Toplevel(self.root)
        self.loading_dialog.title("Processing Excel File")
        self.loading_dialog.geometry("500x350")
        self.loading_dialog.configure(bg=self.theme['bg_dark'])
        self.loading_dialog.transient(self.root)
        self.loading_dialog.grab_set()
        
        self.loading_dialog.geometry("+{}+{}".format(
            int(self.root.winfo_screenwidth()/2 - 250),
            int(self.root.winfo_screenheight()/2 - 175)
        ))
        
        self.loading_dialog.protocol("WM_DELETE_WINDOW", self.close_loading_dialog)
        self.loading_dialog.bind('<Escape>', lambda e: self.close_loading_dialog())
        self.loading_dialog.bind('<Return>', lambda e: self.close_loading_dialog())
        self.loading_dialog.attributes('-topmost', True)
        
        main_frame = tk.Frame(self.loading_dialog, bg=self.theme['bg_dark'])
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        title_label = tk.Label(main_frame,
                              text="Processing Excel File",
                              font=('Inter', 18, 'bold'),
                              fg=self.theme['accent_blue'],
                              bg=self.theme['bg_dark'])
        title_label.pack(pady=(0, 20))
        
        filename = os.path.basename(file_path)
        file_label = tk.Label(main_frame,
                             text=f"File: {filename}",
                             font=('Inter', 12),
                             fg=self.theme['text_secondary'],
                             bg=self.theme['bg_dark'])
        file_label.pack(pady=(0, 20))
        
        log_frame = tk.Frame(main_frame, bg=self.theme['bg_card'])
        log_frame.pack(fill='both', expand=True, pady=(0, 20))
        
        log_header = tk.Label(log_frame,
                             text="Processing Log:",
                             font=('Inter', 12, 'bold'),
                             fg=self.theme['text_primary'],
                             bg=self.theme['bg_card'])
        log_header.pack(anchor='w', padx=15, pady=(15, 5))
        
        self.loading_log = tk.Text(log_frame,
                                  font=('Consolas', 10),
                                  bg=self.theme['bg_dark'],
                                  fg=self.theme['text_primary'],
                                  border=0,
                                  height=10,
                                  wrap='word')
        self.loading_log.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        
        self.process_file_background(file_path)
        
    def process_file_background(self, file_path):
        processing_thread = threading.Thread(
            target=self.process_file_with_logging,
            args=(file_path,),
            daemon=True
        )
        processing_thread.start()
        
    def process_file_with_logging(self, file_path):
        import time
        start_time = time.time()
        
        try:
            result = self.trade_processor.process_file(file_path, self.log_loading)
            
            end_time = time.time()
            processing_time = end_time - start_time
            
            self.trades = result['trades']
            self.daily_pnl = defaultdict(float, result['daily_pnl'])
            self.stats = result['stats']
            
            total_pnl = self.stats.get('total_pnl', 0)
            total_trades = self.stats.get('total_trades', 0)
            win_rate = self.stats.get('win_rate', 0)
            
            self.log_loading(f"Total P&L: ${total_pnl:,.2f}")
            self.log_loading(f"Total Trades: {total_trades}")
            self.log_loading(f"Win Rate: {win_rate:.1f}%")
            self.log_loading(f"Processing completed in {processing_time:.2f} seconds")
            
            self.log_loading("Preparing calendar display...")
            
            self.loading_dialog.after(0, self.finish_processing)
            
        except Exception as e:
            error_msg = str(e)
            self.log_loading(f"ERROR: {error_msg}")
            self.loading_dialog.after(0, lambda: self.show_processing_error(error_msg))
    
    def log_loading(self, message):
        def update_log():
            self.loading_log.insert('end', f"{message}\n")
            self.loading_log.see('end')
            self.loading_dialog.update()
        
        self.loading_dialog.after(0, update_log)
        time.sleep(0.1)
    
    def finish_processing(self):
        self.log_loading("Processing complete!")
        
        self.update_displays()
        
        if self.current_file and self.trades and self.stats:
            self.data_manager.save_trading_data(self.current_file, self.trades, self.daily_pnl, self.stats)
        
        filename = os.path.basename(self.current_file) if self.current_file else "Unknown"
        self.file_status.config(text=f"Loaded: {filename}")
    
    def show_processing_error(self, error_msg):
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
        
        self.loading_dialog.after(100, lambda: messagebox.showerror("Processing Error", f"Failed to process file: {error_msg}"))
    
    def continue_to_calendar(self):
        self.close_loading_dialog()
    
    def close_loading_dialog(self):
        if hasattr(self, 'loading_dialog') and self.loading_dialog:
            self.loading_dialog.destroy()
    
    def update_displays(self):
        self.update_stats_display()
        self.update_calendar()
    
    def update_stats_display(self):
        if not self.stats:
            return
        
        pnl_color = self.theme['accent_green'] if self.stats['total_pnl'] >= 0 else self.theme['accent_red']
        self.stat_cards['total_pnl'].value_label.config(
            text=f"${self.stats['total_pnl']:,.2f}",
            fg=pnl_color
        )
        
        self.stat_cards['win_rate'].value_label.config(
            text=f"{self.stats['win_rate']:.1f}%"
        )
        
        avg_color = self.theme['accent_green'] if self.stats['avg_daily'] >= 0 else self.theme['accent_red']
        self.stat_cards['avg_daily'].value_label.config(
            text=f"${self.stats['avg_daily']:.2f}",
            fg=avg_color
        )
        
        self.stat_cards['total_trades'].value_label.config(
            text=str(self.stats['total_trades'])
        )
    
    def update_calendar(self):
        for widget in self.calendar_frame.winfo_children():
            if int(widget.grid_info()['row']) > 0:
                widget.destroy()
        
        cal = calendar.monthcalendar(self.current_year, self.current_month)
        
        for week_num, week in enumerate(cal):
            for day_num, day in enumerate(week):
                if day == 0:
                    continue
                    
                row = week_num + 1
                col = day_num
                
                date_key = f"{self.current_year}-{self.current_month:02d}-{day:02d}"
                pnl = self.daily_pnl.get(date_key, 0)
                
                self.create_day_cell(day, pnl, row, col)
        
        month_name = calendar.month_name[self.current_month]
        self.month_display.config(text=f"{month_name} {self.current_year}")
    
    def prev_month(self):
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.update_calendar()
    
    def next_month(self):
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.update_calendar()

def main():
    """Entry point for the Strategy Analyzer application."""
    try:
        print("Strategy Analyzer")
        print("=" * 60)
        
        root = tk.Tk()
        app = StrategyAnalyzer(root)
        
        print("Application initialized successfully")
        print("Ready to process TradingView Excel exports")
        print("=" * 60)
        
        root.mainloop()
        
    except ImportError as e:
        print(f"Missing dependency: {e}")
        print("\nThis shouldn't happen as dependencies are auto-installed.")
        print("Please ensure you have internet access and pip is available.")
        input("\nPress Enter to exit...")
        sys.exit(1)
    except Exception as e:
        print(f"Application error: {e}")
        input("\nPress any key to exit...")
        sys.exit(1)

if __name__ == "__main__":
    main()
