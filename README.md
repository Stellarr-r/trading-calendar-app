# Strategy Analyzer

## Quick Install

### Option 1: One-Click Setup (Recommended)
1. **Download** the launcher: [`setup_backtestcalendar.bat`](https://raw.githubusercontent.com/Stellarr-r/trading-calendar-app/main/setup_backtestcalendar.bat)
2. **Run** the downloaded file
3. **Done!** Strategy Analyzer will install and launch automatically

### Option 2: Manual Download
1. Right-click [this link](https://raw.githubusercontent.com/Stellarr-r/trading-calendar-app/main/setup_backtestcalendar.bat) → "Save link as..."
2. Save to your Desktop or Downloads folder
3. Double-click the downloaded file to run

## Requirements

- **Windows** 7/8/10/11
- **Python** 3.7+ (auto-installed if missing)
- **Internet connection** (for initial setup only)

## Supported Data Formats

Strategy Analyzer automatically detects and processes:
- TradingView Excel exports with trade data
- Columns: Date/Time, P&L USD, Trade #, Type
- Sheet names: "List of trades" or variations

## Auto-Update System

Strategy Analyzer features a dual auto-update mechanism:
- **Launcher updates** itself from GitHub
- **Application updates** automatically on each run
- **No manual updates** required - always stay current

## Troubleshooting

### Python Installation
If you see "Python not found":
1. Download Python from [python.org](https://www.python.org/downloads/)
2. **Important**: Check "Add Python to PATH" during installation
3. Restart the launcher

### Download Issues
- Check your internet connection
- Try running as administrator

### Support
For technical issues and feature requests:
- **GitHub Issues**: [Report a problem](https://github.com/Stellarr-r/trading-calendar-app/issues)
- **Documentation**: Check this README for solutions

## Privacy & Security

- **Local Processing**: All data analysis happens on your computer
- **No Data Upload**: Your trading data never leaves your machine
- **Secure Storage**: Data is saved locally in your user directory
- **Open Source**: Full code transparency

## Data Storage

Your data is automatically saved to:
```
%APPDATA%\StrategyAnalyzer\data\
```

### Local Setup
```bash
git clone https://github.com/Stellarr-r/trading-calendar-app.git
cd trading-calendar-app
python trading_calendar.py
```

or download and install the .bat

## License

This project is open source and available under the MIT License.
