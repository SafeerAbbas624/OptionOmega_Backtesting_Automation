# OptionOmega Backtesting Automation

## Overview
This project automates the backtesting process on OptionOmega's platform using Selenium WebDriver. It enables bulk testing of multiple trading strategies by reading parameters from an Excel file and automatically executing backtests, collecting results, and generating comprehensive reports.

## Features
- **Automated Login**: Secure login handling to OptionOmega platform
- **Bulk Strategy Testing**: Process multiple trading strategies from a single Excel file
- **Comprehensive Parameter Support**:
  - Basic trade settings (Entry/Exit conditions, Position sizing)
  - Technical indicators (RSI, SMA, EMA)
  - VIX-based conditions
  - Gap trading parameters
  - Multiple leg options strategies
  - Commission and slippage settings
- **Detailed Results Collection**:
  - P/L metrics
  - CAGR
  - Maximum drawdown
  - Win rate statistics
  - Trade duration metrics
  - Shareable strategy links
- **Professional Report Generation**:
  - Formatted Excel output
  - Alternating row colors for readability
  - Optimized column widths
  - Proper date/time formatting
  - Comprehensive strategy performance metrics

## Prerequisites
- Python 3.x
- Chrome Browser
- OptionOmega account

## Required Python Packages
```bash
pip install selenium pandas openpyxl
```

## Installation
1. Clone the repository:
```bash
git clone https://github.com/yourusername/optionomega-backtester.git
cd optionomega-backtester
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage
1. Prepare your input Excel file following the template structure (see Input File Format section)
2. Update the credentials in the script or use environment variables
3. Run the script:
```bash
python trade_testing.py
```

## Input File Format
The input Excel file should contain two header rows:
1. Main category headers
2. Specific parameter headers

Required columns include:
- Start Date
- End Date
- Ticker
- Strategy parameters (Entry/Exit conditions, Position sizing, etc.)
- Technical indicator settings
- Options leg details

See `example_input.xlsx` for a complete template.

## Output
The script generates a timestamped Excel file containing:
- All input parameters
- Comprehensive backtest results
- Shareable strategy links
- Formatted for readability with:
  - Alternating row colors
  - Optimized column widths
  - Proper number/date formatting

## Error Handling
The script includes robust error handling for:
- Network connectivity issues
- Element loading delays
- Stale elements
- Invalid input parameters
- Session timeouts

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

## License
This project is licensed under the Mozilla License - see the [LICENSE](LICENSE) file for details.

## Disclaimer
This tool is for educational and research purposes only. Always verify backtest results and perform due diligence before trading with real money. The authors are not responsible for any financial losses incurred using this tool.

## Acknowledgments
- OptionOmega platform for providing the backtesting infrastructure
- Selenium WebDriver community
- Contributors and testers

## Support
For support, please open an issue in the GitHub repository or contact [safeerabbas.624@hotmail.com]

## Future Enhancements
- [ ] Multi-threading support for parallel backtesting
- [ ] API integration when available
- [ ] Enhanced reporting features
- [ ] Strategy optimization algorithms
- [ ] Custom indicator support
