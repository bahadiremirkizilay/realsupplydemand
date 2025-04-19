# Supply Demand Analysis Tool

This project is a tool that helps determine supply and demand levels by analyzing stock market data. The project consists of two main components:

1. Stock data retrieval and analysis tool
2. Duplicate value analysis tool

## Features

- Fetch stock data through Yahoo Finance API
- Analyze high and low price levels
- Determine potential supply-demand levels using quadrant analysis
- Detect and count duplicate values
- Save results to Excel files

## Requirements

- Python 3.x
- openpyxl
- yfinance
- pandas
- datetime

To install the required libraries:

```bash
pip install openpyxl yfinance pandas
```

## Usage

### 1. Stock Data Analysis (stock_analysis.py)

```bash
python stock_analysis.py
```

The program will prompt you for the following information:
- Stock symbol (e.g., AAPL, TSLA, ^GSPC)
- Start date (in YYYY-MM-DD format)
- End date (in YYYY-MM-DD format)
- Output file name

### 2. Duplicate Value Analysis (duplicate_analyzer.ipynb)

Open the Jupyter Notebook and run the cells in sequence. Analysis results will be automatically saved to an Excel file.

## Outputs

1. When running `stock_analysis.py`:
   - High and low price levels for the specified stock
   - Quadrant analysis results
   - All data saved to an Excel file

2. When running `duplicate_analyzer.ipynb`:
   - List of duplicate values
   - Count of each duplicate value
   - Results saved to an Excel file

## Contributing

1. Fork this repository
2. Create a new branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Create a Pull Request

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact

Project Owner - [@yourusername](https://github.com/yourusername)

Project Link: [https://github.com/yourusername/realsupplydemand](https://github.com/yourusername/realsupplydemand)

