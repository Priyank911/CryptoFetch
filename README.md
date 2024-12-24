### README.md  

# Live Cryptocurrency Data Fetching and Analysis

This project fetches live cryptocurrency data for the top 50 cryptocurrencies, analyzes the data, and updates an Excel sheet in real time. The program is designed to provide continuous updates every 5 minutes, making it suitable for market monitoring and analysis.

---

## Features

1. **Fetch Live Data**  
   - Utilizes the [CoinGecko API](https://www.coingecko.com/en/api) to retrieve live data for the top 50 cryptocurrencies by market capitalization.  
   - Includes key metrics:
     - Cryptocurrency Name  
     - Symbol  
     - Current Price (USD)  
     - Market Capitalization  
     - 24-hour Trading Volume  
     - 24-hour Price Change (percentage)  

2. **Data Analysis**  
   - Extracts and analyzes key insights from the fetched data:  
     - **Top 5 Cryptocurrencies** by market capitalization.  
     - **Average Price** of the top 50 cryptocurrencies.  
     - **Highest and Lowest 24-hour Price Change** percentages.  

3. **Live Excel Sheet Update**  
   - Automatically updates an Excel sheet (`crypto_live_data.xlsx`) with the latest data.  
   - Overwrites previous data to ensure real-time accuracy.  
   - Compatible with Microsoft Excel.  

4. **Console Reporting**  
   - Displays analysis results in a user-friendly format:  
     - Top 5 cryptocurrencies by market cap.  
     - Average price of the top 50 cryptocurrencies.  
     - Cryptocurrencies with the highest and lowest 24-hour price changes.  

---

## Installation and Usage

1. **Clone the Repository**  
   ```bash
   git clone https://github.com/Priyank911/CryptoFetch.git
   cd crypto-data-fetcher

2. **Install Dependencies**  
   Ensure you have Python 3.x installed and run:  
   ```bash
   pip install requests pandas openpyxl
   ```

3. **Run the Script**  
   Execute the script to start fetching and analyzing live cryptocurrency data:  
   ```bash
   python main.py
   ```

4. **Access Excel Output**  
   The data will be saved in a file named `crypto_live_data.xlsx` in the project directory.

---

## Project Structure

- `main.py` : The main script that fetches data, analyzes it, and updates the Excel sheet.  

---

## Requirements

- Python 3.x  
- Libraries:  
  - `requests`  
  - `pandas`  
  - `openpyxl`  

---

## API Used

- **CoinGecko API**: A free, public API for fetching cryptocurrency data. [Documentation](https://www.coingecko.com/en/api)  

---

## Future Improvements

- **Visualization**: Add graphical dashboards for better insights.  
- **Database Integration**: Store historical data for trend analysis.  
- **Customizable Features**: Allow dynamic interval settings and additional metrics.  

---

## License

This project is licensed under the MIT License. See `LICENSE` for more details.

---

## Contributing

Feel free to fork this repository, submit issues, or create pull requests for improvements.  
```

Let me know if you need further customization!
