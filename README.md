# VBA Overview
The VBA Stock Analysis repository is a collection of Visual Basic for Applications (VBA) code designed to analyze stock market data across multiple years. The system processes large datasets of stock information, calculates key performance metrics, identifies trends, and presents the results in a formatted output.

This overview provides a high-level introduction to the repository's purpose, architecture, and functionality. For detailed information about specific components, please refer to the corresponding wiki pages referenced throughout this document.

## Purpose and Scope
The VBA Stock Analysis system was created to automate the analysis of stock market performance across multiple years. It takes raw stock market data organized in Excel worksheets and processes it to calculate:
- Yearly price changes for each stock ticker
- Percentage changes in stock value
- Total stock trading volume
- Stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume
The system is designed to handle large volumes of data efficiently and present the results in an easily interpretable format with appropriate formatting to highlight positive and negative changes.
## System Architecture Overview
The VBA Stock Analysis system follows a modular architecture with distinct components handling specific responsibilities in the analysis process.
<img width="1216" alt="Screenshot 2025-05-31 at 1 55 44 PM" src="https://github.com/user-attachments/assets/78e834e7-a299-4049-a781-bf85fdaacb4d" />
The system is organized into three primary modules:
1. Data Processing Module - Handles reading stock data from worksheets and organizing it for analysis. 
2. Calculation Module - Performs calculations on the processed data to derive performance metrics.
3. Reporting Module - Formats and presents the analysis results in a readable format.

## Data Flow
The following diagram illustrates how data moves through the VBA Stock Analysis system:
<img width="1519" alt="Screenshot 2025-05-31 at 1 58 14 PM" src="https://github.com/user-attachments/assets/aa83d29d-4ab9-4f8d-b092-34f8bddb78a1" />
The data flow follows these steps:
1. Data Input - Raw stock data (ticker symbol, date, open/close prices, volume) is read from worksheets
2. Processing - Data is processed by ticker symbol across the entire year
3. Calculation - Key metrics are calculated for each ticker
4. Analysis - Stocks with extreme values are identified
5. Formatting - Results are formatted with conditional formatting to highlight performance
6. Output - Results are displayed in summary tables on the output worksheets

## Execution Process
The following sequence diagram shows how the VBA Stock Analysis system executes when triggered by the user:
<img width="732" alt="Screenshot 2025-05-31 at 2 00 34 PM" src="https://github.com/user-attachments/assets/c55dae7e-7b98-4647-954d-142807eecc01" />
This process illustrates the typical workflow when a user runs the stock analysis application.

## Core Functionality
The VBA Stock Analysis system provides the following key functionality:

### Stock Data Processing
- Reads stock data from multiple worksheets
- Processes data chronologically by ticker symbol
- Extracts opening and closing prices for the year
### Performance Calculation
- Calculates yearly price changes (closing price - opening price)
- Computes percentage changes ((yearly change / opening price) * 100)
- Totals stock trading volume for each ticker
### Exceptional Performance Identification
- Identifies stocks with the greatest percentage increase
- Identifies stocks with the greatest percentage decrease
- Identifies stocks with the greatest total volume
### Results Presentation
- Creates summary tables with all ticker symbols and their metrics
- Applies conditional formatting (e.g., green for positive change, red for negative)
- Displays extreme value stocks in a separate table

## Repository Structure
The VBA Stock Analysis repository is organized as follows:
<img width="1303" alt="Screenshot 2025-05-31 at 2 04 16 PM" src="https://github.com/user-attachments/assets/a2debd3c-2a9f-41c4-8d9a-f5e5d7c7112a" />

The repository contains:
- Excel workbooks with VBA code implementation
- Stock data worksheets with multiple years of market data
- Result worksheets where analysis output is displayed
- Documentation in the form of wiki pages

## System Integration
The VBA Stock Analysis system integrates with Excel as shown in the following diagram:
<img width="1303" alt="Screenshot 2025-05-31 at 2 04 16 PM" src="https://github.com/user-attachments/assets/4e3259ae-a8b4-457f-b0ab-14fbb9680e0e" />
This integration allows users to interact with the stock analysis system through the familiar Excel interface.

## Summary
The VBA Stock Analysis repository provides a complete solution for analyzing stock market performance data across multiple years. Its modular architecture separates data processing, calculation, and reporting concerns, making the system maintainable and extensible.
