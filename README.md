# Portfolio Risk Analysis and Optimization using Monte Carlo Simulations

## Project Overview

This project is a Python application designed to analyze financial risk and optimize portfolio allocation using Monte Carlo simulations. The application calculates key financial metrics such as average close price, average daily return, and volatility for various stocks. It also estimates Value at Risk (VaR) and Historical VaR (HVaR) to assist in making informed investment decisions.

## Features

- **Financial Metrics Calculation**: Compute average close price, average daily return, and volatility.
- **Historical Value at Risk (HVaR)**: Calculate HVaR based on historical data to assess potential losses.
- **Monte Carlo Value at Risk (VaR)**: Perform Monte Carlo simulations to estimate potential future losses.
- **Portfolio Allocation**: Allocate portfolio based on calculated risk metrics and optimize investment distribution.

## Technologies Used

- **Python**: The core programming language used for the application.
- **pandas**: For data manipulation and analysis.
- **numpy**: For numerical operations and Monte Carlo simulations.
- **openpyxl**: For reading and writing Excel files.

## Setup Instructions

To set up and run the application locally, follow these steps:

1. **Clone the repository**:
   git clone https://github.com/GarrettPerschino/portfolio-risk-analysis.git
   cd portfolio-risk-analysis

2. **Create and activate a virtual environment**:
python3 -m venv env
source env/bin/activate

3. **Install the required packages**:

pip install pandas numpy openpyxl

4. **Run the application**:

python3 portfolio_analysis.py

**USAGE**

1. **Prepare Your Data**:

Ensure you have an Excel file containing stock data with at least a 'Close' column. Save this file in your project directory or specify the path when prompted.

2. **Run the Application**:

Execute the script and provide the path to the Excel file and the total portfolio worth when prompted:

python portfolio_analysis.py

3. **View Results**:

The application will output the calculated financial metrics, Historical VaR, and Monte Carlo VaR. It will also save the portfolio allocation results to a new Excel file named 'portfolio_allocation.xlsx'.

**Example Output**

_Upon running the application, you will see output similar to this:_

Stock Average Close Average Daily Return Volatility Historical VaR Monte Carlo VaR Allocation
AAPL (1) 46. 968888 0.001247 0.023390 -0. 030465 348. 369679 97.863418
NVDA (1) 7.075146 0.001798 0. 030996 -0.044560 315.649827 81.910102
MSFT 97. 247579 0.000702 0.017089 -0.025056 365. 637389 115. 996389
105. 508884 0.000233 0. 020586 -0. 029962 292.206518 108. 688831
TTWO 60. 249104 0.000774 0. 026761 -0. 036981 275.758311 95. 541260

**Project Structure**
The project repository is structured as follows:

**portfolio_analysis.py: The main script that performs data loading, metric calculation, Monte Carlo simulation, and portfolio allocation.
README.md: Project overview and setup instructions.**
