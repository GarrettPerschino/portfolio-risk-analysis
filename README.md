# Portfolio Allocation Tool

A Python-based GUI application that helps users allocate their investment portfolio based on financial metrics calculated from historical stock data.

## Features

- Load stock data from an Excel file.
- Calculate financial metrics such as Average Close, Average Daily Return, Volatility, Historical VaR, and Monte Carlo VaR.
- Allocate portfolio based on calculated metrics.
- Display allocation results in a tabular format.
- Visualize portfolio allocation with a pie chart.

## Requirements

- Python 3.x
- matplotlib
- pandas
- openpyxl
- numpy
- tkinter

## Setup and Running the Project Locally

### Prerequisites

- Python 3.x installed on your machine. Download from [here](https://www.python.org/downloads/).

### Universal Steps

1. **Clone the Project Repository**:
    ```bash
    git clone https://github.com/GarrettPerschino/portfolio-risk-analysis.git
    cd portfolio-risk-analysis
    ```

2. **Create and Activate a Virtual Environment**:
    ```bash
    python3 -m venv myenv
    ```
    - **Windows**:
      ```bash
      myenv\Scripts\activate
      ```
    - **Mac/Linux**:
      ```bash
      source myenv/bin/activate
      ```

3. **Install Required Packages**:
    ```bash
    pip install matplotlib pandas openpyxl numpy
    ```

4. **Run the Python Script**:
    ```bash
    python portfolio_analysis.py
    ```

### Mac-Specific Instructions

1. **Install Homebrew**:
    ```bash
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    ```

2. **Install pipx and Ensure Path**:
    ```bash
    brew install pipx
    pipx ensurepath
    ```

3. **Follow Universal Steps**:
    - After installing Homebrew and `pipx`, follow the universal steps to create a virtual environment, install required packages, and run your script.

## Using the Project

1. **Load Excel File**:
    - Click the "Browse" button to select an Excel file containing stock data. The Excel file should have a 'Close' column for stock prices.

2. **Enter Portfolio Worth**:
    - Enter the total worth of your portfolio in the provided entry field.

3. **Calculate Allocation**:
    - Click the "Calculate Allocation" button to calculate the portfolio allocation based on the financial metrics.

4. **View Results**:
    - The results will be displayed in the table and a pie chart will visualize the allocation.

## License

This project is licensed under the MIT License.
