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
- Pandas
- NumPy
- Matplotlib
- Tkinter (comes pre-installed with Python)

## Installation

1. **Clone the repository:**

    ```bash
    git clone https://github.com/yourusername/portfolio-allocation-tool.git
    cd portfolio-allocation-tool
    ```

2. **Create a virtual environment:**

    ```bash
    python3 -m venv venv
    ```

3. **Activate the virtual environment:**

    - On macOS/Linux:
      ```bash
      source venv/bin/activate
      ```
    - On Windows:
      ```bash
      venv\Scripts\activate
      ```

4. **Install the required Python packages:**

    ```bash
    pip install pandas numpy matplotlib
    ```

## Usage

1. **Ensure the virtual environment is activated:**

    - On macOS/Linux:
      ```bash
      source venv/bin/activate
      ```
    - On Windows:
      ```bash
      venv\Scripts\activate
      ```

2. **Run the program:**

    ```bash
    python portfolio_allocation.py
    ```

3. **Using the GUI:**

    - **Load an Excel File:** Click on the `Browse` button to select an Excel file containing historical stock data. The file should have one or more sheets, each with a `Close` column representing the closing prices of the stock.
    - **Enter Portfolio Worth:** Input the total worth of your portfolio in the `Portfolio Worth` entry.
    - **Calculate Allocation:** Click the `Calculate Allocation` button to compute the financial metrics and allocate the portfolio.
    - **View Results:** The results will be displayed in a table and a pie chart showing the allocation percentages.

## Example

Here is an example of how the Excel file should be structured:

**Sheet1:**

| Date       | Close |
|------------|-------|
| 2023-01-01 | 100.5 |
| 2023-01-02 | 101.0 |
| ...        | ...   |

**Sheet2:**

| Date       | Close |
|------------|-------|
| 2023-01-01 | 200.5 |
| 2023-01-02 | 201.0 |
| ...        | ...   |

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## Contact

For any questions or suggestions, please open an issue or contact me at gaperschino@gmail.com
