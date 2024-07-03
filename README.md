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
- numpy

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
    pip install matplotlib pandas numpy
    ```

4. **Run the Python Script**:
    ```bash
    python /path/to/your/script.py
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

## Contributing

1. Fork the repository
2. Create a new branch (`git checkout -b feature-branch`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature-branch`)
5. Create a new Pull Request

## License

This project is licensed under the MIT License.
