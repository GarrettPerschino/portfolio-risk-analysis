import pandas as pd
import numpy as np
import openpyxl

def load_data(file_path):
    """
    Load data from an Excel file.

    Parameters:
    file_path (str): The path to the Excel file.

    Returns:
    pd.ExcelFile: An ExcelFile object containing the data.
    """
    try:
        print(f"Loading data from: {file_path}")
        return pd.ExcelFile(file_path)
    except Exception as e:
        raise ValueError(f"Error loading Excel file: {str(e)}")

def check_required_columns(df):
    """
    Check if the required columns are present in the DataFrame.

    Parameters:
    df (pd.DataFrame): The DataFrame to check.

    Returns:
    bool: True if required columns are present, False otherwise.
    """
    required_columns = {'Close'}
    return required_columns.issubset(df.columns)

def calculate_metrics(df):
    """
    Calculate financial metrics: average close price, average daily return, and volatility.

    Parameters:
    df (pd.DataFrame): The DataFrame containing stock data.

    Returns:
    dict: A dictionary containing the calculated metrics.
    """
    metrics = {}
    try:
        df['Close'] = pd.to_numeric(df['Close'], errors='coerce')
        df.dropna(subset=['Close'], inplace=True)
        df['Return'] = df['Close'].pct_change().dropna()
        metrics['average_close'] = df['Close'].mean()
        metrics['average_daily_return'] = df['Return'].mean()
        metrics['volatility'] = df['Return'].std(ddof=0)
    except Exception as e:
        raise ValueError(f"Error calculating metrics: {str(e)}")
    return metrics

def historical_var(df, confidence_level=0.95):
    """
    Calculate Historical Value at Risk (HVaR).

    Parameters:
    df (pd.DataFrame): The DataFrame containing stock data.
    confidence_level (float): The confidence level for VaR calculation.

    Returns:
    float: The calculated HVaR.
    """
    try:
        sorted_returns = df['Return'].sort_values()
        index = int((1 - confidence_level) * len(sorted_returns))
        return sorted_returns.iloc[index]
    except Exception as e:
        raise ValueError(f"Error calculating historical VaR: {str(e)}")

def monte_carlo_var(avg_return, volatility, portfolio_worth, num_simulations=10000, confidence_level=0.95):
    """
    Perform Monte Carlo simulation for VaR.

    Parameters:
    avg_return (float): The average daily return.
    volatility (float): The volatility of returns.
    portfolio_worth (float): The total worth of the portfolio.
    num_simulations (int): The number of simulations to run.
    confidence_level (float): The confidence level for VaR calculation.

    Returns:
    float: The calculated Monte Carlo VaR.
    """
    try:
        simulated_end_values = np.zeros(num_simulations)
        for i in range(num_simulations):
            simulated_returns = np.random.normal(avg_return, volatility, 252)
            simulated_end_values[i] = portfolio_worth * np.prod(1 + simulated_returns)
        var = np.percentile(simulated_end_values, (1 - confidence_level) * 100)
        return var
    except Exception as e:
        raise ValueError(f"Error performing Monte Carlo simulation: {str(e)}")

def allocate_portfolio(metrics_list, portfolio_worth):
    """
    Allocate portfolio based on VaR and Volatility.

    Parameters:
    metrics_list (list): A list of tuples containing sheet names and their corresponding metrics.
    portfolio_worth (float): The total worth of the portfolio.

    Returns:
    list: A list of tuples containing sheet names, metrics, and allocation values.
    """
    try:
        total_inverse_volatility = sum(1 / metrics['volatility'] for _, metrics in metrics_list)
        total_inverse_hvar = sum(1 / metrics['historical_var'] for _, metrics in metrics_list)
        total_inverse_mcvar = sum(1 / metrics['monte_carlo_var'] for _, metrics in metrics_list)

        allocations = []
        for sheet_name, metrics in metrics_list:
            inverse_volatility = 1 / metrics['volatility']
            inverse_hvar = 1 / metrics['historical_var']
            inverse_mcvar = 1 / metrics['monte_carlo_var']
            normalized_metric = (
                inverse_volatility / total_inverse_volatility + 
                inverse_hvar / total_inverse_hvar + 
                inverse_mcvar / total_inverse_mcvar
            ) / 3
            allocation = normalized_metric * portfolio_worth
            allocations.append((sheet_name, metrics, allocation))
        
        return allocations
    except Exception as e:
        raise ValueError(f"Error allocating portfolio: {str(e)}")

def run_allocation(file_path, portfolio_worth):
    """
    Run the portfolio allocation and VaR calculations.

    Parameters:
    file_path (str): The path to the Excel file.
    portfolio_worth (float): The total worth of the portfolio.

    Returns:
    None
    """
    try:
        excel_file = load_data(file_path)
        results = pd.DataFrame(columns=["Stock", "Average Close", "Average Daily Return", "Volatility", "Historical VaR", "Monte Carlo VaR", "Allocation"])
        metrics_list = []

        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            if not df.empty and check_required_columns(df):
                metrics = calculate_metrics(df)
                metrics['historical_var'] = historical_var(df)
                metrics['monte_carlo_var'] = monte_carlo_var(metrics['average_daily_return'], metrics['volatility'], portfolio_worth)
                metrics_list.append((sheet_name, metrics))

        if not metrics_list:
            raise ValueError("The Excel file contains no non-empty sheets with the required columns.")

        allocations = allocate_portfolio(metrics_list, portfolio_worth)

        for sheet_name, metrics, allocation in allocations:
            stock_allocation_df = pd.DataFrame({
                "Stock": [sheet_name], 
                "Average Close": [metrics['average_close']], 
                "Average Daily Return": [metrics['average_daily_return']], 
                "Volatility": [metrics['volatility']], 
                "Historical VaR": [metrics['historical_var']],
                "Monte Carlo VaR": [metrics['monte_carlo_var']],
                "Allocation": [allocation]
            })
            results = pd.concat([results, stock_allocation_df], ignore_index=True)

        results['Monte Carlo VaR'] = results['Monte Carlo VaR'].apply(lambda x: f"${x:,.2f}")
        results['Allocation'] = results['Allocation'].apply(lambda x: f"${x:,.2f}")

        print("Portfolio Allocation:")
        print(results.to_string(index=False))

        output_file = 'portfolio_allocation.xlsx'
        results.to_excel(output_file, index=False)
        print(f"Allocation saved to {output_file}")
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    file_path = input("Enter the path to the Excel file: ")
    portfolio_worth = float(input("Enter the total portfolio worth: "))
    run_allocation(file_path, portfolio_worth)
