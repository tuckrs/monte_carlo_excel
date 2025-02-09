import numpy as np
import pandas as pd
import xlwings as xw

def run_simulation(input_range, output_range, num_iterations=1000):
    """Run Monte Carlo simulation on the input data."""
    wb = xw.books.active
    sheet = wb.sheets.active
    
    # Get input data
    input_data = sheet.range(input_range).options(pd.DataFrame, header=True).value
    
    # Initialize results DataFrame
    results = pd.DataFrame()
    
    # Run simulation for each column
    for col in input_data.columns:
        data = input_data[col].dropna()
        mean = data.mean()
        std = data.std()
        
        # Generate random samples
        samples = np.random.normal(mean, std, num_iterations)
        results[f"{col}_sim"] = samples
    
    # Calculate summary statistics
    summary = pd.DataFrame({
        'Mean': results.mean(),
        'Std Dev': results.std(),
        '5th Percentile': results.quantile(0.05),
        '95th Percentile': results.quantile(0.95)
    })
    
    # Write results to Excel
    sheet.range(output_range).value = summary

if __name__ == '__main__':
    xw.serve()
