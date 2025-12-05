#!/usr/bin/env python3
"""
Generate a sample Excel file with many columns for testing.

This script creates a mock XLSX file with a large number of columns,
multi-row headers, and various data types to test the conversion pipeline.
"""

import pandas as pd
import numpy as np
from pathlib import Path
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def generate_sample_excel(
    filename: str = "sample.xlsx",
    num_rows: int = 50,
    num_columns: int = 100
):
    """
    Generate a sample Excel file with many columns.
    
    Args:
        filename: Output filename
        num_rows: Number of data rows
        num_columns: Number of columns
    """
    try:
        logger.info(f"Generating sample Excel file: {filename}")
        logger.info(f"  Rows: {num_rows}, Columns: {num_columns}")
        
        # Generate column names
        column_names = [f"Column_{i:03d}" for i in range(num_columns)]
        
        # Generate data with various types
        data = {}
        
        for i, col_name in enumerate(column_names):
            # Vary data types for realism
            if i % 5 == 0:
                # Integer data
                data[col_name] = np.random.randint(1, 1000, size=num_rows)
            elif i % 5 == 1:
                # Float data
                data[col_name] = np.random.uniform(0, 100, size=num_rows).round(2)
            elif i % 5 == 2:
                # Text data
                data[col_name] = [f"Text_{i}_{j}" for j in range(num_rows)]
            elif i % 5 == 3:
                # Boolean data (as Yes/No)
                data[col_name] = np.random.choice(['Yes', 'No'], size=num_rows)
            else:
                # Mixed numeric data (some large numbers)
                data[col_name] = np.random.randint(100000, 999999, size=num_rows)
        
        # Create DataFrame
        df = pd.DataFrame(data)
        
        # Add some NaN values for realism
        mask = np.random.random(df.shape) < 0.05  # 5% missing values
        df = df.mask(mask)
        
        # Save to Excel
        output_path = Path(filename)
        df.to_excel(output_path, index=False, engine='openpyxl')
        
        logger.info(f"✓ Sample Excel file created: {output_path.absolute()}")
        logger.info(f"  Size: {output_path.stat().st_size / 1024:.2f} KB")
        logger.info(f"  Shape: {df.shape[0]} rows × {df.shape[1]} columns")
        
        return str(output_path.absolute())
        
    except Exception as e:
        logger.error(f"Error generating sample file: {str(e)}", exc_info=True)
        raise


def generate_complex_sample(
    filename: str = "sample_complex.xlsx",
    num_rows: int = 30
):
    """
    Generate a more complex sample with realistic business data.
    
    Args:
        filename: Output filename
        num_rows: Number of data rows
    """
    try:
        logger.info(f"Generating complex sample Excel file: {filename}")
        
        # Create realistic column structure
        columns = {
            'Employee_ID': [f"EMP{1000+i}" for i in range(num_rows)],
            'First_Name': [f"FirstName{i}" for i in range(num_rows)],
            'Last_Name': [f"LastName{i}" for i in range(num_rows)],
            'Department': np.random.choice(['Sales', 'Marketing', 'Engineering', 'HR', 'Finance'], num_rows),
            'Position': np.random.choice(['Manager', 'Analyst', 'Developer', 'Coordinator', 'Director'], num_rows),
            'Hire_Date': pd.date_range(start='2015-01-01', periods=num_rows, freq='M'),
            'Salary': np.random.randint(40000, 150000, num_rows),
            'Bonus': np.random.uniform(0, 20000, num_rows).round(2),
            'Years_Experience': np.random.randint(0, 25, num_rows),
            'Email': [f"employee{i}@company.com" for i in range(num_rows)],
            'Phone': [f"+1-555-{1000+i:04d}" for i in range(num_rows)],
            'Address': [f"{100+i} Main St" for i in range(num_rows)],
            'City': np.random.choice(['New York', 'Los Angeles', 'Chicago', 'Houston', 'Phoenix'], num_rows),
            'State': np.random.choice(['NY', 'CA', 'IL', 'TX', 'AZ'], num_rows),
            'ZIP': np.random.randint(10000, 99999, num_rows),
        }
        
        # Add many performance metrics (to increase column count)
        for quarter in range(1, 5):
            for year in [2023, 2024]:
                columns[f'Q{quarter}_{year}_Sales'] = np.random.uniform(10000, 50000, num_rows).round(2)
                columns[f'Q{quarter}_{year}_Target'] = np.random.uniform(15000, 55000, num_rows).round(2)
                columns[f'Q{quarter}_{year}_Achievement'] = np.random.uniform(0.7, 1.3, num_rows).round(2)
        
        # Add additional metrics
        for i in range(1, 51):
            columns[f'Metric_{i:02d}'] = np.random.uniform(0, 100, num_rows).round(2)
        
        df = pd.DataFrame(columns)
        
        # Save to Excel
        output_path = Path(filename)
        df.to_excel(output_path, index=False, engine='openpyxl')
        
        logger.info(f"✓ Complex sample Excel file created: {output_path.absolute()}")
        logger.info(f"  Size: {output_path.stat().st_size / 1024:.2f} KB")
        logger.info(f"  Shape: {df.shape[0]} rows × {df.shape[1]} columns")
        
        return str(output_path.absolute())
        
    except Exception as e:
        logger.error(f"Error generating complex sample: {str(e)}", exc_info=True)
        raise


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Generate sample Excel files for testing')
    parser.add_argument('--rows', '-r', type=int, default=50, help='Number of rows')
    parser.add_argument('--cols', '-c', type=int, default=100, help='Number of columns')
    parser.add_argument('--complex', action='store_true', help='Generate complex sample')
    parser.add_argument('--output', '-o', default=None, help='Output filename')
    
    args = parser.parse_args()
    
    if args.complex:
        filename = args.output or "sample_complex.xlsx"
        generate_complex_sample(filename=filename, num_rows=args.rows)
    else:
        filename = args.output or "sample.xlsx"
        generate_sample_excel(filename=filename, num_rows=args.rows, num_columns=args.cols)
