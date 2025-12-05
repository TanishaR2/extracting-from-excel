"""
Column Slicer Module
Handles slicing of DataFrame columns into manageable chunks for pagination.
"""

import logging
from typing import List, Tuple
import pandas as pd

logger = logging.getLogger(__name__)


class ColumnSlicer:
    """
    Slices a DataFrame into multiple DataFrames based on column count.
    Each slice contains a maximum number of columns for optimal PDF rendering.
    """
    
    def __init__(self, max_columns_per_page: int = 12):
        """
        Initialize the ColumnSlicer.
        
        Args:
            max_columns_per_page: Maximum number of columns per page slice
        """
        self.max_columns_per_page = max_columns_per_page
        logger.info(f"ColumnSlicer initialized with max {max_columns_per_page} columns per page")
    
    def slice_dataframe(self, df: pd.DataFrame) -> List[Tuple[pd.DataFrame, int, int]]:
        """
        Slice a DataFrame into multiple DataFrames based on column count.
        
        Args:
            df: Input DataFrame to slice
            
        Returns:
            List of tuples containing (sliced_df, start_col_idx, end_col_idx)
        """
        try:
            total_columns = len(df.columns)
            logger.info(f"Slicing DataFrame with {total_columns} columns")
            
            if total_columns == 0:
                logger.warning("DataFrame has no columns")
                return []
            
            slices = []
            start_idx = 0
            
            # Slice the DataFrame into chunks
            while start_idx < total_columns:
                end_idx = min(start_idx + self.max_columns_per_page, total_columns)
                
                # Extract the slice
                sliced_df = df.iloc[:, start_idx:end_idx].copy()
                
                logger.debug(f"Created slice: columns {start_idx}-{end_idx-1}")
                slices.append((sliced_df, start_idx, end_idx - 1))
                
                start_idx = end_idx
            
            logger.info(f"Created {len(slices)} slices from DataFrame")
            return slices
            
        except Exception as e:
            logger.error(f"Error slicing DataFrame: {str(e)}", exc_info=True)
            raise
    
    def get_slice_info(self, df: pd.DataFrame) -> dict:
        """
        Get information about how the DataFrame will be sliced.
        
        Args:
            df: Input DataFrame
            
        Returns:
            Dictionary with slicing information
        """
        try:
            total_columns = len(df.columns)
            num_pages = (total_columns + self.max_columns_per_page - 1) // self.max_columns_per_page
            
            info = {
                'total_columns': total_columns,
                'max_columns_per_page': self.max_columns_per_page,
                'num_pages': num_pages,
                'last_page_columns': total_columns % self.max_columns_per_page or self.max_columns_per_page
            }
            
            logger.debug(f"Slice info: {info}")
            return info
            
        except Exception as e:
            logger.error(f"Error getting slice info: {str(e)}", exc_info=True)
            raise
