"""
Markdown Exporter Module
Converts DataFrame slices into Markdown format for better data extraction and OCR compatibility.
"""

import logging
from pathlib import Path
from typing import List, Dict, Any
import pandas as pd

logger = logging.getLogger(__name__)


class MarkdownExporter:
    """
    Exports DataFrames as Markdown files for OCR and data extraction.
    Markdown is ideal for preserving table structure and relationships.
    """
    
    def __init__(self, output_dir: str = "output", config: Dict[str, Any] = None):
        """
        Initialize the MarkdownExporter.
        
        Args:
            output_dir: Directory for output Markdown files
            config: Configuration dictionary
        """
        self.output_dir = Path(output_dir)
        self.config = config or {}
        
        # Create output directory if it doesn't exist
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"MarkdownExporter initialized with output directory: {output_dir}")
    
    def dataframe_to_markdown(
        self,
        df: pd.DataFrame,
        page_num: int,
        start_col: int,
        end_col: int,
        include_index: bool = False
    ) -> str:
        """
        Convert a DataFrame to Markdown table format.
        
        Args:
            df: DataFrame to convert
            page_num: Current page number
            start_col: Starting column index
            end_col: Ending column index
            include_index: Whether to include the DataFrame index
            
        Returns:
            Markdown string
        """
        try:
            # Add page header
            page_label = self.config.get('output', {}).get(
                'page_label_format',
                'Page {page_num} – Columns {start_col}-{end_col}'
            ).format(page_num=page_num, start_col=start_col, end_col=end_col)
            
            markdown = f"# {page_label}\n\n"
            markdown += f"**Page:** {page_num} | **Columns:** {start_col} to {end_col}\n\n"
            markdown += "---\n\n"
            
            # Convert DataFrame to Markdown
            # Use pandas built-in to_markdown if available
            try:
                table_md = df.to_markdown(index=include_index)
            except AttributeError:
                # Fallback for older pandas versions
                table_md = self._manual_markdown_conversion(df, include_index)
            
            markdown += table_md
            markdown += "\n\n---\n\n"
            
            logger.debug(f"Converted DataFrame to Markdown for page {page_num}")
            return markdown
            
        except Exception as e:
            logger.error(f"Error converting DataFrame to Markdown: {str(e)}", exc_info=True)
            raise
    
    def _manual_markdown_conversion(self, df: pd.DataFrame, include_index: bool = False) -> str:
        """
        Manually convert DataFrame to Markdown (fallback method).
        
        Args:
            df: DataFrame to convert
            include_index: Whether to include the index
            
        Returns:
            Markdown table string
        """
        # Build header
        if include_index:
            headers = ['Index'] + list(df.columns)
        else:
            headers = list(df.columns)
        
        # Escape pipe characters in headers
        headers = [str(h).replace('|', '\\|') for h in headers]
        
        # Create header row
        markdown = '| ' + ' | '.join(headers) + ' |\n'
        
        # Create separator row
        markdown += '| ' + ' | '.join(['---' for _ in headers]) + ' |\n'
        
        # Add data rows
        for idx, row in df.iterrows():
            if include_index:
                row_data = [str(idx)] + [self._format_cell(val) for val in row]
            else:
                row_data = [self._format_cell(val) for val in row]
            
            # Escape pipe characters
            row_data = [str(val).replace('|', '\\|') for val in row_data]
            markdown += '| ' + ' | '.join(row_data) + ' |\n'
        
        return markdown
    
    def _format_cell(self, value: Any) -> str:
        """
        Format a cell value for Markdown display.
        
        Args:
            value: Cell value to format
            
        Returns:
            Formatted string
        """
        if pd.isna(value):
            return ""
        elif isinstance(value, float):
            # Format floats to avoid scientific notation
            return f"{value:.2f}" if value != int(value) else str(int(value))
        else:
            return str(value)
    
    def export_dataframe_slice(
        self,
        df: pd.DataFrame,
        page_num: int,
        start_col: int,
        end_col: int
    ) -> str:
        """
        Export a single DataFrame slice to Markdown file.
        
        Args:
            df: DataFrame slice to export
            page_num: Page number for file naming
            start_col: Starting column index
            end_col: Ending column index
            
        Returns:
            Path to the generated Markdown file
        """
        try:
            # Generate output filename
            output_filename = f"page_{page_num}.md"
            output_path = self.output_dir / output_filename
            
            # Convert to Markdown
            markdown_content = self.dataframe_to_markdown(
                df=df,
                page_num=page_num,
                start_col=start_col,
                end_col=end_col
            )
            
            # Write to file
            output_path.write_text(markdown_content, encoding='utf-8')
            
            logger.info(f"Successfully created Markdown file: {output_path}")
            return str(output_path)
            
        except Exception as e:
            logger.error(f"Error exporting DataFrame slice to Markdown: {str(e)}", exc_info=True)
            raise
    
    def combine_markdown_files(self, markdown_paths: List[str], output_path: str) -> bool:
        """
        Combine multiple Markdown files into a single file.
        
        Args:
            markdown_paths: List of Markdown file paths to combine
            output_path: Path for the combined Markdown file
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Combining {len(markdown_paths)} Markdown files into {output_path}")
            
            combined_content = "# Combined Excel Data Export\n\n"
            combined_content += f"This document contains {len(markdown_paths)} pages of data.\n\n"
            combined_content += "---\n\n"
            
            # Read and combine all Markdown files
            for md_path in markdown_paths:
                if not Path(md_path).exists():
                    logger.warning(f"Markdown file not found: {md_path}")
                    continue
                
                try:
                    content = Path(md_path).read_text(encoding='utf-8')
                    combined_content += content + "\n\n"
                    logger.debug(f"Added to combined file: {md_path}")
                except Exception as e:
                    logger.error(f"Error reading Markdown file {md_path}: {str(e)}")
                    continue
            
            # Write combined file
            Path(output_path).write_text(combined_content, encoding='utf-8')
            
            logger.info(f"Successfully created combined Markdown file: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error combining Markdown files: {str(e)}", exc_info=True)
            return False
    
    def create_index(self, num_pages: int, sheet_info: Dict[str, Any] = None) -> str:
        """
        Create an index/table of contents in Markdown.
        
        Args:
            num_pages: Total number of pages
            sheet_info: Optional information about the Excel sheet
            
        Returns:
            Markdown string for index
        """
        try:
            markdown = "# Excel Data Extraction Index\n\n"
            
            if sheet_info:
                markdown += "## Source Information\n\n"
                for key, value in sheet_info.items():
                    markdown += f"- **{key}:** {value}\n"
                markdown += "\n"
            
            markdown += "## Pages\n\n"
            markdown += f"Total pages generated: {num_pages}\n\n"
            
            for i in range(1, num_pages + 1):
                markdown += f"- [Page {i}](page_{i}.md)\n"
            
            markdown += "\n---\n\n"
            markdown += "*Generated by Excel to Markdown Converter*\n"
            
            return markdown
            
        except Exception as e:
            logger.error(f"Error creating index: {str(e)}", exc_info=True)
            raise
