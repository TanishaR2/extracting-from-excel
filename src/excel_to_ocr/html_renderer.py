"""
HTML Renderer Module
Converts DataFrame slices into HTML using Jinja2 templates.
"""

import logging
from pathlib import Path
from typing import Dict, Any
import pandas as pd
from jinja2 import Environment, FileSystemLoader, Template

logger = logging.getLogger(__name__)


class HTMLRenderer:
    """
    Renders DataFrames as HTML using Jinja2 templates.
    Handles styling and formatting for OCR-friendly output.
    """
    
    def __init__(self, template_dir: str = "templates", config: Dict[str, Any] = None):
        """
        Initialize the HTMLRenderer.
        
        Args:
            template_dir: Directory containing Jinja2 templates
            config: Configuration dictionary with HTML settings
        """
        self.template_dir = Path(template_dir)
        self.config = config or {}
        
        # Set up Jinja2 environment
        try:
            self.env = Environment(loader=FileSystemLoader(str(self.template_dir)))
            logger.info(f"HTMLRenderer initialized with template directory: {template_dir}")
        except Exception as e:
            logger.error(f"Failed to initialize Jinja2 environment: {str(e)}")
            raise
    
    def render_table(
        self,
        df: pd.DataFrame,
        page_num: int,
        start_col: int,
        end_col: int,
        page_label: str = None
    ) -> str:
        """
        Render a DataFrame slice as HTML.
        
        Args:
            df: DataFrame slice to render
            page_num: Current page number
            start_col: Starting column index
            end_col: Ending column index
            page_label: Optional custom page label
            
        Returns:
            HTML string
        """
        try:
            # Load the template
            template = self.env.get_template('table.html')
            
            # Generate page label if not provided
            if page_label is None:
                page_label = self.config.get('output', {}).get(
                    'page_label_format',
                    'Page {page_num} – Columns {start_col}-{end_col}'
                ).format(page_num=page_num, start_col=start_col, end_col=end_col)
            
            # Convert DataFrame to HTML structure
            headers = df.columns.tolist()
            rows = df.values.tolist()
            
            # Handle NaN values
            rows = [
                [self._format_cell(cell) for cell in row]
                for row in rows
            ]
            
            # Get HTML settings from config
            html_config = self.config.get('html', {})
            pdf_config = self.config.get('pdf', {})
            
            # Render the template
            html_content = template.render(
                page_label=page_label,
                headers=headers,
                rows=rows,
                page_num=page_num,
                start_col=start_col,
                end_col=end_col,
                # Style settings
                font_size=pdf_config.get('font_size', '10pt'),
                font_family=pdf_config.get('font_family', 'Arial, sans-serif'),
                table_border=html_config.get('table_border', '1px solid #333'),
                header_bg_color=html_config.get('header_bg_color', '#4CAF50'),
                header_text_color=html_config.get('header_text_color', '#ffffff'),
                cell_padding=html_config.get('cell_padding', '8px'),
                stripe_rows=html_config.get('stripe_rows', True),
                stripe_color=html_config.get('stripe_color', '#f2f2f2')
            )
            
            logger.debug(f"Rendered HTML for page {page_num}")
            return html_content
            
        except Exception as e:
            logger.error(f"Error rendering HTML for page {page_num}: {str(e)}", exc_info=True)
            raise
    
    def _format_cell(self, value: Any) -> str:
        """
        Format a cell value for HTML display.
        
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
    
    def create_index_page(self, num_pages: int, output_dir: str) -> str:
        """
        Create an index page listing all generated PDF pages.
        
        Args:
            num_pages: Total number of pages
            output_dir: Output directory path
            
        Returns:
            HTML string for index page
        """
        try:
            html = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>PDF Conversion Index</title>
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                        margin: 40px;
                    }}
                    h1 {{
                        color: #333;
                    }}
                    ul {{
                        list-style-type: none;
                        padding: 0;
                    }}
                    li {{
                        padding: 10px;
                        margin: 5px 0;
                        background-color: #f0f0f0;
                        border-radius: 5px;
                    }}
                </style>
            </head>
            <body>
                <h1>PDF Conversion Complete</h1>
                <p>Generated {num_pages} pages</p>
                <ul>
            """
            
            for i in range(1, num_pages + 1):
                html += f"<li>Page {i}: page_{i}.pdf</li>\n"
            
            html += """
                </ul>
                <p><strong>Combined output:</strong> combined_output.pdf</p>
            </body>
            </html>
            """
            
            return html
            
        except Exception as e:
            logger.error(f"Error creating index page: {str(e)}", exc_info=True)
            raise
