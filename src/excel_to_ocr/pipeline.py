"""
Pipeline Module
Orchestrates the complete Excel to PDF conversion process.
"""

import logging
from pathlib import Path
from typing import Dict, Any, Optional
import pandas as pd
import yaml

from .column_slicer import ColumnSlicer
from .html_renderer import HTMLRenderer
from .pdf_exporter import PDFExporter

logger = logging.getLogger(__name__)


class ExcelToPDFPipeline:
    """
    Main pipeline for converting Excel files to OCR-friendly PDFs.
    Orchestrates the entire conversion process from XLSX to multi-page PDFs.
    """
    
    def __init__(self, config_path: str = "src/config.yaml"):
        """
        Initialize the pipeline with configuration.
        
        Args:
            config_path: Path to configuration YAML file
        """
        self.config = self._load_config(config_path)
        self._setup_logging()
        
        # Initialize components
        max_cols = self.config.get('slicing', {}).get('max_columns_per_page', 12)
        output_dir = self.config.get('output', {}).get('output_dir', 'output')
        
        self.slicer = ColumnSlicer(max_columns_per_page=max_cols)
        self.renderer = HTMLRenderer(template_dir="templates", config=self.config)
        self.exporter = PDFExporter(output_dir=output_dir, config=self.config)
        
        logger.info("ExcelToPDFPipeline initialized successfully")
    
    def _load_config(self, config_path: str) -> Dict[str, Any]:
        """
        Load configuration from YAML file.
        
        Args:
            config_path: Path to config file
            
        Returns:
            Configuration dictionary
        """
        try:
            with open(config_path, 'r') as f:
                config = yaml.safe_load(f)
            logger.info(f"Configuration loaded from {config_path}")
            return config
        except Exception as e:
            logger.warning(f"Could not load config from {config_path}: {str(e)}")
            logger.info("Using default configuration")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """
        Get default configuration.
        
        Returns:
            Default configuration dictionary
        """
        return {
            'pdf': {
                'page_size': 'A4',
                'orientation': 'landscape',
                'margin_top': '15mm',
                'margin_right': '10mm',
                'margin_bottom': '15mm',
                'margin_left': '10mm',
                'font_size': '10pt',
                'font_family': 'Arial, sans-serif'
            },
            'slicing': {
                'max_columns_per_page': 12,
                'repeat_header': True,
                'header_rows': 1
            },
            'output': {
                'output_dir': 'output',
                'temp_dir': 'temp',
                'combine_pdfs': True,
                'keep_individual_pages': True,
                'page_label_format': 'Page {page_num} – Columns {start_col}-{end_col}'
            },
            'logging': {
                'level': 'INFO',
                'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            },
            'html': {
                'table_border': '1px solid #333',
                'header_bg_color': '#4CAF50',
                'header_text_color': '#ffffff',
                'cell_padding': '8px',
                'stripe_rows': True,
                'stripe_color': '#f2f2f2'
            }
        }
    
    def _setup_logging(self):
        """
        Set up logging configuration.
        """
        log_config = self.config.get('logging', {})
        log_level = log_config.get('level', 'INFO')
        log_format = log_config.get('format', '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        log_file = log_config.get('file')
        
        # Create logs directory if needed
        if log_file:
            log_path = Path(log_file)
            log_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Configure logging
        handlers = [logging.StreamHandler()]
        if log_file:
            handlers.append(logging.FileHandler(log_file))
        
        logging.basicConfig(
            level=getattr(logging, log_level),
            format=log_format,
            handlers=handlers
        )
    
    def read_excel(self, file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """
        Read Excel file into a DataFrame.
        
        Args:
            file_path: Path to Excel file
            sheet_name: Optional sheet name to read (default: first sheet)
            
        Returns:
            DataFrame containing Excel data
        """
        try:
            logger.info(f"Reading Excel file: {file_path}")
            
            # Read Excel file
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            else:
                df = pd.read_excel(file_path, engine='openpyxl')
            
            logger.info(f"Successfully read Excel file: {df.shape[0]} rows, {df.shape[1]} columns")
            return df
            
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}", exc_info=True)
            raise
    
    def process(self, input_file: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """
        Process Excel file and convert to PDFs.
        
        Args:
            input_file: Path to input Excel file
            sheet_name: Optional sheet name to process
            
        Returns:
            Dictionary with processing results
        """
        try:
            logger.info(f"Starting pipeline processing for: {input_file}")
            
            # Step 1: Read Excel file
            df = self.read_excel(input_file, sheet_name)
            
            # Step 2: Get slice information
            slice_info = self.slicer.get_slice_info(df)
            logger.info(f"Will generate {slice_info['num_pages']} pages")
            
            # Step 3: Slice DataFrame
            slices = self.slicer.slice_dataframe(df)
            
            # Step 4: Process each slice
            pdf_paths = []
            for idx, (slice_df, start_col, end_col) in enumerate(slices, start=1):
                try:
                    logger.info(f"Processing page {idx}/{len(slices)}")
                    
                    # Render HTML
                    html_content = self.renderer.render_table(
                        df=slice_df,
                        page_num=idx,
                        start_col=start_col,
                        end_col=end_col
                    )
                    
                    # Export to PDF
                    pdf_path = self.exporter.export_dataframe_slice(
                        html_content=html_content,
                        page_num=idx
                    )
                    
                    pdf_paths.append(pdf_path)
                    logger.info(f"Page {idx} completed: {pdf_path}")
                    
                except Exception as e:
                    logger.error(f"Error processing page {idx}: {str(e)}", exc_info=True)
                    continue
            
            # Step 5: Combine PDFs if configured
            combined_path = None
            if self.config.get('output', {}).get('combine_pdfs', True) and len(pdf_paths) > 0:
                output_dir = Path(self.config.get('output', {}).get('output_dir', 'output'))
                combined_path = str(output_dir / "combined_output.pdf")
                
                success = self.exporter.combine_pdfs(pdf_paths, combined_path)
                if success:
                    logger.info(f"Combined PDF created: {combined_path}")
                else:
                    logger.warning("Failed to create combined PDF")
                    combined_path = None
            
            # Step 6: Clean up individual pages if configured
            if not self.config.get('output', {}).get('keep_individual_pages', True):
                for pdf_path in pdf_paths:
                    try:
                        Path(pdf_path).unlink()
                        logger.debug(f"Deleted individual page: {pdf_path}")
                    except Exception as e:
                        logger.warning(f"Could not delete {pdf_path}: {str(e)}")
            
            # Prepare results
            results = {
                'success': True,
                'input_file': input_file,
                'total_pages': len(pdf_paths),
                'individual_pdfs': pdf_paths,
                'combined_pdf': combined_path,
                'slice_info': slice_info
            }
            
            logger.info("Pipeline processing completed successfully")
            return results
            
        except Exception as e:
            logger.error(f"Pipeline processing failed: {str(e)}", exc_info=True)
            return {
                'success': False,
                'error': str(e),
                'input_file': input_file
            }
    
    def get_config(self) -> Dict[str, Any]:
        """
        Get current configuration.
        
        Returns:
            Configuration dictionary
        """
        return self.config
