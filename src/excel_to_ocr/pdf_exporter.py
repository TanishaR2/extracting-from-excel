"""
PDF Exporter Module
Converts HTML content to PDF using WeasyPrint and combines multiple PDFs.
"""

import logging
from pathlib import Path
from typing import List, Dict, Any
from weasyprint import HTML, CSS
from PyPDF2 import PdfMerger
import os

logger = logging.getLogger(__name__)


class PDFExporter:
    """
    Exports HTML content to PDF files using WeasyPrint.
    Handles PDF generation and merging of multiple pages.
    """
    
    def __init__(self, output_dir: str = "output", config: Dict[str, Any] = None):
        """
        Initialize the PDFExporter.
        
        Args:
            output_dir: Directory for output PDF files
            config: Configuration dictionary with PDF settings
        """
        self.output_dir = Path(output_dir)
        self.config = config or {}
        
        # Create output directory if it doesn't exist
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"PDFExporter initialized with output directory: {output_dir}")
    
    def html_to_pdf(self, html_content: str, output_path: str) -> bool:
        """
        Convert HTML content to a PDF file.
        
        Args:
            html_content: HTML string to convert
            output_path: Path where PDF should be saved
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.debug(f"Converting HTML to PDF: {output_path}")
            
            # Get PDF configuration
            pdf_config = self.config.get('pdf', {})
            
            # Create CSS for page setup
            css_string = self._generate_page_css(pdf_config)
            
            # Convert HTML to PDF with improved error handling
            html_obj = HTML(string=html_content)
            css_obj = CSS(string=css_string)
            
            # Write PDF with compatibility settings
            html_obj.write_pdf(
                output_path,
                stylesheets=[css_obj],
                optimize_size=('fonts',)
            )
            
            logger.info(f"Successfully created PDF: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error converting HTML to PDF: {str(e)}", exc_info=True)
            # Try alternative method without optimization
            try:
                logger.info("Attempting fallback PDF generation method...")
                html_obj = HTML(string=html_content)
                html_obj.write_pdf(output_path)
                logger.info(f"Successfully created PDF using fallback method: {output_path}")
                return True
            except Exception as e2:
                logger.error(f"Fallback method also failed: {str(e2)}", exc_info=True)
                return False
    
    def _generate_page_css(self, pdf_config: Dict[str, Any]) -> str:
        """
        Generate CSS for PDF page layout.
        
        Args:
            pdf_config: PDF configuration dictionary
            
        Returns:
            CSS string
        """
        orientation = pdf_config.get('orientation', 'landscape')
        page_size = pdf_config.get('page_size', 'A4')
        
        # Determine page size
        if orientation == 'landscape':
            size = f"{page_size} landscape"
        else:
            size = page_size
        
        css = f"""
        @page {{
            size: {size};
            margin-top: {pdf_config.get('margin_top', '15mm')};
            margin-right: {pdf_config.get('margin_right', '10mm')};
            margin-bottom: {pdf_config.get('margin_bottom', '15mm')};
            margin-left: {pdf_config.get('margin_left', '10mm')};
        }}
        
        body {{
            font-family: {pdf_config.get('font_family', 'Arial, sans-serif')};
            font-size: {pdf_config.get('font_size', '10pt')};
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            page-break-inside: auto;
        }}
        
        tr {{
            page-break-inside: avoid;
            page-break-after: auto;
        }}
        
        th, td {{
            page-break-inside: avoid;
        }}
        """
        
        return css
    
    def combine_pdfs(self, pdf_paths: List[str], output_path: str) -> bool:
        """
        Combine multiple PDF files into a single PDF.
        
        Args:
            pdf_paths: List of PDF file paths to combine
            output_path: Path for the combined PDF
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Combining {len(pdf_paths)} PDFs into {output_path}")
            
            # Create PDF merger
            merger = PdfMerger()
            
            # Add each PDF
            for pdf_path in pdf_paths:
                if not os.path.exists(pdf_path):
                    logger.warning(f"PDF file not found: {pdf_path}")
                    continue
                
                try:
                    merger.append(pdf_path)
                    logger.debug(f"Added to merger: {pdf_path}")
                except Exception as e:
                    logger.error(f"Error adding PDF {pdf_path}: {str(e)}")
                    continue
            
            # Write combined PDF
            merger.write(output_path)
            merger.close()
            
            logger.info(f"Successfully combined PDFs into: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error combining PDFs: {str(e)}", exc_info=True)
            return False
    
    def export_dataframe_slice(
        self,
        html_content: str,
        page_num: int
    ) -> str:
        """
        Export a single DataFrame slice to PDF.
        
        Args:
            html_content: HTML content to convert
            page_num: Page number for file naming
            
        Returns:
            Path to the generated PDF file
        """
        try:
            # Generate output filename
            output_filename = f"page_{page_num}.pdf"
            output_path = self.output_dir / output_filename
            
            # Convert to PDF
            success = self.html_to_pdf(html_content, str(output_path))
            
            if success:
                return str(output_path)
            else:
                raise Exception(f"Failed to create PDF for page {page_num}")
                
        except Exception as e:
            logger.error(f"Error exporting DataFrame slice: {str(e)}", exc_info=True)
            raise
    
    def cleanup_temp_files(self, file_patterns: List[str] = None):
        """
        Clean up temporary files.
        
        Args:
            file_patterns: List of file patterns to delete (e.g., ['*.html', '*.tmp'])
        """
        try:
            if file_patterns is None:
                return
            
            temp_dir = Path(self.config.get('output', {}).get('temp_dir', 'temp'))
            
            if not temp_dir.exists():
                return
            
            for pattern in file_patterns:
                for file in temp_dir.glob(pattern):
                    try:
                        file.unlink()
                        logger.debug(f"Deleted temp file: {file}")
                    except Exception as e:
                        logger.warning(f"Could not delete {file}: {str(e)}")
            
            logger.info("Cleanup completed")
            
        except Exception as e:
            logger.error(f"Error during cleanup: {str(e)}", exc_info=True)
