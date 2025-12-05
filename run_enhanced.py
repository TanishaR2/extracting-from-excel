#!/usr/bin/env python3
"""
Enhanced CLI runner script for Excel to OCR-friendly PDF/Markdown conversion.

Usage:
    python run_enhanced.py --input file.xlsx --format markdown
    python run_enhanced.py --input file.xlsx --format pdf --max-cols 12
    python run_enhanced.py --input file.xlsx --format both --sheet "Sheet1"
"""

import argparse
import sys
from pathlib import Path
import logging
import pandas as pd

# Add src to path
sys.path.insert(0, str(Path(__file__).parent))

from src.excel_to_ocr.pipeline import ExcelToPDFPipeline
from src.excel_to_ocr.column_slicer import ColumnSlicer
from src.excel_to_ocr.html_renderer import HTMLRenderer
from src.excel_to_ocr.markdown_exporter import MarkdownExporter

# Configure basic logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Convert Excel files to OCR-friendly PDFs or Markdown',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --input data.xlsx --format markdown
  %(prog)s --input data.xlsx --format pdf --max-cols 15
  %(prog)s --input data.xlsx --format both --sheet "Sheet1"
  %(prog)s --input data.xlsx --format markdown --all-sheets
        """
    )
    
    parser.add_argument(
        '--input',
        '-i',
        required=True,
        help='Path to input Excel (.xlsx) file'
    )
    
    parser.add_argument(
        '--format',
        '-f',
        choices=['pdf', 'markdown', 'md', 'both'],
        default='markdown',
        help='Output format: pdf, markdown/md, or both (default: markdown)'
    )
    
    parser.add_argument(
        '--max-cols',
        '-m',
        type=int,
        default=12,
        help='Maximum number of columns per page (default: 12)'
    )
    
    parser.add_argument(
        '--sheet',
        '-s',
        default=None,
        help='Sheet name to process (default: first sheet)'
    )
    
    parser.add_argument(
        '--all-sheets',
        '-a',
        action='store_true',
        help='Process all sheets in the workbook'
    )
    
    parser.add_argument(
        '--output-dir',
        '-o',
        default='output',
        help='Output directory (default: output)'
    )
    
    parser.add_argument(
        '--list-sheets',
        '-l',
        action='store_true',
        help='List all sheets in the Excel file and exit'
    )
    
    parser.add_argument(
        '--verbose',
        '-v',
        action='store_true',
        help='Enable verbose (DEBUG) logging'
    )
    
    return parser.parse_args()


def list_excel_sheets(file_path: str):
    """
    List all sheets in an Excel file.
    
    Args:
        file_path: Path to Excel file
    """
    try:
        logger.info(f"Reading Excel file: {file_path}")
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        
        print("\n" + "="*70)
        print(f"Excel File: {Path(file_path).name}")
        print("="*70)
        print(f"\nTotal Sheets: {len(excel_file.sheet_names)}\n")
        
        for idx, sheet_name in enumerate(excel_file.sheet_names, start=1):
            # Get sheet info
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            print(f"{idx}. '{sheet_name}'")
            print(f"   - Rows: {df.shape[0]}")
            print(f"   - Columns: {df.shape[1]}")
            print()
        
        excel_file.close()
        
    except Exception as e:
        logger.error(f"Error reading Excel file: {str(e)}", exc_info=True)
        sys.exit(1)


def process_markdown_format(
    df: pd.DataFrame,
    sheet_name: str,
    max_cols: int,
    output_dir: Path
) -> dict:
    """
    Process DataFrame and export to Markdown format.
    
    Args:
        df: DataFrame to process
        sheet_name: Name of the sheet
        max_cols: Maximum columns per page
        output_dir: Output directory
        
    Returns:
        Results dictionary
    """
    try:
        # Create sheet-specific output directory
        sheet_dir = output_dir / f"markdown_{sheet_name.replace(' ', '_')}"
        sheet_dir.mkdir(parents=True, exist_ok=True)
        
        # Initialize components
        slicer = ColumnSlicer(max_columns_per_page=max_cols)
        md_exporter = MarkdownExporter(output_dir=str(sheet_dir))
        
        # Get slice information
        slice_info = slicer.get_slice_info(df)
        logger.info(f"Will generate {slice_info['num_pages']} Markdown pages")
        
        # Slice DataFrame
        slices = slicer.slice_dataframe(df)
        
        # Process each slice
        md_paths = []
        for idx, (slice_df, start_col, end_col) in enumerate(slices, start=1):
            logger.info(f"Processing page {idx}/{len(slices)}")
            
            md_path = md_exporter.export_dataframe_slice(
                df=slice_df,
                page_num=idx,
                start_col=start_col,
                end_col=end_col
            )
            md_paths.append(md_path)
        
        # Combine Markdown files
        combined_path = str(sheet_dir / "combined_output.md")
        md_exporter.combine_markdown_files(md_paths, combined_path)
        
        # Create index
        index_content = md_exporter.create_index(
            num_pages=len(md_paths),
            sheet_info={
                'Sheet Name': sheet_name,
                'Total Rows': df.shape[0],
                'Total Columns': df.shape[1],
                'Pages Generated': len(md_paths)
            }
        )
        index_path = str(sheet_dir / "index.md")
        Path(index_path).write_text(index_content, encoding='utf-8')
        
        return {
            'success': True,
            'format': 'markdown',
            'sheet_name': sheet_name,
            'total_pages': len(md_paths),
            'individual_files': md_paths,
            'combined_file': combined_path,
            'index_file': index_path,
            'output_dir': str(sheet_dir)
        }
        
    except Exception as e:
        logger.error(f"Error processing Markdown format: {str(e)}", exc_info=True)
        return {'success': False, 'error': str(e)}


def main():
    """Main execution function."""
    try:
        args = parse_arguments()
        
        # Set logging level
        if args.verbose:
            logging.getLogger().setLevel(logging.DEBUG)
        
        # Validate input file
        if not Path(args.input).exists():
            logger.error(f"Input file not found: {args.input}")
            sys.exit(1)
        
        # List sheets if requested
        if args.list_sheets:
            list_excel_sheets(args.input)
            sys.exit(0)
        
        logger.info("=" * 70)
        logger.info("Excel to OCR-friendly PDF/Markdown Converter")
        logger.info("=" * 70)
        logger.info(f"Input file: {args.input}")
        logger.info(f"Output format: {args.format}")
        logger.info(f"Output directory: {args.output_dir}")
        
        # Create output directory
        output_dir = Path(args.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Read Excel file
        excel_file = pd.ExcelFile(args.input, engine='openpyxl')
        
        # Determine which sheets to process
        if args.all_sheets:
            sheets_to_process = excel_file.sheet_names
            logger.info(f"Processing all {len(sheets_to_process)} sheets")
        elif args.sheet:
            if args.sheet not in excel_file.sheet_names:
                logger.error(f"Sheet '{args.sheet}' not found in Excel file")
                logger.info(f"Available sheets: {', '.join(excel_file.sheet_names)}")
                sys.exit(1)
            sheets_to_process = [args.sheet]
        else:
            sheets_to_process = [excel_file.sheet_names[0]]
            logger.info(f"Processing first sheet: '{sheets_to_process[0]}'")
        
        # Process each sheet
        all_results = []
        
        for sheet_name in sheets_to_process:
            logger.info("\n" + "-" * 70)
            logger.info(f"Processing sheet: '{sheet_name}'")
            logger.info("-" * 70)
            
            # Read sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            logger.info(f"Sheet size: {df.shape[0]} rows × {df.shape[1]} columns")
            
            # Process based on format
            if args.format in ['markdown', 'md', 'both']:
                logger.info("\nGenerating Markdown output...")
                md_results = process_markdown_format(
                    df=df,
                    sheet_name=sheet_name,
                    max_cols=args.max_cols,
                    output_dir=output_dir
                )
                all_results.append(md_results)
                
                if md_results['success']:
                    logger.info(f"✓ Markdown generation completed!")
                    logger.info(f"  Output directory: {md_results['output_dir']}")
                    logger.info(f"  Total pages: {md_results['total_pages']}")
                    logger.info(f"  Combined file: {md_results['combined_file']}")
                    logger.info(f"  Index file: {md_results['index_file']}")
            
            if args.format in ['pdf', 'both']:
                logger.info("\nGenerating PDF output...")
                # Use the original pipeline for PDF
                pipeline = ExcelToPDFPipeline(config_path='src/excel_to_ocr/config.yaml')
                pipeline.config['slicing']['max_columns_per_page'] = args.max_cols
                pipeline.slicer.max_columns_per_page = args.max_cols
                
                # Set sheet-specific output directory
                sheet_output_dir = output_dir / f"pdf_{sheet_name.replace(' ', '_')}"
                pipeline.config['output']['output_dir'] = str(sheet_output_dir)
                pipeline.exporter.output_dir = sheet_output_dir
                pipeline.exporter.output_dir.mkdir(parents=True, exist_ok=True)
                
                # Process
                pdf_results = pipeline.process(args.input, sheet_name=sheet_name)
                all_results.append(pdf_results)
                
                if pdf_results['success']:
                    logger.info(f"✓ PDF generation completed!")
                    logger.info(f"  Total pages: {pdf_results['total_pages']}")
                    if pdf_results.get('combined_pdf'):
                        logger.info(f"  Combined PDF: {pdf_results['combined_pdf']}")
        
        excel_file.close()
        
        # Final summary
        logger.info("\n" + "=" * 70)
        logger.info("CONVERSION COMPLETE")
        logger.info("=" * 70)
        logger.info(f"Processed {len(sheets_to_process)} sheet(s)")
        logger.info(f"Output location: {output_dir.absolute()}")
        logger.info("=" * 70)
        
        # Display results
        for result in all_results:
            if result.get('success'):
                if result.get('format') == 'markdown':
                    logger.info(f"\n📄 Markdown - Sheet '{result['sheet_name']}':")
                    logger.info(f"   {result['output_dir']}")
                elif result.get('total_pages', 0) > 0:
                    logger.info(f"\n📄 PDF - Sheet '{result.get('sheet_name', 'Unknown')}':")
                    logger.info(f"   {result.get('combined_pdf', 'N/A')}")
        
        logger.info("")
        
    except KeyboardInterrupt:
        logger.info("\nOperation cancelled by user")
        sys.exit(130)
    
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
