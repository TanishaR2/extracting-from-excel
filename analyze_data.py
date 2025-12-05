#!/usr/bin/env python3
"""
Analyze extracted Markdown financial data.
Parse and analyze the financial data from converted Excel files.
"""

import re
import pandas as pd
from pathlib import Path
from typing import List, Dict, Any
import argparse


def parse_markdown_table(md_content: str) -> pd.DataFrame:
    """
    Parse a Markdown table into a pandas DataFrame.
    
    Args:
        md_content: Markdown content containing a table
        
    Returns:
        DataFrame with the parsed table
    """
    # Extract table portion (skip header text)
    lines = md_content.split('\n')
    
    # Find where the table starts
    table_start = None
    for i, line in enumerate(lines):
        if line.strip().startswith('|') and '---' not in line:
            table_start = i
            break
    
    if table_start is None:
        print("No table found in markdown content")
        return pd.DataFrame()
    
    # Get table lines
    table_lines = []
    for line in lines[table_start:]:
        if line.strip().startswith('|'):
            # Skip separator lines
            if not re.match(r'^\|\s*[-:]+', line):
                table_lines.append(line)
        elif table_lines:
            # Table has ended
            break
    
    if len(table_lines) < 2:
        return pd.DataFrame()
    
    # Parse header
    header = table_lines[0]
    headers = [h.strip() for h in header.split('|')[1:-1]]
    
    # Parse data rows
    data = []
    for line in table_lines[2:]:  # Skip header and separator
        if '---' in line:
            continue
        cells = [c.strip() for c in line.split('|')[1:-1]]
        data.append(cells)
    
    # Create DataFrame
    df = pd.DataFrame(data, columns=headers)
    
    # Try to convert numeric columns
    for col in df.columns:
        try:
            # Try converting to numeric
            df[col] = pd.to_numeric(df[col], errors='ignore')
        except:
            pass
    
    return df


def analyze_sheet(sheet_dir: Path) -> Dict[str, Any]:
    """
    Analyze a converted sheet directory.
    
    Args:
        sheet_dir: Path to sheet directory
        
    Returns:
        Analysis results
    """
    results = {
        'sheet_name': sheet_dir.name,
        'files': [],
        'total_columns': 0,
        'total_rows': 0
    }
    
    # Read index file if exists
    index_file = sheet_dir / 'index.md'
    if index_file.exists():
        index_content = index_file.read_text(encoding='utf-8')
        
        # Extract metadata
        match = re.search(r'Total Rows:\*\* (\d+)', index_content)
        if match:
            results['total_rows'] = int(match.group(1))
        
        match = re.search(r'Total Columns:\*\* (\d+)', index_content)
        if match:
            results['total_columns'] = int(match.group(1))
        
        match = re.search(r'Sheet Name:\*\* (.+)', index_content)
        if match:
            results['sheet_display_name'] = match.group(1).strip()
    
    # List page files
    page_files = sorted(sheet_dir.glob('page_*.md'))
    results['num_pages'] = len(page_files)
    results['files'] = [str(f.name) for f in page_files]
    
    # Check for combined file
    combined_file = sheet_dir / 'combined_output.md'
    if combined_file.exists():
        results['combined_file'] = str(combined_file)
        results['combined_size_kb'] = combined_file.stat().st_size / 1024
    
    return results


def search_in_markdown(sheet_dir: Path, search_term: str, case_sensitive: bool = False) -> List[Dict[str, Any]]:
    """
    Search for a term in all markdown files.
    
    Args:
        sheet_dir: Path to sheet directory
        search_term: Term to search for
        case_sensitive: Whether search is case sensitive
        
    Returns:
        List of matches with context
    """
    matches = []
    combined_file = sheet_dir / 'combined_output.md'
    
    if not combined_file.exists():
        print(f"Combined file not found in {sheet_dir}")
        return matches
    
    content = combined_file.read_text(encoding='utf-8')
    lines = content.split('\n')
    
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(re.escape(search_term), flags)
    
    for i, line in enumerate(lines):
        if pattern.search(line):
            # Get context (2 lines before and after)
            context_start = max(0, i - 2)
            context_end = min(len(lines), i + 3)
            context = '\n'.join(lines[context_start:context_end])
            
            matches.append({
                'line_number': i + 1,
                'line': line.strip(),
                'context': context
            })
    
    return matches


def extract_financial_metrics(sheet_dir: Path, metric_name: str) -> pd.DataFrame:
    """
    Extract specific financial metrics across years.
    
    Args:
        sheet_dir: Path to sheet directory
        metric_name: Name of the metric to extract
        
    Returns:
        DataFrame with metric values over time
    """
    combined_file = sheet_dir / 'combined_output.md'
    if not combined_file.exists():
        return pd.DataFrame()
    
    content = combined_file.read_text(encoding='utf-8')
    
    # Find rows matching the metric name
    matches = search_in_markdown(sheet_dir, metric_name, case_sensitive=False)
    
    print(f"\nFound {len(matches)} occurrences of '{metric_name}'")
    
    for match in matches[:5]:  # Show first 5
        print(f"\nLine {match['line_number']}: {match['line']}")
    
    return pd.DataFrame()


def main():
    """Main execution function."""
    parser = argparse.ArgumentParser(description='Analyze extracted financial data')
    parser.add_argument('--dir', '-d', required=True, help='Output directory path')
    parser.add_argument('--search', '-s', help='Search for a term')
    parser.add_argument('--metric', '-m', help='Extract specific metric')
    parser.add_argument('--summary', action='store_true', help='Show summary of all sheets')
    
    args = parser.parse_args()
    
    output_dir = Path(args.dir)
    if not output_dir.exists():
        print(f"Directory not found: {output_dir}")
        return
    
    # Find all sheet directories
    sheet_dirs = [d for d in output_dir.iterdir() if d.is_dir() and d.name.startswith('markdown_')]
    
    if not sheet_dirs:
        print(f"No sheet directories found in {output_dir}")
        return
    
    print("="*70)
    print("Financial Data Analysis")
    print("="*70)
    
    if args.summary:
        # Show summary of all sheets
        for sheet_dir in sheet_dirs:
            analysis = analyze_sheet(sheet_dir)
            print(f"\n📊 Sheet: {analysis.get('sheet_display_name', analysis['sheet_name'])}")
            print(f"   Rows: {analysis['total_rows']}")
            print(f"   Columns: {analysis['total_columns']}")
            print(f"   Pages: {analysis['num_pages']}")
            if analysis.get('combined_size_kb'):
                print(f"   Combined File Size: {analysis['combined_size_kb']:.1f} KB")
    
    if args.search:
        # Search for term
        print(f"\n🔍 Searching for: '{args.search}'")
        for sheet_dir in sheet_dirs:
            matches = search_in_markdown(sheet_dir, args.search)
            if matches:
                print(f"\n📄 Sheet: {sheet_dir.name}")
                print(f"   Found {len(matches)} matches")
                for match in matches[:3]:  # Show first 3
                    print(f"\n   Line {match['line_number']}: {match['line']}")
    
    if args.metric:
        # Extract specific metric
        print(f"\n📈 Extracting metric: '{args.metric}'")
        for sheet_dir in sheet_dirs:
            print(f"\n📄 Sheet: {sheet_dir.name}")
            extract_financial_metrics(sheet_dir, args.metric)
    
    print("\n" + "="*70)


if __name__ == "__main__":
    main()
