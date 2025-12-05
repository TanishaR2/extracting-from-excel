# Financial Excel Data Extraction - Project Pandion

## 📊 Extraction Summary

Successfully extracted financial data from complex Excel file with multiple sheets and many columns.

### Source File
- **File:** `Project_Pandion_Copy__Financial_Model_(February_2025).xlsx`
- **Total Sheets:** 2
- **Extraction Date:** December 5, 2025

### Extracted Sheets

#### 1. PF Consolidated (Consolidated Financials)
- **Rows:** 321
- **Columns:** 63
- **Pages Generated:** 5 (15 columns per page)
- **Content:** Pro Forma consolidated financials, Income Statement, Revenue projections, COGS analysis
- **Time Period:** 2023-2029 with TTM data

#### 2. DAH (DAH Financials)
- **Rows:** 191
- **Columns:** 60
- **Pages Generated:** 4 (15 columns per page)
- **Content:** Detailed financial metrics and projections

## 📁 Output Structure

```
extracting-from-excel/output/
├── markdown_PF_Consolidated/
│   ├── index.md                 # Navigation index with sheet info
│   ├── combined_output.md       # All pages combined (419 KB)
│   ├── page_1.md                # Columns 0-14
│   ├── page_2.md                # Columns 15-29
│   ├── page_3.md                # Columns 30-44
│   ├── page_4.md                # Columns 45-59
│   └── page_5.md                # Columns 60-62
│
└── markdown_DAH/
    ├── index.md                 # Navigation index with sheet info
    ├── combined_output.md       # All pages combined (244 KB)
    ├── page_1.md                # Columns 0-14
    ├── page_2.md                # Columns 15-29
    ├── page_3.md                # Columns 30-44
    └── page_4.md                # Columns 45-59
```

## 🎯 Key Features of Extraction

### ✅ Preserved Data Relationships
- All column relationships maintained
- Row-column structure intact
- Multi-year financial projections preserved
- Header rows and metadata included

### ✅ OCR-Compatible Format
- Clean Markdown tables
- Human-readable text format
- Easy to parse programmatically
- Compatible with LLM processing

### ✅ Structured Organization
- Logical page breaks (15 columns per page)
- Clear page labels indicating column ranges
- Individual pages for focused analysis
- Combined files for complete overview

## 📖 How to Use the Extracted Data

### 1. View Individual Pages
Each page contains a slice of columns for easier review:
```bash
cd extracting-from-excel/output/markdown_PF_Consolidated
cat page_1.md  # View first 15 columns
```

### 2. View Complete Data
The combined file contains all pages in one document:
```bash
cat markdown_PF_Consolidated/combined_output.md
```

### 3. Navigate Using Index
Start with the index file to understand the structure:
```bash
cat markdown_PF_Consolidated/index.md
cat markdown_DAH/index.md
```

### 4. Process with Scripts
The Markdown format is easy to parse with Python, awk, or other tools:
```python
import pandas as pd

# Read Markdown table back into pandas
with open('page_1.md', 'r') as f:
    content = f.read()
    # Extract table portion
    # Parse with pandas.read_html() or custom parser
```

## 💡 Data Insights from PF Consolidated Sheet

### Financial Metrics Extracted:
- **Revenue Streams:** VPS, DAH, SLI
- **Time Period:** 2023-2029 projections
- **Key Metrics:**
  - Total Net Revenue
  - YoY Growth %
  - Cost of Goods Sold (COGS)
  - Gross Profit
  - Margin Analysis

### Sample Data Points (Page 1):
- DAH Revenue 2024: $32,111,654.47
- DAH Revenue 2025 Forecast: $30,231,785.12
- DAH COGS % of Revenue 2024: 87.21%
- DAH Gross Profit Margin 2024: 12.79%

## 🔧 Technical Details

### Conversion Settings Used:
- **Format:** Markdown
- **Max Columns per Page:** 15
- **Column Slicing:** Automatic horizontal split
- **Encoding:** UTF-8
- **Table Format:** GitHub-flavored Markdown

### Why Markdown?
1. **OCR-Compatible:** Pure text format, perfect for OCR and LLM processing
2. **Relationship Preservation:** Tables maintain exact structure
3. **Human-Readable:** Easy to review and edit
4. **Version Control Friendly:** Can be tracked in Git
5. **Programmatic Access:** Simple to parse and analyze
6. **No Data Loss:** All values preserved exactly

## 🚀 Next Steps for Data Analysis

### Option 1: Re-import to Pandas
```python
import pandas as pd

# If you need to work with the data again
df = pd.read_excel('Project_Pandion_Copy__Financial_Model_(February_2025).xlsx', 
                   sheet_name='PF Consolidated')
```

### Option 2: Parse Markdown Tables
```python
# Use the markdown files for analysis
import re

def parse_markdown_table(md_file):
    with open(md_file, 'r') as f:
        content = f.read()
    # Parse table using regex or markdown parser
    return parsed_data
```

### Option 3: Use for LLM Analysis
- Feed markdown files directly to GPT-4, Claude, or other LLMs
- Ask questions about the financial data
- Generate reports and insights
- Extract specific metrics

### Option 4: Create Visualizations
```python
import matplotlib.pyplot as plt
import pandas as pd

# Extract revenue data from markdown
# Create charts showing:
# - Revenue trends 2023-2029
# - Margin analysis
# - YoY growth comparisons
```

## 📊 Column Distribution

### PF Consolidated (63 columns):
- **Page 1 (Cols 0-14):** Basic financials, revenue, COGS
- **Page 2 (Cols 15-29):** Operating expenses, departmental costs
- **Page 3 (Cols 30-44):** Additional financial metrics
- **Page 4 (Cols 45-59):** Extended projections
- **Page 5 (Cols 60-62):** Final columns

### DAH (60 columns):
- **Page 1 (Cols 0-14):** Primary financial data
- **Page 2 (Cols 15-29):** Detailed breakdowns
- **Page 3 (Cols 30-44):** Additional metrics
- **Page 4 (Cols 45-59):** Extended data

## ⚙️ Regenerating or Customizing Output

### Change Columns per Page
```bash
cd /path/to/newone
python3 run_enhanced.py \
    --input "extracting-from-excel/Project_Pandion_Copy__Financial_Model_(February_2025).xlsx" \
    --format markdown \
    --all-sheets \
    --max-cols 20 \
    --output-dir extracting-from-excel/output_20col
```

### Extract Specific Sheet Only
```bash
python3 run_enhanced.py \
    --input "extracting-from-excel/Project_Pandion_Copy__Financial_Model_(February_2025).xlsx" \
    --format markdown \
    --sheet "PF Consolidated" \
    --max-cols 12 \
    --output-dir extracting-from-excel/output_pf_only
```

### Generate PDF Instead (if WeasyPrint working)
```bash
python3 run_enhanced.py \
    --input "extracting-from-excel/Project_Pandion_Copy__Financial_Model_(February_2025).xlsx" \
    --format pdf \
    --all-sheets \
    --max-cols 12 \
    --output-dir extracting-from-excel/output_pdf
```

### Generate Both Formats
```bash
python3 run_enhanced.py \
    --input "extracting-from-excel/Project_Pandion_Copy__Financial_Model_(February_2025).xlsx" \
    --format both \
    --all-sheets \
    --max-cols 15 \
    --output-dir extracting-from-excel/output_both
```

## 🎓 Understanding the Financial Data

### Key Sections in PF Consolidated:
1. **Header Rows:** Model info, time periods, forecast types
2. **Revenue Section:** VPS, DAH, SLI revenue streams
3. **COGS Section:** Cost breakdown, D&A, margins
4. **Gross Profit:** Profitability analysis
5. **Operating Expenses:** (in subsequent pages)

### Financial Periods:
- **2023:** Actual data
- **2024:** Mix (partial actual, partial forecast)
- **2025-2029:** Forecast data
- **TTM:** Trailing Twelve Months

## 📈 Data Quality Notes

- ✅ All 321 rows from PF Consolidated extracted
- ✅ All 191 rows from DAH extracted
- ✅ All 63 and 60 columns respectively preserved
- ✅ Numeric precision maintained
- ✅ Empty cells represented as blank
- ✅ Special characters handled
- ✅ Column relationships intact

## 🔍 Search and Grep Examples

Find specific financial metrics:
```bash
# Search for revenue data
grep -i "revenue" markdown_PF_Consolidated/combined_output.md

# Find margin percentages
grep "%" markdown_PF_Consolidated/combined_output.md

# Search for specific years
grep "2025" markdown_PF_Consolidated/combined_output.md

# Find COGS data
grep -i "cogs" markdown_PF_Consolidated/combined_output.md
```

## 📝 Notes

- Files are UTF-8 encoded for maximum compatibility
- Markdown tables use GitHub-flavored Markdown syntax
- Pipe characters (|) in data are escaped to prevent table breaking
- NaN/empty values shown as blank cells
- Float values formatted to 2 decimal places where appropriate
- Large numbers preserved without scientific notation

## 🤝 Support

For questions about:
- **Data Extraction:** Review the conversion logs
- **Missing Data:** Check individual page files
- **Custom Extraction:** Modify `run_enhanced.py` parameters
- **Re-processing:** Use the commands in "Regenerating" section

---

**Extraction Tool:** Excel to Markdown/PDF Converter  
**Version:** 1.0  
**Date:** December 5, 2025  
**Status:** ✅ Complete - All data successfully extracted
