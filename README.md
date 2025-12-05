# Excel to OCR-Friendly Converter

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![UV](https://img.shields.io/badge/uv-package%20manager-green.svg)](https://github.com/astral-sh/uv)

Convert complex Excel XLSX files with 100+ columns into OCR-friendly formats (Markdown/PDF) by automatically splitting wide tables horizontally.

## 🎯 Perfect For

- **Financial Models** with many columns
- **Data Extraction** from complex spreadsheets  
- **OCR Processing** of wide tables
- **AI/LLM Analysis** of Excel data
- **Report Generation** from financial data

## ✨ Key Features

- ✅ **Automatic Column Slicing** - Splits wide tables horizontally into manageable pages
- ✅ **OCR-Friendly Output** - Large, readable text (≥10pt), no shrink-to-fit
- ✅ **Multiple Formats** - Markdown (perfect for OCR/AI) and PDF support
- ✅ **Preserves Relationships** - All data connections maintained
- ✅ **Configurable** - Adjust columns per page, styling, and more
- ✅ **CLI Interface** - Easy command-line usage
- ✅ **Production-Ready** - Comprehensive logging and error handling

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/TanishaR2/extracting-from-excel.git
cd extracting-from-excel

# Install with UV (recommended)
uv sync

# Or with pip
pip install -e .
```

### Basic Usage

```bash
# List sheets in Excel file
python run_enhanced.py --input your_file.xlsx --list-sheets

# Convert to Markdown (best for OCR/AI)
python run_enhanced.py --input your_file.xlsx --format markdown --all-sheets

# Convert to PDF
python run_enhanced.py --input your_file.xlsx --format pdf --all-sheets

# Convert with custom column count
python run_enhanced.py --input your_file.xlsx --format markdown --max-cols 20
```

## 📊 Example Use Case

**Input:** Excel file with 63 columns × 321 rows (financial model)  
**Output:** 5 Markdown pages (15 columns each) or combined PDF  
**Result:** 100% data preserved, OCR-ready, AI-compatible

## 🎯 How It Works

```
Excel File (100+ columns)
    ↓
Column Slicer (splits into chunks)
    ↓
Renderer (HTML or Markdown)
    ↓
Exporter (PDF or .md files)
    ↓
Combined Output (multi-page document)
```

## 📖 Documentation

- [Complete Documentation](docs/README.md) - Full project documentation
- [Quick Start Guide](output/QUICK_START_GUIDE.md) - Get started immediately
- [Extraction Guide](output/EXTRACTION_README.md) - Understanding your data
- [Project Summary](docs/PROJECT_SUMMARY.md) - Technical overview

## 💡 Use Cases

### 1. Financial Data Extraction
Extract data from complex financial models with many columns for:
- AI/LLM analysis (ChatGPT, Claude)
- Data processing pipelines
- Report generation
- Audit trails

### 2. OCR Processing
Convert wide Excel tables to OCR-friendly formats:
- Large, readable text
- Proper table structure
- No image quality loss
- Vector text in PDFs

### 3. AI/LLM Integration
Feed extracted data directly to AI models:
- Markdown format preferred by LLMs
- Preserves table structure
- Searchable and parseable
- Token-efficient

## 🛠️ Technology Stack

- **Python 3.8+** - Core language
- **pandas & openpyxl** - Excel file handling
- **Jinja2** - HTML templating
- **WeasyPrint** - PDF generation
- **PyPDF2** - PDF manipulation
- **UV** - Fast Python package manager

## 📦 Project Structure

```
extracting-from-excel/
├── src/excel_to_ocr/       # Core conversion modules
│   ├── column_slicer.py    # Horizontal table splitting
│   ├── html_renderer.py    # HTML template rendering
│   ├── pdf_exporter.py     # PDF generation
│   ├── markdown_exporter.py # Markdown export
│   ├── pipeline.py         # Orchestration
│   └── config.yaml         # Configuration
├── templates/              # Jinja2 templates
├── output/                 # Generated files
├── docs/                   # Documentation
├── examples/               # Example files
├── run_enhanced.py         # Main CLI tool
├── analyze_data.py         # Data analysis utility
├── generate_sample.py      # Test file generator
└── pyproject.toml          # Project configuration
```

## 🔧 Configuration

Edit `src/excel_to_ocr/config.yaml` to customize:

```yaml
pdf:
  page_size: "A4"
  orientation: "landscape"
  font_size: "10pt"

slicing:
  max_columns_per_page: 12

output:
  output_dir: "output"
  combine_pdfs: true
```

## 📊 Supported Formats

### Output Formats
- **Markdown** (.md) - Best for OCR, AI/LLM, text processing
- **PDF** - For printing, sharing, archival
- **Both** - Generate both formats simultaneously

### Input Requirements
- Excel files (.xlsx)
- Any number of columns (tested with 100+)
- Multiple sheets supported
- Merged cells handled

## 🎓 Advanced Features

### Multiple Sheet Processing
```bash
# Process all sheets
python run_enhanced.py --input file.xlsx --all-sheets

# Process specific sheet
python run_enhanced.py --input file.xlsx --sheet "Sheet1"
```

### Custom Styling
```bash
# More columns per page
python run_enhanced.py --input file.xlsx --max-cols 20

# Custom output directory
python run_enhanced.py --input file.xlsx --output-dir my_output
```

### Data Analysis
```bash
# Show summary
python analyze_data.py --dir output --summary

# Search for terms
python analyze_data.py --dir output --search "Revenue"
```

## 🧪 Testing

```bash
# Generate test files
python generate_sample.py --rows 50 --cols 100

# Generate complex sample
python generate_sample.py --complex --rows 30

# Test conversion
python run_enhanced.py --input sample.xlsx --format markdown
```

## 📈 Real-World Example

Successfully processed a financial model with:
- **2 sheets** (PF Consolidated, DAH)
- **63 columns** in first sheet
- **321 rows** of financial data
- **Result:** 9 total Markdown pages, 663 KB total output
- **Status:** 100% data preserved, OCR-ready

## 🤝 Contributing

Contributions welcome! This project helps with:
- Complex data extraction
- OCR workflows
- AI/LLM data preparation
- Financial analysis automation

## 📝 License

MIT License - see LICENSE file for details

## 🙏 Acknowledgments

Built for extracting data from complex financial Excel files with many columns, making them OCR-compatible and AI-ready.

## 📞 Support

- **Issues:** [GitHub Issues](https://github.com/TanishaR2/extracting-from-excel/issues)
- **Documentation:** See `docs/` folder
- **Examples:** See `examples/` folder

## 🌟 Star History

If this project helps you, please consider giving it a star! ⭐

---

**Status:** Production-Ready ✅  
**Tested:** Financial models with 60-100+ columns ✅  
**Quality:** 100% data preservation ✅