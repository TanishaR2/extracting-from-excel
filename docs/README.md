# Excel to OCR-Friendly PDF Converter

A complete, production-ready solution for converting complex Excel XLSX files with very large numbers of columns into OCR-friendly PDFs by automatically splitting tables horizontally into multiple pages.

## 🎯 Features

- **Automatic Column Slicing**: Splits wide Excel tables into manageable chunks
- **OCR-Friendly Output**: Large, readable text (≥10pt) - no tiny, shrink-to-fit text
- **Multi-Page PDF Generation**: Each page contains a configurable number of columns
- **Repeated Headers**: Header rows are repeated on each page for clarity
- **HTML → PDF Pipeline**: Uses WeasyPrint for high-quality vector PDF output
- **Configurable Layout**: A4 landscape pages with customizable margins and styling
- **PDF Combination**: Automatically merges individual pages into a single PDF
- **Robust Error Handling**: Comprehensive logging and error recovery
- **Flexible Configuration**: YAML-based configuration with CLI overrides

## 📋 Requirements

- Python 3.10 or higher
- System dependencies for WeasyPrint (see Installation section)

## 🚀 Installation

### 1. Install System Dependencies

#### Ubuntu/Debian:
```bash
sudo apt-get update
sudo apt-get install -y python3-pip python3-cffi python3-brotli libpango-1.0-0 libpangoft2-1.0-0 libharfbuzz0b libpangocairo-1.0-0
```

#### macOS:
```bash
brew install python3 pango cairo libffi
```

#### Windows:
Download and install GTK+ runtime from: https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases

### 2. Install Python Dependencies

```bash
pip install -r requirements.txt
```

Or install manually:
```bash
pip install pandas openpyxl Jinja2 WeasyPrint PyPDF2 PyYAML numpy
```

## 📁 Project Structure

```
.
├── src/
│   ├── __init__.py
│   ├── config.yaml              # Configuration file
│   ├── column_slicer.py         # Column slicing logic
│   ├── html_renderer.py         # HTML template rendering
│   ├── pdf_exporter.py          # PDF generation and merging
│   └── pipeline.py              # Main pipeline orchestration
├── templates/
│   └── table.html               # Jinja2 HTML template
├── output/                      # Generated PDFs (created automatically)
├── logs/                        # Log files (created automatically)
├── run.py                       # CLI runner script
├── generate_sample.py           # Sample Excel file generator
├── requirements.txt             # Python dependencies
└── README.md                    # This file
```

## 🎮 Usage

### Quick Start

1. **Generate a sample Excel file**:
```bash
python generate_sample.py --rows 50 --cols 100
```

2. **Convert to PDF**:
```bash
python run.py --input sample.xlsx --max-cols 12
```

3. **Check output**:
```bash
ls -lh output/
```

### Command Line Options

```bash
python run.py --help
```

**Basic usage**:
```bash
python run.py --input myfile.xlsx
```

**With custom column count**:
```bash
python run.py --input myfile.xlsx --max-cols 15
```

**Process specific sheet**:
```bash
python run.py --input myfile.xlsx --sheet "Sheet1"
```

**Custom output directory**:
```bash
python run.py --input myfile.xlsx --output-dir my_output
```

**Don't combine PDFs**:
```bash
python run.py --input myfile.xlsx --no-combine
```

**Verbose logging**:
```bash
python run.py --input myfile.xlsx --verbose
```

**Custom configuration**:
```bash
python run.py --input myfile.xlsx --config my_config.yaml
```

### Generate Sample Files

**Simple sample** (100 columns):
```bash
python generate_sample.py --rows 50 --cols 100 --output sample.xlsx
```

**Complex sample** (realistic business data):
```bash
python generate_sample.py --complex --rows 30 --output sample_complex.xlsx
```

## ⚙️ Configuration

Edit `src/config.yaml` to customize the conversion process:

### PDF Layout Settings
```yaml
pdf:
  page_size: "A4"              # Page size (A4, Letter, etc.)
  orientation: "landscape"      # landscape or portrait
  margin_top: "15mm"           # Top margin
  margin_right: "10mm"         # Right margin
  margin_bottom: "15mm"        # Bottom margin
  margin_left: "10mm"          # Left margin
  font_size: "10pt"            # Base font size (≥10pt for OCR)
  font_family: "Arial, sans-serif"
```

### Column Slicing Settings
```yaml
slicing:
  max_columns_per_page: 12     # Max columns per page
  repeat_header: true          # Repeat headers on each page
  header_rows: 1               # Number of header rows
```

### Output Settings
```yaml
output:
  output_dir: "output"         # Output directory
  temp_dir: "temp"             # Temporary files directory
  combine_pdfs: true           # Combine pages into single PDF
  keep_individual_pages: true  # Keep individual page PDFs
  page_label_format: "Page {page_num} – Columns {start_col}-{end_col}"
```

### HTML Styling
```yaml
html:
  table_border: "1px solid #333"
  header_bg_color: "#4CAF50"
  header_text_color: "#ffffff"
  cell_padding: "8px"
  stripe_rows: true
  stripe_color: "#f2f2f2"
```

## 🔧 Pipeline Details

The conversion process follows these steps:

1. **Excel Reading**: Load XLSX file using pandas and openpyxl
2. **Column Analysis**: Determine how many pages are needed
3. **Column Slicing**: Split DataFrame into column chunks
4. **HTML Rendering**: Convert each chunk to HTML using Jinja2
5. **PDF Generation**: Convert HTML to PDF using WeasyPrint
6. **PDF Merging**: Combine individual pages using PyPDF2

### Example Column Slicing

For a 100-column Excel file with `max_cols = 12`:
- **Page 1**: Columns 0-11
- **Page 2**: Columns 12-23
- **Page 3**: Columns 24-35
- ...
- **Page 9**: Columns 96-99

Each page includes:
- Page label (e.g., "Page 1 – Columns 0-11")
- Repeated header row
- All data rows for those columns

## 📊 Output

After conversion, the `output/` directory contains:

```
output/
├── page_1.pdf          # First 12 columns
├── page_2.pdf          # Columns 13-24
├── page_3.pdf          # Columns 25-36
├── ...
└── combined_output.pdf # All pages combined
```

## 🐛 Troubleshooting

### WeasyPrint Installation Issues

**Error: "cannot load library 'gobject-2.0-0'"**
- Install system dependencies (see Installation section)
- On Windows, ensure GTK+ runtime is in PATH

**Error: "cairo library not found"**
```bash
# Ubuntu/Debian
sudo apt-get install libcairo2

# macOS
brew install cairo
```

### Memory Issues with Large Files

For very large Excel files (1000+ columns, 10000+ rows):
- Increase Python memory limit
- Process in smaller batches
- Reduce `max_columns_per_page`

### PDF Text Not OCR-Friendly

- Ensure `font_size` ≥ 10pt in config
- Check PDF viewer zoom level
- Verify vector text (not rasterized) with PDF inspector

## 📝 Logging

Logs are written to:
- Console (INFO level by default)
- `logs/conversion.log` (if configured)

Change log level in `src/config.yaml`:
```yaml
logging:
  level: "DEBUG"  # DEBUG, INFO, WARNING, ERROR, CRITICAL
```

Or use `--verbose` flag:
```bash
python run.py --input sample.xlsx --verbose
```

## 🧪 Testing

Generate test files and verify output:

```bash
# Generate test file
python generate_sample.py --rows 100 --cols 150 --output test.xlsx

# Convert with different settings
python run.py --input test.xlsx --max-cols 10
python run.py --input test.xlsx --max-cols 15
python run.py --input test.xlsx --max-cols 20

# Verify output
ls -lh output/
```

## 🎨 Customization

### Custom HTML Template

Edit `templates/table.html` to customize the PDF appearance:
- Change colors, fonts, spacing
- Add company logos or watermarks
- Modify header/footer layout

### Custom Page Labels

In `src/config.yaml`:
```yaml
output:
  page_label_format: "Sheet 1 - Part {page_num} (Cols {start_col}-{end_col})"
```

### Portrait Orientation

For narrow tables:
```yaml
pdf:
  orientation: "portrait"
  max_columns_per_page: 6  # Fewer columns for portrait
```

## 🚦 Performance Tips

- **Large files**: Use `--verbose` to monitor progress
- **Memory optimization**: Process sheets separately
- **Speed**: Disable individual page PDFs if not needed (`--no-combine`)
- **Storage**: Clean up temp files after conversion

## 📜 License

This project is provided as-is for production use. Modify as needed for your requirements.

## 🤝 Contributing

Feel free to extend this project with:
- Support for merged cells
- Multi-row headers
- Custom column selection
- Excel formulas preservation
- Conditional formatting

## 📞 Support

For issues:
1. Check logs in `logs/conversion.log`
2. Verify system dependencies are installed
3. Test with generated sample files first
4. Enable verbose logging (`--verbose`)

## ✨ Features Checklist

- ✅ Automatic column slicing
- ✅ OCR-friendly text (≥10pt, vector)
- ✅ Multi-page PDF generation
- ✅ Repeated headers
- ✅ HTML → PDF pipeline
- ✅ A4 landscape layout
- ✅ Page labels
- ✅ PDF combination
- ✅ Configurable settings
- ✅ Comprehensive logging
- ✅ Error handling
- ✅ CLI interface
- ✅ Sample file generation
- ✅ Complete documentation

---

**Ready to convert your wide Excel tables to OCR-friendly PDFs!** 🎉
