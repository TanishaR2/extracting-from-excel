#!/bin/bash
# Quick Start Script for Excel to OCR Converter

set -e  # Exit on error

echo "=================================================="
echo "Excel to OCR Converter - Quick Start"
echo "=================================================="
echo ""

# Check if UV is installed
if ! command -v uv &> /dev/null; then
    echo "❌ UV not found. Installing UV..."
    curl -LsSf https://astral.sh/uv/install.sh | sh
    echo "✅ UV installed successfully"
    echo ""
fi

echo "📦 Installing dependencies..."
uv sync
echo "✅ Dependencies installed"
echo ""

echo "🎯 Generating sample Excel file..."
uv run python generate_sample.py
echo "✅ Sample file created: sample.xlsx"
echo ""

echo "🔄 Converting sample to Markdown..."
uv run python run_enhanced.py --input sample.xlsx --format markdown --max-cols 15
echo "✅ Conversion complete!"
echo ""

echo "📊 Output location: output/markdown_Sheet1/"
echo ""
echo "View the results:"
echo "  cat output/markdown_Sheet1/index.md"
echo "  cat output/markdown_Sheet1/page_1.md"
echo ""
echo "=================================================="
echo "Next Steps:"
echo "=================================================="
echo ""
echo "1. Convert your own Excel file:"
echo "   uv run python run_enhanced.py --input YOUR_FILE.xlsx --format markdown"
echo ""
echo "2. Process all sheets:"
echo "   uv run python run_enhanced.py --input YOUR_FILE.xlsx --all-sheets"
echo ""
echo "3. Adjust columns per page:"
echo "   uv run python run_enhanced.py --input YOUR_FILE.xlsx --max-cols 20"
echo ""
echo "4. Search extracted data:"
echo "   uv run python analyze_data.py --dir output --search 'Revenue'"
echo ""
echo "5. List sheets in a file:"
echo "   uv run python run_enhanced.py --input YOUR_FILE.xlsx --list-sheets"
echo ""
echo "📚 Full documentation: docs/README.md"
echo "💡 Examples: examples/EXAMPLE_USAGE.md"
echo ""
