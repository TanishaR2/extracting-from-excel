# Example: Converting Financial Excel to OCR-Friendly Format

This example demonstrates converting a financial model with many columns to Markdown format.

## Input File
- **Name:** Financial_Model.xlsx
- **Sheets:** 2 (PF Consolidated, DAH)
- **Columns:** 60-63 per sheet
- **Rows:** 191-321 per sheet

## Command Used

```bash
python run_enhanced.py \
    --input Financial_Model.xlsx \
    --format markdown \
    --all-sheets \
    --max-cols 15 \
    --output-dir output
```

## Output Structure

```
output/
├── markdown_PF_Consolidated/
│   ├── index.md              # Sheet summary and navigation
│   ├── combined_output.md    # All data in one file (419 KB)
│   ├── page_1.md             # Columns 0-14
│   ├── page_2.md             # Columns 15-29
│   ├── page_3.md             # Columns 30-44
│   ├── page_4.md             # Columns 45-59
│   └── page_5.md             # Columns 60-62
└── markdown_DAH/
    ├── index.md
    ├── combined_output.md    # All data in one file (244 KB)
    ├── page_1.md
    ├── page_2.md
    ├── page_3.md
    └── page_4.md
```

## Results

- ✅ **Total Pages:** 9 (5 for PF Consolidated, 4 for DAH)
- ✅ **Data Preservation:** 100%
- ✅ **Format:** OCR-compatible Markdown tables
- ✅ **Size:** 663 KB total output
- ✅ **Processing Time:** ~3 seconds

## Data Sample

### From PF Consolidated - Page 1 (Columns 0-14)

```markdown
| Unnamed: 0 | VSH II Pro Forma Entity | Model | 2023 | 2024 | 2025 |
|------------|------------------------|-------|------|------|------|
| nan        | Revenue:               | nan   | nan  | nan  | nan  |
| nan        | DAH                    | nan   | 21590737.63 | 32111654.47 | 30231785.12 |
| nan        | Total Net Revenue      | nan   | nan  | nan  | nan  |
```

## Analyzing the Output

```bash
# View summary
python analyze_data.py --dir output --summary

# Search for metrics
python analyze_data.py --dir output --search "Revenue"
python analyze_data.py --dir output --search "Gross Profit"

# View specific sections
grep -i "revenue" output/markdown_PF_Consolidated/combined_output.md
grep "2025" output/markdown_DAH/combined_output.md
```

## Using with AI/LLM

Copy the markdown tables and paste into ChatGPT/Claude:

**Prompt Example:**
```
Analyze this financial data and identify:
1. Revenue growth trends from 2023-2029
2. Margin improvement opportunities
3. Key financial metrics

[paste markdown table here]
```

## Converting to PDF

If you prefer PDF format:

```bash
python run_enhanced.py \
    --input Financial_Model.xlsx \
    --format pdf \
    --all-sheets \
    --max-cols 12 \
    --output-dir output_pdf
```

## Additional Options

```bash
# Process only one sheet
python run_enhanced.py --input file.xlsx --sheet "PF Consolidated"

# Change columns per page
python run_enhanced.py --input file.xlsx --max-cols 20

# Generate both formats
python run_enhanced.py --input file.xlsx --format both
```

## Success Metrics

- ✅ All 512 rows preserved
- ✅ All 123 columns preserved
- ✅ All relationships intact
- ✅ OCR-ready format
- ✅ AI-compatible structure
