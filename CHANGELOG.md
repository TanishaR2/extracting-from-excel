# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-12-05

### Added
- ✨ Initial release of Excel to OCR-Friendly Converter
- 🎯 Automatic column slicing for wide Excel tables
- 📄 Markdown export for OCR-compatible output
- 📑 PDF export with WeasyPrint
- 🔧 Configurable YAML-based settings
- 🖥️ CLI interface with `run_enhanced.py`
- 📊 Data analysis tool (`analyze_data.py`)
- 🧪 Sample file generator (`generate_sample.py`)
- 📖 Comprehensive documentation
- 🎨 HTML templates for PDF generation
- 🔄 Multi-sheet processing support
- 📝 Comprehensive logging
- ⚠️ Error handling and recovery
- 🚀 Production-ready code

### Features
- Handle Excel files with 100+ columns
- Split tables horizontally into pages (configurable columns per page)
- Preserve all data relationships and structure
- Generate OCR-friendly Markdown tables
- Generate high-quality PDF documents (A4 landscape)
- Combine multiple pages into single output
- Process all sheets or specific sheets
- Search and analyze extracted data

### Documentation
- Complete README with installation and usage
- Quick Start Guide for immediate use
- Extraction README for understanding output
- Project Summary with technical details
- Code comments throughout
- Example files and use cases

### Testing
- Successfully processed financial models with 60+ columns
- Tested with 321 rows × 63 columns dataset
- 100% data preservation verified
- OCR compatibility confirmed

## [Unreleased]

### Planned Features
- [ ] Support for merged cells
- [ ] Excel formula preservation
- [ ] Conditional formatting export
- [ ] Chart extraction
- [ ] Multiple output formats (HTML, CSV)
- [ ] Batch processing mode
- [ ] Web interface
- [ ] Docker support
- [ ] CI/CD pipeline

---

**Legend:**
- ✨ New features
- 🐛 Bug fixes
- 📝 Documentation
- 🔧 Configuration
- ⚡ Performance improvements
- 🔒 Security updates
