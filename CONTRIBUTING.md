# Contributing to Excel to OCR-Friendly Converter

Thank you for your interest in contributing! This document provides guidelines for contributing to this project.

## How to Contribute

### Reporting Issues
- Use GitHub Issues to report bugs or suggest features
- Include clear description and reproduction steps
- Provide sample Excel files if possible (anonymized data)

### Pull Requests
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Test thoroughly
5. Commit with clear messages (`git commit -m 'Add amazing feature'`)
6. Push to your branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

### Development Setup

```bash
# Clone your fork
git clone https://github.com/YOUR_USERNAME/extracting-from-excel.git
cd extracting-from-excel

# Install with UV
uv sync

# Install dev dependencies
uv pip install -e ".[dev]"
```

### Code Style
- Follow PEP 8 guidelines
- Use type hints where applicable
- Add docstrings to functions and classes
- Keep functions focused and testable

### Testing
```bash
# Run tests
pytest

# Generate test data
python generate_sample.py

# Test conversion
python run_enhanced.py --input sample.xlsx --format markdown
```

## Areas for Contribution

- 🐛 Bug fixes
- ✨ New features (e.g., support for more formats)
- 📝 Documentation improvements
- 🧪 Additional tests
- 🎨 UI/UX enhancements for CLI
- 🌐 Internationalization

## Questions?

Feel free to open an issue for questions or discussions!
