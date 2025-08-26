# files2md Tool Guidelines

## Commands
- **Install**: `make install` (installs to ~/bin)
- **Run**: `python3 files2md.py [options] <paths>` 
- **Test**: Test manually with various file types
- **Lint**: `flake8 files2md.py`
- **Type Check**: `mypy --strict files2md.py`

## Code Style
- **Python**: Follow PEP 8 conventions
- **Imports**: Standard library first, then third-party libraries, then local modules
- **Docstrings**: Google style with Args/Returns sections
- **Function Names**: snake_case
- **Variable Names**: snake_case
- **Type Hints**: Prefer explicit type annotations
- **Error Handling**: Use try/except for expected errors, with detailed error messages
- **String Formatting**: Use f-strings for string interpolation
- **Line Length**: Maximum 100 characters
- **Comments**: Use for complex logic explanations only

## Features
- Converts files (PDF, DOCX, XLSX, plain text) to Markdown
- Supports recursive directory searching
- Filename pattern matching via regex